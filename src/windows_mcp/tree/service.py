from windows_mcp.tree.config import (
    INTERACTIVE_CONTROL_TYPE_NAMES,
    DOCUMENT_CONTROL_TYPE_NAMES,
    INFORMATIVE_CONTROL_TYPE_NAMES,
    DEFAULT_ACTIONS,
    THREAD_MAX_RETRIES,
    TREE_STATE_TIMEOUT_S,
    TREE_STATE_TIMEOUT_DOM_S,
    ROOT_ENUM_TIMEOUT_S,
    UIA_TIMEOUT_COOLDOWN_S,
    UIA_MAX_WORKERS,
)
from windows_mcp.tree.views import TreeElementNode, ScrollElementNode, TextElementNode, Center, BoundingBox, TreeState, DOMInfo
from uiautomation import Control,ImageControl,ScrollPattern,WindowControl,Rect,GetRootControl,PatternId
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError as FuturesTimeoutError
from windows_mcp.tree.utils import random_point_within_bounding_box
from PIL import Image, ImageFont, ImageDraw
from typing import TYPE_CHECKING,Optional
from windows_mcp.desktop.views import App
from windows_mcp.desktop.config import EXCLUDED_APPS
from time import sleep
import time
import logging
import random
import os
import uiautomation as uia

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
handler = logging.StreamHandler()
formatter = logging.Formatter('[%(levelname)s] %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

if TYPE_CHECKING:
    from windows_mcp.desktop.service import Desktop


def _init_uia_in_worker_thread() -> None:
    """
    UIAutomation (COM) must be initialized per thread when using `uiautomation`
    from a ThreadPoolExecutor. Without this, calls may fail with:
      [WinError -2147221008] CoInitialize has not been called
    """
    try:
        uia.InitializeUIAutomationInCurrentThread()
    except Exception as e:
        # Don't hard-fail the worker thread; the caller will surface partial results.
        logger.debug(f"UIAutomation init in worker thread failed: {e}")
    
class Tree:
    def __init__(self,desktop:'Desktop'):
        self.desktop=desktop
        self.screen_size=self.desktop.get_screen_size()
        self.dom_info:Optional[DOMInfo]=None
        self.dom_bounding_box:BoundingBox=None
        # Persistent executor to avoid hanging on shutdown when any UIA traversal thread blocks.
        self._executor = ThreadPoolExecutor(
            max_workers=UIA_MAX_WORKERS,
            thread_name_prefix="windows-mcp-uia",
            initializer=_init_uia_in_worker_thread,
        )
        # If UIAutomation gets into a bad state (common when a provider hangs), we temporarily
        # disable UIA enumeration to avoid leaking stuck threads on repeated calls.
        self._uia_disabled_until: float = 0.0
        self.screen_box=BoundingBox(
            top=0, left=0, bottom=self.screen_size.height, right=self.screen_size.width,
            width=self.screen_size.width, height=self.screen_size.height 
        )

    def get_state(self,active_app:App,other_apps:list[App],use_dom:bool=False)->TreeState:
        t0 = time.perf_counter()
        timeout_s = TREE_STATE_TIMEOUT_DOM_S if use_dom else TREE_STATE_TIMEOUT_S
        root_timeout_s = min(ROOT_ENUM_TIMEOUT_S, timeout_s)

        now = time.monotonic()
        if now < self._uia_disabled_until:
            remaining = self._uia_disabled_until - now
            warnings = [
                f"UIAutomation enumeration temporarily disabled for {remaining:.1f}s after a prior timeout; returning partial results."
            ]
            logger.warning("State-Tool(Tree): UIA disabled (%.1fs remaining), returning partial.", remaining)
            return TreeState(
                dom_info=None,
                interactive_nodes=[],
                scrollable_nodes=[],
                dom_informative_nodes=[],
                is_partial=True,
                warnings=warnings,
            )

        other_apps_handle=set(map(lambda other_app: other_app.handle,other_apps))
        exclude_system_windows = os.getenv("WINDOWS_MCP_EXCLUDE_SYSTEM_WINDOWS", "false").strip().lower() in {
            "1", "true", "yes", "y"
        }

        root_warnings: list[str] = []
        root_partial = False

        def collect_targets() -> list[Control]:
            root = GetRootControl()
            children = root.GetChildren()
            targets: list[Control] = []
            for child in children:
                try:
                    hwnd = child.NativeWindowHandle
                except Exception:
                    continue
                if hwnd in other_apps_handle:
                    continue
                if active_app and child.ClassName == "Progman":
                    # Keep the desktop out of the listing when an active app exists (matches previous behavior).
                    continue
                if exclude_system_windows and child.ClassName in EXCLUDED_APPS:
                    continue
                targets.append(child)
            return targets

        apps: list[Control] = []
        try:
            fut = self._executor.submit(collect_targets)
            apps = fut.result(timeout=root_timeout_s)
        except FuturesTimeoutError:
            root_partial = True
            root_warnings.append(
                f"Root UIAutomation enumeration timed out after {root_timeout_s:.1f}s; falling back to active window only."
            )
            self._uia_disabled_until = max(self._uia_disabled_until, time.monotonic() + UIA_TIMEOUT_COOLDOWN_S)
            try:
                fut.cancel()
            except Exception:
                pass
        except Exception as e:
            root_partial = True
            root_warnings.append("Root UIAutomation enumeration failed; falling back to active window only.")
            logger.debug(f"Root enumeration error: {e}")

        # Fallback: at least enumerate the active app if root enumeration is unavailable.
        if not apps and active_app and active_app.handle:
            try:
                fut = self._executor.submit(lambda: uia.ControlFromHandle(active_app.handle))
                apps = [fut.result(timeout=min(1.0, root_timeout_s))]
            except Exception:
                root_partial = True
                root_warnings.append("Active window UIAutomation handle lookup failed; returning partial state.")

        t_root = time.perf_counter()
        logger.info(
            "State-Tool(Tree): targets=%d root=%.3fs use_dom=%s root_partial=%s",
            len(apps),
            (t_root - t0),
            use_dom,
            root_partial,
        )

        interactive_nodes,scrollable_nodes,dom_informative_nodes,is_partial,warnings=self.get_appwise_nodes(apps=apps,use_dom=use_dom)
        is_partial = bool(is_partial or root_partial)
        warnings = (warnings or []) + root_warnings
        t_end = time.perf_counter()
        logger.info(
            "State-Tool(Tree): done total=%.3fs interactive=%d scroll=%d partial=%s",
            (t_end - t0),
            len(interactive_nodes),
            len(scrollable_nodes),
            is_partial,
        )
        return TreeState(
            dom_info=self.dom_info,
            interactive_nodes=interactive_nodes,
            scrollable_nodes=scrollable_nodes,
            dom_informative_nodes=dom_informative_nodes,
            is_partial=is_partial,
            warnings=warnings,
        )

    def get_appwise_nodes(self,apps:list[Control],use_dom:bool=False)-> tuple[list[TreeElementNode],list[ScrollElementNode],list[TextElementNode],bool,list[str]]:
        interactive_nodes, scrollable_nodes,dom_informative_nodes = [], [], []
        is_partial = False
        warnings: list[str] = []

        timeout_s = TREE_STATE_TIMEOUT_DOM_S if use_dom else TREE_STATE_TIMEOUT_S
        deadline = time.monotonic() + timeout_s

        retry_counts = {app: 0 for app in apps}
        future_to_app = {
            # Determine browser-ness inside the worker thread to avoid blocking on UIA properties here.
            self._executor.submit(self.get_nodes, app, None, use_dom): app
            for app in apps
        }
        total_submitted = len(future_to_app)
        completed = 0

        while future_to_app:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                is_partial = True
                warnings.append(f"UI tree enumeration timed out after {timeout_s:.1f}s; returning partial results.")
                break
            try:
                for future in as_completed(list(future_to_app), timeout=remaining):
                    app = future_to_app.pop(future)
                    try:
                        result = future.result()
                        if result:
                            element_nodes, scroll_nodes, informative_nodes = result
                            interactive_nodes.extend(element_nodes)
                            scrollable_nodes.extend(scroll_nodes)
                            dom_informative_nodes.extend(informative_nodes)
                        completed += 1
                    except Exception as e:
                        retry_counts[app] += 1
                        # Avoid touching UIA properties here (e.g. app.Name) — they can block too.
                        logger.debug(f"Error in processing UIA subtree (retry {retry_counts[app]}/{THREAD_MAX_RETRIES}): {e}")
                        if retry_counts[app] < THREAD_MAX_RETRIES and (deadline - time.monotonic()) > 0:
                            new_future = self._executor.submit(
                                self.get_nodes, app, None, use_dom
                            )
                            future_to_app[new_future] = app
                        else:
                            is_partial = True
                            warnings.append("Some UIA subtrees failed; returning partial results.")
            except FuturesTimeoutError:
                is_partial = True
                warnings.append(f"UI tree enumeration timed out after {timeout_s:.1f}s; returning partial results.")
                break

        # Best-effort: cancel pending work (won't stop running threads, but avoids queue buildup).
        for fut in list(future_to_app):
            fut.cancel()

        if is_partial:
            logger.info(
                "State-Tool(Tree): subtree_partial=%s completed=%d/%d nodes=%d scroll=%d",
                is_partial,
                completed,
                total_submitted,
                len(interactive_nodes),
                len(scrollable_nodes),
            )

        return interactive_nodes, scrollable_nodes, dom_informative_nodes, is_partial, warnings
    
    def iou_bounding_box(self,window_box: Rect,element_box: Rect,) -> BoundingBox:
        # Step 1: Intersection of element and window (existing logic)
        intersection_left = max(window_box.left, element_box.left)
        intersection_top = max(window_box.top, element_box.top)
        intersection_right = min(window_box.right, element_box.right)
        intersection_bottom = min(window_box.bottom, element_box.bottom)

        # Step 2: Clamp to screen boundaries (new addition)
        intersection_left = max(self.screen_box.left, intersection_left)
        intersection_top = max(self.screen_box.top, intersection_top)
        intersection_right = min(self.screen_box.right, intersection_right)
        intersection_bottom = min(self.screen_box.bottom, intersection_bottom)

        # Step 3: Validate intersection
        if (intersection_right > intersection_left and intersection_bottom > intersection_top):
            bounding_box = BoundingBox(
                left=intersection_left,
                top=intersection_top,
                right=intersection_right,
                bottom=intersection_bottom,
                width=intersection_right - intersection_left,
                height=intersection_bottom - intersection_top
            )
        else:
            # No valid visible intersection (either outside window or screen)
            bounding_box = BoundingBox(
                left=0,
                top=0,
                right=0,
                bottom=0,
                width=0,
                height=0
            )
        return bounding_box

    def get_nodes(self, node: Control, is_browser:bool|None=None, use_dom:bool=False) -> tuple[list[TreeElementNode],list[ScrollElementNode]]:
        if is_browser is None:
            try:
                is_browser = self.desktop.is_app_browser(node)
            except Exception:
                is_browser = False
        window_bounding_box=node.BoundingRectangle

        def is_element_visible(node:Control,threshold:int=0):
            is_control=node.IsControlElement
            box=node.BoundingRectangle
            if box.isempty():
                return False
            width=box.width()
            height=box.height()
            area=width*height
            is_offscreen=(not node.IsOffscreen) or node.ControlTypeName in ['EditControl']
            return area > threshold and is_offscreen and is_control
    
        def is_element_enabled(node:Control):
            try:
                return node.IsEnabled
            except Exception:
                return False
            
        def is_default_action(node:Control):
            legacy_pattern=node.GetLegacyIAccessiblePattern()
            default_action=legacy_pattern.DefaultAction.title()
            if default_action in DEFAULT_ACTIONS:
                return True
            return False
        
        def is_element_image(node:Control):
            if isinstance(node,ImageControl):
                if node.LocalizedControlType=='graphic' or not node.IsKeyboardFocusable:
                    return True
            return False
        
        def is_element_text(node:Control):
            try:
                if node.ControlTypeName in INFORMATIVE_CONTROL_TYPE_NAMES:
                    if is_element_visible(node) and is_element_enabled(node) and not is_element_image(node):
                        return True
            except Exception:
                return False
            return False
            
        def is_window_modal(node:WindowControl):
            try:
                window_pattern=node.GetWindowPattern()
                return window_pattern.IsModal
            except Exception:
                return False
            
        def is_keyboard_focusable(node:Control):
            try:
                if node.ControlTypeName in set(['EditControl','ButtonControl','CheckBoxControl','RadioButtonControl','TabItemControl']):
                    return True
                return node.IsKeyboardFocusable
            except Exception:
                return False
            
        def element_has_child_element(node:Control,control_type:str,child_control_type:str):
            if node.LocalizedControlType==control_type:
                first_child=node.GetFirstChildControl()
                if first_child is None:
                    return False
                return first_child.LocalizedControlType==child_control_type
            
        def group_has_no_name(node:Control):
            try:
                if node.ControlTypeName=='GroupControl':
                    if not node.Name.strip():
                        return True
                return False
            except Exception:
                return False
            
        def is_element_scrollable(node:Control):
            try:
                if (node.ControlTypeName in INTERACTIVE_CONTROL_TYPE_NAMES|INFORMATIVE_CONTROL_TYPE_NAMES) or node.IsOffscreen:
                    return False
                scroll_pattern:ScrollPattern=node.GetPattern(PatternId.ScrollPattern)
                if scroll_pattern is None:
                    return False
                return scroll_pattern.VerticallyScrollable
            except Exception:
                return False
            
        def is_element_interactive(node:Control):
            try:
                if is_browser and node.ControlTypeName in set(['DataItemControl','ListItemControl']) and not is_keyboard_focusable(node):
                    return False
                elif not is_browser and node.ControlTypeName=="ImageControl" and is_keyboard_focusable(node):
                    return True
                elif node.ControlTypeName in INTERACTIVE_CONTROL_TYPE_NAMES|DOCUMENT_CONTROL_TYPE_NAMES:
                    return is_element_visible(node) and is_element_enabled(node) and (not is_element_image(node) or is_keyboard_focusable(node))
                elif node.ControlTypeName=='GroupControl':
                    if is_browser:
                        return is_element_visible(node) and is_element_enabled(node) and (is_default_action(node) or is_keyboard_focusable(node))
                    # else:
                    #     return is_element_visible and is_element_enabled(node) and is_default_action(node)
            except Exception:
                return False
            return False
        
        def dom_correction(node:Control):
            if element_has_child_element(node,'list item','link') or element_has_child_element(node,'item','link'):
                dom_interactive_nodes.pop()
                return None
            elif node.ControlTypeName=='GroupControl':
                dom_interactive_nodes.pop()
                if is_keyboard_focusable(node):
                    child=node
                    try:
                        while child.GetFirstChildControl() is not None:
                            if child.ControlTypeName in INTERACTIVE_CONTROL_TYPE_NAMES:
                                return None
                            child=child.GetFirstChildControl()
                    except Exception:
                        return None
                    if child.ControlTypeName!='TextControl':
                        return None
                    legacy_pattern=node.GetLegacyIAccessiblePattern()
                    value=legacy_pattern.Value
                    element_bounding_box = node.BoundingRectangle
                    bounding_box=self.iou_bounding_box(self.dom_bounding_box,element_bounding_box)
                    center = bounding_box.get_center()
                    is_focused=node.HasKeyboardFocus
                    dom_interactive_nodes.append(TreeElementNode(**{
                        'name':child.Name.strip(),
                        'control_type':node.LocalizedControlType,
                        'value':value,
                        'shortcut':node.AcceleratorKey,
                        'bounding_box':bounding_box,
                        'xpath':'',
                        'center':center,
                        'app_name':app_name,
                        'is_focused':is_focused
                    }))
            elif element_has_child_element(node,'link','heading'):
                dom_interactive_nodes.pop()
                node=node.GetFirstChildControl()
                control_type='link'
                legacy_pattern=node.GetLegacyIAccessiblePattern()
                value=legacy_pattern.Value
                element_bounding_box = node.BoundingRectangle
                bounding_box=self.iou_bounding_box(self.dom_bounding_box,element_bounding_box)
                center = bounding_box.get_center()
                is_focused=node.HasKeyboardFocus
                dom_interactive_nodes.append(TreeElementNode(**{
                    'name':node.Name.strip(),
                    'control_type':control_type,
                    'value':node.Name.strip(),
                    'shortcut':node.AcceleratorKey,
                    'bounding_box':bounding_box,
                    'xpath':'',
                    'center':center,
                    'app_name':app_name,
                    'is_focused':is_focused
                }))
            
        def tree_traversal(node: Control,is_dom:bool=False,is_dialog:bool=False):
            # Checks to skip the nodes that are not interactive
            if node.IsOffscreen and (node.ControlTypeName not in set(["GroupControl","EditControl","TitleBarControl"])) and node.ClassName not in set(["Popup","Windows.UI.Core.CoreComponentInputSource"]):
                return None
            
            if is_element_scrollable(node):
                scroll_pattern:ScrollPattern=node.GetPattern(PatternId.ScrollPattern)
                box = node.BoundingRectangle
                # Get the center
                x,y=random_point_within_bounding_box(node=node,scale_factor=0.8)
                center = Center(x=x,y=y)
                scrollable_nodes.append(ScrollElementNode(**{
                    'name':node.Name.strip() or node.AutomationId or node.LocalizedControlType.capitalize() or "''",
                    'app_name':app_name,
                    'control_type':node.LocalizedControlType.title(),
                    'bounding_box':BoundingBox(**{
                        'left':box.left,
                        'top':box.top,
                        'right':box.right,
                        'bottom':box.bottom,
                        'width':box.width(),
                        'height':box.height()
                    }),
                    'center':center,
                    'xpath':'',
                    'horizontal_scrollable':scroll_pattern.HorizontallyScrollable,
                    'horizontal_scroll_percent':scroll_pattern.HorizontalScrollPercent if scroll_pattern.HorizontallyScrollable else 0,
                    'vertical_scrollable':scroll_pattern.VerticallyScrollable,
                    'vertical_scroll_percent':scroll_pattern.VerticalScrollPercent if scroll_pattern.VerticallyScrollable else 0,
                    'is_focused':node.HasKeyboardFocus
                }))
                    
            if is_element_interactive(node):
                legacy_pattern=node.GetLegacyIAccessiblePattern()
                value=legacy_pattern.Value.strip() if legacy_pattern.Value is not None else ""
                is_focused=node.HasKeyboardFocus
                name=node.Name.strip()
                element_bounding_box = node.BoundingRectangle
                if is_browser and is_dom:
                    bounding_box=self.iou_bounding_box(self.dom_bounding_box,element_bounding_box)
                    center = bounding_box.get_center()
                    tree_node=TreeElementNode(**{
                        'name':name,
                        'control_type':node.LocalizedControlType.title(),
                        'value':value,
                        'shortcut':node.AcceleratorKey,
                        'bounding_box':bounding_box,
                        'center':center,
                        'xpath':'',
                        'app_name':app_name,
                        'is_focused':is_focused
                    })
                    dom_interactive_nodes.append(tree_node)
                    dom_correction(node=node)
                else:
                    bounding_box=self.iou_bounding_box(window_bounding_box,element_bounding_box)
                    center = bounding_box.get_center()
                    tree_node=TreeElementNode(**{
                        'name':name,
                        'control_type':node.LocalizedControlType.title(),
                        'value':value,
                        'shortcut':node.AcceleratorKey,
                        'bounding_box':bounding_box,
                        'center':center,
                        'xpath':'',
                        'app_name':app_name,
                        'is_focused':is_focused
                    })
                    interactive_nodes.append(tree_node)
            elif is_element_text(node):
                dom_informative_nodes.append(TextElementNode(
                    text=node.Name.strip(),
                ))
            
            children=node.GetChildren()

            # Recursively traverse the tree the right to left for normal apps and for DOM traverse from left to right
            for child in (children if is_dom else children[::-1]):
                # Incrementally building the xpath
                
                # Check if the child is a DOM element
                if is_browser and child.AutomationId == "RootWebArea":
                    bounding_box=child.BoundingRectangle
                    self.dom_bounding_box=BoundingBox(left=bounding_box.left,top=bounding_box.top,
                    right=bounding_box.right,bottom=bounding_box.bottom,width=bounding_box.width(),
                    height=bounding_box.height())
                    scroll_pattern=child.GetPattern(PatternId.ScrollPattern)
                    self.dom_info=DOMInfo(
                        horizontal_scrollable=scroll_pattern.HorizontallyScrollable,
                        horizontal_scroll_percent=scroll_pattern.HorizontalScrollPercent if scroll_pattern.HorizontallyScrollable else 0,
                        vertical_scrollable=scroll_pattern.VerticallyScrollable,
                        vertical_scroll_percent=scroll_pattern.VerticalScrollPercent if scroll_pattern.VerticallyScrollable else 0
                    )
                    # enter DOM subtree
                    tree_traversal(child, is_dom=True, is_dialog=is_dialog)
                # Check if the child is a dialog
                elif isinstance(child,WindowControl):
                    if not child.IsOffscreen:
                        if is_dom:
                            bounding_box=child.BoundingRectangle
                            if bounding_box.width() > 0.8*self.dom_bounding_box.width:
                                # Because this window element covers the majority of the screen
                                dom_interactive_nodes.clear()
                        else:
                            if is_window_modal(child):
                                # Because this window element is modal
                                interactive_nodes.clear()
                    # enter dialog subtree
                    tree_traversal(child, is_dom=is_dom, is_dialog=True)
                else:
                    # normal non-dialog children
                    tree_traversal(child, is_dom=is_dom, is_dialog=is_dialog)

        interactive_nodes, dom_interactive_nodes, scrollable_nodes, dom_informative_nodes = [], [], [], []
        app_name=node.Name.strip()
        match node.ClassName:
            case "Progman":
                app_name="Desktop"
            case 'Shell_TrayWnd'|'Shell_SecondaryTrayWnd':
                app_name="Taskbar"
            case 'Microsoft.UI.Content.PopupWindowSiteBridge':
                app_name="Context Menu"
            case _:
                pass
        tree_traversal(node,is_dom=False,is_dialog=False)

        logger.debug(f'Interactive nodes:{len(interactive_nodes)}')
        logger.debug(f'DOM interactive nodes:{len(dom_interactive_nodes)}')
        logger.debug(f'Scrollable nodes:{len(scrollable_nodes)}')

        if use_dom:
            if is_browser:
                return (dom_interactive_nodes,scrollable_nodes,dom_informative_nodes)
            else:
                return ([],[],[])
        else:
            return (interactive_nodes+dom_interactive_nodes,scrollable_nodes,dom_informative_nodes)

    def get_annotated_screenshot(self, nodes: list[TreeElementNode],scale:float=1.0) -> Image.Image:
        screenshot = self.desktop.get_screenshot()
        sleep(0.10)
        
        original_width = screenshot.width
        original_height = screenshot.height

        scaled_width = int(original_width * scale)
        scaled_height = int(original_height * scale)
        screenshot = screenshot.resize((scaled_width, scaled_height), Image.Resampling.LANCZOS)
        
        # Add padding
        padding = 5
        width = int(screenshot.width + (1.5 * padding))
        height = int(screenshot.height + (1.5 * padding))
        padded_screenshot = Image.new("RGB", (width, height), color=(255, 255, 255))
        padded_screenshot.paste(screenshot, (padding, padding))

        draw = ImageDraw.Draw(padded_screenshot)
        font_size = 12
        try:
            font = ImageFont.truetype('arial.ttf', font_size)
        except IOError:
            font = ImageFont.load_default()

        def get_random_color():
            return "#{:06x}".format(random.randint(0, 0xFFFFFF))

        def draw_annotation(label, node: TreeElementNode):
            box = node.bounding_box
            color = get_random_color()

            # Scale and pad the bounding box coordinates
            adjusted_box = (
                int(box.left * scale) + padding,
                int(box.top * scale) + padding,
                int(box.right * scale) + padding,
                int(box.bottom * scale) + padding
            )
            # Draw bounding box
            draw.rectangle(adjusted_box, outline=color, width=2)

            # Label dimensions
            label_width = draw.textlength(str(label), font=font)
            label_height = font_size
            left, top, right, bottom = adjusted_box

            # Label position above bounding box
            label_x1 = right - label_width
            label_y1 = top - label_height - 4
            label_x2 = label_x1 + label_width
            label_y2 = label_y1 + label_height + 4

            # Draw label background and text
            draw.rectangle([(label_x1, label_y1), (label_x2, label_y2)], fill=color)
            draw.text((label_x1 + 2, label_y1 + 2), str(label), fill=(255, 255, 255), font=font)

        # Draw annotations in parallel
        with ThreadPoolExecutor() as executor:
            executor.map(draw_annotation, range(len(nodes)), nodes)
        return padded_screenshot