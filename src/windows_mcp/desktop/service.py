from windows_mcp.desktop.config import BROWSER_NAMES, PROCESS_PER_MONITOR_DPI_AWARE
from windows_mcp.desktop.views import DesktopState, App, Size, Status
from windows_mcp.tree.service import Tree
from locale import getpreferredencoding
from contextlib import contextmanager
from typing import Optional,Literal
from markdownify import markdownify
from fuzzywuzzy import process
from psutil import Process
from time import sleep
from PIL import Image
import win32process
import subprocess
import win32gui
import win32con
import requests
import logging
import base64
import ctypes
import csv
import re
import os
import io
import time

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
handler = logging.StreamHandler()
formatter = logging.Formatter('[%(levelname)s] %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

try:  
    ctypes.windll.shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
except Exception:  
    ctypes.windll.user32.SetProcessDPIAware()  

import uiautomation as uia
import pyautogui as pg
import time

pg.FAILSAFE=False
pg.PAUSE=1.0

class Desktop:
    def __init__(self):
        self.encoding=getpreferredencoding()
        self.tree=Tree(self)
        self.desktop_state=None
        # Ensure UIAutomation COM is initialized in the main thread too (some APIs depend on it).
        try:
            uia.InitializeUIAutomationInCurrentThread()
        except Exception:
            pass
        
    def get_resolution(self)->tuple[int,int]:
        return pg.size()
        
    def get_state(self,use_vision:bool=False,use_dom:bool=False,as_bytes:bool=False,scale:float=1.0)->DesktopState:
        sleep(0.1)
        t0 = time.perf_counter()
        apps=self.get_apps()
        t_apps = time.perf_counter()
        active_app=self.get_active_app()
        t_active = time.perf_counter()
        if active_app is not None:
            # Remove active app by handle (depth/name may differ, so dataclass equality is unreliable here).
            apps = [app for app in apps if app.handle != active_app.handle]
        logger.info(
            "State-Tool: apps=%d (%.3fs), active=%s (%.3fs), use_dom=%s, use_vision=%s",
            len(apps),
            (t_apps - t0),
            (active_app.name if active_app else "None"),
            (t_active - t_apps),
            use_dom,
            use_vision,
        )
        tree_state=self.tree.get_state(active_app,apps,use_dom=use_dom)
        t_tree = time.perf_counter()
        if use_vision:
            screenshot=self.tree.get_annotated_screenshot(tree_state.interactive_nodes,scale=scale)
            if as_bytes:
                bytes_io=io.BytesIO()
                screenshot.save(bytes_io,format='PNG')
                screenshot=bytes_io.getvalue()
        else:
            screenshot=None
        t_end = time.perf_counter()
        logger.info(
            "State-Tool: tree_nodes=%d scroll=%d partial=%s tree=%.3fs total=%.3fs",
            len(tree_state.interactive_nodes),
            len(tree_state.scrollable_nodes),
            getattr(tree_state, "is_partial", False),
            (t_tree - t_active),
            (t_end - t0),
        )
        self.desktop_state=DesktopState(apps= apps,active_app=active_app,screenshot=screenshot,tree_state=tree_state)
        return self.desktop_state
    
    def get_window_element_from_element(self,element:uia.Control)->uia.Control|None:
        while element is not None:
            if uia.IsTopLevelWindow(element.NativeWindowHandle):
                return element
            element = element.GetParentControl()
        return None
    
    def get_active_app(self)->App|None:
        try:
            handle = win32gui.GetForegroundWindow()
            if not handle:
                return None
            name = (win32gui.GetWindowText(handle) or "").strip() or "Foreground Window"
            left, top, right, bottom = win32gui.GetWindowRect(handle)
            width, height = max(0, right - left), max(0, bottom - top)
            _, pid = win32process.GetWindowThreadProcessId(handle)
            is_min = bool(win32gui.IsIconic(handle))
            if hasattr(win32gui, "IsZoomed"):
                is_max = bool(win32gui.IsZoomed(handle))  # type: ignore[attr-defined]
            else:
                # Fallback for environments where win32gui lacks IsZoomed()
                try:
                    # (flags, showCmd, ptMinPosition, ptMaxPosition, rect)
                    show_cmd = win32gui.GetWindowPlacement(handle)[1]
                    is_max = show_cmd == win32con.SW_SHOWMAXIMIZED
                except Exception:
                    is_max = False
            is_visible = bool(win32gui.IsWindowVisible(handle))
            status = (
                Status.MINIMIZED if is_min
                else Status.MAXIMIZED if is_max
                else Status.NORMAL if is_visible
                else Status.HIDDEN
            )
            return App(
                name=name,
                depth=0,
                status=status,
                size=Size(width=width, height=height),
                handle=handle,
                process_id=pid,
            )
        except Exception as ex:
            logger.error(f"Error in get_active_app: {ex}")
        return None
    
    def get_app_status(self,control:uia.Control)->Status:
        if uia.IsIconic(control.NativeWindowHandle):
            return Status.MINIMIZED
        elif uia.IsZoomed(control.NativeWindowHandle):
            return Status.MAXIMIZED
        elif uia.IsWindowVisible(control.NativeWindowHandle):
            return Status.NORMAL
        else:
            return Status.HIDDEN
    
    def get_cursor_location(self)->tuple[int,int]:
        position=pg.position()
        return (position.x,position.y)
    
    def get_element_under_cursor(self)->uia.Control:
        return uia.ControlFromCursor()
    
    def get_apps_from_start_menu(self)->dict[str,str]:
        command='Get-StartApps | ConvertTo-Csv -NoTypeInformation'
        apps_info,_=self.execute_command(command)
        reader=csv.DictReader(io.StringIO(apps_info))
        return {row.get('Name').lower():row.get('AppID') for row in reader}
    
    def execute_command(self,command:str)->tuple[str,int]:
        try:
            encoded = base64.b64encode(command.encode("utf-16le")).decode("ascii")
            result = subprocess.run(
                ['powershell', '-NoProfile', '-EncodedCommand', encoded], 
                capture_output=True, 
                errors='ignore',
                timeout=25,
                cwd=os.path.expanduser(path='~')
            )
            stdout=result.stdout
            stderr=result.stderr
            return (stdout or stderr,result.returncode)
        except subprocess.TimeoutExpired:
            return ('Command execution timed out', 1)
        except Exception as e:
            return ('Command execution failed', 1)
        
    def is_app_browser(self,node:uia.Control):
        process=Process(node.ProcessId)
        return process.name() in BROWSER_NAMES
    
    def get_default_language(self)->str:
        command="Get-Culture | Select-Object Name,DisplayName | ConvertTo-Csv -NoTypeInformation"
        response,_=self.execute_command(command)
        reader=csv.DictReader(io.StringIO(response))
        return "".join([row.get('DisplayName') for row in reader])
    
    def resize_app(self,size:tuple[int,int]=None,loc:tuple[int,int]=None)->tuple[str,int]:
        active_app=self.desktop_state.active_app
        if active_app is None:
            return "No active app found",1
        if active_app.status==Status.MINIMIZED:
            return f"{active_app.name} is minimized",1
        elif active_app.status==Status.MAXIMIZED:
            return f"{active_app.name} is maximized",1
        else:
            app_control=uia.ControlFromHandle(active_app.handle)
            if loc is None:
                x=app_control.BoundingRectangle.left
                y=app_control.BoundingRectangle.top
                loc=(x,y)
            if size is None:
                width=app_control.BoundingRectangle.width()
                height=app_control.BoundingRectangle.height()
                size=(width,height)
            x,y=loc
            width,height=size
            app_control.MoveWindow(x,y,width,height)
            return (f'{active_app.name} resized to {width}x{height} at {x},{y}.',0)
    
    def is_app_running(self,name:str)->bool:
        apps={app.name:app for app in self.get_apps()}
        return process.extractOne(name,list(apps.keys()),score_cutoff=60) is not None
    
    def app(self,mode:Literal['launch','switch','resize'],name:Optional[str]=None,loc:Optional[tuple[int,int]]=None,size:Optional[tuple[int,int]]=None):
        match mode:
            case 'launch':
                response,status=self.launch_app(name)
                sleep(1.25)
                if status!=0:
                    return response
                consecutive_waits=10
                for _ in range(consecutive_waits):
                    if not self.is_app_running(name):
                        sleep(1.25)
                    else:
                        return f'{name.title()} launched.'
                return f'Launching {name.title()} wait for it to come load.'
            case 'resize':
                response,status=self.resize_app(size=size,loc=loc)
                if status!=0:
                    return response
                else:
                    return response
            case 'switch':
                response,status=self.switch_app(name)
                if status!=0:
                    return response
                else:
                    return response
        
    def launch_app(self,name:str)->tuple[str,int]:
        apps_map=self.get_apps_from_start_menu()
        matched_app=process.extractOne(name,apps_map.keys(),score_cutoff=70)
        if matched_app is None:
            return (f'{name.title()} not found in start menu.',1)
        app_name,_=matched_app
        appid=apps_map.get(app_name)
        if appid is None:
            return (f'{name.title()} not found in start menu.',1)
        if appid.endswith('.exe'):
            command=f"Start-Process '{appid}'"
        else:
            command=f"Start-Process shell:AppsFolder\\{appid}"
        response,status=self.execute_command(command)
        return response,status
    
    def switch_app(self,name:str):
        apps={app.name:app for app in [self.desktop_state.active_app]+self.desktop_state.apps if app is not None}
        matched_app:Optional[tuple[str,float]]=process.extractOne(name,list(apps.keys()),score_cutoff=70)
        if matched_app is None:
            return (f'Application {name.title()} not found.',1)
        app_name,_=matched_app
        app=apps.get(app_name)
        target_handle=app.handle

        # Always attempt a full activation sequence. Restoring alone can bring a window
        # on-screen without giving it keyboard focus.
        if win32gui.IsIconic(target_handle):
            try:
                win32gui.ShowWindow(target_handle, win32con.SW_RESTORE)
            except Exception:
                # fall back to uiautomation wrapper if needed
                try:
                    uia.ShowWindow(target_handle, win32con.SW_RESTORE)
                except Exception:
                    pass
        self.bring_window_to_top(target_handle)
        content=f'Switched to {app_name.title()} window.'
        return content,0
    
    def bring_window_to_top(self,target_handle:int):
        """
        Best-effort "activate + focus" for a top-level window.

        Notes:
        - Windows actively prevents background processes from stealing focus.
        - `BringWindowToTop` only adjusts Z-order; it doesn't guarantee keyboard focus.
        - We attach input threads to improve the odds that `SetForegroundWindow` succeeds.
        """
        foreground_handle=win32gui.GetForegroundWindow()
        foreground_thread,_=win32process.GetWindowThreadProcessId(foreground_handle)
        target_thread,_=win32process.GetWindowThreadProcessId(target_handle)
        current_thread = ctypes.windll.kernel32.GetCurrentThreadId()
        try:
            ctypes.windll.user32.AllowSetForegroundWindow(-1)
            # Attach our thread to both the foreground and target threads.
            try:
                win32process.AttachThreadInput(current_thread, foreground_thread, True)
            except Exception:
                pass
            try:
                win32process.AttachThreadInput(current_thread, target_thread, True)
            except Exception:
                pass

            # Ensure the window is shown (SW_SHOW is a no-op if already visible).
            try:
                win32gui.ShowWindow(target_handle, win32con.SW_SHOW)
            except Exception:
                pass

            # Try to foreground + focus.
            win32gui.SetForegroundWindow(target_handle)
            try:
                win32gui.SetActiveWindow(target_handle)
            except Exception:
                pass
            try:
                win32gui.SetFocus(target_handle)
            except Exception:
                pass
            try:
                win32gui.BringWindowToTop(target_handle)
            except Exception:
                pass

            # Extra nudge: try UIAutomation focus on the window element.
            try:
                uia.ControlFromHandle(target_handle).SetFocus()
            except Exception:
                pass

            # Optional last resort: click the title bar to force focus (off by default).
            if os.getenv("WINDOWS_MCP_FORCE_FOCUS_CLICK", "false").strip().lower() in {"1", "true", "yes", "y"}:
                try:
                    if win32gui.GetForegroundWindow() != target_handle:
                        left, top, right, bottom = win32gui.GetWindowRect(target_handle)
                        x = int((left + right) / 2)
                        y = int(top + 12)
                        pg.click(x, y, clicks=1, duration=0.05)
                        time.sleep(0.05)
                except Exception:
                    pass
        except Exception as e:
            logger.error(f'Failed to bring window to top: {e}')
        finally:
            try:
                win32process.AttachThreadInput(current_thread, target_thread, False)
            except Exception:
                pass
            try:
                win32process.AttachThreadInput(current_thread, foreground_thread, False)
            except Exception:
                pass
    
    def get_element_handle_from_label(self,label:int)->uia.Control:
        tree_state=self.desktop_state.tree_state
        element_node=tree_state.interactive_nodes[label]
        xpath=element_node.xpath
        element_handle=self.get_element_from_xpath(xpath)
        return element_handle
    
    def get_coordinates_from_label(self,label:int)->tuple[int,int]:
        element_handle=self.get_element_handle_from_label(label)
        bounding_rectangle=element_handle.BoundingRectangle
        return bounding_rectangle.xcenter(),bounding_rectangle.ycenter()
        
    def click(self,loc:tuple[int,int],button:str='left',clicks:int=2):
        x,y=loc
        pg.click(x,y,button=button,clicks=clicks,duration=0.1)

    def type(self,loc:tuple[int,int],text:str,caret_position:Literal['start','end','none']='none',clear:Literal['true','false']='false',press_enter:Literal['true','false']='false'):
        x,y=loc
        pg.leftClick(x,y)
        if caret_position == 'start':
            pg.press('home')
        elif caret_position == 'end':
            pg.press('end')
        else:
            pass
        if clear=='true':
            pg.sleep(0.5)
            pg.hotkey('ctrl','a')
            pg.press('backspace')
        pg.typewrite(text,interval=0.02)
        if press_enter=='true':
            pg.press('enter')

    def scroll(self,loc:tuple[int,int]=None,type:Literal['horizontal','vertical']='vertical',direction:Literal['up','down','left','right']='down',wheel_times:int=1)->str|None:
        if loc:
            self.move(loc)
        match type:
            case 'vertical':
                match direction:
                    case 'up':
                        uia.WheelUp(wheel_times)
                    case 'down':
                        uia.WheelDown(wheel_times)
                    case _:
                        return 'Invalid direction. Use "up" or "down".'
            case 'horizontal':
                match direction:
                    case 'left':
                        pg.keyDown('Shift')
                        pg.sleep(0.05)
                        uia.WheelUp(wheel_times)
                        pg.sleep(0.05)
                        pg.keyUp('Shift')
                    case 'right':
                        pg.keyDown('Shift')
                        pg.sleep(0.05)
                        uia.WheelDown(wheel_times)
                        pg.sleep(0.05)
                        pg.keyUp('Shift')
                    case _:
                        return 'Invalid direction. Use "left" or "right".'
            case _:
                return 'Invalid type. Use "horizontal" or "vertical".'
        return None
    
    def drag(self,loc:tuple[int,int]):
        x,y=loc
        pg.sleep(0.5)
        pg.dragTo(x,y,duration=0.6)

    def move(self,loc:tuple[int,int]):
        x,y=loc
        pg.moveTo(x,y,duration=0.1)

    def shortcut(self,shortcut:str):
        shortcut=shortcut.split('+')
        if len(shortcut)>1:
            pg.hotkey(*shortcut)
        else:
            pg.press(''.join(shortcut))

    def multi_select(self,press_ctrl:Literal['true','false']='false',elements:list[tuple[int,int]|int]=[]):
        if press_ctrl=='true':
            pg.keyDown('ctrl')
        for element in elements:
            x,y=element
            pg.click(x,y,duration=0.2)
            pg.sleep(0.5)
        pg.keyUp('ctrl')
    
    def multi_edit(self,elements:list[tuple[int,int,str]|tuple[int,str]]):
        for element in elements:
            x,y,text=element
            self.type((x,y),text=text,clear='true')
    
    def scrape(self,url:str)->str:
        response=requests.get(url,timeout=10)
        html=response.text
        content=markdownify(html=html)
        return content
    
    def get_app_size(self,control:uia.Control):
        window=control.BoundingRectangle
        if window.isempty():
            return Size(width=0,height=0)
        return Size(width=window.width(),height=window.height())
    
    def is_app_visible(self,app)->bool:
        is_minimized=self.get_app_status(app)!=Status.MINIMIZED
        size=self.get_app_size(app)
        area=size.width*size.height
        is_overlay=self.is_overlay_app(app)
        return not is_overlay and is_minimized and area>10
    
    def is_overlay_app(self,element:uia.Control) -> bool:
        no_children = len(element.GetChildren()) == 0
        is_name = "Overlay" in element.Name.strip()
        return no_children or is_name
        
    def get_apps(self) -> list[App]:
        """
        Enumerate user-facing top-level windows using Win32 APIs.

        This intentionally avoids UIAutomation enumeration, which can intermittently hang on some providers.
        """
        apps: list[App] = []
        try:
            def enum_cb(hwnd: int, _lparam):
                try:
                    if not hwnd or not win32gui.IsWindow(hwnd):
                        return True
                    if not win32gui.IsWindowVisible(hwnd):
                        return True

                    title = (win32gui.GetWindowText(hwnd) or "").strip()
                    if not title:
                        return True

                    # Prefer skipping known noisy system windows by class name.
                    try:
                        class_name = win32gui.GetClassName(hwnd)
                    except Exception:
                        class_name = ""
                    from windows_mcp.desktop.config import EXCLUDED_APPS
                    if class_name in EXCLUDED_APPS:
                        return True

                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    if not (style & win32con.WS_MINIMIZEBOX and style & win32con.WS_MAXIMIZEBOX):
                        return True

                    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                    width, height = max(0, right - left), max(0, bottom - top)
                    if width * height <= 10:
                        return True

                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    is_min = bool(win32gui.IsIconic(hwnd))
                    if hasattr(win32gui, "IsZoomed"):
                        is_max = bool(win32gui.IsZoomed(hwnd))  # type: ignore[attr-defined]
                    else:
                        try:
                            show_cmd = win32gui.GetWindowPlacement(hwnd)[1]
                            is_max = show_cmd == win32con.SW_SHOWMAXIMIZED
                        except Exception:
                            is_max = False
                    is_visible = bool(win32gui.IsWindowVisible(hwnd))
                    status = (
                        Status.MINIMIZED if is_min
                        else Status.MAXIMIZED if is_max
                        else Status.NORMAL if is_visible
                        else Status.HIDDEN
                    )
                    apps.append(App(
                        name=title,
                        depth=len(apps),
                        status=status,
                        size=Size(width=width, height=height),
                        handle=hwnd,
                        process_id=pid,
                    ))
                except Exception:
                    # Best-effort enumeration: ignore individual window failures.
                    pass
                return True

            win32gui.EnumWindows(enum_cb, None)
        except Exception as ex:
            logger.error(f"Error in get_apps: {ex}")
        return apps
    
    def get_xpath_from_element(self,element:uia.Control):
        current=element
        if current is None:
            return ""
        path_parts=[]
        while current is not None:
            parent=current.GetParentControl()
            if parent is None:
                # we are at the root node
                path_parts.append(f'{current.ControlTypeName}')
                break
            children=parent.GetChildren()
            same_type_children=["-".join(map(lambda x:str(x),child.GetRuntimeId())) for child in children if child.ControlType==current.ControlType]
            index=same_type_children.index("-".join(map(lambda x:str(x),current.GetRuntimeId())))
            if same_type_children:
                path_parts.append(f'{current.ControlTypeName}[{index+1}]')
            else:
                path_parts.append(f'{current.ControlTypeName}')
            current=parent
        path_parts.reverse()
        xpath="/".join(path_parts)
        return xpath

    def get_element_from_xpath(self,xpath:str)->uia.Control:
        pattern = re.compile(r'(\w+)(?:\[(\d+)\])?')
        parts=xpath.split("/")
        root=uia.GetRootControl()
        element=root
        for part in parts[1:]:
            match=pattern.fullmatch(part)
            if match is None:
                continue
            control_type, index=match.groups()
            index=int(index) if index else None
            children=element.GetChildren()
            same_type_children=list(filter(lambda x:x.ControlTypeName==control_type,children))
            if index:
                element=same_type_children[index-1]
            else:
                element=same_type_children[0]
        return element
    
    def get_windows_version(self)->str:
        response,status=self.execute_command("(Get-CimInstance Win32_OperatingSystem).Caption")
        if status==0:
            return response.strip()
        return "Windows"
    
    def get_user_account_type(self)->str:
        response,status=self.execute_command("(Get-LocalUser -Name $env:USERNAME).PrincipalSource")
        return "Local Account" if response.strip()=='Local' else "Microsoft Account" if status==0 else "Local Account"
    
    def get_dpi_scaling(self):
        user32 = ctypes.windll.user32
        dpi = user32.GetDpiForSystem()
        return dpi / 96.0
    
    def get_screen_size(self)->Size:
        width, height = uia.GetScreenSize()
        return Size(width=width,height=height)

    def get_screenshot(self)->Image.Image:
        return pg.screenshot()
    
    @contextmanager
    def auto_minimize(self):
        try:
            handle = uia.GetForegroundWindow()
            uia.ShowWindow(handle, win32con.SW_MINIMIZE)
            yield
        finally:
            uia.ShowWindow(handle, win32con.SW_RESTORE)