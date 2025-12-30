import os

INTERACTIVE_CONTROL_TYPE_NAMES=set([
    'ButtonControl',
    'ListItemControl',
    'MenuItemControl',
    'EditControl',
    'CheckBoxControl',
    'RadioButtonControl',
    'ComboBoxControl',
    'HyperlinkControl',
    'SplitButtonControl',
    'TabItemControl',
    'TreeItemControl',
    'DataItemControl',
    'HeaderItemControl',
    'TextBoxControl',
    'SpinnerControl',
    'ScrollBarControl'
])

DOCUMENT_CONTROL_TYPE_NAMES=set([
    'DocumentControl'
])

STRUCTURAL_CONTROL_TYPE_NAMES = set([
    'PaneControl',
    'GroupControl',
    'CustomControl'
])

INFORMATIVE_CONTROL_TYPE_NAMES=set([
    'TextControl',
    'ImageControl',
    'StatusBarControl',
    # 'ProgressBarControl',
    # 'ToolTipControl',
    # 'TitleBarControl',
    # 'SeparatorControl',
    # 'HeaderControl',
    # 'HeaderItemControl',
])

DEFAULT_ACTIONS=set([
    'Click',
    'Press',
    'Jump',
    'Check',
    'Uncheck',
    'Double Click'
])

THREAD_MAX_RETRIES = 3

# ---- Performance / safety knobs (env overridable) ----
#
# The UIAutomation tree can block on some windows/providers. These timeouts ensure
# `State-Tool` returns in bounded time with partial results rather than hanging.
TREE_STATE_TIMEOUT_S = float(os.getenv("WINDOWS_MCP_TREE_STATE_TIMEOUT_S", "6"))
TREE_STATE_TIMEOUT_DOM_S = float(os.getenv("WINDOWS_MCP_TREE_STATE_TIMEOUT_DOM_S", "10"))

# Timeout for the initial "desktop root -> children" enumeration. This is on the critical path
# and can hang intermittently on some systems/providers.
ROOT_ENUM_TIMEOUT_S = float(os.getenv("WINDOWS_MCP_ROOT_ENUM_TIMEOUT_S", "2.0"))

# If root enumeration times out, temporarily disable UIA enumeration to avoid leaking stuck threads.
UIA_TIMEOUT_COOLDOWN_S = float(os.getenv("WINDOWS_MCP_UIA_TIMEOUT_COOLDOWN_S", "15"))

# Threadpool size used for app-wise UIA traversal.
UIA_MAX_WORKERS = max(1, int(os.getenv("WINDOWS_MCP_UIA_MAX_WORKERS", "8")))

# Cap how many elements we stringify in the State-Tool output (does not limit traversal).
MAX_INTERACTIVE_ROWS = max(1, int(os.getenv("WINDOWS_MCP_MAX_INTERACTIVE_ROWS", "200")))
MAX_SCROLLABLE_ROWS = max(1, int(os.getenv("WINDOWS_MCP_MAX_SCROLLABLE_ROWS", "120")))