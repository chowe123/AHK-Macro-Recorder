import ctypes
from ctypes import wintypes

# Windows Constants
WM_COPYDATA = 0x4A

class COPYDATASTRUCT(ctypes.Structure):
    _fields_ = [
        ('dwData', ctypes.c_size_t),
        ('cbData', wintypes.DWORD),
        ('lpData', ctypes.c_void_p)
    ]

# Define Win32 function signatures with robust 64-bit types
user32 = ctypes.windll.user32
# Use c_void_p for all window handles (HWND) to prevent signed-integer overflow
user32.GetClassNameW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int
user32.GetWindowTextLengthW.argtypes = [ctypes.c_void_p]
user32.GetWindowTextLengthW.restype = ctypes.c_int
user32.GetWindowTextW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int
user32.EnumWindows.argtypes = [ctypes.WINFUNCTYPE(wintypes.BOOL, ctypes.c_void_p, wintypes.LPARAM), wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL
# Use c_size_t for WPARAM and LPARAM to ensure 64-bit pointer compatibility
user32.SendMessageW.argtypes = [ctypes.c_void_p, wintypes.UINT, ctypes.c_size_t, ctypes.c_size_t]
user32.SendMessageW.restype = ctypes.c_void_p
# Needed for PID discovery
user32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD

def load_ahk_profile(profile_name):
    # Prepare the command payload
    payload = f"LOAD_PROFILE:{profile_name}"
    
    # We'll search for all windows with "Config Manager" in the title
    # and windows belonging to AutoHotkey class
    target_titles = ["config manager", "config manager.ahk", "config manager.exe"]
    target_classes = ["AutoHotkey", "AutoHotkeyGUI"]
    
    found_hwnds = []

    def enum_windows_proc(hwnd, lParam):
        try:
            # Check Class
            class_buff = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, class_buff, 256)
            cls = class_buff.value
            
            if cls in target_classes:
                # Check Title
                length = user32.GetWindowTextLengthW(hwnd)
                buff = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buff, length + 1)
                title = buff.value
                
                # Broad match on title
                if any(t in title.lower() for t in target_titles):
                    if hwnd not in found_hwnds:
                        # Prioritize class "AutoHotkey" by putting it at the front
                        if cls == "AutoHotkey":
                            found_hwnds.insert(0, hwnd)
                        else:
                            found_hwnds.append(hwnd)
        except Exception:
            pass
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, ctypes.c_void_p, wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(enum_windows_proc), 0)

    if not found_hwnds:
        print("Error: No 'Config Manager' instances found running.")
        return False
        
    # Use ctypes to create a utf-16le string buffer
    encoded_string = payload.encode('utf-16le')
    string_buffer = ctypes.create_string_buffer(encoded_string + b'\x00\x00')
    
    cds = COPYDATASTRUCT()
    cds.dwData = 0
    cds.cbData = len(encoded_string) + 2
    cds.lpData = ctypes.addressof(string_buffer)
    
    # De-duplicate: one message per PID, prioritizing the Main window
    final_hwnds = {}
    for hwnd in found_hwnds:
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        p_id = pid.value
        class_buff = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buff, 256)
        curr_cls = class_buff.value
        
        if p_id not in final_hwnds or curr_cls == "AutoHotkey":
            final_hwnds[p_id] = hwnd

    success_count = 0
    for p_id, hwnd in final_hwnds.items():
        # Send profile change (cast address to size_t)
        result = user32.SendMessageW(hwnd, WM_COPYDATA, 0, ctypes.addressof(cds))
        if result == 1:
            success_count += 1
    
    if success_count > 0:
        print(f"Successfully broadcasted to {success_count} instance(s) for profile: {profile_name}")
        return True
    else:
        print("Instances found but none responded (Message filtered or Profile not found?)")
        return False

# Usage Example:
#load_ahk_profile("Default")

load_ahk_profile("Test Profile")
