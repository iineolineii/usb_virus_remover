import os
import subprocess
import tempfile
import winreg
from typing import List, Optional, Tuple

ProfileList  = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList"
WSH_Settings = "Software\\Microsoft\\Windows Script Host\\Settings"

FULL_PROFILE_SIDS: List[str] = []


def key_exists(hive: int, key: str) -> bool:
    """
    Checks if a key exists in a given registry hive.
    """
    try:
        with winreg.OpenKey(hive, key, 0, winreg.KEY_READ) as key: # type: ignore
            return True
    except:
        return False


def disable_wsh(SID: Optional[str] = None) -> Tuple[bool, str]:
    """
    Disables Windows Script Host for the Local Machine or a given SID.
    """

    # If SID is None, disable Script Host system-wide
    hive = winreg.HKEY_LOCAL_MACHINE if not SID else winreg.HKEY_USERS
    prefix = "" if not SID else SID + "\\"
    key_path = prefix + WSH_Settings

    # Create the Settings key if it doesn't exist
    if not key_exists(hive, key_path):
        if SID and not key_exists(hive, SID):
            return False, "SID owner is not logged in."

        with winreg.OpenKey(hive, prefix + "Software\\Microsoft", 0, winreg.KEY_WRITE):
            winreg.CreateKey(hive, key_path)

    # Disable Windows Script Host
    try:
        with winreg.OpenKey(hive, key_path, 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "Enabled", 0, winreg.REG_DWORD, 0)
            return True, "Successfully disabled Windows Script Host."

    except PermissionError:
        return False, "Permission denied."

    except FileNotFoundError:
        return False, "Enabled key not found."

    except OSError as e:
        return False, str(e)


def restrict_users():
    succeeded: List[str] = []

    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, ProfileList) as key:
        for i in range(winreg.QueryInfoKey(key)[0]):
            sid = winreg.EnumKey(key, i)

            # Filter out non-user profiles
            result, message = is_user_profile(sid)
            if not result:
                continue

            FULL_PROFILE_SIDS.append(sid)
            print(f"\n[INFO] Found SID: {repr(sid)}")

            result, message = process_sid(sid)
            if result:
                succeeded.append(sid)
                print(f"[SUCCESS] {message}")
            else:
                print(f"[ERROR] {message}")

    if not succeeded:
        print(f"\n🚫 Could not disable Windows Script Host for any users.")
    else:
        print(f"\n[SUCCESS] Successfully disabled Windows Script Host for {len(succeeded)}/{len(FULL_PROFILE_SIDS)} SID(s).")


def restrict_machine():
    print(f"\n[INFO] Disabling Windows Script Host system-wide...")

    result, message = disable_wsh()
    if result:
        print(f"[SUCCESS] {message}")
    else:
        print(f"[ERROR] {message}")


def load_ntuser_dat(sid: str, profile_image_path: str) -> Tuple[bool, str]:
    """
    Loads NTUSER.DAT for a specific SID into HKU.
    """
    try:
        subprocess.run(
            ["reg", "load", f"HKU\\{sid}", f"{profile_image_path}\\NTUSER.DAT"],
            check=True,
            capture_output=True,
            text=True,
            encoding="cp866",
        )
    except subprocess.CalledProcessError as e:
        return False, e.stderr
    else:
        return True, "Successfully loaded NTUSER.DAT."


def unload_ntuser_dat(sid: str) -> Tuple[bool, str]:
    """
    Unloads NTUSER.DAT for a specific SID from HKU.
    """
    try:
        subprocess.run(
            ["reg", "unload", f"HKU\\{sid}"],
            check=True,
            capture_output=True,
            text=True,
            encoding="cp866",
        )
    except subprocess.CalledProcessError as e:
        return False, e.stderr
    else:
        return True, "Successfully unloaded NTUSER.DAT."


def is_user_profile(sid: str) -> Tuple[bool, str]:
    """
    Processes a single SID, attempting to disable Windows Script Host.
    """
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"{ProfileList}\\{sid}") as sid_key:
            sid_value, _ = winreg.QueryValueEx(sid_key, "Sid")

            if (
                not isinstance(sid_value, bytes)
                or not sid_value.startswith(
                    b'\x01\x05\x00\x00\x00\x00\x00\x05'
                    b'\x15\x00\x00\x00'
                )
            ):
                return False, "Not a FullProfile."

            return True, sid

    except FileNotFoundError:
        return False, "Not a FullProfile."

    except PermissionError:
        return False, "Permission denied."

    except OSError as e:
        return False, str(e)


def process_sid(sid: str) -> Tuple[bool, str]:
    """
    Processes a single SID, attempting to disable Windows Script Host.
    """
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"{ProfileList}\\{sid}") as sid_key:
        result, message = disable_wsh(sid)
        if result:
            return True, message
        print("[ERROR]", message)

        profile_image_path, _ = winreg.QueryValueEx(sid_key, "ProfileImagePath")
        result, message = load_ntuser_dat(sid, profile_image_path)
        if not result:
            return False, message
        print("[SUCCESS]", message)

        try:
            result, message = disable_wsh(sid)
            return result, message
        finally:
            unload_ntuser_dat(sid)


def wsh_is_enabled(title: str, message: str, script_path: str = tempfile.mktemp(suffix=".vbs")):
    """
    Checks if Windows Script Host and displays a warning message box if not.
    """
    # Write a test script displaying the warning message
    with open(script_path, "w") as f:
        f.write(f'MsgBox "{message}", vbExclamation, "{title}"')

    # Try to execute the test script
    try:
        subprocess.check_call(script_path, shell=True)
    except subprocess.CalledProcessError:
        return False
    else:
        return True
    # Delete the test script anyway
    finally:
        os.remove(script_path)
