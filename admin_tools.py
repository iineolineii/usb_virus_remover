import ctypes
import os
from pathlib import Path
import shutil
from enum import IntEnum
import sys
from tempfile import mkdtemp
from typing import Optional


# Source: https://stackoverflow.com/a/42787518
class SW(IntEnum):
    HIDE = 0
    MAXIMIZE = 3
    MINIMIZE = 6
    RESTORE = 9
    SHOW = 5
    SHOWDEFAULT = 10
    SHOWMAXIMIZED = 3
    SHOWMINIMIZED = 2
    SHOWMINNOACTIVE = 7
    SHOWNA = 8
    SHOWNOACTIVATE = 4
    SHOWNORMAL = 1

class Errors(IntEnum):
    ZERO = 0
    FILE_NOT_FOUND = 2
    PATH_NOT_FOUND = 3
    BAD_FORMAT = 11
    ACCESS_DENIED = 5
    ASSOC_INCOMPLETE = 27
    DDE_BUSY = 30
    DDE_FAIL = 29
    DDE_TIMEOUT = 28
    DLL_NOT_FOUND = 32
    NO_ASSOC = 31
    OOM = 8
    SHARE = 26

SUCCESS_EXIT_CODE = 32


def is_admin() -> bool:
    """Check if current user has Administrator privelegies"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def restart_as_admin(script_path: str, *, window_title: Optional[str] = None) -> str:
    """
    Restart the provided script with the Administrator privelegies.
    Returns a path to a temporary directory with a copy of the script.

    Example:
    ```python
    from admin_tools import is_admin, restart_as_admin

    if not is_admin():
        print("Restarting as administrator...")
        restart_as_admin(__file__)
        exit()

    # Rest of your code here...
    ```
    """

    # Make a temporary copy of the script module
    tempdir = mkdtemp()
    temporary_copy = shutil.copy(script_path, tempdir)
    shutil.copy(__file__, tempdir)

    # Make a restart command
    python   = f'"{str(Path(sys.executable).with_name("python.exe"))}"' # IDLE Support
    script   = f'"{temporary_copy}"'
    username = f'"{os.getlogin() }"'
    command = [python, script, username]

    # If a window title is provided, set it
    if window_title:
        command.insert(0, f"title {window_title} &&")

    restart = " ".join(command)

    # Request UAC promotion to restart the script
    print("\nRequesting UAC promotion to perform a restart:", restart)
    execute(f'"{restart}"', persistent_window=True, run_as_admin=True)
    return tempdir


def execute(
    command: str,
    *,
    persistent_window: bool = False,
    run_as_admin: bool = False,
) -> None:
    """
    Execute given command in a new cmd.exe window.

    Parameters:
        command (`str`): \
            The command to be executed.

        persistent_window (`bool`): \
            Whether to remain the command prompt window after the command execution.\
            Defaults to False.

        run_as_admin (`bool`): \
            Whether to request UAC promotion to run the command with Administrator privileges.\
            Defaults to False.
    """
    # Build a command based on the given arguments
    window_view = SW.SHOWNORMAL if persistent_window else SW.HIDE
    command = ("/k" if persistent_window else "/c") + command
    mode =  "runas" if run_as_admin else "open"

    # Execute the command
    exit_code = ctypes.windll.shell32.ShellExecuteW(
        None, mode, "cmd.exe", command, None, window_view
    )

    # Check for errors
    if exit_code > SUCCESS_EXIT_CODE:
        print(f"CMD exited with code: {exit_code} (success)")
        return

    error = Errors(exit_code).name
    raise RuntimeError(f"Failed to execute command. Error: {error} (exit code: {exit_code})")
