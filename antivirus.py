import os
import shutil
import subprocess
import sys
from pathlib import Path

import wsh_disabler
from admin_tools import is_admin, restart_as_admin
from wsh_disabler import wsh_is_enabled, restrict_machine, restrict_users


# Restart the antivirus with the Administrator privelegies
if not is_admin():
    tempdir = restart_as_admin(__file__, window_title="ANTIVIRUS")
    # Bring along dependencies
    shutil.copy(wsh_disabler.__file__, tempdir)
    sys.exit()


def delete_virus_folder(username: str) -> None:
    virus_folder = Path(f"C:\\Users\\{username}\\AppData\\Roaming\\WindowsServices")
    # NOTE: For some reason Python gets permission denied
    # to delete this folder so we're using cmd here instead
    if virus_folder.is_dir():
        command = f'rmdir /S /Q "{virus_folder}"')
    elif virus_folder.exists():
        command = f'rm "{virus_folder}"') # In case if WindowsServices is a file
    else:
        return print(f'[ERROR] Virus folder "{virus_folder}" does not exist')

    execute(
        command,
        'Virus folder "{virus_folder}" was successfully removed',
        'Failed to remove virus folder "{virus_folder}"',
    )


def clear_startup(username: str) -> None:
    startup_path = Path(f"C:\\Users\\{username}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup")
    possible_viruses = startup_path.glob("helper*")

    for virus_startup in possible_viruses:
        try:
            virus_startup.unlink()
            print(f"[SUCCESS] {virus_startup.name} was successfully removed from shell:startup.")
        except Exception as e:
            print(f"[ERROR] Failed to remove {virus_startup} from shell:startup: {e}")

    # Open shell:startup
    os.system(f"explorer.exe {startup_path}")


def execute(
    command: str,
    success_message: str = "",
    failure_message: str = "",
    *,
    shell = False
):
    """
    Execute a command and return its output.
    """
    try:
        result = subprocess.run(command, shell=shell, encoding="cp866", text=True, check=True, capture_output=True)
    except subprocess.CalledProcessError as e:
        stderr = e.stderr.strip() if e.stderr else ""
        print("\n[ERROR] ", failure_message, ": ", stderr, sep="")
    else:
        stdout = result.stdout.strip() if result.stdout else ""
        print("\n[SUCCESS] ", success_message, ": ", stdout, sep="")


if __name__ == "__main__":
    # Previous username should be pointed to argv on restarting
    username = sys.argv[1]

    # Remove previous user from the Administrators group
    execute(
        f"net localgroup \"Administrators\" /delete \"{username}\"",
        f"{username} removed from the Administrators group",
        "Failed to remove user from the Administrators group"
    )

    # Disable password expiry
    execute(
        "net accounts /maxpwage:unlimited",
        "Password expiry disabled",
        "Failed to disable password expiry"
    )

    # Kill Windows Script Host processes
    execute(
        "taskkill /F /IM wscript.exe",
        "wscript.exe was successfully terminated",
        "Windows Script Host is not running"
    )


    # Delete the virus folder
    delete_virus_folder(username)

    # Clear shell:startup
    clear_startup(username)


    # Disable Windows Script Host for all local users
    restrict_users()

    # Disable Windows Script Host system-wide
    restrict_machine()

    # Check disabled Windows Script Host
    if wsh_is_enabled(
        title   = "Security Warning",
        message = (
            "Please note that Windows Script Host is enabled on this computer. "
            "This may create additional security risks."
        )
    ):
        raise RuntimeError("Windows Script Host was not disabled!")
