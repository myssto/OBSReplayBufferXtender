from abc import ABC, abstractmethod
import os
import platform
import subprocess
import obspython as o


class WindowNameResolver(ABC):
    """
    Base class for a window name resolution handler
    """

    disallowed_chars: list[str] = []

    def __init__(self) -> None:
        pass

    @abstractmethod
    def get_focused_window_name(self) -> str:
        """
        Gets the name of the window currently in focus
        """
        pass

    def sanitize_string(self, txt: str) -> str:
        """
        Removes chars from strings that are disallowed in file names
        """
        for char in self.disallowed_chars:
            txt = txt.replace(char, "")
        return txt.strip()


class WindowsWindowNameResolver(WindowNameResolver):
    """
    Handles window name resolution on Windows platforms
    """

    def __init__(self) -> None:
        super().__init__()

        # Load windows-specific modules
        import psutil
        import win32api as wapi
        import win32gui as wgui
        import win32process as wproc

        self.psutil = psutil
        self.wapi = wapi
        self.wgui = wgui
        self.wproc = wproc

        self.disallowed_chars = ["\\", "/", ":", '*',
                                 "?", '"', "<", ">", "|", ".exe", "$"]

    def get_focused_window_name(self) -> str:
        return self.sanitize_string(self.get_application_name() or self.get_window_display_name())

    def get_application_name(self) -> str:
        """
        Gets the name of the window in focus by inspecting file version info.
        Generally yields better results than get_window_display_name when available.

        See https://stackoverflow.com/a/31119785
        """
        application_name = ""

        # Get the process ID of the window in focus
        _, pid = self.wproc.GetWindowThreadProcessId(
            self.wgui.GetForegroundWindow())

        # Get the path to the window's executable
        exe_path = self.psutil.Process(pid).exe()

        try:
            language, codepage = self.wapi.GetFileVersionInfo(
                exe_path, '\\VarFileInfo\\Translation')[0]
            stringFileInfo = u'\\StringFileInfo\\%04X%04X\\%s' % (
                language, codepage, "FileDescription")
            application_name = self.wapi.GetFileVersionInfo(
                exe_path, stringFileInfo)
        except:
            pass

        return application_name

    def get_window_display_name(self) -> str:
        """
        Gets the literal display name of the window in focus using win32api
        """
        return self.wgui.GetWindowText(self.wgui.GetForegroundWindow())


class LinuxWindowNameResolver(WindowNameResolver):
    """
    Handles window name resolution on Linux platforms
    """

    IS_FLATPAK = "FLATPAK_ID" in os.environ or os.path.exists("/.flatpak-info")

    def __init__(self) -> None:
        super().__init__()

        self.disallowed_chars = ["/"]

    def get_focused_window_name(self) -> str:
        # If running from the flatpak sandbox, we need to run xprop from the host shell
        args = ["flatpak-spawn", "--host"] if self.IS_FLATPAK else []
        window_title = ""

        try:
            # Get active window ID
            window_id = subprocess.check_output(
                [*args, "xprop", "-root", "_NET_ACTIVE_WINDOW"]).decode().split()[-1]
            # Get the window title, sanitize UTF-8 and escaped strings
            window_title = subprocess.check_output(
                [*args, "xprop", "-id", window_id, "WM_NAME"]).decode().strip().split('=', 1)[-1].strip(' "')
        except:
            pass

        return self.sanitize_string(window_title)


class ReplayBufferXtender:
    """
    All main script logic lives here
    """

    # Script settings
    base_dir = None
    prepend_window_name = True

    def __init__(self) -> None:
        if (platform.system() == "Linux"):
            self.window_resolver = LinuxWindowNameResolver()
        else:
            self.window_resolver = WindowsWindowNameResolver()

    def event_handler(self, event, *_) -> None:
        """
        Internal, handles OBS events
        """
        if event == o.OBS_FRONTEND_EVENT_REPLAY_BUFFER_SAVED:
            try:
                self.move_video()
            except BaseException as e:
                print(e)

    def get_last_replay_path(self) -> str:
        """
        Retrieve last replay buffer output
        """
        replay_buffer = o.obs_frontend_get_replay_buffer_output()
        cd = o.calldata_create()
        ph = o.obs_output_get_proc_handler(replay_buffer)

        o.proc_handler_call(
            handler=ph,
            name="get_last_replay",
            params=cd
        )
        path = o.calldata_string(data=cd, name="path")

        o.obs_output_release(replay_buffer)
        return path

    def move_video(self) -> None:
        """
        Moves the last file the Replay Buffer created\n
        into `base_dir\\{cur_win_name}\\base_name`
        """
        # Get last replay path
        lr_orig_fullpath = self.get_last_replay_path()
        # Parse into directory and file name
        lr_orig_dir, lr_fname = os.path.split(lr_orig_fullpath)

        # First try to get the name from the executable's version information resource
        # If that fails, try to get the name from the window text
        sub_dir = self.window_resolver.get_focused_window_name() or "Unknown"

        # Replace the word "Replay" in the file name with the window name
        if self.prepend_window_name:
            lr_fname = lr_fname.replace("Replay", sub_dir)

        # Determine base directory
        if self.base_dir:
            lr_base_dir = self.base_dir
        else:
            lr_base_dir = lr_orig_dir

        # Create final directory
        lr_dir = os.path.join(lr_base_dir, sub_dir)
        if not os.path.exists(lr_dir):
            os.mkdir(lr_dir)

        # Move the file
        os.rename(lr_orig_fullpath, os.path.join(lr_dir, lr_fname))


inst = ReplayBufferXtender()


def script_description() -> str:
    return """
    <center><h2>ReplayBufferXtender</h2>
    <p>This is an OBS Python script that automatically renames and organizes video
    files generated by the Replay Buffer based on the window in focus. Attempts to 
    emulate the file naming conventions of Nvidia Shadowplay as closely as possible!</p>

    <p><br><b>Author: Myssto</b>
    <br><a href=https://github.com/myssto/OBSReplayBufferXtender>Github</a></p>
    </center>
    """


def script_load(_) -> None:
    o.obs_frontend_add_event_callback(on_event)


def script_unload() -> None:
    o.obs_frontend_remove_event_callback(on_event)


def script_properties() -> any:
    p = o.obs_properties_create()

    # Option to overide existing OBS Recording output path for Replay Buffer only
    o.obs_properties_add_path(
        props=p,
        name="baseSavePath",
        description="Base Save Path",
        type=o.OBS_PATH_DIRECTORY,
        filter=None,
        default_path=None
    )

    # Option to prepend the window name to the replay file like Shadowplay
    # Will replace the default "Replay" text from the file name
    o.obs_properties_add_bool(
        props=p,
        name="prependWindowName",
        description="Prepend Window Name"
    )

    return p


def script_defaults(s) -> None:
    o.obs_data_set_default_bool(
        data=s,
        name="prependWindowName",
        val=ReplayBufferXtender.prepend_window_name
    )


def script_update(s) -> None:
    inst.base_dir = o.obs_data_get_string(s, "baseSavePath")
    inst.prepend_window_name = o.obs_data_get_bool(s, "prependWindowName")


def on_event(event, *_) -> None:
    """
    Called by OBS when events are fired\n
    This exists because you cannot bind instance methods as callbacks in OBS
    """

    inst.event_handler(event)
