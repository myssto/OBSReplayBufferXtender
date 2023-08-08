import os

import obspython as o
import win32gui as wgui


class ReplayBufferXtender:
    """
    All main script logic lives here
    """

    base_dir = None
    prepend_window_name = True
    use_windowsapps = True
    disallowed_chars = ["\\", "/", ":", '*',
                        "?", '"', "<", ">", "|", ".exe", "$"]

    def __init__(self) -> None:
        pass

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

    def get_focused_window_name(self) -> str:
        """
        Uses the win32api to grab the name of the currently focused window
        """

        # Get the window text
        w_text = wgui.GetWindowText(wgui.GetForegroundWindow())

        # Sanitize window text
        for char in self.disallowed_chars:
            if char in w_text:
                w_text = w_text.replace(char, "")

        return w_text.strip()

    def move_video(self) -> None:
        """
        Moves the last file the Replay Buffer created\n
        into `base_dir\\{cur_win_name}\\base_name`
        """
        # Get replay path and focused window
        lr_orig_fullpath = self.get_last_replay_path()
        lr_orig_dir, lr_fname = os.path.split(lr_orig_fullpath)
        sub_dir = self.get_focused_window_name()

        if not sub_dir:
            if self.use_windowsapps:
                sub_dir = "Windowsapps"
            elif self.base_dir:
                os.rename(lr_orig_fullpath, os.path.join(self.base_dir, lr_fname))
                return
            return

        if self.prepend_window_name:
            lr_fname = lr_fname.replace("Replay", sub_dir)

        # Parse out directories
        if self.base_dir:
            lr_base_dir = self.base_dir
        else:
            lr_base_dir = lr_orig_dir

        lr_dir = os.path.join(lr_base_dir, sub_dir)
        if not os.path.exists(lr_dir):
            os.mkdir(lr_dir)

        os.rename(lr_orig_fullpath, os.path.join(lr_dir, lr_fname))


inst = ReplayBufferXtender()


def script_description() -> str:
    return """
    <center><h2>ReplayBufferXtender</h2>
    <p>This is an OBS Python script that automatically renames video files generated 
    by the Replay Buffer based on the window in focus when they are created using the win32api.
    Attempts to emulate the file naming conventions of Nvidia Shadowplay as closely as possible!</p>

    <p><br><b>Author: Myssto</b>
    <br><a href=https://github.com/myssto/OBSReplayBufferXtender/tree/main#obsreplaybufferxtender>Github Documentation</a></p>
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

    # Option to send unknown window names to a Windowsapps directory
    o.obs_properties_add_bool(
        props=p,
        name="useWindowsapps",
        description="Use Windowsapps for Unknown Programs"
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
        name="useWindowsapps",
        val=ReplayBufferXtender.use_windowsapps
    )

    o.obs_data_set_default_bool(
        data=s,
        name="prependWindowName",
        val=ReplayBufferXtender.prepend_window_name
    )


def script_update(s) -> None:
    inst.base_dir = o.obs_data_get_string(s, "baseSavePath")
    inst.use_windowsapps = o.obs_data_get_bool(s, "useWindowsapps")
    inst.prepend_window_name = o.obs_data_get_bool(s, "prependWindowName")


def on_event(event, *_) -> None:
    """
    Called by OBS when events are fired\n
    This exists because you cannot bind instance methods as callbacks in OBS
    """

    inst.event_handler(event)
