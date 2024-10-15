""" EDMC plugin for reading out game chat. """
from concurrent.futures import ThreadPoolExecutor
import os
import sys
import tempfile
import logging
import tkinter as tk
from tkinter import ttk
from collections import namedtuple
from enum import Enum
from typing import Optional, Dict, Any

import myNotebook as nb
from config import config
from config import appname

_plugin_name = os.path.basename(os.path.dirname(__file__))
_logger = logging.getLogger(f'{appname}.{_plugin_name}')

_VERSION = 'dev'

# Created with
#  mkdir libs
#  pip install edge-tts -t libs
#  python -m zipapp . -o ../deps.pyz -m edge_tts.utils:main
sys.path.append(os.path.join(os.path.dirname(__file__), 'deps.pyz'))
# sys.path.append(os.path.join(os.path.dirname(__file__), 'lib'))
import edge_tts  # noqa: E402 pylint: disable=wrong-import-position,import-error,wrong-import-order


# Define a namedtuple type called 'Person'
_TTSConfig = namedtuple('TTSConfig', ['voice', 'volume', 'rate'])


class _EdgeTTSEngine:
    """ TTS engine based on edge-tts. """

    def create(self, message, sound_file, tts_config):
        """ Create sound file from text message. """
        communicate = edge_tts.Communicate(
            message,
            tts_config.voice,
            volume=tts_config.volume,
            rate=tts_config.rate
        )
        for chunk in communicate.stream_sync():
            if chunk["type"] == "audio":
                sound_file.write(chunk["data"])


class _WindowsPlaybackEngine:
    """ Playback engine using the Windows MCI functions. """

    def __init__(self):
        from ctypes import windll, wintypes  # pylint: disable=import-outside-toplevel
        self.__mci_send_string_w = windll.winmm.mciSendStringW
        self.__mci_send_string_w.argtypes = [
            wintypes.LPCWSTR,
            wintypes.LPWSTR,
            wintypes.UINT, wintypes.HANDLE
        ]
        self.__mci_send_string_w.restype = wintypes.UINT

    def __send_command(self, command):
        """ Send command to MCI system. """
        error_code = self.__mci_send_string_w(command, None, 0, 0)
        if error_code != 0:
            raise Exception(f'Failed to execute "{command}": {error_code}')

    def play(self, sound_file):
        """ Play sound. """
        try:
            self.__send_command(f'open "{sound_file}"')
            self.__send_command(f'play "{sound_file}" wait')
        finally:
            self.__send_command(f'close "{sound_file}"')


class _MessagePlayer:
    """ Message player, hanlding tts conversion and playback. """

    def __init__(self):
        self.__tts_pool = ThreadPoolExecutor(max_workers=1)
        self.__play_pool = ThreadPoolExecutor(max_workers=1)
        self.__tts_engine = _EdgeTTSEngine()
        self.__play_engine = _WindowsPlaybackEngine()

    def shutdown(self, wait=True):
        """ Shutdown player and all thread pools. """
        self.__tts_pool.shutdown(wait=wait)
        self.__play_pool.shutdown(wait=wait)

    def play_message(self, message, tts_config):
        """ Convert text message to voice and play. """
        self.__tts_pool.submit(self.__tts, message, tts_config)

    def __create_temp_file(self):
        return tempfile.NamedTemporaryFile(
            prefix="edmc-chat-tts-",
            suffix=".mp3",
            delete=False
        )

    def __tts(self, message, tts_config):
        try:
            with self.__create_temp_file() as sound_file:
                _logger.info(f'Writing to {sound_file.name}')
                self.__tts_engine.create(message, sound_file, tts_config)
            self.__play_pool.submit(self.__play, sound_file)
        except Exception:
            _logger.exception('TTS failed')

    def __play(self, sound_file):
        try:
            _logger.info(f'Playing {sound_file.name}')
            self.__play_engine.play(sound_file.name)
            os.unlink(sound_file.name)
        except Exception:
            _logger.exception('playback failed')


class _PluginConfigs(Enum):
    """ Plugin configuration. """

    SPEAK_NPC_CHAT = 'chat_tts_speak_npc_chat', False
    SPEAK_SYSTEM_CHAT = 'chat_tts_speak_system_chat', True
    SPEAK_LOCAL_CHAT = 'chat_tts_speak_local_chat', True
    SPEAK_SQUADRON_CHAT = 'chat_tts_speak_squadron_chat', True
    SPEAK_DIRECT_CHAT = 'chat_tts_speak_direct_chat', True
    VOICE_NAME = 'chat_tts_voice_name', 'en-GB-SoniaNeural'
    VOICE_VOLUME_ADJUST = 'chat_tts_voice_volume_adjust', '-20%'
    VOICE_RATE_ADJUST = 'chat_tts_voice_rate_adjust', '+10%'

    def get_str(self):
        """ Return string setting value. """
        (cfg_name, cfg_default) = self.value
        return config.get_str(cfg_name, default=cfg_default)

    def get_int(self):
        """ Return integer setting value. """
        (cfg_name, cfg_default) = self.value
        return config.get_int(cfg_name, default=cfg_default)

    def get_bool(self):
        """ Return boolean setting value. """
        (cfg_name, cfg_default) = self.value
        return config.get_bool(cfg_name, default=cfg_default)

    def get_bool_as_int(self):
        """ Return boolean setting but as int (0/1). """
        (cfg_name, cfg_default) = self.value
        val = config.get_bool(cfg_name, default=cfg_default)
        return 1 if val else 0

    def set(self, config_value):
        """ Set new value for setting. """
        (cfg_name, _) = self.value
        config.set(cfg_name, config_value)

    def delete(self):
        """ Delete setting. """
        (cfg_name, _) = self.value
        config.delete(cfg_name)


class _AutoRow:
    """ Helper class for automatically incrementing a row counter. """

    def __init__(self):
        self.__cur_row = -1

    def next(self):
        """ Return next row. """
        self.__cur_row += 1
        return self.__cur_row

    def cur(self):
        """ Return current row. """
        return self.__cur_row


class _PluginPrefs:  # pylint: disable=too-many-instance-attributes
    """ Plugin preferences. """

    def __init__(self, message_player):
        self.__player = message_player
        self.__speak_npc_chat = tk.IntVar(value=_PluginConfigs.SPEAK_NPC_CHAT.get_bool_as_int())
        self.__speak_system_chat = tk.IntVar(value=_PluginConfigs.SPEAK_SYSTEM_CHAT.get_bool_as_int())
        self.__speak_local_chat = tk.IntVar(value=_PluginConfigs.SPEAK_LOCAL_CHAT.get_bool_as_int())
        self.__speak_squadron_chat = tk.IntVar(value=_PluginConfigs.SPEAK_SQUADRON_CHAT.get_bool_as_int())
        self.__speak_direct_chat = tk.IntVar(value=_PluginConfigs.SPEAK_DIRECT_CHAT.get_bool_as_int())
        self.__voice_name = tk.StringVar(value=_PluginConfigs.VOICE_NAME.get_str())
        self.__voice_volume_adjust = tk.StringVar(value=_PluginConfigs.VOICE_VOLUME_ADJUST.get_str())
        self.__voice_rate_adjust = tk.StringVar(value=_PluginConfigs.VOICE_RATE_ADJUST.get_str())
        self.__voices = self.__load_voices()

    def __load_voices(self):
        """ Fixed list with English voices. """
        return [
            'en-AU-NatashaNeural',
            'en-AU-WilliamNeural',
            'en-CA-ClaraNeural',
            'en-CA-LiamNeural',
            'en-HK-SamNeural',
            'en-HK-YanNeural',
            'en-IN-NeerjaExpressiveNeural',
            'en-IN-NeerjaNeural',
            'en-IN-PrabhatNeural',
            'en-IE-ConnorNeural',
            'en-IE-EmilyNeural',
            'en-KE-AsiliaNeural',
            'en-KE-ChilembaNeural',
            'en-NZ-MitchellNeural',
            'en-NZ-MollyNeural',
            'en-NG-AbeoNeural',
            'en-NG-EzinneNeural',
            'en-PH-JamesNeural',
            'en-PH-RosaNeural',
            'en-SG-LunaNeural',
            'en-SG-WayneNeural',
            'en-ZA-LeahNeural',
            'en-ZA-LukeNeural',
            'en-TZ-ElimuNeural',
            'en-TZ-ImaniNeural',
            'en-GB-LibbyNeural',
            'en-GB-MaisieNeural',
            'en-GB-RyanNeural',
            'en-GB-SoniaNeural',
            'en-GB-ThomasNeural',
            'en-US-AvaMultilingualNeural',
            'en-US-AndrewMultilingualNeural',
            'en-US-EmmaMultilingualNeural',
            'en-US-BrianMultilingualNeural',
            'en-US-AvaNeural',
            'en-US-AndrewNeural',
            'en-US-EmmaNeural',
            'en-US-BrianNeural',
            'en-US-AnaNeural',
            'en-US-AriaNeural',
            'en-US-ChristopherNeural',
            'en-US-EricNeural',
            'en-US-GuyNeural',
            'en-US-JennyNeural',
            'en-US-MichelleNeural',
            'en-US-RogerNeural',
            'en-US-SteffanNeural',
        ]

    def create_frame(self, parent: nb.Notebook):
        """ Create and return preferences frame. """
        padx = 10
        pady = 4
        boxy = 2
        auto_row = _AutoRow()

        frame = nb.Frame(parent)
        frame.columnconfigure(0, weight=0)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=0)

        nb.Label(frame, text='Speak:').grid(
            column=0, row=auto_row.next(), padx=padx, pady=pady, sticky=tk.W
        )

        fields = (
            ('NPC chat', self.__speak_npc_chat),
            ('System chat', self.__speak_system_chat),
            ('Local chat', self.__speak_local_chat),
            ('Squadron chat', self.__speak_squadron_chat),
            ('Direct messages', self.__speak_direct_chat)
        )
        for label, variable in fields:
            button = nb.Checkbutton(frame, text=label, variable=variable)
            button.grid(column=1, row=auto_row.next(), padx=padx, pady=pady, sticky=tk.W)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(
            columnspan=3, padx=padx, pady=pady, sticky=tk.EW, row=auto_row.next()
        )

        nb.Label(frame, text='Voice:').grid(
            column=0, row=auto_row.next(), padx=padx, pady=pady, sticky=tk.W
        )
        nb.OptionMenu(frame, self.__voice_name, self.__voice_name.get(), *self.__voices).grid(
            column=1, row=auto_row.cur(), padx=padx, pady=boxy, sticky=tk.EW
        )
        ttk.Button(frame, text='Test', command=self.__on_test_voice).grid(
            column=2, row=auto_row.cur(), padx=padx, sticky=tk.EW
        )

        nb.Label(frame, text='Voice volume (+/- in %):').grid(
            column=0, row=auto_row.next(), padx=padx, pady=pady, sticky=tk.W
        )
        nb.EntryMenu(frame, takefocus=False, textvariable=self.__voice_volume_adjust).grid(
            column=1, row=auto_row.cur(), padx=padx, pady=boxy, sticky=tk.EW
        )

        nb.Label(frame, text='Voice rate (+/- in %):').grid(
            column=0, row=auto_row.next(), padx=padx, pady=pady, sticky=tk.W
        )
        nb.EntryMenu(frame, takefocus=False, textvariable=self.__voice_rate_adjust).grid(
            column=1, row=auto_row.cur(), padx=padx, pady=boxy, sticky=tk.EW
        )

        nb.Label(frame, text=f'Version: {_VERSION}', fg='grey').grid(
            column=0, columnspan=3, row=auto_row.next(), padx=padx, pady=pady, sticky=tk.SE
        )

        # Configure rows
        for row in range(auto_row.cur()):
            frame.rowconfigure(row, weight=0)
        frame.rowconfigure(auto_row.cur(), weight=1)

        return frame

    def __on_test_voice(self):
        """ Test button is pressed. Play test message. """
        message = 'Message from test: How does this sound to you?'
        tts_config = _TTSConfig(
            self.__voice_name.get().strip(),
            self.__voice_volume_adjust.get().strip(),
            self.__voice_rate_adjust.get().strip()
        )
        self.__player.play_message(message, tts_config)

    def on_change(self):
        """ Preferences need to get applied. """
        _PluginConfigs.SPEAK_NPC_CHAT.set(self.__speak_npc_chat.get())
        _PluginConfigs.SPEAK_SYSTEM_CHAT.set(self.__speak_system_chat.get())
        _PluginConfigs.SPEAK_LOCAL_CHAT.set(self.__speak_local_chat.get())
        _PluginConfigs.SPEAK_SQUADRON_CHAT.set(self.__speak_squadron_chat.get())
        _PluginConfigs.SPEAK_DIRECT_CHAT.set(self.__speak_direct_chat.get())
        _PluginConfigs.VOICE_NAME.set(self.__voice_name.get().strip())
        _PluginConfigs.VOICE_VOLUME_ADJUST.set(self.__voice_volume_adjust.get().strip())
        _PluginConfigs.VOICE_RATE_ADJUST.set(self.__voice_rate_adjust.get().strip())


class _PluginApp:
    """ Plugin application. """

    def __init__(self, message_player):
        self.__player = message_player
        self.reload_settings()

    def reload_settings(self):
        """ Reload settings. """
        channel_config = {
            'npc': _PluginConfigs.SPEAK_NPC_CHAT.get_bool(),
            'system': _PluginConfigs.SPEAK_SYSTEM_CHAT.get_bool(),
            'local': _PluginConfigs.SPEAK_LOCAL_CHAT.get_bool(),
            'squadron': _PluginConfigs.SPEAK_SQUADRON_CHAT.get_bool(),
            'player': _PluginConfigs.SPEAK_DIRECT_CHAT.get_bool()
        }
        self.__allowed_channels = {ch for ch, allowed in channel_config.items() if allowed}
        self.__tts_config = _TTSConfig(
            _PluginConfigs.VOICE_NAME.get_str(),
            _PluginConfigs.VOICE_VOLUME_ADJUST.get_str(),
            _PluginConfigs.VOICE_RATE_ADJUST.get_str())

    def on_receive_text(self, sender, channel, message):
        """ Called when a text message is received. """
        # Ignore NPC default message
        if channel == 'npc' and message.startswith('$') and message.endswith(';'):
            return
        # Ignore channels according to settings
        if channel not in self.__allowed_channels:
            return
        # Speak message
        # We could remember the last commander and not repeat
        # from whom it is if it is near the time of his last message.
        text = f'Message from {sender}: {message}'
        self.__player.play_message(text, self.__tts_config)


_message_player = _MessagePlayer()
_plugin_prefs = _PluginPrefs(_message_player)
_plugin_app = _PluginApp(_message_player)


def plugin_start3(_plugin_dir: str) -> str:
    """ Called by EDMC to start plugin. """
    return _plugin_name


def plugin_prefs(parent: nb.Notebook, _cmdr: str, _is_beta: bool) -> Optional[tk.Frame]:
    """ Called by EDMC when showing preferences. """
    return _plugin_prefs.create_frame(parent)


def prefs_changed(_cmdr: str, _is_beta: bool) -> None:
    """ Called by EDMC when preferences are applied. """
    _plugin_prefs.on_change()
    _plugin_app.reload_settings()


def journal_entry(
    _cmdr: str,
    _is_beta: bool,
    _system: str,
    _station: str,
    entry: Dict[str, Any],
    _state: Dict[str, Any]
) -> Optional[str]:
    """ Called by EDMC for every new journal entry. """
    if entry['event'] == 'ReceiveText':
        # Log entry for testing. This can be removed later
        _logger.info(f'entry: {entry}')
        _plugin_app.on_receive_text(
            entry['From'],
            entry['Channel'],
            entry['Message']
        )
