import sys
import os
import subprocess
import time
import json
import logging
# config settings validation
from pathlib import Path
# watch for processes
import wmi
import pythoncom
# ui & threads
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QThread, QObject, pyqtSignal, QTimer, QEventLoop, QSettings, QCoreApplication
from PyQt5.QtGui import QIcon
# show notifications
from plyer import notification

# ------------------------------------------------------------------------------------------

# path to the app config json file
APPCONFIGJSON = 'Config.json'
TRAY_TOOLTIP = 'EnforceAudioDevice'
TRAY_ICON = 'EnforceAudioDevice.ico'
ALERT_ICON = 'EnforceAudioDeviceAlert.ico'
APP_NAME = 'EnforceAudioDevice'
LOG_FILE = 'EnforceAudioDevice.log'
RUN_PATH = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"
# need to define base path due to app launching via windows autostart
BASE_PATH = os.path.dirname(sys.argv[0])

# ------------------------------------------------------------------------------------------

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = BASE_PATH

    return os.path.join(base_path, relative_path)

# ------------------------------------------------------------------------------------------

# setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(BASE_PATH, LOG_FILE), "w"),
        logging.StreamHandler(sys.stdout)
    ]
)

############################################################################################
# ProcesWatcher
############################################################################################

class ProcessWatcher(QThread):
    """watches for wmi process creation or deletion events and signals when an event arrives"""
    watcher_signal = pyqtSignal(str, int)

    # ------------------------------------------------------------------------------------------

    def __init__(self, Type: str):
        QThread.__init__(self)
        self.Type = Type

    # ------------------------------------------------------------------------------------------

    def run(self):
        self.continue_run = True

        pythoncom.CoInitialize()
        c = wmi.WMI()
        if self.Type == "creation" or self.Type == "deletion":
            try:
                watcher = c.Win32_Process.watch_for(self.Type)
                while self.continue_run:
                    event = watcher()
                    self.watcher_signal.emit(event.Caption, event.ProcessID)
            except Exception as e:
                print(e)
        else:
            logging.error(
                f"Tried to create process listener with invalid type '{self.Type}'. Valid types are: creation, deletion")

    # ------------------------------------------------------------------------------------------

    def stop(self):
        self.continue_run = False
        # disconnect all slots from the signal as it might take a little bit until this thread actually terminates
        try:
            self.watcher_signal.disconnect()
        except Exception as e:
            logging.error(e)

############################################################################################
# ProcesWorker
############################################################################################


class ProcessWorker(QThread):
    # dictionary of processes to check, will be filled from json on init
    process_dict = {}
    # the parent app containing config data
    app = None
    # timers that delay the set audio device command (user defined in config)
    delayedCommandTimers = []

    # ------------------------------------------------------------------------------------------

    def __init__(self, parent=None, app=None):
        QObject.__init__(self, parent=parent)
        self.app = app

    # ------------------------------------------------------------------------------------------

    def run(self):
        self.loop = QEventLoop()
        self.loop.exec_()

    # ------------------------------------------------------------------------------------------

    def stop(self):
        # stop all running timers if the application should quit
        for t in self.delayedCommandTimers:
            t.stop()
        # stop the loop to quit this thread
        self.loop.quit()

    # ------------------------------------------------------------------------------------------

    def add_app(self, application, data):
        """add to or update an app in the process list"""
        if not bool(data):
            logging.warning(
                f'Application \'{application}\' is missing parameters. Apps require a \'Device\' parameter defining the audio output device.')
            return

        app_name = application.lower()
        if not app_name.endswith('.exe'):
            app_name += '.exe'

        device = ''
        if 'Device' in data:
            device = data['Device']
        # check if the device is valid
        if not device in self.app.valid_devices:
            logging.warning(
                f'Application \'{application}\' has no or invalid \'Device\' configured \'{device}\'. Allowed devices are: {self.app.valid_devices}')
            return

        delay = 1
        if 'Delay' in data:
            try:
                delay = max(min(float(data['Delay']), 60.0), 0.0)
            except ValueError:
                logging.warning(f'Delay of \'{application}\' is not a number!')

        already_contains_app = app_name in self.process_dict

        if already_contains_app:
            if self.process_dict[app_name]['AudioDevice'] == device:
                # nothing changed, just return
                return

        # either the app hasn't been added yet or the device changed
        self.process_dict[app_name] = {'State': False,
                                       'AudioDevice': device, 'Delay': delay}

        logging.info(
            ('Updated' if already_contains_app else 'Added') + ' app: ' + application)
        # check if the process is already running and handle it
        self.check_processes(app_name)
        return

    # -------------------------------------------------------------------------------------------

    def check_processes(self, process_name):
        pythoncom.CoInitialize()
        c = wmi.WMI()

        for process in c.Win32_Process(name=process_name):
            self.process_started(process_name, process.ProcessID)

    # ------------------------------------------------------------------------------------------

    def process_started(self, name: str, id: int):
        process_name = name.lower()
        if process_name in self.process_dict:
            # already running, ignore this process
            if self.process_dict[process_name]['State']:
                return
            # add the process with its process id
            else:
                logging.info(f"Found new process running: '{process_name}'")
                self.process_dict[process_name]['State'] = True
                self.process_dict[process_name]['ID'] = id
                delay = self.process_dict[process_name]['Delay']
                self.set_audio_device(process_name, delay)

    # ------------------------------------------------------------------------------------------

    def process_ended(self, name: str, id: int):
        process_name = name.lower()
        if process_name in self.process_dict:
            # this process is running and has the same id
            if self.process_dict[process_name]['State'] and self.process_dict[process_name]['ID'] == id:
                self.process_dict[process_name]['State'] = False
                self.process_dict[process_name]['ID'] = 0
                logging.info(f"Process '{process_name}' has ended")

    # ------------------------------------------------------------------------------------------

    def set_audio_device(self, application: str, delay: float):
        """sets the audio device for the application after the defined delay"""
        if application in self.process_dict:
            audio_device = self.process_dict[application]['AudioDevice']
            command = f'{self.app.sound_volume_view_path} /SetAppDefault "{audio_device}" 0 "{application}"'
            # queue the command via timer
            self.set_command_timer(lambda: self.run_command(
                command, application, audio_device), int(delay * 1000))

    # ------------------------------------------------------------------------------------------

    def run_command(self, command, application_name, audio_device):
        #res = os.system(command)
        res = subprocess.call(command, shell=False)
        if res == 0:
            logging.info(
                f'Set audio device of application \'{application_name}\' to \'{audio_device}\'')
        else:
            logging.warning(
                f'SoundVolumeView failed to set audio device \'{audio_device}\' for application \'{application_name}\'. Error code: {res}')

    # ------------------------------------------------------------------------------------------

    def set_command_timer(self, event, delay_msec: int):
        timer = QTimer()
        self.delayedCommandTimers.append(timer)
        timer.setSingleShot(True)
        timer.timeout.connect(event)
        timer.timeout.connect(lambda: self.delayedCommandTimers.remove(timer))
        timer.start(delay_msec)

############################################################################################
# EnforceAudioDeviceApp
############################################################################################


class EnforceAudioDeviceApp(QApplication):
    stop_signal = pyqtSignal()

    # a set of valid audio output devices
    valid_devices = set()
    # path to the sound volume view tool to actually run the audio device command
    sound_volume_view_path = 'SoundVolumeView.exe'
    # the thread the worker is running in
    thread: ProcessWorker = None

    # ------------------------------------------------------------------------------------------

    def __init__(self, argv) -> None:
        super().__init__(argv)
        self.create_settings()
        self.load_config_and_start_worker()
        self.trayIcon = EnforceAudioDeviceTrayIcon(self)

    # ------------------------------------------------------------------------------------------

    def create_settings(self):
        QCoreApplication.setApplicationName(APP_NAME)
        self.settings = QSettings(RUN_PATH, QSettings.NativeFormat)

    # ------------------------------------------------------------------------------------------

    def load_config_and_start_worker(self):
        self.create_worker_threads()
        if self.load_config_json():
            self.start_worker_thread()
            logging.info(
                'Successfully loaded config and started process monitoring worker')
            return True
        else:
            logging.warning(
                'Failed to load config and start process monitoring')
            return False

    # ------------------------------------------------------------------------------------------

    def create_worker_threads(self):
        self.thread = ProcessWorker(app=self)
        self.stop_signal.connect(self.thread.stop)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

        self.create_listener = ProcessWatcher("creation")
        self.stop_signal.connect(self.create_listener.stop)
        self.create_listener.finished.connect(self.create_listener.deleteLater)
        self.create_listener.watcher_signal.connect(
            self.thread.process_started)
        self.create_listener.start()

        self.delete_listener = ProcessWatcher("deletion")
        self.stop_signal.connect(self.delete_listener.stop)
        self.delete_listener.finished.connect(self.delete_listener.deleteLater)
        self.delete_listener.watcher_signal.connect(self.thread.process_ended)
        self.delete_listener.start()

    # ------------------------------------------------------------------------------------------

    def start_worker_thread(self):
        if not self.thread is None:
            self.thread.start()

    # ------------------------------------------------------------------------------------------

    def start_reload_config(self):
        logging.info('Reloading config file...')
        self.stop_signal.emit()
        self.thread.destroyed.connect(self.finish_reload_config)

    # ------------------------------------------------------------------------------------------

    def finish_reload_config(self):
        self.load_config_and_start_worker()

    # ------------------------------------------------------------------------------------------

    def start_quit(self):
        logging.info('Exiting...')
        # end running watcher threads
        self.stop_signal.emit()
        self.thread.destroyed.connect(self.finish_quit)

    # ------------------------------------------------------------------------------------------

    def finish_quit(self):
        logging.info('Exit')
        self.quit()

    # ------------------------------------------------------------------------------------------

    def load_config_json(self):
        """loads the apps from the apps json file"""
        # check if the config exists, if not, create one filled with example data
        config_path = os.path.join(BASE_PATH, APPCONFIGJSON)
        if os.path.exists(config_path):
            config = {}
            # load the config file data
            with open(config_path, "r", encoding='UTF-8') as file:
                try:
                    config = json.load(file)
                except json.JSONDecodeError:
                    logging.error(
                        f'Failed to load \'{APPCONFIGJSON}\', aborting!')
                finally:
                    file.close()
            # load audio valid audio devices, the SoundVolumeView path and apps. Exit if any of these fail.
            if not bool(config) or not self.load_config_data(config) or not self.load_valid_audio_devices() or not self.get_apps_from_config(config):
                return False
        else:
            default_config = {'Config': {'SoundVolumeViewPath': "SoundVolumeView.exe"}, 'Apps': {
                'MyExampleApp1.exe': "MyExampleAudioDevice", 'MyExampleApp2.exe': "MyExampleAudioDevice", }}
            data = json.dumps(default_config, indent=2)
            # write a default config file if the config doesn't exist
            with open(APPCONFIGJSON, "w", encoding='UTF-8') as outfile:
                outfile.write(data)
            outfile.close()
            logging.info(
                f'Created: \'{config_path}\'. Please add your apps to the file and reload the config.')
            self.send_notify("Enforce Audio Device Info",
                            f'Created: \'{APPCONFIGJSON}\'.\nPlease add your apps to the file and reload the config.', resource_path(ALERT_ICON))
        return True

    # ------------------------------------------------------------------------------------------

    def load_config_data(self, config):
        """loads general config data from the config file"""

        has_config = 'Config' in config
        if not has_config:
            logging.warning(
                f'Couldn\'t find \'Config\' section in the config file.')

        # checks if the sound volume view tool path is valid and points to a file
        path_valid = False
        if has_config and 'SoundVolumeViewPath' in config['Config']:
            self.sound_volume_view_path = config['Config']['SoundVolumeViewPath']
        if Path(self.sound_volume_view_path).is_file():
            path_valid = True

        if not path_valid:
            logging.error(
                f'Invalid Sound Volume View path \'{self.sound_volume_view_path}\'. Make sure the path is set correctly in the Config.json.')
            self.send_notify("Enforce Audio Device Error",
                            f'Invalid Sound Volume View path \'{self.sound_volume_view_path}\'.\nMake sure the path is set correctly in the Config.json.', resource_path(ALERT_ICON))
            return False

        return True

    # ------------------------------------------------------------------------------------------

    def load_valid_audio_devices(self):
        """fills a dictionary of valid audio devices that can be used"""
        script_path = os.path.abspath(os.path.dirname(__file__))
        devices_json_path = Path(script_path + '\\ValidDevices.json')

        # if this file already exists, remove it
        if devices_json_path.is_file():
            os.remove(devices_json_path)

        # wait until the file is removed
        while devices_json_path.is_file():
            time.sleep(1)

        # call the soundVolumeView tool and export all audio devices to a json file
        command = f'{self.sound_volume_view_path} /sjson {devices_json_path}'
        #res = os.system(command)
        res = subprocess.call(command, shell=False)

        if res == 0:
            # try to open the files 10 times, fail if not possible
            for _ in range(10):
                try:
                    with open(devices_json_path, 'r', encoding='UTF-16') as file:
                        data = file.read()
                        device_dump = json.loads(data)
                        self.load_audio_devices_from_device_json(device_dump)
                        file.close()
                        os.remove(devices_json_path)
                        break
                except IOError:
                    time.sleep(1)
            else:
                logging.error(
                    f'Failed to access default devices in {devices_json_path}, SoundVolumeView failed to create the file.')
                return False
        else:
            logging.error(
                f'Finding valid audio devices failed using {command}. Error code = {res}')
            return False
        return True

    # ------------------------------------------------------------------------------------------

    def load_audio_devices_from_device_json(self, device_dict):
        """reads the device dictionary and picks valid output devices"""
        for device in device_dict:
            if device['Direction'] == 'Render' and device['Type'] == 'Device':
                self.valid_devices.add(device['Name'])
        return

    # ------------------------------------------------------------------------------------------

    def get_apps_from_config(self, config):
        """gets all configured apps from the config"""
        if 'Apps' in config:
            apps = config['Apps']
            if bool(apps):
                for app in apps:
                    self.thread.add_app(app, apps[app])
                return True

        logging.warning(
            f'No Apps defined in \'{os.path.abspath(APPCONFIGJSON)}\'. Please add apps to the config file and reload the config via the system tray.')
        self.send_notify("Enforce Audio Device Error",
                            f'No Apps defined in \'{APPCONFIGJSON}\'.\nPlease add apps to the config file and reload the config via the system tray.', resource_path(ALERT_ICON))
        return False

    # ------------------------------------------------------------------------------------------

    def send_notify(self, title : str, message : str, icon, duration : int = 10):
        notification.notify(
            title=title,
            message=message,
            app_icon=icon,
            app_name=APP_NAME,
            timeout=duration
        )

############################################################################################
# EnforceAudioDeviceTrayIcon
############################################################################################


class EnforceAudioDeviceTrayIcon(QSystemTrayIcon):

    def __init__(self, app: EnforceAudioDeviceApp):
        super(EnforceAudioDeviceTrayIcon, self).__init__(app)
        self.app = app
        self.create_tray_menu()
        self.autostart.setChecked(self.app.settings.contains(APP_NAME))

    def create_tray_menu(self):
        icon = QIcon(resource_path(TRAY_ICON))
        self.setToolTip(TRAY_TOOLTIP)
        self.setIcon(icon)
        self.setVisible(True)

        # Creating the options
        self.menu = QMenu()

        # Create reload option
        self.reload = self.menu.addAction(
            "Reload Config", self.app.start_reload_config)
        # Create autostart option
        self.autostart = self.menu.addAction(
            "Launch on Boot", self.toggle_autostart_state)
        self.autostart.setCheckable(True)
        self.menu.addSeparator()
        # Create quit option
        self.quit = self.menu.addAction("Quit", self.app.start_quit)
        # Adding options to the System Tray
        self.setContextMenu(self.menu)

    def toggle_autostart_state(self):
        current_state = self.app.settings.contains(APP_NAME)
        new_state = not current_state
        if current_state:
            self.app.settings.remove(APP_NAME)
        else:
            self.app.settings.setValue(APP_NAME, sys.argv[0])
        self.autostart.setChecked(new_state)
        logging.info(
            f'Added {APP_NAME} to autostart' if new_state else f'Removed {APP_NAME} from autostart')

# ------------------------------------------------------------------------------------------

def check_already_running():
    c = wmi.WMI()
    process_name = os.path.basename(sys.argv[0])
    process_count = 0
    for process in c.Win32_Process(name=process_name):
        process_count = process_count + 1
        # two processes are from us, if there are more than 2, another instance is already running
        if process_count > 2:
            return True
    return False

# ------------------------------------------------------------------------------------------

if __name__ == '__main__':
    if not check_already_running():
        app = EnforceAudioDeviceApp(sys.argv)
        sys.exit(app.exec_())