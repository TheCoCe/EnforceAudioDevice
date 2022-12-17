import sys
import os
import subprocess
import time
import json
import logging
# watch for processes
import wmi
import pythoncom
# ui & threads
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QThread, QObject, pyqtSignal, QTimer, QEventLoop, QSettings, QCoreApplication, Qt
from PyQt5.QtGui import QIcon
# show notifications
from plyer import notification

# ------------------------------------------------------------------------------------------


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    # PyInstaller creates a temp folder and stores the path in _MEIPASS, so if we are in a packaged build we
    # use the _MEIPASS as the relative base path, otherwise we'll use the working directory
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))

    return os.path.join(base_path, relative_path)

# ------------------------------------------------------------------------------------------


def app_path(relative_path):
    """Get the absolute path to the file relative to the executable"""
    if getattr(sys, 'frozen', False):
        # if we are in a bundled app, we need to use the path of the executable
        # as this is where files like the config or log should go
        base_path = os.path.dirname(sys.executable)
    else:
        # if we are just executing the script as dev we need to use the path of the file
        # as this is where files like the config or log should go
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

# ------------------------------------------------------------------------------------------


# name of the application
APP_NAME = 'EnforceAudioDevice'
# files
CONFIG_FILE_PATH = app_path('EnforceAudioDevice.json')
LOG_FILE_PATH = app_path('EnforceAudioDevice.log')
VALID_DEVICES_FILE_PATH = app_path('ValidDevices.json')
# resources
TRAY_ICON_FILE_PATH = resource_path('EnforceAudioDevice.ico')
ALERT_ICON_FILE_PATH = resource_path('EnforceAudioDeviceAlert.ico')
# strings
TRAY_TOOLTIP = 'EnforceAudioDevice'
# registry key
REG_RUN_PATH = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"

# ------------------------------------------------------------------------------------------

# setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE_PATH, "w"),
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
    # timers that delay the set audio device command (delay is user defined in config)
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
        self.check_process(app_name)
        return

    # -------------------------------------------------------------------------------------------

    def check_process(self, process_name):
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

    # ------------------------------------------------------------------------------------------

    def stop_all_command_timers(self):
        for t in self.delayedCommandTimers:
            t.stop()
        self.delayedCommandTimers.clear()

    # ------------------------------------------------------------------------------------------

    def reset_process_states(self):
        # cancel all pending timers
        self.stop_all_command_timers()

        pythoncom.CoInitialize()
        c = wmi.WMI()

        # reset current state of all processes and set the device for any active ones again
        for p in self.process_dict:
            self.process_dict[p]['State'] = False
            for process in c.Win32_Process(name=p):
                self.process_started(p, process.ProcessID)
                break;

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
        self.settings = QSettings(REG_RUN_PATH, QSettings.NativeFormat)

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

    def reset_processes(self):
        self.thread.reset_process_states()

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
        if os.path.exists(CONFIG_FILE_PATH):
            config = {}
            # load the config file data
            with open(CONFIG_FILE_PATH, "r", encoding='UTF-8') as file:
                try:
                    config = json.load(file)
                except json.JSONDecodeError:
                    logging.error(
                        f'Failed to load \'{CONFIG_FILE_PATH}\', aborting!')
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
            with open(CONFIG_FILE_PATH, "w", encoding='UTF-8') as outfile:
                outfile.write(data)
            outfile.close()
            logging.info(
                f'Created: \'{CONFIG_FILE_PATH}\'. Please add your apps to the file and reload the config.')
            self.send_notify("Enforce Audio Device Info",
                             f'Created: \'{os.path.basename(CONFIG_FILE_PATH)}\'.\nPlease add your apps to the file and reload the config.', ALERT_ICON_FILE_PATH)
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
        if os.path.isfile(self.sound_volume_view_path):
            path_valid = True

        if not path_valid:
            logging.error(
                f'Invalid Sound Volume View path \'{self.sound_volume_view_path}\'. Make sure the path is set correctly in the Config.json.')
            self.send_notify("Enforce Audio Device Error",
                             f'Invalid Sound Volume View path \'{self.sound_volume_view_path}\'.\nMake sure the path is set correctly in the Config.json.', ALERT_ICON_FILE_PATH)
            return False

        return True

    # ------------------------------------------------------------------------------------------

    def load_valid_audio_devices(self):
        """fills a dictionary of valid audio devices that can be used"""
        # if this file already exists, remove it
        if os.path.isfile(VALID_DEVICES_FILE_PATH):
            os.remove(VALID_DEVICES_FILE_PATH)

        # call the soundVolumeView tool and export all audio devices to a json file
        command = f'{self.sound_volume_view_path} /sjson {VALID_DEVICES_FILE_PATH}'
        #res = os.system(command)
        res = subprocess.call(command, shell=False)

        if res == 0:
            # try to open the files 10 times, fail if not possible
            for _ in range(10):
                try:
                    with open(VALID_DEVICES_FILE_PATH, 'r', encoding='UTF-16') as file:
                        data = file.read()
                        device_dump = json.loads(data)
                        self.load_audio_devices_from_device_json(device_dump)
                        file.close()
                        os.remove(VALID_DEVICES_FILE_PATH)
                        break
                except IOError:
                    time.sleep(1)
            else:
                logging.error(
                    f'Failed to access default devices in {VALID_DEVICES_FILE_PATH}, SoundVolumeView failed to create the file.')
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
            f'No Apps defined in \'{CONFIG_FILE_PATH}\'. Please add apps to the config file and reload the config via the system tray.')
        self.send_notify("Enforce Audio Device Error",
                         f'No Apps defined in \'{os.path.basename(CONFIG_FILE_PATH)}\'.\nPlease add apps to the config file and reload the config via the system tray.', ALERT_ICON_FILE_PATH)
        return False

    # ------------------------------------------------------------------------------------------

    def send_notify(self, title: str, message: str, icon, duration: int = 10):
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
        self.act_autostart.setChecked(self.app.settings.contains(APP_NAME))

    # ------------------------------------------------------------------------------------------

    def create_tray_menu(self):
        icon = QIcon(TRAY_ICON_FILE_PATH)
        self.setToolTip(TRAY_TOOLTIP)
        self.setIcon(icon)
        self.setVisible(True)

        # Creating the options
        self.menu = QMenu("Options")
        self.menu.setWindowFlags(self.menu.windowFlags() | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        self.menu.setAttribute(Qt.WA_TranslucentBackground)
        self.menu.setStyleSheet("""
            QMenu{
                  background-color: #ffffff;
                  border-image: url("P:/Stuff/_Projects/Programming/Python/EnforceAudioDevice/ContextMenu.png") 1 stretch;
                  border-radius: 10px;
            }
            QMenu::item {
                    background-color: transparent;
                    padding: 5px 5px;
                    margin: 10px 10px;
            }
            QMenu::item:selected 
            { 
                background-color: #fc8c29;
                border-radius: 5px
            }
            QMenu::item:disabled {
                background-color: transparent;
                color: #ffffff;
                font-weight: bold;
            }
        """)

        self.act_device = self.menu.addAction("Enforce Audio Device")
        self.act_device.setEnabled(False)

        # Creating config sub menu
        self.config_menu = QMenu("Config")
        self.config_menu.setWindowFlags(self.menu.windowFlags() | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        self.config_menu.setAttribute(Qt.WA_TranslucentBackground)
        self.config_menu.setStyleSheet("""
            QMenu{
                  background-color: #ffffff;
                  border: 5px solid #4c241d;
                  border-radius: 10px;
            }
            QMenu::item {
                    background-color: transparent;
                    padding: 5px 5px;
                    margin: 5px 5px;
            }
            QMenu::item:selected 
            { 
                background-color: #fc8c29;
                border-radius: 5px
            }
            QMenu::item:disabled {
                background-color: transparent;
                color: #ffffff;
                font-weight: bold;
            }
        """)

        # Create reload config option
        self.act_reload = self.config_menu.addAction(
            "Reload Config", self.app.start_reload_config)

        self.config_menu.addSeparator()

        # Create open config file option
        self.act_open_config = self.config_menu.addAction(
            "Open Config", self.open_config_file)

        # Create open config folder option
        self.act_open_config_dir = self.config_menu.addAction(
            "Go to Config", self.open_config_folder)

        self.config_menu.addSeparator()

        # Create open log option
        self.act_open_log = self.config_menu.addAction(
            "Open Log", self.open_log_file)

        # add the config menu to the menu
        self.act_open_config_menu = self.menu.addMenu(self.config_menu)

        # Create reset audio devices button
        self.act_reset = self.menu.addAction(
            "Reset audio devices", self.app.reset_processes)

        # Create autostart option
        self.act_autostart = self.menu.addAction(
            "Launch on Boot", self.toggle_autostart_state)
        self.act_autostart.setCheckable(True)

        self.menu.addSeparator()

        # Create quit option
        self.act_quit = self.menu.addAction("Quit", self.app.start_quit)

        # Adding options to the System Tray
        self.setContextMenu(self.menu)

    # ------------------------------------------------------------------------------------------

    def toggle_autostart_state(self):
        current_state = self.app.settings.contains(APP_NAME)
        new_state = not current_state
        if current_state:
            self.app.settings.remove(APP_NAME)
        else:
            self.app.settings.setValue(APP_NAME, sys.argv[0])
        self.act_autostart.setChecked(new_state)
        logging.info(
            f'Added {APP_NAME} to autostart' if new_state else f'Removed {APP_NAME} from autostart')

    # ------------------------------------------------------------------------------------------

    def open_config_folder(self):
        path = os.path.dirname(CONFIG_FILE_PATH)
        os.startfile(path)

    # ------------------------------------------------------------------------------------------

    def open_config_file(self):
        os.startfile(CONFIG_FILE_PATH)

    # ------------------------------------------------------------------------------------------

    def open_log_file(self):
        os.startfile(LOG_FILE_PATH)


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
