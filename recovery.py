import os
import psutil
import wmi
import winreg
import win32api
import win32con
import win32com.client
import win32evtlog
import argparse
import sqlite3
import time
import datetime
import browser_history
from datetime import date
from colors import *

timestamp = time.time()
time_dict = {
    "day": 86_400,
    "week": 604_800.02,
    "month": 2_629_800,
    "year": 31_556_952
}

# CurrentVersionRun es una clave del Registro de Windows que se utiliza
# para iniciar programas automáticamente cuando se inicia el sistema operativo.
def get_currentversionrun():
    print(color("Current Version Run Changes:", fg="blue"))
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
        num_values = winreg.QueryInfoKey(key)[1]
        for i in range(num_values):
            value_name, value_data, value_type = winreg.EnumValue(key, i)
            last_modified = winreg.QueryInfoKey(key)[2] # As 100’s of nanoseconds since Jan 1, 1601
            if not last_modified == None:
                if timestamp < (last_modified / 1e7 - 11644473600):
                    print(f"\t{value_name} - {date.fromtimestamp(last_modified / 1e7 - 11644473600)}")

def get_running_programs():
    print(color("Running programs:", fg="blue"))
    for proc_pid in psutil.pids():
        proc = psutil.Process(proc_pid)
        timestamp_date = proc.create_time()
        if (timestamp_date > timestamp):
            print(f"\t{proc.name()} - {date.fromtimestamp(timestamp_date)}")

def get_recent_files():
    recent_directory = win32com.client.Dispatch("WScript.Shell").SpecialFolders("Recent")
    print(color(f"Recent files (based on - {recent_directory}):", fg="yellow"))
    access_files = os.listdir(recent_directory)
    for file in access_files:
        #print(os.path.getatime(str(recent_directory) + '/' +str(file)))
        timestamp_date = os.path.getmtime(str(recent_directory) + '/' +str(file))
        if (file != "" and timestamp_date > timestamp):
            if (file.endswith(".lnk")):
                file = file[:-4:]
            print(f"\t{file} - {date.fromtimestamp(timestamp_date)}")
        

def get_eventlog_of(logtype:str):
    good = False
    print(color(f"\t{logtype}:", fg="red"))
    try:
        hand = win32evtlog.OpenEventLog(None,logtype)
    except Exception as err:
        return False
    try:
        total=win32evtlog.GetNumberOfEventLogRecords(hand)
        flags = win32evtlog.EVENTLOG_BACKWARDS_READ|win32evtlog.EVENTLOG_SEQUENTIAL_READ
        events = win32evtlog.ReadEventLog(hand,flags,0)
        for ev_obj in events:
            the_time = ev_obj.TimeGenerated #'12/23/99 15:54:09'
            timestamp_evt = date.fromtimestamp(datetime.datetime.strptime(str(the_time), '%Y-%m-%d %H:%M:%S').timestamp())
            print(f"\t\t{ev_obj.SourceName} - {timestamp_evt}")
        good = True

    except Exception as err:
        good = False
    finally:
        return good

def print_eventlog_of(logtype: str):
    if not get_eventlog_of(logtype):
        print("\t\tRestricted access, open with administrator mode")

def get_eventlog():
    print(color("Try new things:", fg="red"))
    print_eventlog_of('Security')
    print_eventlog_of('System')
    print_eventlog_of('Error')
    
def get_installed_programs():
    print(color("Get installed Programs:", fg="green"))
    wmi_obj = win32com.client.GetObject("winmgmts:\\\\.\\root\\cimv2")

    # Consulta la clase Win32_Product para obtener información sobre los programas instalados
    programs = wmi_obj.ExecQuery("SELECT * FROM Win32_Product")

    # Itera sobre los programas y muestra información relevante
    for program in programs:
        install_date = program.InstallDate
        if install_date is not None and isinstance(install_date, str):
            install_date = datetime.datetime.strptime(install_date, "%Y%m%d")
            if install_date.date() > date.fromtimestamp(timestamp):
                print(f"\tProgram: {program.Name} - {install_date.date()}")


def get_history():
    ###
    # See suported browsers on https://browser-history.readthedocs.io/en/latest/browsers.html
    ###
    print(color("Get Browser History:", fg="magenta"))
    histories = browser_history.get_history().histories
    for history_entry_date, history_entry_url in histories:
        if history_entry_date.timestamp() > timestamp:
            print(f"\tURL - {history_entry_url} on {history_entry_date.date()}")

def get_connected_devices():
    print(color("Get Connected Devices (no date):", fg="green"))
    wmi = win32com.client.GetObject('winmgmts:')
    devices = wmi.InstancesOf('Win32_PnPEntity')
    for device in devices:
        if device.InstallDate != None and isinstance(device.InstallDate, datetime.datetime):
            print(f"\t{device.Name} - {device.PNPDeviceID} - {device.InstallDate.date()}")
        else:
            print(f"\t{device.Name} - {device.PNPDeviceID} - No info date")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Forensics data, may be interesting for you")
    parser.add_argument("time", choices=["day", "week", "month", "year"], help="Sets how old the data can be")
    args = parser.parse_args()
    timestamp -= time_dict[args.time]

    get_currentversionrun()
    get_recent_files()
    get_running_programs()
    get_installed_programs()
    get_history()
    get_connected_devices()
    get_eventlog()