import wmi
import win32com
import os
import subprocess

def main():

    # Local PC
    client = wmi.WMI()

    for info in client.Win32_ComputerSystem():
        print("SystemFamily: {}".format(info.SystemFamily))
        # => SystemFamily: ThinkPad X201s

    for info in client.Win32_OperatingSystem():
        print("Caption: {}".format(info.Caption))
        # => Caption: Microsoft Windows 10 Home
        
    for info in client.Win32_NetworkAdapterConfiguration():
        print("Description: {}".format(info.Description))
        # => Description: Intel(R) 82577LM Gigabit Network Connection
        # => Description: Intel(R) Centrino(R) Advanced-N 6250 AGN

    # Remote PC
    remote = wmi.WMI("remote_host")
    for info in remote.Win32_OperatingSystem():
        print("Caption: {}".format(info.Caption))
        
        
    # MS Office Product key using `ospp.vbs`
    excel = win32com.client.Dispatch("Excel.Application")
    office_dir = excel.Path
    excel.Quit()
    
    ospp_path = os.path.join(office_dir, "ospp.vbs")
    
    # local
    subprocess.run(["cscript", ospp_path, "/dstatus"])
    # remote
    subprocess.run(["cscript", ospp_path, "/dstatus", "remote_host"])

    

if __name__ == "__main__":
    main()