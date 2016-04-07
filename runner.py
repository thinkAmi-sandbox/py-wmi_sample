import wmi
import win32com     # MS Officeのインストールパスを取得するのに使用
import os           # パスまわりで使用
import subprocess   # ospp.vbsを使うのに使用

class PC:
    def __init__(self, hostname=None):
        self.client = wim.WMI(hostname) if hostname else wmi.WMI()


    def show(self, wmi_object, propperties):
        for property in propperties:
            try:
                value = getattr(wmi_object, property)
                self.print_with_format(property, value)
                
            except UnicodeEncodeError as e:
                # Win32_ProductではUnicodeEncodeErrorが発生する可能性があるため、
                # その場合にはutf-8でエンコードする
                # UnicodeEncodeError: 'cp932' codec can't encode character '\xe9' in position 12: illegal multibyte sequence
                self.print_with_format(property, value.encode("utf-8"))
            except:
                print("{property}: Disabled by OS version".format(property=property))
                
    def show_with_conversion(self, wmi_object, property, conversion_dict):
        try:
            key = getattr(wmi_object, property)
            value = conversion_dict[key]
            self.print_with_format(property, value)
        except:
            print("{property}: disabled by OS version".format(property=property))
        
    def print_with_format(self, name, value):
        print("{name}: {value}".format(name=name, value=value))
        
    def print_title(self, title):
        print("-" * len(title))
        print(title)
        print("-" * len(title))
        

    def base_board(self):
        PROPERTIES = [
            "Caption",
            "ConfigOptions",
            "Description",
            "Manufacturer",
            "Model",
            "Name",
            "OtherIdentifyingInfo",
            "PartNumber",
            "Product",
            "RequirementsDescription",
            "SerialNumber",
            "SKU",
            "Tag",
            "Version",
        ]

        self.print_title("Win32_BaseBoard")
        for info in self.client.Win32_BaseBoard():
            self.show(info, PROPERTIES)
            print("---")
        
        
    def cd_rom_drive(self):
        PROPERTIES = [
            "Caption",
            "CompressionMethod",
            "Description",
            "Drive",
            "Manufacturer",
            "MediaType",
            "MfrAssignedRevisionLevel",
            "Name",
            "PNPDeviceID",
            "RevisionLevel",
            "SerialNumber",
            "Status",
            "TransferRate",
            "VolumeSerialNumber",
        ]
        
        self.print_title("Win32_CDROMDrive")
        for info in self.client.Win32_CDROMDrive():
            self.show(info, PROPERTIES)
            print("---")
            
    
    def computer_system(self):
        PC_SYSTEM_TYPE = {
            0: "Unspecified",
            1: "Desktop",
            2: "Mobile",
            3: "Workstation",
            4: "Enterprise Server",
            5: "SOHO Server",
            6: "Appliance PC",
            7: "Performance Server",
            8: "Maximum",
        }
        
        PC_SYSTEM_TYPE_EX = {
            0: "Unspecified",
            1: "Desktop",
            2: "Mobile",
            3: "Workstation",
            4: "Enterprise Server",
            5: "SOHO Server",
            6: "Appliance PC",
            7: "Performance Server",
            8: "Slate",
            9: "Maximum",
        }
        
        PROPERTIES = [
            "Caption",
            "ChassisSKUNumber", # Win10
            "Description",
            "DNSHostName",
            "Domain",
            "Manufacturer",
            "Model",
            "Name",
            "OEMStringArray",
            "SystemFamily",     # Win10
            "SystemSKUNumber",  # Win10
            "SystemType",
            "TotalPhysicalMemory",
        ]
        
        self.print_title("Win32_ComputerSystem")
        for info in self.client.Win32_ComputerSystem():
            self.show(info, PROPERTIES)

            self.show_with_conversion(info, "PCSystemType", PC_SYSTEM_TYPE)
            self.show_with_conversion(info, "PCSystemTypeEx", PC_SYSTEM_TYPE_EX)
            print("---")


    def desktop_monitor(self):
        DISPLAY_TYPE = {
            0: "Unknown",
            1: "Other",
            2: "Multiscan Color",
            3: "Multiscan Monochrome",
            4: "Fixed Frequency Color",
            5: "Fixed Frequency Monochrome",
        }
        
        PROPERTIES = [
            "Caption",
            "Bandwidth",
            "Description",
            "DeviceID",
            "MonitorManufacturer",
            "MonitorType",
            "Name",
            "PixelsPerXLogicalInch",
            "PixelsPerYLogicalInch",
            "ScreenHeight",
            "ScreenWidth",
            "Status",
        ]
        
        self.print_title("Win32_DesktopMonitor")
        for info in self.client.Win32_DesktopMonitor():
            self.show(info, PROPERTIES)
            self.show_with_conversion(info, "DisplayType", DISPLAY_TYPE)
            print("---")
    
    
    def disk_drive(self):
        PROPERTIES = [
            "Caption",
            "BytesPerSector",
            "CompressionMethod",
            "DefaultBlockSize",
            "Description",
            "DeviceID",
            "InterfaceType",
            "Manufacturer",
            "MediaType",
            "Model",
            "Name",
            "Partitions",
            "PNPDeviceID",
            "SectorsPerTrack",
            "SerialNumber",
            "Size",
            "Status",
            "TotalCylinders",
            "TotalHeads",
            "TotalSectors",
            "TotalTracks",
            "TracksPerCylinder",
        ]
        
        self.print_title("Win32_DiskDrive")
        for info in self.client.Win32_DiskDrive():
            self.show(info, PROPERTIES)
            print("---")
    
    
    def fan(self):
        PROPERTIES = [
            "Caption",
            "ConfigManagerErrorCode",
            "Description",
            "DesiredSpeed",
            "ErrorDescription",
            "Name",
            "PowerManagementCapabilities",
            "Status",
        ]
        
        self.print_title("Win32_Fan")
        for info in self.client.Win32_Fan():
            self.show(info, PROPERTIES)
            print("---")
    
    
    def network_adapter_configuration(self):
        PROPERTIES = [
            "Caption",
            "DefaultIPGateway",
            "Description",
            "DHCPEnabled",
            "DHCPServer",
            "IPAddress",
            "IPSubnet",
            "IPEnabled",
            "DNSServerSearchOrder",
            "MACAddress",
            "ServiceName",
            "TcpWindowSize",
        ]
        
        self.print_title("Win32_NetworkAdapterConfiguration")
        for info in self.client.Win32_NetworkAdapterConfiguration():
            self.show(info, PROPERTIES)
            print("---")


    def operating_system(self):
        PROPERTIES = [
            "Name",
            "Caption",
            "CSDVersion",
            "OSArchitecture",
            "CSName",
        ]
        
        self.print_title("Win32_OperatingSystem")
        for info in self.client.Win32_OperatingSystem():
            self.show(info, PROPERTIES)
            print("---")
            
            
    def physical_memory(self):
        MEMORY_TYPE = {
            0: "Unknown",
            1: "Other",
            2: "DRAM",
            3: "Synchronous DRAM",
            4: "Cache DRAM",
            5: "EDO",
            6: "EDRAM",
            7: "VRAM",
            8: "SRAM",
            9: "RAM",
            10: "ROM",
            11: "Flash",
            12: "EEPROM",
            13: "FEPROM",
            14: "EPROM",
            15: "CDRAM",
            16: "3DRAM",
            17: "SDRAM",
            18: "SGRAM",
            19: "RDRAM",
            20: "DDR",
            21: "DDR2",
            22: "DDR2",
            23: "DDR2 FB-DIMM",
            24: "DDR3",
            25: "FBD2",
        }
        
        PROPERTIES = [
            "Caption",
            "ConfiguredClockSpeed", # Win10
            "DataWidth",
            "Description",
            "DeviceLocator",
            "Manufacturer",
            "Model",
            "Name",
            "OtherIdentifyingInfo",
            "PartNumber",
            "SerialNumber",
            "SKU",
            "Speed",
            "Tag",
        ]
        
        self.print_title("Win32_PhysicalMemory")
        for info in self.client.Win32_PhysicalMemory():
            self.show(info, PROPERTIES)
            
            self.show_with_conversion(info, "MemoryType", MEMORY_TYPE)
            print("---")
            
            
    def processor(self):
        PROPERTIES = [
            "Name",
            "SerialNumber", # Win10
            "SystemName",
        ]
    
        self.print_title("Win32_Processor")
        for info in self.client.Win32_Processor():
            self.show(info, PROPERTIES)
            print("---")
            
                
    def video_controller(self):
        PROPERTIES = [
            "Caption",
            "AdapterCompatibility",
            "AdapterDACType",
            "AdapterRAM",
            "ColorTableEntries",
            "CurrentBitsPerPixel",
            "CurrentHorizontalResolution",
            "CurrentNumberOfColors",
            "CurrentNumberOfColumns",
            "CurrentNumberOfRows",
            "CurrentRefreshRate",
            "CurrentScanMode",
            "CurrentVerticalResolution",
            "Description",
            "DeviceSpecificPens",
            "DriverVersion",
            "MaxMemorySupported",
            "MaxNumberControlled",
            "MaxRefreshRate",
            "MinRefreshRate",
            "Monochrome",
            "Name",
            "ProtocolSupported",
            "Status",
            "VideoArchitecture",
            "VideoMemoryType",
            "VideoMode",
            "VideoModeDescription",
            "VideoProcessor",
        ]
        
        self.print_title("Win32_VideoController")
        for info in self.client.Win32_VideoController():
            self.show(info, PROPERTIES)
            print("---")


    def installed_software(self):
        # 実行に時間がかかることに注意
        PROPERTIES = [
            "Name",
            "Version",
            "InstallLocation",
            "InstallSource",
            "LocalPackage",
            "PackageName",
        ]
        
        products = { s.Name: s for s in self.client.Win32_Product() }

        self.print_title("Win32_Product")
        for k, v in sorted(products.items()):
            self.show(v, PROPERTIES)
            print("---")
            
            
    def get_office_installed_path(self):
        PROPERTIES = [
            "Path",
            "Version",
            "Build",
        ]
        
        excel = win32com.client.Dispatch("Excel.Application")
        self.show(excel, PROPERTIES)

        result = excel.Path        
        excel.Quit()
    
        return result
        
    
    def show_by_ospp(self):
        office_path = self.get_office_installed_path()
        ospp_path = os.path.join(office_path, "ospp.vbs")
        
        if os.path.isfile(ospp_path):
            # Python3.5以降
            subprocess.run(['cscript', ospp_path, '/dstatus'], shell=True)
        else:
            print("{}がありません".format(ospp_path))


if __name__ == "__main__":
    pc = PC()
    
    # Hardware info:
    pc.base_board()
    pc.cd_rom_drive()
    pc.computer_system()
    pc.desktop_monitor()
    pc.disk_drive()
    pc.fan()
    pc.network_adapter_configuration()
    pc.operating_system()
    pc.processor()
    pc.physical_memory()
    pc.video_controller()
    
    # Software info:
    pc.installed_software()
    
    # MS Office info:
    pc.show_by_ospp()