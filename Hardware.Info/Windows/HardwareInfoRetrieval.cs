using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using WmiLight;

// https://docs.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-osversioninfoexa

namespace Hardware.Info.Windows
{
    // https://docs.microsoft.com/en-us/windows/win32/api/sysinfoapi/ns-sysinfoapi-memorystatusex

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    internal class MEMORYSTATUSEX
    {
        public uint dwLength;
        public uint dwMemoryLoad;
        public ulong ullTotalPhys;
        public ulong ullAvailPhys;
        public ulong ullTotalPageFile;
        public ulong ullAvailPageFile;
        public ulong ullTotalVirtual;
        public ulong ullAvailVirtual;
        public ulong ullAvailExtendedVirtual;

        public MEMORYSTATUSEX()
        {
            dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX));
        }
    }

    internal class HardwareInfoRetrieval : HardwareInfoBase, IHardwareInfoRetrieval
    {
        private readonly MEMORYSTATUSEX _memoryStatusEx = new MEMORYSTATUSEX();

        private readonly MemoryStatus _memoryStatus = new MemoryStatus();

        private readonly OS _os = new OS();

        public bool UseAsteriskInWMI { get; set; }

        private readonly string _managementScope = "\\\\.\\root\\cimv2";
        private readonly string _managementScopeWmi = "\\\\.\\root\\wmi";
        private readonly EnumeratorBehaviorOption _enumerationOptions = EnumeratorBehaviorOption.ReturnImmediately;

        public HardwareInfoRetrieval()
        {
            GetOs();
        }

        [DllImport("ntdll.dll", SetLastError = true)]
        private static extern int RtlGetVersion([In, Out] ref OSVERSIONINFOEX lpVersionInformation);

        [StructLayout(LayoutKind.Sequential)]
        private struct OSVERSIONINFOEX
        {
            public uint dwOSVersionInfoSize;
            public uint dwMajorVersion;
            public uint dwMinorVersion;
            public uint dwBuildNumber;
            public uint dwPlatformId;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szCSDVersion;

            public ushort wServicePackMajor;
            public ushort wServicePackMinor;
            public ushort wSuiteMask;
            public byte wProductType;
            public byte wReserved;
        }

        public static Version? GetOsVersionByRtlGetVersion()
        {
            OSVERSIONINFOEX info = new OSVERSIONINFOEX();
            info.dwOSVersionInfoSize = (uint)Marshal.SizeOf(info);

            int result = RtlGetVersion(ref info);

            return (result == 0) // STATUS_SUCCESS
                ? new Version((int)info.dwMajorVersion, (int)info.dwMinorVersion, (int)info.dwBuildNumber)
                : null;
        }

        public void GetOs()
        {
            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_OperatingSystem"
                                                  : "SELECT Caption, Version FROM Win32_OperatingSystem";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                _os.Name = GetPropertyString(mo["Caption"]);
                _os.VersionString = GetPropertyString(mo["Version"]);

                if (Version.TryParse(_os.VersionString, out Version version))
                    _os.Version = version;
            }

            if (string.IsNullOrEmpty(_os.Name))
            {
                _os.Name = "Windows";
            }

            if (string.IsNullOrEmpty(_os.VersionString))
            {
                Version? version = GetOsVersionByRtlGetVersion();

                if (version != null)
                {
                    _os.Version = version;
                    _os.VersionString = version.ToString();
                }
            }
        }

        public OS GetOperatingSystem()
        {
            return _os;
        }

        // https://docs.microsoft.com/en-us/windows/win32/api/sysinfoapi/nf-sysinfoapi-globalmemorystatusex

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX lpBuffer);

        public MemoryStatus GetMemoryStatus()
        {
            if (GlobalMemoryStatusEx(_memoryStatusEx))
            {
                _memoryStatus.TotalPhysical = _memoryStatusEx.ullTotalPhys;
                _memoryStatus.AvailablePhysical = _memoryStatusEx.ullAvailPhys;
                _memoryStatus.TotalPageFile = _memoryStatusEx.ullTotalPageFile;
                _memoryStatus.AvailablePageFile = _memoryStatusEx.ullAvailPageFile;
                _memoryStatus.TotalVirtual = _memoryStatusEx.ullTotalVirtual;
                _memoryStatus.AvailableVirtual = _memoryStatusEx.ullAvailVirtual;
                _memoryStatus.AvailableExtendedVirtual = _memoryStatusEx.ullAvailExtendedVirtual;
            }

            return _memoryStatus;
        }

        public static T GetPropertyValue<T>(object obj) where T : struct
        {
            return (obj == null) ? default : (T)obj;
        }

        public static T[] GetPropertyArray<T>(object obj)
        {
            return (obj is T[] array) ? array : Array.Empty<T>();
        }

        public static string GetPropertyString(object obj)
        {
            return (obj is string str) ? str : string.Empty;
        }

        public static string GetStringFromUInt16Array(ushort[] array)
        {
            try
            {
                if (array.Length == 0)
                    return string.Empty;

                byte[] byteArray = new byte[array.Length * 2];
                Buffer.BlockCopy(array, 0, byteArray, 0, byteArray.Length);

                string str = Encoding.Unicode.GetString(byteArray).Trim('\0');

                return str;
            }
            catch
            {
                return string.Empty;
            }
        }

        // https://docs.microsoft.com/en-us/dotnet/api/system.management.managementpath.defaultpath?view=netframework-4.8

        public List<Battery> GetBatteryList()
        {
            List<Battery> batteryList = new List<Battery>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_Battery"
                                                  : "SELECT FullChargeCapacity, DesignCapacity, BatteryStatus, EstimatedChargeRemaining, EstimatedRunTime, ExpectedLife, MaxRechargeTime, TimeOnBattery, TimeToFullCharge FROM Win32_Battery";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                Battery battery = new Battery
                {
                    FullChargeCapacity = GetPropertyValue<uint>(mo["FullChargeCapacity"]),
                    DesignCapacity = GetPropertyValue<uint>(mo["DesignCapacity"]),
                    BatteryStatus = GetPropertyValue<ushort>(mo["BatteryStatus"]),
                    EstimatedChargeRemaining = GetPropertyValue<ushort>(mo["EstimatedChargeRemaining"]),
                    EstimatedRunTime = GetPropertyValue<uint>(mo["EstimatedRunTime"]),
                    ExpectedLife = GetPropertyValue<uint>(mo["ExpectedLife"]),
                    MaxRechargeTime = GetPropertyValue<uint>(mo["MaxRechargeTime"]),
                    TimeOnBattery = GetPropertyValue<uint>(mo["TimeOnBattery"]),
                    TimeToFullCharge = GetPropertyValue<uint>(mo["TimeToFullCharge"])
                };

                batteryList.Add(battery);
            }

            return batteryList;
        }

        public List<BIOS> GetBiosList()
        {
            List<BIOS> biosList = new List<BIOS>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_BIOS"
                                                  : "SELECT Caption, Description, Manufacturer, Name, ReleaseDate, SerialNumber, SoftwareElementID, Version FROM Win32_BIOS";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                BIOS bios = new BIOS
                {
                    Caption = GetPropertyString(mo["Caption"]),
                    Description = GetPropertyString(mo["Description"]),
                    Manufacturer = GetPropertyString(mo["Manufacturer"]),
                    Name = GetPropertyString(mo["Name"]),
                    ReleaseDate = GetPropertyString(mo["ReleaseDate"]),
                    SerialNumber = GetPropertyString(mo["SerialNumber"]),
                    SoftwareElementID = GetPropertyString(mo["SoftwareElementID"]),
                    Version = GetPropertyString(mo["Version"])
                };

                biosList.Add(bios);
            }

            return biosList;
        }

        public List<CPU> GetCpuList(bool includePercentProcessorTime = true)
        {
            List<CPU> cpuList = new List<CPU>();

                List<CpuCore> cpuCoreList = new List<CpuCore>();

                ulong percentProcessorTime = 0ul;

                using WmiConnection con = new WmiConnection(_managementScope);

                if (includePercentProcessorTime)
                {
                    string queryString = UseAsteriskInWMI
                        ? "SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name != '_Total'"
                        : "SELECT Name, PercentProcessorTime FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name != '_Total'";
                    var percentProcessorTimeQuery = con.CreateQuery(queryString, _enumerationOptions);

                    queryString = UseAsteriskInWMI
                        ? "SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name = '_Total'"
                        : "SELECT PercentProcessorTime FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name = '_Total'";
                    var totalPercentProcessorTimeQuery = con.CreateQuery(queryString, _enumerationOptions);

                    try
                    {
                        foreach (WmiObject mo in percentProcessorTimeQuery)
                        {
                            CpuCore core = new CpuCore
                            {
                                Name = GetPropertyString(mo["Name"]),
                                PercentProcessorTime = GetPropertyValue<ulong>(mo["PercentProcessorTime"])
                            };

                            cpuCoreList.Add(core);
                        }

                        foreach (WmiObject mo in totalPercentProcessorTimeQuery)
                        {
                            percentProcessorTime = GetPropertyValue<ulong>(mo["PercentProcessorTime"]);
                        }
                    }
                    catch (WmiException)
                    {
                        // https://github.com/Jinjinov/Hardware.Info/issues/30
                    }

                    if (percentProcessorTime == 0ul)
                    {
                        queryString = UseAsteriskInWMI
                            ? "SELECT * FROM Win32_Processor"
                            : "SELECT LoadPercentage FROM Win32_Processor";
                        var loadPercentageQuery = con.CreateQuery(queryString, _enumerationOptions);

                        foreach (WmiObject mo in loadPercentageQuery)
                        {
                            percentProcessorTime = GetPropertyValue<ushort>(mo["LoadPercentage"]);
                        }
                    }
                }

                bool isAtLeastWin8 = (_os.Version.Major == 6 && _os.Version.Minor >= 2) || (_os.Version.Major > 6);

                string query = UseAsteriskInWMI
                    ? "SELECT * FROM Win32_Processor"
                    : isAtLeastWin8
                        ? "SELECT Caption, CurrentClockSpeed, Description, L2CacheSize, L3CacheSize, Manufacturer, MaxClockSpeed, Name, NumberOfCores, NumberOfLogicalProcessors, ProcessorId, SecondLevelAddressTranslationExtensions, SocketDesignation, VirtualizationFirmwareEnabled, VMMonitorModeExtensions FROM Win32_Processor"
                        : "SELECT Caption, CurrentClockSpeed, Description, L2CacheSize, L3CacheSize, Manufacturer, MaxClockSpeed, Name, NumberOfCores, NumberOfLogicalProcessors, ProcessorId, SocketDesignation FROM Win32_Processor";
                var mos = con.CreateQuery(query, _enumerationOptions);

                float processorPerformance = 100f;

                try
                {
                    using PerformanceCounter cpuCounter =
                        new PerformanceCounter("Processor Information", "% Processor Performance", "_Total");
                    processorPerformance = cpuCounter.NextValue();
                    System.Threading.Thread.Sleep(1); // the first call to NextValue() always returns 0
                    processorPerformance = cpuCounter.NextValue();
                }
                catch
                {
                    // Ignore performance counter errors and just assume that it's at 100 %
                }

                uint L1InstructionCacheSize = 0;
                uint L1DataCacheSize = 0;
                // L1 = 3
                // L2 = 4
                // L3 = 5
                query = UseAsteriskInWMI
                    ? "SELECT * FROM Win32_CacheMemory WHERE Level = 3"
                    : "SELECT CacheType, MaxCacheSize FROM Win32_CacheMemory WHERE Level = 3";

                var Win32_CacheMemory = con.CreateQuery(query, _enumerationOptions);

                // Other = 1
                // Unknown = 2
                // Instruction = 3
                // Data = 4
                // Unified = 5
                foreach (WmiObject mo in Win32_CacheMemory)
                {
                    ushort CacheType = GetPropertyValue<ushort>(mo["CacheType"]);
                    uint MaxCacheSize = 1024 * GetPropertyValue<uint>(mo["MaxCacheSize"]);

                    // if CacheType is Other or Unknown
                    if (L1InstructionCacheSize == 0)
                        L1InstructionCacheSize = MaxCacheSize;

                    // if CacheType is Other or Unknown
                    if (L1DataCacheSize == 0)
                        L1DataCacheSize = MaxCacheSize;

                    if (CacheType == 3) // Instruction
                        L1InstructionCacheSize = MaxCacheSize;

                    if (CacheType == 4) // Data
                        L1DataCacheSize = MaxCacheSize;
                }

                foreach (WmiObject mo in mos)
                {
                    uint maxClockSpeed = GetPropertyValue<uint>(mo["MaxClockSpeed"]);

                    uint currentClockSpeed = (uint)(maxClockSpeed * (processorPerformance / 100));

                    CPU cpu = new CPU
                    {
                        Caption = GetPropertyString(mo["Caption"]),
                        //CurrentClockSpeed = GetPropertyValue<uint>(mo["CurrentClockSpeed"]), https://stackoverflow.com/questions/61802420/unable-to-get-current-cpu-frequency-in-powershell-or-python
                        CurrentClockSpeed = currentClockSpeed,
                        Description = GetPropertyString(mo["Description"]),
                        L1InstructionCacheSize = L1InstructionCacheSize,
                        L1DataCacheSize = L1DataCacheSize,
                        L2CacheSize = 1024 * GetPropertyValue<uint>(mo["L2CacheSize"]),
                        L3CacheSize = 1024 * GetPropertyValue<uint>(mo["L3CacheSize"]),
                        Manufacturer = GetPropertyString(mo["Manufacturer"]),
                        MaxClockSpeed = maxClockSpeed,
                        Name = GetPropertyString(mo["Name"]),
                        NumberOfCores = GetPropertyValue<uint>(mo["NumberOfCores"]),
                        NumberOfLogicalProcessors = GetPropertyValue<uint>(mo["NumberOfLogicalProcessors"]),
                        ProcessorId = GetPropertyString(mo["ProcessorId"]),
                        SocketDesignation = GetPropertyString(mo["SocketDesignation"]),
                        PercentProcessorTime = percentProcessorTime,
                        CpuCoreList = cpuCoreList
                    };

                    if (isAtLeastWin8)
                    {
                        cpu.SecondLevelAddressTranslationExtensions =
                            GetPropertyValue<bool>(mo["SecondLevelAddressTranslationExtensions"]);
                        cpu.VirtualizationFirmwareEnabled = GetPropertyValue<bool>(mo["VirtualizationFirmwareEnabled"]);
                        cpu.VMMonitorModeExtensions = GetPropertyValue<bool>(mo["VMMonitorModeExtensions"]);
                    }

                    cpuList.Add(cpu);
                }

                return cpuList;
        }

        public override List<Drive> GetDriveList()
        {
            List<Drive> driveList = new List<Drive>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_DiskDrive"
                                                  : "SELECT Caption, Description, DeviceID, FirmwareRevision, Index, Manufacturer, Model, Name, Partitions, SerialNumber, Size FROM Win32_DiskDrive";

            using WmiConnection con = new WmiConnection(_managementScope);

            var Win32_DiskDrive = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject DiskDrive in Win32_DiskDrive)
            {
                Drive drive = new Drive
                {
                    Caption = GetPropertyString(DiskDrive["Caption"]),
                    Description = GetPropertyString(DiskDrive["Description"]),
                    FirmwareRevision = GetPropertyString(DiskDrive["FirmwareRevision"]),
                    Index = GetPropertyValue<uint>(DiskDrive["Index"]),
                    Manufacturer = GetPropertyString(DiskDrive["Manufacturer"]),
                    Model = GetPropertyString(DiskDrive["Model"]),
                    Name = GetPropertyString(DiskDrive["Name"]),
                    Partitions = GetPropertyValue<uint>(DiskDrive["Partitions"]),
                    SerialNumber = GetPropertyString(DiskDrive["SerialNumber"]),
                    Size = GetPropertyValue<ulong>(DiskDrive["Size"])
                };

                string queryString1 = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + DiskDrive["DeviceID"] + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition";
               var Win32_DiskPartition = con.CreateQuery(queryString1, _enumerationOptions);

                foreach (WmiObject DiskPartition in Win32_DiskPartition)
                {
                    Partition partition = new Partition
                    {
                        Bootable = GetPropertyValue<bool>(DiskPartition["Bootable"]),
                        BootPartition = GetPropertyValue<bool>(DiskPartition["BootPartition"]),
                        Caption = GetPropertyString(DiskPartition["Caption"]),
                        Description = GetPropertyString(DiskPartition["Description"]),
                        DiskIndex = GetPropertyValue<uint>(DiskPartition["DiskIndex"]),
                        Index = GetPropertyValue<uint>(DiskPartition["Index"]),
                        Name = GetPropertyString(DiskPartition["Name"]),
                        PrimaryPartition = GetPropertyValue<bool>(DiskPartition["PrimaryPartition"]),
                        Size = GetPropertyValue<ulong>(DiskPartition["Size"]),
                        StartingOffset = GetPropertyValue<ulong>(DiskPartition["StartingOffset"])
                    };

                    string queryString2 = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + DiskPartition["DeviceID"] + "'} WHERE AssocClass = Win32_LogicalDiskToPartition";
                    var Win32_LogicalDisk = con.CreateQuery(queryString2, _enumerationOptions);

                    foreach (WmiObject LogicalDisk in Win32_LogicalDisk)
                    {
                        Volume volume = new Volume
                        {
                            Caption = GetPropertyString(LogicalDisk["Caption"]),
                            Compressed = GetPropertyValue<bool>(LogicalDisk["Compressed"]),
                            Description = GetPropertyString(LogicalDisk["Description"]),
                            FileSystem = GetPropertyString(LogicalDisk["FileSystem"]),
                            FreeSpace = GetPropertyValue<ulong>(LogicalDisk["FreeSpace"]),
                            Name = GetPropertyString(LogicalDisk["Name"]),
                            Size = GetPropertyValue<ulong>(LogicalDisk["Size"]),
                            VolumeName = GetPropertyString(LogicalDisk["VolumeName"]),
                            VolumeSerialNumber = GetPropertyString(LogicalDisk["VolumeSerialNumber"])
                        };

                        partition.VolumeList.Add(volume);
                    }

                    drive.PartitionList.Add(partition);
                }

                driveList.Add(drive);
            }

            return driveList;
        }

        public List<Keyboard> GetKeyboardList()
        {
            List<Keyboard> keyboardList = new List<Keyboard>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_Keyboard"
                                                  : "SELECT Caption, Description, Name, NumberOfFunctionKeys FROM Win32_Keyboard";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                Keyboard keyboard = new Keyboard
                {
                    Caption = GetPropertyString(mo["Caption"]),
                    Description = GetPropertyString(mo["Description"]),
                    Name = GetPropertyString(mo["Name"]),
                    NumberOfFunctionKeys = GetPropertyValue<ushort>(mo["NumberOfFunctionKeys"])
                };

                keyboardList.Add(keyboard);
            }

            return keyboardList;
        }

        public List<Memory> GetMemoryList()
        {
            List<Memory> memoryList = new List<Memory>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_PhysicalMemory"
                                                  : _os.Version.Major >= 10 ? "SELECT BankLabel, Capacity, FormFactor, Manufacturer, MaxVoltage, MinVoltage, PartNumber, SerialNumber, Speed FROM Win32_PhysicalMemory"
                                                                            : "SELECT BankLabel, Capacity, FormFactor, Manufacturer, PartNumber, SerialNumber, Speed FROM Win32_PhysicalMemory";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                Memory memory = new Memory
                {
                    BankLabel = GetPropertyString(mo["BankLabel"]),
                    Capacity = GetPropertyValue<ulong>(mo["Capacity"]),
                    FormFactor = (FormFactor)GetPropertyValue<ushort>(mo["FormFactor"]),
                    Manufacturer = GetPropertyString(mo["Manufacturer"]),
                    PartNumber = GetPropertyString(mo["PartNumber"]),
                    SerialNumber = GetPropertyString(mo["SerialNumber"]),
                    Speed = GetPropertyValue<uint>(mo["Speed"])
                };

                if (_os.Version.Major >= 10)
                {
                    memory.MaxVoltage = GetPropertyValue<uint>(mo["MaxVoltage"]);
                    memory.MinVoltage = GetPropertyValue<uint>(mo["MinVoltage"]);
                }

                memoryList.Add(memory);
            }

            return memoryList;
        }

        public List<Monitor> GetMonitorList()
        {
            List<Monitor> monitorList = new List<Monitor>();

            string win32PnpEntityQuery = UseAsteriskInWMI ? "SELECT * FROM Win32_PnPEntity WHERE PNPClass='Monitor'"
                                                          : "SELECT DeviceId FROM Win32_PnPEntity WHERE PNPClass='Monitor'";

            using WmiConnection con = new WmiConnection(_managementScope);
            using WmiConnection conWmi = new WmiConnection(_managementScopeWmi);

            var win32PnpEntityMos = con.CreateQuery(win32PnpEntityQuery, _enumerationOptions);

            foreach (WmiObject win32PnpEntityMo in win32PnpEntityMos)
            {
                string deviceId = GetPropertyString(win32PnpEntityMo["DeviceId"]);
                string win32DesktopMonitorQuery = UseAsteriskInWMI ? $"SELECT * FROM Win32_DesktopMonitor WHERE PNPDeviceId='{deviceId}'"
                                                                   : $"SELECT Caption, Description, MonitorManufacturer, MonitorType, Name, PixelsPerXLogicalInch, PixelsPerYLogicalInch FROM Win32_DesktopMonitor WHERE PNPDeviceId='{deviceId}'";
                var win32DesktopMonitorMos = con.CreateQuery(win32DesktopMonitorQuery.Replace(@"\", @"\\"), _enumerationOptions);

                using WmiObject? desktopMonitorMo = win32DesktopMonitorMos.FirstOrDefault();

                Monitor monitor = new Monitor();

                if (desktopMonitorMo != null)
                {
                    monitor.Caption = GetPropertyString(desktopMonitorMo["Caption"]);
                    monitor.Description = GetPropertyString(desktopMonitorMo["Description"]);
                    monitor.MonitorManufacturer = GetPropertyString(desktopMonitorMo["MonitorManufacturer"]);
                    monitor.MonitorType = GetPropertyString(desktopMonitorMo["MonitorType"]);
                    monitor.Name = GetPropertyString(desktopMonitorMo["Name"]);
                    monitor.PixelsPerXLogicalInch = GetPropertyValue<uint>(desktopMonitorMo["PixelsPerXLogicalInch"]);
                    monitor.PixelsPerYLogicalInch = GetPropertyValue<uint>(desktopMonitorMo["PixelsPerYLogicalInch"]);
                }

                try
                {
                    string wmiMonitorIdQuery = UseAsteriskInWMI
                        ? $"SELECT * FROM WmiMonitorID WHERE InstanceName LIKE '{deviceId}%'"
                        : $"SELECT Active, ProductCodeID, SerialNumberID, ManufacturerName, UserFriendlyName, WeekOfManufacture, YearOfManufacture FROM WmiMonitorID WHERE InstanceName LIKE '{deviceId}%'";
                    var wmiMonitorIdMos = conWmi.CreateQuery(wmiMonitorIdQuery.Replace(@"\", "_"), _enumerationOptions);

                    using WmiObject? wmiMonitorIdMo = wmiMonitorIdMos.FirstOrDefault();

                    if (wmiMonitorIdMo != null)
                    {
                        monitor.Active = GetPropertyValue<bool>(wmiMonitorIdMo["Active"]);
                        monitor.ProductCodeID =
                            GetStringFromUInt16Array(GetPropertyArray<ushort>(wmiMonitorIdMo["ProductCodeID"]));
                        monitor.UserFriendlyName =
                            GetStringFromUInt16Array(GetPropertyArray<ushort>(wmiMonitorIdMo["UserFriendlyName"]));
                        monitor.SerialNumberID =
                            GetStringFromUInt16Array(GetPropertyArray<ushort>(wmiMonitorIdMo["SerialNumberID"]));
                        monitor.ManufacturerName =
                            GetStringFromUInt16Array(GetPropertyArray<ushort>(wmiMonitorIdMo["ManufacturerName"]));
                        monitor.WeekOfManufacture = GetPropertyValue<byte>(wmiMonitorIdMo["WeekOfManufacture"]);
                        monitor.YearOfManufacture = GetPropertyValue<ushort>(wmiMonitorIdMo["YearOfManufacture"]);
                    }
                }
                catch (Exception ex) when (ex is COMException || ex is WmiException)
                {
                    // Ignore WmiMonitorID errors
                }

                monitorList.Add(monitor);
            }

            return monitorList;
        }

        public List<Motherboard> GetMotherboardList()
        {
            List<Motherboard> motherboardList = new List<Motherboard>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_BaseBoard"
                                                  : "SELECT Manufacturer, Product, SerialNumber FROM Win32_BaseBoard";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                Motherboard motherboard = new Motherboard
                {
                    Manufacturer = GetPropertyString(mo["Manufacturer"]),
                    Product = GetPropertyString(mo["Product"]),
                    SerialNumber = GetPropertyString(mo["SerialNumber"])
                };

                motherboardList.Add(motherboard);
            }

            return motherboardList;
        }

        public List<Mouse> GetMouseList()
        {
            List<Mouse> mouseList = new List<Mouse>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_PointingDevice"
                                                  : "SELECT Caption, Description, Manufacturer, Name, NumberOfButtons FROM Win32_PointingDevice";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                Mouse mouse = new Mouse
                {
                    Caption = GetPropertyString(mo["Caption"]),
                    Description = GetPropertyString(mo["Description"]),
                    Manufacturer = GetPropertyString(mo["Manufacturer"]),
                    Name = GetPropertyString(mo["Name"]),
                    NumberOfButtons = GetPropertyValue<byte>(mo["NumberOfButtons"])
                };

                mouseList.Add(mouse);
            }

            return mouseList;
        }

        public override List<NetworkAdapter> GetNetworkAdapterList(bool includeBytesPersec = true, bool includeNetworkAdapterConfiguration = true)
        {
            List<NetworkAdapter> networkAdapterList = new List<NetworkAdapter>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True AND MACAddress IS NOT NULL"
                                                  : "SELECT AdapterType, Caption, Description, DeviceID, MACAddress, Manufacturer, Name, NetConnectionID, ProductName, Speed, InterfaceIndex FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True AND MACAddress IS NOT NULL";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                NetworkAdapter networkAdapter = new NetworkAdapter
                {
                    AdapterType = GetPropertyString(mo["AdapterType"]),
                    Caption = GetPropertyString(mo["Caption"]),
                    Description = GetPropertyString(mo["Description"]),
                    MACAddress = GetPropertyString(mo["MACAddress"]),
                    Manufacturer = GetPropertyString(mo["Manufacturer"]),
                    Name = GetPropertyString(mo["Name"]),
                    NetConnectionID = GetPropertyString(mo["NetConnectionID"]),
                    ProductName = GetPropertyString(mo["ProductName"]),
                    Speed = GetPropertyValue<ulong>(mo["Speed"])
                };

                var interfaceIndex =  GetPropertyValue<uint>(mo["InterfaceIndex"]);

                if (includeBytesPersec)
                {
                    // https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.performancecounter.instancename

                    string name = networkAdapter.Name.Replace('(', '[').Replace(')', ']').Replace('#', '_').Replace('\\', '_').Replace('/', '_');

                    string query = UseAsteriskInWMI ? $"SELECT * FROM Win32_PerfFormattedData_Tcpip_NetworkAdapter WHERE Name = '{name}'"
                                                    : $"SELECT BytesSentPersec, BytesReceivedPersec, CurrentBandwidth FROM Win32_PerfFormattedData_Tcpip_NetworkAdapter WHERE Name = '{name}'";
                    var managementObjectSearcher = con.CreateQuery(query, _enumerationOptions);

                    foreach (WmiObject managementObject in managementObjectSearcher)
                    {
                        networkAdapter.BytesSentPersec = GetPropertyValue<ulong>(managementObject["BytesSentPersec"]);
                        networkAdapter.BytesReceivedPersec = GetPropertyValue<ulong>(managementObject["BytesReceivedPersec"]);

                        if (networkAdapter.Speed == 0 || networkAdapter.Speed == long.MaxValue)
                        {
                            networkAdapter.Speed = GetPropertyValue<ulong>(managementObject["CurrentBandwidth"]);
                        }
                    }
                }

                if (includeNetworkAdapterConfiguration && mo is WmiObject networkAdapterWmiObject)
                {
                    IPAddress address;

                    string query = UseAsteriskInWMI ? $"SELECT * FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = {interfaceIndex}"
                                                    : $"SELECT DefaultIPGateway, DHCPServer, DNSServerSearchOrder, IPAddress, IPSubnet FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = {interfaceIndex}";
                    var adapterConfigurationQuery = con.CreateQuery(query, _enumerationOptions);

                    foreach (WmiObject configuration in adapterConfigurationQuery)
                    {
                        foreach (string str in GetPropertyArray<string>(configuration["DefaultIPGateway"]))
                            if (IPAddress.TryParse(str, out address))
                                networkAdapter.DefaultIPGatewayList.Add(address);

                        if (IPAddress.TryParse(GetPropertyString(configuration["DHCPServer"]), out address))
                            networkAdapter.DHCPServer = address;

                        foreach (string str in GetPropertyArray<string>(configuration["DNSServerSearchOrder"]))
                            if (IPAddress.TryParse(str, out address))
                                networkAdapter.DNSServerSearchOrderList.Add(address);

                        foreach (string str in GetPropertyArray<string>(configuration["IPAddress"]))
                            if (IPAddress.TryParse(str, out address))
                                networkAdapter.IPAddressList.Add(address);

                        foreach (string str in GetPropertyArray<string>(configuration["IPSubnet"]))
                            if (IPAddress.TryParse(str, out address))
                                networkAdapter.IPSubnetList.Add(address);
                    }
                }

                networkAdapterList.Add(networkAdapter);
            }

            return networkAdapterList;
        }

        public List<Printer> GetPrinterList()
        {
            List<Printer> printerList = new List<Printer>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_Printer"
                                                  : "SELECT Caption, Default, Description, HorizontalResolution, Local, Name, Network, Shared, VerticalResolution FROM Win32_Printer";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                Printer printer = new Printer
                {
                    Caption = GetPropertyString(mo["Caption"]),
                    Default = GetPropertyValue<bool>(mo["Default"]),
                    Description = GetPropertyString(mo["Description"]),
                    HorizontalResolution = GetPropertyValue<uint>(mo["HorizontalResolution"]),
                    Local = GetPropertyValue<bool>(mo["Local"]),
                    Name = GetPropertyString(mo["Name"]),
                    Network = GetPropertyValue<bool>(mo["Network"]),
                    Shared = GetPropertyValue<bool>(mo["Shared"]),
                    VerticalResolution = GetPropertyValue<uint>(mo["VerticalResolution"])
                };

                printerList.Add(printer);
            }

            return printerList;
        }

        public List<SoundDevice> GetSoundDeviceList()
        {
            List<SoundDevice> soundDeviceList = new List<SoundDevice>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_SoundDevice WHERE NOT Manufacturer='Microsoft'"
                                                  : "SELECT Caption, Description, Manufacturer, Name, ProductName FROM Win32_SoundDevice WHERE NOT Manufacturer='Microsoft'";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                SoundDevice soundDevice = new SoundDevice
                {
                    Caption = GetPropertyString(mo["Caption"]),
                    Description = GetPropertyString(mo["Description"]),
                    Manufacturer = GetPropertyString(mo["Manufacturer"]),
                    Name = GetPropertyString(mo["Name"]),
                    ProductName = GetPropertyString(mo["ProductName"])
                };

                soundDeviceList.Add(soundDevice);
            }

            return soundDeviceList;
        }

        public List<VideoController> GetVideoControllerList()
        {
            List<VideoController> videoControllerList = new List<VideoController>();

            string queryString = UseAsteriskInWMI ? "SELECT * FROM Win32_VideoController"
                                                  : "SELECT AdapterCompatibility, AdapterRAM, Caption, CurrentBitsPerPixel, CurrentHorizontalResolution, CurrentNumberOfColors, CurrentRefreshRate, CurrentVerticalResolution, Description, DriverDate, DriverVersion, MaxRefreshRate, MinRefreshRate, Name, PNPDeviceID, VideoModeDescription, VideoProcessor FROM Win32_VideoController";

            using WmiConnection con = new WmiConnection(_managementScope);

            var mos = con.CreateQuery(queryString, _enumerationOptions);

            foreach (WmiObject mo in mos)
            {
                VideoController videoController = new VideoController
                {
                    Manufacturer = GetPropertyString(mo["AdapterCompatibility"]),
                    AdapterRAM = GetPropertyValue<uint>(mo["AdapterRAM"]),
                    Caption = GetPropertyString(mo["Caption"]),
                    CurrentBitsPerPixel = GetPropertyValue<uint>(mo["CurrentBitsPerPixel"]),
                    CurrentHorizontalResolution = GetPropertyValue<uint>(mo["CurrentHorizontalResolution"]),
                    CurrentNumberOfColors = GetPropertyValue<ulong>(mo["CurrentNumberOfColors"]),
                    CurrentRefreshRate = GetPropertyValue<uint>(mo["CurrentRefreshRate"]),
                    CurrentVerticalResolution = GetPropertyValue<uint>(mo["CurrentVerticalResolution"]),
                    Description = GetPropertyString(mo["Description"]),
                    DriverDate = GetPropertyString(mo["DriverDate"]),
                    DriverVersion = GetPropertyString(mo["DriverVersion"]),
                    MaxRefreshRate = GetPropertyValue<uint>(mo["MaxRefreshRate"]),
                    MinRefreshRate = GetPropertyValue<uint>(mo["MinRefreshRate"]),
                    Name = GetPropertyString(mo["Name"]),
                    VideoModeDescription = GetPropertyString(mo["VideoModeDescription"]),
                    VideoProcessor = GetPropertyString(mo["VideoProcessor"])
                };

                try
                {
                    string deviceID = GetPropertyString(mo["PNPDeviceID"]);

                    if (string.IsNullOrEmpty(deviceID))
                        continue;

                    object? driverObject = Microsoft.Win32.Registry.GetValue(@$"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\{deviceID}", "Driver", default(string));

                    if (driverObject is string driver && !string.IsNullOrEmpty(driver))
                    {
                        object? qwMemorySizeObject = Microsoft.Win32.Registry.GetValue(@$"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{driver}", "HardwareInformation.qwMemorySize", default(long));

                        if (qwMemorySizeObject is long qwMemorySize && qwMemorySize != 0L)
                        {
                            videoController.AdapterRAM = (ulong)qwMemorySize;
                        }
                    }
                }
                catch (SecurityException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }

                videoControllerList.Add(videoController);
            }

            return videoControllerList;
        }
    }
}
