﻿using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Win32;
using System.Xml.Linq;
using System.Linq;

namespace os_collect_stats_win
{
    class Program
    {
        private static string _tempFolderPath = Path.Combine(Directory.GetCurrentDirectory(),"collect_stats"); 
        private static string _targetZipFile = Path.Combine(Directory.GetCurrentDirectory(), "outsystems_data_" + DateTimeToTimestamp(DateTime.Now) + ".zip");
        private static string _osInstallationFolder = @"c:\Program Files\OutSystems\Platform Server";
        //private static string _osLogFolder = Path.Combine(_osInstallationFolder, "logs");
        private static string _osServerRegistry = @"SOFTWARE\OutSystems\Installer\Server";
        private static string _SSLProtocolsRegistryPath = @"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\SecurityProviders\Schannel\Protocols";
        private static string _IISRegistryPath = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\InetStp";
        private static string _NetFrameworkRegistryPath = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP";
        private static string _OutSystemsPlatformRegistryPath = @"HKEY_LOCAL_MACHINE\SOFTWARE\OutSystems";
        private static string _IISApplicationHostPath = @"C:\Windows\System32\inetsrv\config\applicationHost.config";


        static void Main(string[] args)
        {
            // Change console encoding to support all characters
            Console.OutputEncoding = Encoding.UTF8;

            // Initialize helper classes
            FileSystemHelper fsHelper = new FileSystemHelper();
            CmdHelper cmdHelper = new CmdHelper();
            WindowsEventLogHelper welHelper = new WindowsEventLogHelper();

            // Finding Installation folder
            try
            {
                Console.WriteLine("Finding OutSystems Platform Installation Path...");
                RegistryKey OSPlatformInstaller = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(_osServerRegistry);
                
                _osInstallationFolder = (string) OSPlatformInstaller.GetValue("");
                Console.WriteLine("Found it on: \"{0}\"", _osInstallationFolder);
            }
            catch (Exception e)
            {
                Console.WriteLine(" * Unable to find OutSystems Platform Server Installation... * ");
                Console.WriteLine(e.ToString());
                WriteExitLines();
                return;
            }

            Object obj = Registry.GetRegistryValue(_osServerRegistry, ""); // The "Defaut" values are empty strings.

            // Delete temporary directory and all contents if it already exists (e.g.: error runs)
            if (Directory.Exists(_tempFolderPath))
            {
                Directory.Delete(_tempFolderPath, true);
            }

            // Create temporary directory
            Directory.CreateDirectory(_tempFolderPath);

            // Process copy files
            CopyAllFiles();

            // Generate Event Viewer Logs
            Console.Write("Generating log files... ");
            welHelper.GenerateLogFiles(_tempFolderPath);
            Console.WriteLine("DONE");

            //TODO: Get IIS access logs -- WORK IN PROGRESS
            try
            {
                // Loading Xml text from the file. Note: Build must be in 64bit. More info: https://stackoverflow.com/questions/38316427/filenotfound-when-reading-the-iis-applicationhost-config-file
                var xmlString = XDocument.Load(_IISApplicationHostPath);

                // Querying the data and finding the Access logs path
                var query = from p in xmlString.Descendants("siteDefaults")
                            select new
                            {
                                LogsFilePath = p.Element("logFile").Attribute("directory").Value,
                            };
                
                string iisAccessLogsPath = query.First().LogsFilePath;

                // TODO: Copy all the contents from the path iisAcessLogsPath, including contents in subfolder 
                // Copying the contents from the Access logs path


                //---------------------------------------------------------------------------------
                Console.WriteLine("Retrieved IIS Access logs:");
            }
            catch (Exception e)
            {
                Console.WriteLine("Attempted to retrieve IIS Access logs but failed..." + e.Message);
            }

            ExecuteCommands();

            // Generate zip file
            Console.WriteLine();
            Console.Write("Creating zip file... ");
            fsHelper.CreateZipFromDirectory(_tempFolderPath, _targetZipFile, true);
            Console.WriteLine("DONE");

            // Delete temp folder
            Directory.Delete(_tempFolderPath, true);

            

            //TODO: memdump

            // Print process end
            PrintEnd();
            WriteExitLines();
        }

        // write a generic exit line and wait for user input
        private static void WriteExitLines()
        {
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        private static void CopyAllFiles()
        {
            // List of OS services and components
            IDictionary<string, string> osServiceNames = new Dictionary<string, string> {
                { "ConfigurationTool", "Configuration Tool" },
                { "LogServer", "Log Service" },
                { "CompilerService", "Deployment Controller Service" },
                { "DeployService", "Deployment Service" },
                { "Scheduler", "Scheduler Service" },
                { "SMSConnector", "SMS Service" }
            };

            // Initialize dictionary with all the files that we need to get and can be accessed directly
            IDictionary<string, string> files = new Dictionary<string, string> {
                { "ServerHSConf", Path.Combine(_osInstallationFolder, "server.hsconf") },
                { "OSVersion", Path.Combine(_osInstallationFolder, "version.txt") }
            };

            // Add OS log and configuration files
            foreach (KeyValuePair<string, string> serviceFileEntry in osServiceNames)
            {
                string confFilePath = Path.Combine(_osInstallationFolder, serviceFileEntry.Key + ".exe.config");

                // Get log file location from conf file
                OSServiceConfigFileParser confParser = new OSServiceConfigFileParser(serviceFileEntry.Value, confFilePath);
                string logPath = confParser.LogFilePath;

                // Add properties file
                files.Add(serviceFileEntry.Value + " config", confFilePath);

                // Add log file
                files.Add(serviceFileEntry.Value + " log", logPath);
            }

            // Copy all files to the temporary folder
            foreach (KeyValuePair<string, string> fileEntry in files)
            {
                String filepath = fileEntry.Value;
                String fileAlias = fileEntry.Key;

                Console.Write("Copying " + fileAlias + "... ");
                if (File.Exists(filepath))
                {
                    String realFilename = Path.GetFileName(filepath);
                    File.Copy(filepath, Path.Combine(_tempFolderPath, realFilename));

                    Console.WriteLine("DONE");
                }
                else
                {
                    Console.WriteLine("FAIL (File does not exist)");
                }
            }
        }

        private static void ExecuteCommands()
        {
            IDictionary<string, CmdLineCommand> commands = new Dictionary<string, CmdLineCommand>
            {
                { "dir_outsystems", new CmdLineCommand(string.Format("dir /s /a \"{0}\"", _osInstallationFolder),Path.Combine(_tempFolderPath, "dir_outsystems")) },
                { "tasklist", new CmdLineCommand("tasklist /v",Path.Combine(_tempFolderPath, "tasklist")) },
                { "cpu_info", new CmdLineCommand("wmic cpu",Path.Combine(_tempFolderPath, "cpu_info")) },
                { "memory_info", new CmdLineCommand("wmic memphysical",Path.Combine(_tempFolderPath, "mem_info")) },
                { "mem_cache", new CmdLineCommand("wmic memcache",Path.Combine(_tempFolderPath, "mem_cache")) },
                { "net_protocol", new CmdLineCommand("wmic netprotocol",Path.Combine(_tempFolderPath, "net_protocol")) },
                { "env_info", new CmdLineCommand("wmic environment",Path.Combine(_tempFolderPath, "env_info")) },
                { "os_info", new CmdLineCommand("wmic os",Path.Combine(_tempFolderPath, "os_info")) },
                { "pagefile", new CmdLineCommand("wmic pagefile",Path.Combine(_tempFolderPath, "pagefile")) },
                { "partition", new CmdLineCommand("wmic partition",Path.Combine(_tempFolderPath, "partition")) },
                { "startup", new CmdLineCommand("wmic startup",Path.Combine(_tempFolderPath, "startup")) },
                { "NetFramework", new CmdLineCommand("ExportRegistry.bat " + "\"" + _NetFrameworkRegistryPath + "\" " + Path.Combine(_tempFolderPath, "NetFrameworkVersion.txt")) },
                { "OutSystems_Info", new CmdLineCommand("ExportRegistry.bat " + "\"" + _OutSystemsPlatformRegistryPath + "\" " + Path.Combine(_tempFolderPath, "OutSystemsPlatform.txt")) },
                { "SSLProtocols", new CmdLineCommand("ExportRegistry.bat " + "\"" + _SSLProtocolsRegistryPath + "\" " + Path.Combine(_tempFolderPath, "SSLProtocols.txt")) },
                { "IISVersion", new CmdLineCommand("ExportRegistry.bat " + "\"" + _IISRegistryPath + "\" " + Path.Combine(_tempFolderPath, "IISVersion.txt")) }
            };

            foreach (KeyValuePair<string, CmdLineCommand> commandEntry in commands)
            {
                Console.Write("Getting {0}...", commandEntry.Key);
                commandEntry.Value.Execute();
                Console.WriteLine("DONE");
            }
        }

        private static void PrintEnd()
        {
            Console.WriteLine();
            Console.WriteLine("collect_stats has finished. Resulting zip file path:");
            Console.WriteLine(_targetZipFile);
            Console.WriteLine();
        }

        private static string DateTimeToTimestamp(DateTime dateTime)
        {
            return dateTime.ToString("yyyyMMdd_HHmm");
        }
    }
}