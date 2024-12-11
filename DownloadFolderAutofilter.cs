using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace DownloadSorter
{
    public partial class DownloadFolderAutofilter : ServiceBase
    {
        private System.Timers.Timer timer;
        private EventLog eventLog;    // for debuging purposes


        public DownloadFolderAutofilter()
        {
            InitializeComponent();

            eventLog = new EventLog();                                                            // for debuging purposes
            if (!EventLog.SourceExists("DownloadFolderAutofilterSource"))
            {
                EventLog.CreateEventSource("DownloadFolderAutofilterSource", "Application");
            }
            eventLog.Source = "DownloadFolderAutofilterSource";
            eventLog.Log = "Application";
        }


        private string userName;
        protected override void OnStart(string[] args)
        {
            timer = new System.Timers.Timer(1000D); // Set timer interval 
            timer.Elapsed += new ElapsedEventHandler(Process); // Attach event handler
            timer.Start();

            // for debuging purposes
            if (timer.Enabled)                
            {
                eventLog.WriteEntry("Timer started successfully.", EventLogEntryType.Information);
            }
            else
            {
                eventLog.WriteEntry("Timer failed to start.", EventLogEntryType.Error);
            }
            userName = getUserName();
            CreateSettigns();
            InsertSettings();
            // Keep the service running
            Task.Run(() => KeepServiceRunning());
            base.OnStart(args);
        }

        private void CreateSettigns()
        {
            if (!Directory.Exists($"C:\\Users\\{userName}\\Downloads\\Settings"))
            {
                Directory.CreateDirectory($"C:\\Users\\{userName}\\Downloads\\Settings");
            }
            if (!File.Exists($"C:\\Users\\{userName}\\Downloads\\Settings\\Settings.txt"))
            {
                string[] lines = 
                    {
                        "#######################################################################################################################################################################################################",
                        "# <- this is comment character, if you want to put notes or something MAKE SURE you put # as first character (without space on start), code will ignore that line and will not work with it           #",
                        "#   -Make sure before EACH edit you Uninstall service, edit settings and then install again because this txt is run once at start                                                                     #",
                        "#   -Strictly follow format as it is written here, if you want to put some custom extensions like \"*.PARTY.PNG\" make sure it is above folder with extension PNG otherwise it will go to PNG folder    #",
                        "#    because it scans if file contains extension and if it does it automatically sends it to that folder ignoring any other custom extensions after.                                                  #",
                        "#   -Below this line you can put your custom Extensions or edit whatever you want even name, make sure they are above original one.                                                                   #",
                        "#######################################################################################################################################################################################################",
                        "Images;*.JPG,*.JPEG,*.PNG,*.GIF,*.BMP,*.TIFF,*.TIF,*.WEBP,*.SVG,*.ICO",
                        "Audios;*.MP3,*.WAV,*.FLAC,*.AAC,*.OGG,*.WMA,*.M4A,*.AIFF",
                        "Videos;*.MP4,*.AVI,*.MKV,*.MOV,*.WMV,*.FLV,*.WEBM,*.MPEG,*.MPG,*.3GP",
                        "Documents;*.DOC,*.DOCX,*.PDF,*.TXT,*.RTF,*.ODT,*.XLS,*.XLSX,*.CSV,*.PPT,*.PPTX",
                        "Compressed Files;*.ZIP,*.RAR,*.7Z,*.TAR,*.GZ,*.BZ2",
                        "Exe and Scripts;*.EXE,*.DLL,*.MSI,*.BAT,*.CMD,*.SH,*.PY,*.RB,*.JS,*.JAR",
                        "Code Markups;*.HTML,*.HTM,*.CSS,*.JS,*.PHP,*.ASP,*.ASPX,*.CS,*.C,*.CPP,*.H,*.JAVA,*.PY,*.RB,*.GO,*.SWIFT,*.TS,*.XML,*.JSON,*.YML,*.MD",
                        "Disk Images;*.ISO,*.IMG,*.BIN,*.NRG,*.DMG",
                        "Fonts;*.TTF,*.OTF,*.FNT,*.WOFF,*.WOFF2",
                        "3D and CAD;*.OBJ,*.STL,*.FBX,*.DAE,*.BLEND,*.3DS,*.DXF",
                        "Miscs;*.LOG,*.CFG,*.INI,*.DAT,*.BAK"
                    };
                File.WriteAllLines($"C:\\Users\\{userName}\\Downloads\\Settings\\Settings.txt", lines);

            }
        }

        private void InsertSettings()
        {
            string[] files = File.ReadAllLines($"C:\\Users\\{userName}\\Downloads\\Settings\\Settings.txt");
            eventLog.WriteEntry(files[0], EventLogEntryType.Error);
            foreach (string line in files)
            {
                char c = line[0];
                if (line[0].Equals('#'))
                {
                    Console.WriteLine("Comment registered");
                }
                else
                {
                    string folder = line.Split(';').First();
                    string[] exts = line.Split(';').Last().Split(',');
                    FileTypes.Add(folder);
                    AllExtensionsHere.Add(exts);
                }

            }
        }
        

        private void KeepServiceRunning()
        {
            // You could do something like this to keep the service alive in the background
            while (true)
            {
                // Sleep for a while, letting other threads (like the timer) run.
                Task.Delay(1000).Wait();
            }
        }

        protected override void OnStop()
        {
            if (timer != null && timer.Enabled)
            {
                timer.Stop();
                eventLog.WriteEntry("Timer stopped.", EventLogEntryType.Information); // for debuging purposes
            }
            timer.Stop();
            base.OnStop();
        }
        public static List<string> FileTypes = new List<string>() ;

        public static List<string[]> AllExtensionsHere = new List<string[]>();
        private static string getUserName()
        {
            SelectQuery query = new SelectQuery(@"Select * from Win32_Process");
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
            {
                foreach (System.Management.ManagementObject Process in searcher.Get())
                {
                    if (Process["ExecutablePath"] != null && string.Equals(Path.GetFileName(Process["ExecutablePath"].ToString()), "explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        string[] OwnerInfo = new string[2];
                        Process.InvokeMethod("GetOwner", (object[])OwnerInfo);
                        return OwnerInfo[0];
                    }
                }
            }

            return "";
        }
        private void Process(object sender, ElapsedEventArgs e)
        {
            eventLog.WriteEntry("Timer ticked at Process at " + DateTime.Now, EventLogEntryType.Information); // for debuging purposes - LOGS
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();

            // Run the process logic asynchronously so the timer thread isn't blocked
            Task.Run(() => Starts());

            // Ensure Process() doesn't block the timer thread
            Task.Run(() => FileProcessing());
        }

        private async Task FileProcessing()
        {
            //eventLog.WriteEntry("User is set to:" + userName, EventLogEntryType.Information); // for debuging purposes - LOGS
            string dir = $@"C:\Users\{userName}\Downloads\";
            var files = Directory.GetFiles(dir);
            /*eventLog.WriteEntry("FileProcessing at " + DateTime.Now, EventLogEntryType.Information);  // for debuging purposes - LOGS
            eventLog.WriteEntry("Directory set to: " + dir, EventLogEntryType.Information);*/
            foreach (string file in files)
            {
                string fileExtension = Path.GetExtension(file).ToUpper();
                string fileExtensionLowerCase = fileExtension.ToLower();
                string fileName = file.Split('\\').Last();
                //string fileName1 = fileName.Split('.').First();
                // Process the file, ensure it's non-blocking
                await Task.Run(() => MoveFileBasedOnType(file, fileName, fileExtension));
            }
        }
        private string RenameFile(string fileName, string Extension, string destination)
        {
            eventLog.WriteEntry("Ticked at RenameFile with results:", EventLogEntryType.Information);  // for debuging purposes - LOGS
            //eventLog.WriteEntry("FileName is: " + fileName + " Extension is: " + Extension, EventLogEntryType.Information);  // for debuging purposes - LOGS
            DateTime time = DateTime.Now;
            fileName = fileName.Replace(Extension.ToLower(), "");
            //eventLog.WriteEntry("new FileName is: " + fileName, EventLogEntryType.Information);  // for debuging purposes - LOGS
            fileName = fileName +"__" + time.ToString("dd") + "." + time.ToString("MM") + "." + time.ToString("yyyy") + "_" + time.ToString("HH") + "-" + time.ToString("mm") + "-" + time.ToString("ss") + Extension.ToLower();
            eventLog.WriteEntry("Final fileName is: " + fileName, EventLogEntryType.Information);  // for debuging purposes - LOGS
            return destination += fileName;
        }
        private void MoveFileBasedOnType(string file, string fileName, string fileExtension)
        {
            int FolderIndex = 0;
            //eventLog.WriteEntry("FileName origo: " + fileName + "FileName bez konce: " + fileName1 + "filename s datem a koncem: " + fileName2 , EventLogEntryType.Information);  // for debuging purposes - LOGS
            string StarAndfileExtension = fileExtension.Insert(0, "*");
            foreach (string[] AllExt in AllExtensionsHere)
            {
                foreach(string CusExt in AllExt)
                {
                    if (fileName.Contains(CusExt.Split('*').Last().ToLower()))
                    {
                        StarAndfileExtension = CusExt;
                    }
                }
                eventLog.WriteEntry("Current file extensions scanning: " + AllExt, EventLogEntryType.Information);  // for debuging purposes - LOGS
                eventLog.WriteEntry("Ext to scan: " + StarAndfileExtension + " Default Ext: " + fileExtension, EventLogEntryType.Information);  // for debuging purposes - LOGS
                eventLog.WriteEntry("AllExt Format Check: " + AllExt[0],EventLogEntryType.Information);  // for debuging purposes - LOGS
                if (AllExt.Contains(StarAndfileExtension) == true)
                {
                    eventLog.WriteEntry("If statement true", EventLogEntryType.Information);  // for debuging purposes - LOGS
                    foreach (string TypeExt in AllExt)
                    {
                        eventLog.WriteEntry("If statement: TypeExt is: " + TypeExt + "e fileExtension is: " + StarAndfileExtension + "e", EventLogEntryType.Information);  // for debuging purposes - LOGS
                        if (TypeExt.Equals(StarAndfileExtension))
                        {
                            eventLog.WriteEntry("If statement true", EventLogEntryType.Information);  // for debuging purposes - LOGS
                            string destination = $@"C:\Users\{userName}\Downloads\{FileTypes[FolderIndex]}\{fileName}";
                            eventLog.WriteEntry("Original destination is: " + destination, EventLogEntryType.Information);  // for debuging purposes - LOGS
                            if (File.Exists(destination))
                            {
                                destination = $@"C:\Users\{userName}\Downloads\{FileTypes[FolderIndex]}\";
                                destination = RenameFile(fileName, fileExtension, destination);
                                eventLog.WriteEntry("Modified destination is: " + destination, EventLogEntryType.Information);  // for debuging purposes - LOGS
                            }
                            eventLog.WriteEntry("File: " + file + " Destination: " + destination, EventLogEntryType.Information);  // for debuging purposes - LOGS
                            File.Move(file, destination);
                        }
                    }
                    
                }
                FolderIndex++;
            }
                
        }


        /*private void Process(object sender, ElapsedEventArgs e)    -- base kinda messy code which didnt work with service, but does work with console app if u want just click and sort but u will have to insert it into IDE w/ C# - P.S.: Not updated to current version of code
        {

            Starts();

            Console.WriteLine(userName);
            string dir = $@"C:\Users\{userName}\Downloads\";
            var files = Directory.GetFiles($@"C:\Users\{userName}\Downloads\");

            foreach (string file in files)
            {
                string fileExtension = Path.GetExtension(file).ToUpper();

                string fileName = file.Split('\\').Last();

                foreach (string imgExt in ImageExtensions)
                {
                    if (fileExtension == imgExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[0]}\{fileName}");
                    }
                }

                foreach (string audioExt in AudioExtensions)
                {
                    if (fileExtension == audioExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[1]}\{fileName}");
                    }
                }

                foreach (string videoExt in VideoExtensions)
                {
                    if (fileExtension == videoExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[2]}\{fileName}");
                    }
                }

                foreach (string docExt in DocumentExtensions)
                {
                    if (fileExtension == docExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[3]}\{fileName}");
                    }
                }

                foreach (string compressExt in CompressExtensions)
                {
                    if (fileExtension == compressExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[4]}\{fileName}");
                    }
                }

                foreach (string exeScriptExt in Exe_ScriptExtensions)
                {
                    if (fileExtension == exeScriptExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[5]}\{fileName}");
                    }
                }

                foreach (string codeMarkupExt in Code_MarkupExtensions)
                {
                    if (fileExtension == codeMarkupExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[6]}\{fileName}");
                    }
                }

                foreach (string diskImageExt in DiskImageExtensions)
                {
                    if (fileExtension == diskImageExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[7]}\{fileName}");
                    }
                }

                foreach (string fontExt in FontExtensions)
                {
                    if (fileExtension == fontExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[8]}\{fileName}");
                    }
                }

                foreach (string threeDExt in ThreeD_CADExtensions)
                {
                    if (fileExtension == threeDExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[9]}\{fileName}");
                    }
                }

                foreach (string miscExt in MiscExtensions)
                {
                    if (fileExtension == miscExt.TrimStart('*').ToUpper())
                    {
                        File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[10]}\{fileName}");
                    }
                }
            }
            files = Directory.GetFiles($@"C:\Users\{userName}\Downloads\");
            foreach (string file in files)
            {
                string fileName = file.Split('\\').Last();
                File.Move(file, $@"C:\Users\{userName}\Downloads\{FileTypes[11]}\{fileName}");
            }
        }*/

        private void Starts()
        {

            foreach (string type in FileTypes)
            {
                string folderPath = $@"C:\Users\{userName}\Downloads\{type}";
                if (!Directory.Exists(folderPath))
                {
                    // Create the folder if it doesn't exist
                    Directory.CreateDirectory(folderPath);
                }
            }

        }
    }
}
