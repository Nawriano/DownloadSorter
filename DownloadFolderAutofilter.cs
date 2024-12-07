using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
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
            timer = new System.Timers.Timer(1000); // Set timer interval (1 second)
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
            // Keep the service running
            Task.Run(() => KeepServiceRunning());
            base.OnStart(args);
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
        
        public static readonly List<string> FileTypes = new List<string>() { "Images", "Audios", "Videos", "Documents", "Compressed Files", "Exe and Scripts", "Code Markups", "Disk Images", "Fonts", "3D and CAD", "Miscs", "Others" };
        public static readonly List<string> ImageExtensions = new List<string> { "*.JPG", "*.JPEG", "*.PNG", "*.GIF", "*.BMP", "*.TIFF", "*.TIF", "*.WEBP", "*.SVG", "*.ICO" };
        public static readonly List<string> AudioExtensions = new List<string> { "*.MP3", "*.WAV", "*.FLAC", "*.AAC", "*.OGG", "*.WMA", "*.M4A", "*.AIFF" };
        public static readonly List<string> VideoExtensions = new List<string> { "*.MP4", "*.AVI", "*.MKV", "*.MOV", "*.WMV", "*.FLV", "*.WEBM", "*.MPEG", "*.MPG", "*.3GP" };
        public static readonly List<string> DocumentExtensions = new List<string> { "*.DOC", "*.DOCX", "*.PDF", "*.TXT", "*.RTF", "*.ODT", "*.XLS", "*.XLSX", "*.CSV", "*.PPT", "*.PPTX" };
        public static readonly List<string> CompressExtensions = new List<string> { "*.ZIP", "*.RAR", "*.7Z", "*.TAR", "*.GZ", "*.BZ2" };
        public static readonly List<string> Exe_ScriptExtensions = new List<string> { "*.EXE", "*.DLL", "*.MSI", "*.BAT", "*.CMD", "*.SH", "*.PY", "*.RB", "*.JS", "*.JAR" };
        public static readonly List<string> Code_MarkupExtensions = new List<string> { "*.HTML", "*.HTM", "*.CSS", "*.JS", "*.PHP", "*.ASP", "*.ASPX", "*.CS", "*.C", "*.CPP", "*.H", "*.JAVA", "*.PY", "*.RB", "*.GO", "*.SWIFT", "*.TS", "*.XML", "*.JSON", "*.YML", "*.MD" };
        public static readonly List<string> DiskImageExtensions = new List<string> { "*.ISO", "*.IMG", "*.BIN", "*.NRG", "*.DMG" };
        public static readonly List<string> FontExtensions = new List<string> { "*.TTF", "*.OTF", "*.FNT", "*.WOFF", "*.WOFF2" };
        public static readonly List<string> ThreeD_CADExtensions = new List<string> { "*.OBJ", "*.STL", "*.FBX", "*.DAE", "*.BLEND", "*.3DS", "*.DXF" };
        public static readonly List<string> MiscExtensions = new List<string> { "*.LOG", "*.CFG", "*.INI", "*.DAT", "*.BAK" };

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
            //eventLog.WriteEntry("Timer ticked at Process at " + DateTime.Now, EventLogEntryType.Information); // for debuging purposes
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();

            // Run the process logic asynchronously so the timer thread isn't blocked
            Task.Run(() => Starts());

            // Ensure Process() doesn't block the timer thread
            Task.Run(() => FileProcessing());
        }

        private async Task FileProcessing()
        {
            //eventLog.WriteEntry("User is set to:" + userName, EventLogEntryType.Information); // for debuging purposes
            string dir = $@"C:\Users\{userName}\Downloads\";
            var files = Directory.GetFiles(dir);
            /*eventLog.WriteEntry("FileProcessing at " + DateTime.Now, EventLogEntryType.Information);  // for debuging purposes
            eventLog.WriteEntry("Directory set to: " + dir, EventLogEntryType.Information);*/
            foreach (string file in files)
            {
                string fileExtension = Path.GetExtension(file).ToUpper();
                string fileName = file.Split('\\').Last();
                string fileName1 = fileName.Split('.').First();
                DateTime time = DateTime.Now;
                string fileName2 = fileName1 + '_' + time.ToString("dd") + "_" + time.ToString("MM") + "_" + time.ToString("yyyy") + "_" + time.ToString("HH") + "_" + time.ToString("mm") + "_" + time.ToString("ss") + fileExtension.ToLower();


                // Process the file, ensure it's non-blocking
                await Task.Run(() => MoveFileBasedOnType(file, fileName2, fileExtension, fileName1, fileName2));
            }
        }

        private void MoveFileBasedOnType(string file, string fileName, string fileExtension, string fileName1, string fileName2)
        {
            //eventLog.WriteEntry("FileName origo: " + fileName + "FileName bez konce: " + fileName1 + "filename s datem a koncem: " + fileName2 , EventLogEntryType.Information);  // for debuging purposes
            foreach (string imgExt in ImageExtensions)
            {
                if (fileExtension == imgExt.TrimStart('*').ToUpper())
                {
                    string destination = $@"C:\Users\{userName}\Downloads\{FileTypes[0]}\{fileName}";
                    //eventLog.WriteEntry("File: "+ file + " Destination: " + destination, EventLogEntryType.Information);  // for debuging purposes
                    File.Move(file, destination);
                }
            }
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


        /*private void Process(object sender, ElapsedEventArgs e)    -- base kinda messy code which didnt work with service, but does work with console app
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
