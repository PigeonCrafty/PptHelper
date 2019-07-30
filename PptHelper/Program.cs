using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace PptHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathInput;
            string langInput;
            
            Console.WriteLine(DateTime.Now + " " + "Developed by Moravia Publishing & Media team. All rights reserved.");
            Console.WriteLine("========================================");

            // 1. Input folder or file path
            Console.WriteLine(">>> Please input the complete Folder or single File Path: ");
            PathInput:
            pathInput = Console.ReadLine();
            if (string.IsNullOrEmpty(pathInput) && string.IsNullOrWhiteSpace(pathInput))
            {
                Console.WriteLine("<!> Empty or invalid directory, please re-enter: ");
                goto PathInput;
            }

            // 2. Input language name or code e.g 'Japanese' or 'Japan' or 'JAP'
            Console.WriteLine("\n"+ ">>> Please input Target Language Name or Code: ");
            LangInput:
            langInput = Console.ReadLine();
            if (string.IsNullOrEmpty(langInput) && string.IsNullOrWhiteSpace(langInput))
            {
                Console.WriteLine("<!> Unable to identify, please double check and re-enter: (e.g.\"Japanese\" or \"Jap\"): ");
                goto LangInput;
            }

            //TODO: detect language
            TargetLang = langInput;

            // Transfer input path to DirectoryInfo
            if (!File.Exists(pathInput) && Directory.Exists(pathInput))
            {
                if (!pathInput.EndsWith("\\")) pathInput += "\\";

                var dirInfo = new DirectoryInfo(pathInput);

                try
                {
                    SortLocFiles(dirInfo);
                }
                catch (IOException e)
                {
                    Console.WriteLine("Fail to sort files: " + e.Message);
                    return;
                }
                /*
                //Get all subfolder names and add into LocalLanguage List
                try
                {
                    ListLangFolders(dirInfo);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error when create folders list： " + ex.Message);
                    return;
                }
                foreach (var lang in ListLang)
                    try
                    {
                        var langDirInfo = new DirectoryInfo(pathInput + lang + "\\");
                        SortLocFiles(langDirInfo);
                    }
                    catch (IOException e)
                    {
                        Console.WriteLine("Fail to sort files: " + e.Message);
                        return;
                    }
                */

            }
            else if (File.Exists(pathInput))
            {
                var fileInfo = new FileInfo(pathInput);

                switch (fileInfo.Extension)
                {
                    
                    case ".pptx":
                    //case ".potx":
                    //case ".potm":
                        ListPpt.Add(fileInfo.FullName);
                        break;
                }
            }

            KillProcess();

            Console.WriteLine("========================================");
            Console.WriteLine(">>> Succeeded to load " + ListPpt.Count.ToString() + " files.");
            Console.WriteLine(">>> Please hold on until progress complete. It may take a while...");
            Console.WriteLine("========================================");

            // PowerPointHandler
            if (ListPpt.Count > 0)
                foreach (var f in ListPpt)
                {
                    if (f.Contains("~$")) continue;
                    Console.Write("\r\n" + "Processing: " + f);
                    Common.WriteLine("\r\n" + f + "\r\n");
                    var objPpt = new PowerPointHandler(f, TargetLang);
                    objPpt.PptMain(f);
                    Common.WriteSglText(f);
                }

            // Write all Console content to text
            Common.WriteAllText(pathInput);

            // End
            Console.WriteLine("========================================");
            Console.WriteLine("All files process complete! You are good to go!");
            Console.ReadLine();
        }

        public static void ListLangFolders(DirectoryInfo rootInput)
        {
            try
            {
                var folders = rootInput.GetDirectories(); // To get folder list
                foreach (var fd in folders)
                    ListLang.Add(fd.Name);
                // Console.WriteLine(fd.Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error!" + ex.Message);
                throw;
            }
        }

        public static void SortLocFiles(DirectoryInfo langFolder)
        {
            if (!langFolder.Exists) return;

            var files = langFolder.GetFiles();

            foreach (var fil in files)
                switch (fil.Extension)
                {
                    case ".potx":
                    case ".pptx":
                    case ".potm":
                        ListPpt.Add(fil.FullName);
                        // Console.WriteLine(fil.Directory.Name); // To get the file language folder name
                        break;
                }
        }
        public static void KillProcess()
        {
            if (Process.GetProcessesByName("POWERPNT").Any())
                foreach (var proc in Process.GetProcessesByName("POWERPNT"))
                    proc.Kill();

            //if (Process.GetProcessesByName("WINWORD").Any())
            //    foreach (var proc in Process.GetProcessesByName("WINWORD"))
            //        proc.Kill();

            //if (Process.GetProcessesByName("EXCEL").Any())
            //    foreach (var proc in Process.GetProcessesByName("EXCEL"))
            //        proc.Kill();
        }

        public static string TargetLang { get; set; }
        public static List<string> ListLang { get; set; } = new List<string>();
        //public static List<string> ListWord { get; set; } = new List<string>();
        public static List<string> ListPpt { get; set; } = new List<string>();
        //public static List<string> ListExcel { get; set; } = new List<string>();

    }
}
