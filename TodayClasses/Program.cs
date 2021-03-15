using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Shell32;
using IWshRuntimeLibrary;
using System.Threading;
using Microsoft.Win32;

namespace TodayClasses
{
    class Program
    {
        static string applicationLocation { get => (new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)).Directory.FullName; }
        static string globalShortcutURL { get => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Classes.lnk"); }
        static string currentClassesFolderURL { get => Path.Combine(applicationLocation, "CurrentClasses\\"); }
        static string classesListURL { get => Path.Combine(applicationLocation, "CurrentClasses\\My classes.txt"); }

        static Task midnightWait;

        static void Main(string[] args)
        {
            CheckEnv();

            FileSystemWatcher fs = new FileSystemWatcher();
            fs.Path = currentClassesFolderURL;
            fs.Filter = "My Classes.txt";
            fs.EnableRaisingEvents = true;
            fs.Changed += new FileSystemEventHandler(OnChanged);
            SystemEvents.PowerModeChanged += (PowerModeChangedEventHandler)((object s, PowerModeChangedEventArgs e) => 
            {
                Console.WriteLine("Power mode changed");
                Check();
            });

            midnightWait = new Task((Action)(() => Run()));
            midnightWait.Start();
            midnightWait.Wait();
        }
        private static void OnChanged(object source, FileSystemEventArgs e)
        {
            Console.WriteLine("Classes modified");
            Thread.Sleep(new TimeSpan(0, 0, 2));
            Check();
        }

        static void CheckEnv()
        {
            if (!CurrentClassesFolderExists()) CreateCurrentClassesFolder();
            if (!GlobalShortcutExists()) CreateGlobalShortcut();
            if (!MyClassesExists()) CreateMyClassesFile();
        }

        static void Check()
        {
            Console.WriteLine("Updating...");
            CheckEnv();
            UpdateCurrentClasses();
        }

        static void Run()
        {
            Check();

            while(true)
            {
                Console.WriteLine("Time to wait : " + (60 - DateTime.Now.TimeOfDay.Minutes) + "m");
                Thread.Sleep(new TimeSpan(0, 0, 60 - DateTime.Now.TimeOfDay.Seconds));
                Check();
            }
        }

        static void CreateShortcut(string origin, string target)
        {
            var wsh = new IWshShell_Class();
            IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(origin) as IWshRuntimeLibrary.IWshShortcut;
            shortcut.TargetPath = target;
            shortcut.Save();
        }

        static bool MyClassesExists()
        {
            return System.IO.File.Exists(classesListURL);
        }

        static void CreateMyClassesFile()
        {
            var sw = System.IO.File.CreateText(classesListURL);
            sw.WriteLine("# File containing the classes informations");
            sw.WriteLine("# Lines starting by '#' will be ignored");
            sw.WriteLine("# ");
            sw.WriteLine("# To add a class, add a line as following :");
            sw.WriteLine("# 'Monday=C:\\My classes\\English\\'");
            sw.Flush();
            sw.Close();
        }

        static string[] ClassesOfTheDay()
        {
            List<string> classes = new List<string>();

            string day = System.DateTime.Now.DayOfWeek.ToString();
            string[] lines = System.IO.File.ReadAllLines(classesListURL);

            foreach (string line in lines)
            {
                if (line == "") continue;
                else if (line[0] == '#')
                {
                    continue;
                }
                else if (line.Contains("="))
                {
                    if (line.Split('=')[0] == day)
                    {
                        string directory = line.Split('=')[1];
                        if (System.IO.Directory.Exists(directory))
                        {
                            classes.Add(directory);
                        }
                        else Console.WriteLine("Directory \"" + directory + "\" doesn't exist, cannot create shortcut !");
                    }
                }
            }

            return classes.ToArray();
        }

        static string[] ToShortcuts(string[] arr)
        {
            string[] shortcuts = new string[arr.Length];
            for(int i = 0; i < arr.Length; i++)
            {
                shortcuts[i] = ClassShortcutPath(arr[i]);
            }
            return shortcuts;
        }

        static void UpdateCurrentClasses()
        {
            string[] currentClasses = ClassesOfTheDay();
            string[] currentClassesShortcuts = ToShortcuts(currentClasses);
            List<string> classesAlreadyPresent = new List<string>();


            foreach(string file in System.IO.Directory.GetFiles(currentClassesFolderURL))
            {
                FileInfo fi = new FileInfo(file);
                if(fi.Extension == "lnk" || fi.Extension == ".lnk")
                {
                    if(!currentClassesShortcuts.Contains(file))
                    {
                        System.IO.File.Delete(file);
                    }
                    else
                    {
                        classesAlreadyPresent.Add(file);
                    }
                }
            }
            for(int i = 0; i < currentClassesShortcuts.Length; i++)
            {
                string file = currentClassesShortcuts[i];
                if (classesAlreadyPresent.Contains(file)) continue;
                else
                {
                    string path = currentClasses[i];
                    CreateShortcut(file, path);
                    classesAlreadyPresent.Add(file);
                }
            }
        }

        static string ClassShortcutPath(string myClass)
        {
            FileInfo fi = new FileInfo(myClass);
            return Path.Combine(currentClassesFolderURL, fi.Name + ".lnk");
        }

        static bool CurrentClassesFolderExists()
        {
            return System.IO.File.Exists(currentClassesFolderURL);
        }

        static void CreateCurrentClassesFolder()
        {
            try
            {
                System.IO.Directory.CreateDirectory(currentClassesFolderURL);
            }
            catch(Exception e)
            {
                Console.WriteLine("Current classes folder already exists");
            }
        }

        static bool GlobalShortcutExists()
        {
            return System.IO.File.Exists(globalShortcutURL);
        }

        static void CreateGlobalShortcut()
        {
            CreateShortcut(globalShortcutURL, currentClassesFolderURL);
        }
    }
}
