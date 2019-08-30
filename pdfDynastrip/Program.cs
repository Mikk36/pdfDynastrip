using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using Acrobat;

namespace pdfDynastrip
{
    class Program
    {
        private static string inputPath = null;
        private static string outputPath = null;
        private static string copyPath = null;
        private static string archivePath = null;

        private static ManualResetEvent _quitEvent = new ManualResetEvent(false);

        private static AcroAVDoc g_AVDoc = null;

        private static Regex pageNameRegex = new Regex(@".*?(?=:|$)", RegexOptions.IgnoreCase);

        static void Main(string[] args)
        {
            if (CheckHelp(args))
            {
                return;
            }

            try
            {
                SetupPaths(args);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }

            Console.WriteLine("Watching folder: {0}", inputPath);
            Console.WriteLine("Output folder: {0}", outputPath);
            Console.WriteLine("Copy folder: {0}", copyPath);
            if (archivePath != null)
            {
                Console.WriteLine("Archive folder: {0}", archivePath);
            }

            FileSystemWatcher watcher = new FileSystemWatcher(inputPath, "*.pdf");

            Console.CancelKeyPress += (sender, eArgs) =>
            {
                _quitEvent.Set();
                eArgs.Cancel = true;
                watcher.Dispose();
                Console.WriteLine("Shutting down");
            };

            watcher.Created += new FileSystemEventHandler(NewFileHandler);
            watcher.EnableRaisingEvents = true;

            ParseFolder().ForEach(ProcessFile);

            _quitEvent.WaitOne();
        }

        private static bool CheckHelp(string[] args)
        {
            if (args.Length > 0 && "/?".Equals(args[0]))
            {
                Console.WriteLine("Default input path is working directory");
                Console.WriteLine("Output is \"out\" directory inside working directory");
                Console.WriteLine("Copy path is \"copy\" directory inside working directory");
                Console.WriteLine("First parameter: input path");
                Console.WriteLine("Second parameter: output path");
                Console.WriteLine("Third parameter: copy path");
                Console.WriteLine("Fourth parameter: archive path, optional");
                return true;
            }
            return false;
        }

        private static void SetupPaths(string[] args)
        {
            if (args.Length > 0)
            {
                inputPath = args[0];
                if (!Path.IsPathRooted(inputPath))
                {
                    inputPath = Path.GetFullPath(inputPath);
                }
            }
            else
            {
                inputPath = Path.GetFullPath(".");
            }

            if (args.Length > 1)
            {
                outputPath = args[1];
                if (!Path.IsPathRooted(outputPath))
                {
                    outputPath = Path.GetFullPath(outputPath);
                }
            }
            else
            {
                outputPath = inputPath + @"\out";
            }

            if (args.Length > 2)
            {
                copyPath = args[2];
                if (!Path.IsPathRooted(copyPath))
                {
                    copyPath = Path.GetFullPath(copyPath);
                }
            }
            else
            {
                copyPath = inputPath + @"\copy";
            }

            if (args.Length > 3)
            {
                archivePath = args[3];
                if (!Path.IsPathRooted(archivePath))
                {
                    archivePath = Path.GetFullPath(archivePath);
                }
            }

            inputPath += @"\";
            outputPath += @"\";
            copyPath += @"\";
            if (archivePath != null)
            {
                archivePath += @"\";
            }

            if (!Directory.Exists(inputPath) || !Directory.Exists(outputPath))
            {
                throw new Exception("Some of the paths does not exist");
            }
        }

        private static void NewFileHandler(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine("{2}: {0}: {1}", e.ChangeType, e.Name, DateTime.Now);
            while (IsFileLocked(e.FullPath))
            {
                Thread.Sleep(1000);
            }
            Console.WriteLine("{0}: Unlocked: {1}", DateTime.Now, e.Name);
            ProcessFile(e.FullPath);
        }

        private static bool IsFileLocked(string file)
        {
            FileInfo info = new FileInfo(file);
            FileStream stream = null;
            try
            {
                stream = info.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException e)
            {

                return true;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }
            return false;
        }

        private static void ProcessFile(string file)
        {
            if (g_AVDoc != null)
            {
                g_AVDoc.Close(0);
            }
            g_AVDoc = new AcroAVDoc();
            g_AVDoc.Open(file, "");
            while (!g_AVDoc.IsValid())
            {
                Thread.Sleep(1000);
            }
            if (g_AVDoc.IsValid())
            {
                ExportPages();

                g_AVDoc.Close(1);
            }
            try
            {
                Console.WriteLine("{0}: Copying file", DateTime.Now);
                File.Copy(file, copyPath + Path.GetFileName(file));
                if(archivePath != null)
                {
                    File.Copy(file, archivePath + Path.GetFileName(file));
                }
                File.Delete(file);
                Console.WriteLine("{0}: File deleted", DateTime.Now);
            }
            catch (IOException e)
            {
                Console.WriteLine("Error copying/deleting file {0}: {1}", file, e.Message);
            }
        }

        private static void ExportPages()
        {
            CAcroPDDoc pdDoc = (CAcroPDDoc)g_AVDoc.GetPDDoc();
            Object jsObj = pdDoc.GetJSObject();
            Type T = jsObj.GetType();

            int pages = GetPageCount();
            Console.WriteLine("The file {0} contains {1} pages", GetFileName(), pages);

            for (int i = 0; i < pages; i++)
            {
                Console.WriteLine("Page {0} name: {1}", i + 1, GetPageName(i));
                SavePage(i, GetPageName(i), GetFileName());
            }
        }

        private static void SavePage(int page, string pageName, string fileName)
        {
            CAcroPDDoc pdDoc = (CAcroPDDoc)g_AVDoc.GetPDDoc();
            Object jsObj = pdDoc.GetJSObject();
            Type T = jsObj.GetType();

            object[] parameters = {
                page,
                page,
                outputPath + fileName + pageName + ".pdf"
            };
            T.InvokeMember(
                "extractPages",
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                jsObj,
                parameters
                );
        }

        private static string GetFileName()
        {
            CAcroPDDoc pdDoc = (CAcroPDDoc)g_AVDoc.GetPDDoc();
            Object jsObj = pdDoc.GetJSObject();
            Type T = jsObj.GetType();

            return ((string)T.InvokeMember(
                "documentFileName",
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                jsObj,
                null
                )).Replace(".pdf", "");
        }

        private static string GetPageName(int page)
        {
            CAcroPDDoc pdDoc = (CAcroPDDoc)g_AVDoc.GetPDDoc();
            Object jsObj = pdDoc.GetJSObject();
            Type T = jsObj.GetType();

            object[] parameters = { page };
            string fullName = (string)T.InvokeMember(
                "getPageLabel",
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                jsObj,
                parameters
                );

            Match match = pageNameRegex.Match(fullName);
            if (match.Success)
            {
                return match.Value;
            }

            return fullName;
        }

        private static int GetPageCount()
        {
            CAcroPDDoc pdDoc = (CAcroPDDoc)g_AVDoc.GetPDDoc();
            Object jsObj = pdDoc.GetJSObject();
            Type T = jsObj.GetType();

            return Convert.ToInt32((double)T.InvokeMember(
                "numPages",
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                jsObj,
                null
                ));
        }

        private static List<string> ParseFolder()
        {
            List<string> files = null;
            try
            {
                files = new List<string>(Directory.EnumerateFiles(inputPath, "*.pdf"));
                files.ForEach(delegate (string name)
                {
                    Console.WriteLine(name);
                });
            }
            catch (UnauthorizedAccessException UAEx)
            {
                Console.WriteLine(UAEx.Message);
            }
            catch (PathTooLongException PathEx)
            {
                Console.WriteLine(PathEx.Message);
            }

            return files;
        }
    }
}
