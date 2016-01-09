using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Extensibility;

namespace ConsoleApplication1
{

    class Program
    {
        
        // totally messed up code , but its working as for POC
        //copied stuff from mostafa

        //constracting the location of the log file to be zipped
        static string path1 = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        //below is path for logs files used by Office 16.0 (SkypeFB) change the 16.0 to 15.0 for (lync)
        static string path2_16 = @"AppData\Local\Microsoft\Office\16.0\Lync\Tracing\";
        static string path2_15 = @"AppData\Local\Microsoft\Office\15.0\Lync\Tracing\";
        static string _logLocation_16 = Path.Combine(path1, path2_16);
        static string _logLocation_15 = Path.Combine(path1, path2_15);
        //location where i want to save the zipped file
        static string _folderToZip = @"c:\temp\tempSkype4b\logs";
        static string _zippedLogs = @"c:\temp\tempSkype4b\Skype4b_logs.zip";
        static string _lyncLogCopy = @"c:\temp\tempSkype4b\logs\lynclogfile.uccapilog";
        


         static void Main(string[] args)
        {

            if (Directory.Exists(_logLocation_16))
            {

                string[] _getLogFiles_16 = Directory.GetFiles(_logLocation_16, "*.uccapilog");
                CopyLogFile(_getLogFiles_16);
            }


            else if (Directory.Exists(_logLocation_15))
            {
                string[] _getLogFiles_15 = Directory.GetFiles(_logLocation_15, "*.uccapilog");
                CopyLogFile(_getLogFiles_15);
            }
            
            string s2 = null;
            if (args.Length > 0)
            {
                Console.WriteLine("User: {0}", args[0]);
            }
            if (args.Length > 1)
            {
                string s = args[1].ToString().Split(':')[1];
                int i = s.Length;
                s2 = s.Substring(0, i - 1);
                Console.WriteLine("Contact: {0}", s2);
                Console.WriteLine("Contact: {0}", args[1]);
            }

            //
            try
            {
                // Create the major UI Automation objects.
                Automation _Automation = LyncClient.GetAutomation();

                // Create a dictionary object to contain AutomationModalitySettings data pairs.
                Dictionary<AutomationModalitySettings, object> _ModalitySettings = new Dictionary<AutomationModalitySettings, object>();

                AutomationModalities _ChosenMode = AutomationModalities.FileTransfer | AutomationModalities.InstantMessage;
                //AutomationModalities _ChosenMode =  AutomationModalities.InstantMessage| AutomationModalities.FileTransfer;

                // Store the file path as an object using the generic List class.
                string myFileTransferPath = string.Empty;
                // Edit this to provide a valid file path.
                myFileTransferPath = @"C:\Temp\tempSkype4b\Skype4b_logs.zip";

                // Create a generic List object to contain a contact URI.

                String[] invitees = {s2};
                //String[] invitees = { "jpoindexter@roseninspection.net" };



                // Adds text to toast and local user IMWindow text entry control.
                _ModalitySettings.Add(AutomationModalitySettings.FirstInstantMessage, "Hello attached you will get my Skype4B logfile");
                //_ModalitySettings.Add(AutomationModalitySettings.);
                _ModalitySettings.Add(AutomationModalitySettings.SendFirstInstantMessageImmediately, true);

                // Add file transfer conversation context type
                _ModalitySettings.Add(AutomationModalitySettings.FilePathToTransfer, myFileTransferPath);

                /*
                need to add something here to make a confirmation before sending the file
                */
                // Start the conversation.
                if (invitees != null)
                {
                    IAsyncResult ar = _Automation.BeginStartConversation(
                        _ChosenMode
                        , invitees
                        , _ModalitySettings
                        , null
                        , null);

                    // Block UI thread until conversation is started.
                    _Automation.EndStartConversation(ar);


                }
                //Console.ReadLine();
            }

            catch
            {
                Console.WriteLine("Error");
                Console.ReadLine();
            }
        }


         static void CopyLogFile(string[] args)
         {
             foreach (string file in args)
             {
                 //copy the file and wait for the copy to finish
                 Directory.Delete(@"C:\temp\tempSkype4b",true);
                 Directory.CreateDirectory(Path.GetDirectoryName(_lyncLogCopy));
                 Task CopyLogFile = Task.Factory.StartNew(() => File.Copy(file, _lyncLogCopy, true));
                 CopyLogFile.Wait();
                 //zip the copied log file
                 ZipFile.CreateFromDirectory(_folderToZip, _zippedLogs, CompressionLevel.Optimal, false);
             }


         }
    }
}