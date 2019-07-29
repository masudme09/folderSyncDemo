using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using System.Xml;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.Management.Automation;

namespace folderSyncDemo
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer(); // name space(using System.Timers;)
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 5000; //number in milisecinds  
            timer.Enabled = true;
        }
        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            //WriteToFile("Service is recall at " + DateTime.Now);
            List<string> fileInfo = new List<string>();
            fileInfo = getFileInfo(System.Configuration.ConfigurationManager.AppSettings.Get("testInputPath"));

            foreach (string a in fileInfo)
            {
                WriteToFile(a);

            }

            //xml file writer
            XmlDocument xmlDoc =new XmlDocument();
            XmlNode rootNode;
            if (!File.Exists(System.Configuration.ConfigurationManager.AppSettings.Get("logPath") + "\\LogXML.xml"))
            {
                // Create an XML declaration. 
                XmlDeclaration xmldecl;
                xmldecl = xmlDoc.CreateXmlDeclaration("1.0", null, null);
                xmldecl.Encoding = "UTF-8";
                xmldecl.Standalone = "yes";

                rootNode = xmlDoc.CreateElement("FileInfo");
                xmlDoc.AppendChild(rootNode);
            }
            else
            {
                xmlDoc.Load(System.Configuration.ConfigurationManager.AppSettings.Get("logPath") + "\\LogXML.xml");
                rootNode = xmlDoc.GetElementsByTagName("FileInfo")[0];
            }

            

            //Loop through each file info and add them to xml
            foreach (string a in fileInfo)
            {
                //splitString = ((a.Split('|')[1]).Split(':')[1]).Trim();//UserName
                //splitString = (a.Split('|')[0]).Trim();//Date and Time
                //splitString = ((a.Split('|')[2]).Split(':')[1]).Trim();//File Name
                //splitString = ((a.Split('|')[3]).Split(':')[1]).Trim();//File Size
                XmlNode fileNode = xmlDoc.CreateElement("Files");
                rootNode.AppendChild(fileNode);
                xmlNodeAppend(xmlDoc, "DateTime", "dateTime", (a.Split('|')[0]).Trim(), "Date Time", fileNode);
                xmlNodeAppend(xmlDoc, "UserDetails", "userName", ((a.Split('|')[1]).Split(':')[1]).Trim(), "User Name", fileNode);
                xmlNodeAppend(xmlDoc, "FileName", "fileName", ((a.Split('|')[2]).Split(':')[1]).Trim(), "File Name", fileNode);
                xmlNodeAppend(xmlDoc, "FileSize", "fileSize", ((a.Split('|')[3]).Split(':')[1]).Trim(), "File Size", fileNode);
                xmlDoc.Save(System.Configuration.ConfigurationManager.AppSettings.Get("logPath") + "\\LogXML.xml");
            }
            
            //move the files to destination
            moveFile(System.Configuration.ConfigurationManager.AppSettings.Get("testInputPath"), 
                System.Configuration.ConfigurationManager.AppSettings.Get("testInput1Path"));

            //Monitor the second file
            //WriteToFile(matchAndReturnUsernameFromXML(xmlDoc, "61 PC Pass.txt"));
            string CurrentUser = GetUsername(Process.GetCurrentProcess().SessionId);
            WriteToFile(CurrentUser);
        }
        

        public void copyFileFromTestInputToTestInput1(string testInputPath, string testInput1Path)
        {

        }

        //Get file infos from the directory
        public List<string> getFileInfo(string filePath)
        {
            long fileSizeOnDisc;
            string fileName;
            string userName, appendFileInfo;
            List<string> fileInfo=new List<string>();
            string[] filesOnDirectory = Directory.GetFiles(filePath);
            foreach (string a in filesOnDirectory)
            {
                fileSizeOnDisc = GetFileSizeOnDisk(a);//File size on the disk
                fileName = Path.GetFileName(a);
                userName = System.Environment.GetEnvironmentVariable("UserName");
                appendFileInfo = DateTime.Now + "\t|UserName:" + userName + "\t|File:" + fileName + "\t|Size: " + fileSizeOnDisc/1000 + "KB";
                fileInfo.Add(appendFileInfo);
            }

            return fileInfo;
        }

        public void moveFile(string srcFolderPath, string desFolderPath)
        {
            string[] filesOnDirectory = Directory.GetFiles(srcFolderPath);
            foreach (string a in filesOnDirectory)
            {
                File.Move(a, desFolderPath + "\\" + Path.GetFileName(a));
            }

        }

        //Return file size on the disc
        public static long GetFileSizeOnDisk(string file)
        {
            FileInfo info = new FileInfo(file);
            uint dummy, sectorsPerCluster, bytesPerSector;
            int result = GetDiskFreeSpaceW(info.Directory.Root.FullName, out sectorsPerCluster, out bytesPerSector, out dummy, out dummy);
            if (result == 0) throw new Win32Exception();
            uint clusterSize = sectorsPerCluster * bytesPerSector;
            uint hosize;
            uint losize = GetCompressedFileSizeW(file, out hosize);
            long size;
            size = (long)hosize << 32 | losize;
            return ((size + clusterSize - 1) / clusterSize) * clusterSize;
        }

        public void xmlNodeAppend(XmlDocument xmlDoc, string nodeName, string attributeName, string attributeValue,
            string toAppendNodeInnertext, XmlNode rootNode)
        {
            //Creation of username node and assigning value
            XmlNode user = xmlDoc.CreateElement(nodeName);
            XmlAttribute userName = xmlDoc.CreateAttribute(attributeName);
            userName.Value = attributeValue;//Need to provide value
            user.Attributes.Append(userName);
            user.InnerText = toAppendNodeInnertext;
            rootNode.AppendChild(user);
        }

        //Take xmlDocument and file name to search the file and return the username. If not found then username will be empty
        public string matchAndReturnUsernameFromXML(XmlDocument xmlDoc, string fileName)
        {
            string username="";
            XmlNodeList xmlFileNodes = xmlDoc.GetElementsByTagName("Files");

            foreach(XmlNode FilesNode in xmlFileNodes)
            {

                string atbFileName = FilesNode.SelectSingleNode("//FileName").Attributes[0].Value;
                string atbUserName = FilesNode.SelectSingleNode("//UserDetails").Attributes[0].Value;
                
                if (atbFileName == fileName)
                {
                    username = atbUserName;
                }
            }

            return username;
        }

        //Create folder with username and set permission
        public void createFolderWithUserNameandSetPerm(string username)
        {
            
            string dirName = System.Configuration.ConfigurationManager.AppSettings.Get("outputPath");
            string userToRemove = System.Configuration.ConfigurationManager.AppSettings.Get("removeUser");
            dirName=createFolder(dirName, username);
            RemoveFileSecurity(dirName, userToRemove, FileSystemRights.FullControl, AccessControlType.Allow);
            
        }

        private string createFolder(string dirName, string folderName)//create folder
        {
            // If directory does not exist, create it. 
            if (!Directory.Exists(dirName))
            {
                Directory.CreateDirectory(dirName);
            }
            string folderDir = @dirName + "\\" + folderName;
            if (!Directory.Exists(folderDir))
            {
                Directory.CreateDirectory(folderDir);
            }
            return folderDir;
        }


        // Removes an ACL entry on the specified file for the specified account.
        public static void RemoveFileSecurity(string fileName, string account,
            FileSystemRights rights, AccessControlType controlType)
        {
            List<AccessRule> modifiedRulesCol = new List<AccessRule>();
            // Get a FileSecurity object that represents the
            // current security settings.
            FileSecurity fSecurity = File.GetAccessControl(fileName);

            //Get all rules collection including inherited ones
            AuthorizationRuleCollection ruleCol = fSecurity.GetAccessRules(true, true, typeof(NTAccount));

            try
            {
                //Removing inherited rules
                fSecurity.SetAccessRuleProtection(true, false);


                //Creating a list of rules except that need to be removed
                foreach (AccessRule a in ruleCol)
                {
                    if (string.Compare(a.IdentityReference.ToString(), account, true) != 0)
                    {
                        modifiedRulesCol.Add(a);
                    }
                }

                // Remove the FileSystemAccessRule from the security settings.
                fSecurity.RemoveAccessRule(new FileSystemAccessRule(account,
                    rights, controlType));

                foreach (AccessRule a in modifiedRulesCol)
                {
                    fSecurity.AddAccessRule(new FileSystemAccessRule(a.IdentityReference,
                    FileSystemRights.FullControl, a.AccessControlType));
                }

                fSecurity.AddAccessRule(new FileSystemAccessRule(Environment.UserName,
               FileSystemRights.FullControl, AccessControlType.Allow));

                // Set the new access settings.
                File.SetAccessControl(fileName, fSecurity);
            }

            catch (Exception e)
            {
                WriteToFile(e.ToString());
                
            }

        }
        public static void WriteToFile(string Message)
        {
            string path = System.Configuration.ConfigurationManager.AppSettings.Get("logPath");
            //string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string filepath = path + "\\log.txt";
                if (!File.Exists(filepath))
                {
                    File.Create(filepath);
                    // Create a file to write to.   
                    using (StreamWriter sw = new System.IO.StreamWriter(filepath))
                    {
                        sw.WriteLine(Message);
                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(filepath))
                    {
                        sw.WriteLine(Message);
                    }
                }
            }
            catch (Exception e)
            {
                WriteToFile(e.ToString());
            }
        }

        // Adds an ACL entry on the specified file for the specified account.
        public static void AddFileSecurity(string fileName, string account,
            FileSystemRights rights, AccessControlType controlType)
        {


            // Get a FileSecurity object that represents the
            // current security settings.
            FileSecurity fSecurity = File.GetAccessControl(fileName);

            // Add the FileSystemAccessRule to the security settings.
            fSecurity.AddAccessRule(new FileSystemAccessRule(account,
                rights, controlType));

            // Set the new access settings.
            File.SetAccessControl(fileName, fSecurity);

        }


        public static string GetUsername(int sessionID)
        {
            try
            {
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                Pipeline pipeline = runspace.CreatePipeline();
                pipeline.Commands.AddScript("Quser");
                pipeline.Commands.Add("Out-String");

                Collection<PSObject> results = pipeline.Invoke();

                runspace.Close();

                StringBuilder stringBuilder = new StringBuilder();
                foreach (PSObject obj in results)
                {
                    stringBuilder.AppendLine(obj.ToString());
                }

                foreach (string User in stringBuilder.ToString().Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Skip(1))
                {
                    string[] UserAttributes = User.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                    if (UserAttributes.Length == 6)
                    {
                        if (int.Parse(UserAttributes[1].Trim()) == sessionID)
                        {
                            return UserAttributes[0].Replace(">", string.Empty).Trim();
                        }
                    }
                    else
                    {
                        if (int.Parse(UserAttributes[2].Trim()) == sessionID)
                        {
                            return UserAttributes[0].Replace(">", string.Empty).Trim();
                        }
                    }
                }

            }
            catch (Exception exp)
            {
                // Error handling
            }

            return "Undefined";
        }

        [DllImport("kernel32.dll")]
        static extern uint GetCompressedFileSizeW([In, MarshalAs(UnmanagedType.LPWStr)] string lpFileName,
           [Out, MarshalAs(UnmanagedType.U4)] out uint lpFileSizeHigh);

        [DllImport("kernel32.dll", SetLastError = true, PreserveSig = true)]
        static extern int GetDiskFreeSpaceW([In, MarshalAs(UnmanagedType.LPWStr)] string lpRootPathName,
           out uint lpSectorsPerCluster, out uint lpBytesPerSector, out uint lpNumberOfFreeClusters,
           out uint lpTotalNumberOfClusters);
    }

}
