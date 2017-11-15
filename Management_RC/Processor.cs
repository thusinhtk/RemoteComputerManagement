using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Management_RC
{
    class Processor
    {
        //Singleton 
        //private static Processor instance = new Processor();

        //private Processor()
        //{
        //}
        //public static Processor getInstance()
        //{
        //    return instance;
        //}

        //Declare variables
        public static string _path = System.Configuration.ConfigurationManager.ConnectionStrings["TextPath"].ConnectionString;

        public static string _computerName = Environment.MachineName;
        public static string _userName = Environment.UserName;       


        public void Run()
        {            
            bool isUpdate = false;

            string[] allFileName = Directory.GetFiles(_path, "*.txt");
            string[] stringseparators = new string[] { "___" };

            int index = 0;
            foreach (string line in allFileName)
            {
                bool iscontain = line.Contains(_computerName);
                if (iscontain)
                {
                    StringBuilder savetemp = new StringBuilder();
                    //update lines
                    string[] line_child = line.Split(stringseparators, StringSplitOptions.None);

                    string dateTime = DateTime.Now.ToString().Replace(":", ".").Replace("/", "-");

                    savetemp.Append(line_child[0]).Append("___")
                                                  .Append(_userName)
                                                  .Append("___")
                                                  .Append(dateTime)
                                                  .Append(".txt");

                    //delete file to update;
                    File.Delete(line);

                    //create new file
                    string pathNew = savetemp + "";
                    File.Create(pathNew).Close();
                    isUpdate = true;
                    break;
                }
                ++index;
            }
            if(!isUpdate)
            {
                StringBuilder tempPath  = new StringBuilder();

                string dateTime = DateTime.Now.ToString().Replace(":", ".").Replace("/", "-");

                tempPath.Append(_computerName)
                           .Append("___")
                           .Append(_userName)
                           .Append("___")
                           .Append(dateTime)
                           .Append(".txt");

                string newPath = _path + tempPath;
                File.Create(newPath).Close();
            }
            Console.WriteLine("Done");
        }
    }
}
