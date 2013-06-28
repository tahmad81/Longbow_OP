using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CasePlugin
{
    public class ClassService
    {
        public ClassService()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\wordService.ini";
            IniFile iniFile = new IniFile(path);
            iniFile.Load(path);

            URL = iniFile.GetKeyValue("MAIN", "URL");
            Username = iniFile.GetKeyValue("MAIN", "Username");
            Password = iniFile.GetKeyValue("MAIN", "Password");

            MSGPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\_msg";
            if (!Directory.Exists(MSGPath))
            {
                Directory.CreateDirectory(MSGPath);
            }
            if (!MSGPath.EndsWith("\\"))
            {
                MSGPath += "\\";
            }
        }

        public void SaveSetting()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\wordService.ini";
            IniFile iniFile = new IniFile();

            iniFile.Load(path);
            IniFile.IniSection section = iniFile.GetSection("MAIN");
            section.AddKey("URL").Value = URL;
            section.AddKey("Username").Value = Username;
            section.AddKey("Password").Value = Password;

            iniFile.Save(path);
            
        }

        public int[] GetCaseNumber()
        {            
            XCaseService.CaseService caseService = new XCaseService.CaseService();
            int[] caseNumber = caseService.GetCaseNumber();

            return caseNumber;
        }

        public string URL { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }

        public string MSGPath;

    }
}
