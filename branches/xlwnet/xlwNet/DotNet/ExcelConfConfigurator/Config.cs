using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Reflection;
using System.Xml;
using System.Configuration;



namespace Updater
{

    class Configurator
    {

        struct AssembleyRecord
        {
            public string name;
            public string path;
            public string publickeytoken;
            public string version;
            public string culture;
        }


        public static void Config(string original, string installPath, string destination, string[] AssembleyList)
        {

                System.Text.Encoding enc = System.Text.Encoding.Unicode;
                List<AssembleyRecord> theAssemblies = new List<AssembleyRecord>();
                for (int i = 0; i < AssembleyList.Length; ++i)
                {
                    
                    AssembleyRecord theRecord = new AssembleyRecord();
                    string theAssembly = AssembleyList[i];

                        theRecord.path = installPath + @"\" + theAssembly;
                        Console.WriteLine("Creating record for "+theRecord.path);

                        Assembly thePhysical = Assembly.LoadFile(installPath + @"\" + theAssembly);
                        string theInfo = thePhysical.FullName;
                        string[] theTokens = theInfo.Split(',');

                        theRecord.name = theTokens[0];
                        theRecord.version = theTokens[1].Split('=')[1];
                        theRecord.culture = theTokens[2].Split('=')[1];
                        theRecord.publickeytoken = theTokens[3].Split('=')[1];

                        theAssemblies.Add(theRecord);

                }

                XmlDocument doc = new XmlDocument();
                doc.Load(installPath + @"\" + original);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                string ns = "urn:schemas-microsoft-com:asm.v1";
                nsmgr.AddNamespace("ns", ns);
                XmlNode binding = doc.SelectSingleNode("//ns:assemblyBinding", nsmgr);
                Console.WriteLine("Building Excel.exe.config");
                foreach (AssembleyRecord record in theAssemblies)
                {
                    XmlElement newAssembley = doc.CreateElement("dependentAssembly", ns);
                    XmlNode theAssembley = binding.AppendChild(newAssembley);

                    Console.WriteLine("Regsistering " + record.name);
                    XmlElement theIdentity = doc.CreateElement("assemblyIdentity", ns);
                    theIdentity.SetAttribute("name", record.name);
                    theIdentity.SetAttribute("publicKeyToken", record.publickeytoken);
                    theIdentity.SetAttribute("culture", record.culture);
                    theAssembley.AppendChild(theIdentity);

                    XmlElement theCodeBase = doc.CreateElement("codeBase", ns);
                    theCodeBase.SetAttribute("version", record.version);
                    theCodeBase.SetAttribute("href", @"file://" + record.path);
                    theAssembley.AppendChild(theCodeBase);

                }

                XmlTextWriter theWriter = new XmlTextWriter(destination, Encoding.ASCII);

                doc.WriteTo(theWriter);
                theWriter.Close();
               
        }
    }
}
