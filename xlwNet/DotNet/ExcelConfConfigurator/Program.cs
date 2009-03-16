using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

namespace ExcelConfConfigurator
{
    class Program
    {
        static int  Main(string[] args)
        {
            try
            {
                if (args.Length != 3)
                {
                    Console.WriteLine("Usage: ExcelConfConfigurator <excel.exe.conf template> <assembley list file> ");
                    return 0;
                }
                string ListFileName = args[1];
                string fname;
                List<string> AssembleyList = new List<string>();
                using (StreamReader ListFile = new StreamReader(ListFileName))
                {
                    while ((fname = ListFile.ReadLine()) != null)
                    {
                        AssembleyList.Add(fname);
                    }

                }
                Updater.Configurator.Config(args[0], args[2], "Excel.exe.conf", AssembleyList.ToArray());

                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine("** " + e.Message);
                return -1;
            }
        }
    }
}