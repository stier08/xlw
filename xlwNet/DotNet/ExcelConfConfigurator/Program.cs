/*
 Copyright (C) 2008 2009  Narinder S Claire

 This file is part of XLWDOTNET, a free-software/open-source C# wrapper of the
 Excel C API - http://xlw.sourceforge.net/
 
 XLWDOTNET is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/
 
 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net
 
 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Reflection;


namespace ExcelConfConfigurator
{
    class Program
    {
        static int  Main(string[] args)
        {

            try
            {
                if (args.Length < 3 || args.Length > 4)
                {
                    Console.WriteLine("Usage: ExcelConfConfigurator <excel.exe.conf template> <assembley> <outputDir> [<assembley list file>]\n   [<assembley list file>] is optional\n");
                    return -1;
                }
                List<string> AssembleyList = new List<string>();
                if (args.Length == 4)
                {
                    string ListFileName = args[3];
                    string fname;

                    using (StreamReader ListFile = new StreamReader(ListFileName))
                    {
                        while ((fname = ListFile.ReadLine()) != null)
                        {
                            fname = fname.Trim();
                            if (fname != "")
                            {
                                AssembleyList.Add(fname);
                            }
                        }

                    }
                }
                string dllName = args[1];
                Console.WriteLine("Attempting to Load Assembley " + dllName);
                AssembleyList.Add(dllName);
                Assembly theLoadedAssembley = Assembly.LoadFrom(dllName);
                Console.WriteLine("Attempting to get all referenced aseemblies of " + dllName + " that are in the same directory\n");
                AssemblyName[] ReferencedAssemblies =
                    theLoadedAssembley.GetReferencedAssemblies();
                string AssemblyPath = Path.GetDirectoryName(theLoadedAssembley.Location);

                foreach (AssemblyName s in ReferencedAssemblies)
                {
                    string tryDll = s.Name + ".dll";
                    string tryFullDLLPath = Path.Combine(AssemblyPath, tryDll);
                    if (File.Exists(tryFullDLLPath))
                    {
                        Console.WriteLine("Referenced Assembly " + s.Name + " found\n");
                        AssembleyList.Add(tryDll);
                    }

                }
               

                Updater.Configurator.Config(args[0], args[2], "Excel.exe.config", AssembleyList.ToArray());

                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine("********************\n" + e.Message + "\n********************\n");
                return -1;
            }
        }
    }
}