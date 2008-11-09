
/*
 Copyright (C) 2008 Narinder Claire

 This file is part of XLW.NET 

 XLW.NET is free software: you can redistribute it and/or modify it under the
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
using xlwDotNet;
using xlwDotNet.xlwTypes;


namespace TemplateProject
{
    public class Class1
    {
     

        [ExcelExport("echoes an array")]
        public static MyArray EchoArray(
                             [Parameter(" argument to be echoed")] MyArray Echoee
                            )
        {
            return Echoee;
        }
       
    }
}

