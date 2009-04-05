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
using xlwDotNet;
using xlwDotNet.xlwTypes;


namespace Example
{
    public class Class1
    {
     
        [ExcelExport("tests empty args ")]
        public static String EmptyArgFunction()
        {
            return "this function is useless except for testing.";
        }

        [ExcelExport("echoes a short ")]
        public static short EchoShort(
                             [Parameter("number to be echoed")] short x
                            )
        {
            return x;
        }

        [ExcelExport("echoes a matrix")]
        public static MyMatrix EchoMat(
                             [Parameter("argument to be echoed")] MyMatrix Echoee
                            )
        {
            return Echoee;
        }
        [ExcelExport("echoes an array")]
        public static MyArray EchoArray(
                             [Parameter(" argument to be echoed")] MyArray Echoee
                            )
        {
            return Echoee;
        }
        [ExcelExport("echoes a  CellMatrix")]
        public static CellMatrix EchoCells(
                             [Parameter(" argument to be echoed")] CellMatrix Echoee
                            )
        {
            return Echoee;
        }

        [ExcelExport("computes the circumference of a circle ")]
        public static double Circ(
                             [Parameter("the circle's diameter")] double Diameter
                            )
        {
            return Diameter * 3.14159;
        }

        [ExcelExport("Concatenates two strings")]
        public static string Concat(
                             [Parameter("first string")] string str1,
                             [Parameter("second string")] string str2  
                            )
        {
            return str1+str2;
        }
        [ExcelExport("computes mean and variance of a range ")]
        public static MyArray stats(
                             [Parameter("input for computation")] MyArray data
                            )
        {
            double total=0.0;
	        double totalsq=0.0;

	        if (data.size < 2)
		        throw(new Exception("At least data points are needed"));

	        for (int i=0; i < data.size; i++)
	            {
		            total+=data[i];
		            totalsq+=data[i]*data[i];
	            }
    	
	        MyArray values = new MyArray(2);
	        values[0] = total/data.size;
	        values[1] = totalsq/data.size - values[0] *values[0] ;

	        return values;

        }


        [ExcelExport("throws an exception")]
        public static double Exception()
        {
            throw (new ArgumentException("OH NO"));
            return 1.1;
        }

        [ExcelExport("says hello name")]
        public static string HelloWorldAgain(
                             [Parameter("name to be echoed")] string name
                            )
        {
            return "hello " + name;
        }



        [ExcelExport("echoes arg list")]
        public static CellMatrix EchoArgList(
                             [Parameter(" arguments to echo")] ArgumentList args
                            )
        {
            return args.AllData();
        }

///////////////////////////////////////////////////////////////////////////////////
/////////////// C# Specific test /////////////////////////////////////////////////
        [ExcelExport("takes double[]")]
        public static MyArray CastToCSArray (
                             [Parameter(" double Array")] double[] csarray
                            )
        {
            return csarray;
        }

        [ExcelExport("takes double[]")]
        public static double[] CastToCSArrayTwice(
                             [Parameter(" double Array")] double[] csarray
                            )
        {
            return csarray;
        }

        [ExcelExport("takes double[,]")]
        public static MyMatrix CastToCSMatrix(
                             [Parameter(" double Array")] double[,] csmatrix
                            )
        {
            return csmatrix;
        }

        [ExcelExport("takes double[,]")]
        public static double[,] CastToCSMatrixTwice(
                             [Parameter(" double Array")] double[,] csmatrix
                            )
        {
            return csmatrix;
        }

        ///////////////////////////////////////////////////////////////////////////////////
        /////////////// Throwing Exceptions ///////////////////////////////////////////////

        [ExcelExport("throws an exception of type ArgumentNullException")]
        public static double throwString(
                             [Parameter("Just any random string ")] string err
                            )
        {
            throw(new ArgumentNullException(err,err));
            return 0;
        }
        [ExcelExport("throws an exception of type cellMatrixException")]
        public static double throwCellMatrix(
                             [Parameter("Just any random string ")] string err
                            )
        {
            ArgumentList theContent = new ArgumentList("EXCEPTION");
            theContent.add("Error", "Don't worry, this is a dummy Error :-)");
            throw (new cellMatrixException("WOW its warm up here ",theContent.AllData()));
            return 0;
        }

        [ExcelExport("makes the C Runtime throw an exception")]
        public static double throwCError(
                             [Parameter("Just any random string ")] string err
                            )
        {
            MyMatrix theMatrix = new MyMatrix(2,2);
            return theMatrix[3,3];
        }
    }
}

