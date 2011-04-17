//
//
//                                    main.cpp
//
//
/*
Copyright (C) 2006, 2008 Mark Joshi
Copyright (C) 2011  Narinder Claire

This file is part of XLW, a free-software/open-source C++ wrapper of the
Excel C API - http://xlw.sourceforge.net/

XLW is free software: you can redistribute it and/or modify it under the
terms of the XLW license.  You should have received a copy of the
license along with this program; if not, please email xlw-users@lists.sf.net

This program is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
FOR A PARTICULAR PURPOSE.  See the license for more details.
*/
#ifdef _MSC_VER
#if _MSC_VER < 1250
#pragma warning(disable:4786)
#endif
#pragma warning(disable:4258)
#endif
#include <iostream>
#include <fstream>
#include "Functionizer.h"
#include "FunctionModel.h"
#include "FunctionType.h"
#include "Outputter.h"
#include "ManagedOutputter.h"
#include "ParserData.h"
#include "Tokenizer.h"
#include "Strip.h"
using namespace std;


int main(int argc, char *argv[])
{
  try
  {

    std::vector<std::string> args;
    std::vector<std::string> options;

    for (int i=1; i < argc; ++i)
    {
      std::string arg(argv[i]);
      if (arg.size()>0 && arg[0] =='-')
        options.push_back(arg);
      else
        args.push_back(arg);

    }

    if (args.size() < 1 || args.size() > 2)
		throw("usage is :\n "
		      "          inputfile \n"
		      "          inputfile outputfile \n"
			  "      -m  inputfile ");
    std::string inputfile(args[0]);

    bool clw = false;
	bool managed = false;

    for (std::vector<std::string>::const_iterator it = options.begin(); it != options.end(); ++it)
    {
      if (*it == "-c")
	  {
        clw = true;
	  }
	  else if (*it == "-m")
	  {
		  managed = true;
	  }
	  else
	  {
        std::cerr << "unknown option ignored: " << *it << "\n";
	  }

    }


    std::string outputfile;
	std::string managed_outputfile_h;
	std::string managed_outputfile_cpp;

	std::string dirname(getdir(inputfile));
	

    if (args.size()==2)
	{
	 
	  if (managed) throw ("specifying output filename not allowed for managed mode ");
      outputfile = args[1];
	}
    else
    {
      if (clw)
        outputfile= "clw";
      else
	  {

        outputfile = "xlw";
		managed_outputfile_h = "mxlw";
        
	  }
	  std::string libName(strip(inputfile));
      for (unsigned long i=0; i < libName.size(); i++)
      {
        if (libName[i] == '.')
          break;
        PushBack(outputfile,libName[i]);
		PushBack(managed_outputfile_h,libName[i]);
      }

      outputfile += ".cpp";
	  managed_outputfile_cpp = managed_outputfile_h + ".cpp";
	  managed_outputfile_h +=".h";
    }

    ifstream input(inputfile.c_str());
    if (!input)
      throw("input file not found :"+inputfile+"\n");

    std::vector<char> inputvector;

    char c;
    while (input.get(c))
    {
      int i = static_cast<int>(c);
      if (i<32 && c!='\n')
        c=' '; // strip out special characters
      inputvector.push_back(c);
    }
    std::cout << "file has been read in\n";

    std::vector<Token> tokenVector1(Tokenize(inputvector));
    std::cout << "file has been tokenized\n";

    std::vector<Token> tokenVector2(Strip(tokenVector1));
    std::cout << "file has been stripped\n";

    std::string LibraryName(inputfile);// use input file name as default library name

    std::vector<FunctionModel> modelVector(ConvertToFunctionModel(tokenVector2,LibraryName));

    std::cout << "file has been function modeled\n";

    std::vector<FunctionDescription> functionVector(FunctionTyper(modelVector));

    std::cout << "file has been function described\n";


    std::vector<char> outputVector_h;
	std::vector<char> outputVector_cpp;


    if(managed)
	{
		OutputFileCreatorMan(functionVector,inputfile,LibraryName,outputVector_h,outputVector_cpp);
		inputfile = managed_outputfile_h;
      
		std::cout << " .. writing " << (dirname+"/" + managed_outputfile_h) <<  "\n";
		writeOutputFile(dirname+"/" + managed_outputfile_h,outputVector_h);
		std::cout << " .. writing " << (dirname+"/" + managed_outputfile_cpp) <<  "\n";
		writeOutputFile(dirname+"/" + managed_outputfile_cpp,outputVector_cpp);

		outputfile = dirname+"/" + outputfile;

	}




    if (clw)
     outputVector_cpp = OutputFileCreatorCL(functionVector,
                                          inputfile);
    else
     outputVector_cpp = OutputFileCreator(functionVector,
                                          inputfile,LibraryName);

	std::cout << " .. writing " << outputfile << "\n";
    writeOutputFile(outputfile,outputVector_cpp);

    std::cout << "new file is a vector\n";

   
    std::cout << "all done\n";

  }
  catch (const char *c)
  {
    std::cout << "***ERROR***\n" << c << "\n***ERROR***\n";
  }
  catch (std::string c)
  {
    std::cout << "***ERROR***\n" << c << "\n***ERROR***\n";
  }
  catch (...)
  {
    std::cout << "***ERROR***\n" << "exception thrown" << "\n***ERROR***\n";
  }
  //char d;
  //std::cin >> d;
  return 0;
}
