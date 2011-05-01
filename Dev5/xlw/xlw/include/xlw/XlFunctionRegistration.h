
/*
 Copyright (C) 2006 Mark Joshi
 Copyright (C) 2007, 2008 Eric Ehlers
 Copyright (C) 2011 Narinder S Claire

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#if !defined(REGISTER_XL_FUNCTION_H)
#define REGISTER_XL_FUNCTION_H


#include <string>
#include <vector>
#include <list>
#include <xlw/XlfExcel.h>
#include <xlw/Singleton.h>

namespace xlw {

    namespace XLRegistration
    {

    struct Arg { const char * ArgumentName; const char * ArgumentDescription; const char * ArgumentType; };

    class XLFunctionRegistrationData
    {
    public:
          XLFunctionRegistrationData(const std::string& FunctionName_,
                         const std::string& ExcelFunctionName_,
                         const std::string& FunctionDescription_,
                         const std::string& Library_,
                         const Arg Arguments[],
                         int NoOfArguments_,
                         bool Volatile_,
                         bool Threadsafe_,
                         const std::string &ReturnTypeCode_,
                         const std::string &HelpID_,
                         bool Asynchronous_,
                         bool MacroSheetEquivalent_,
                         bool ClusterSafe_);

        std::string GetFunctionName() const;
        std::string GetExcelFunctionName() const;
        std::string GetFunctionDescription() const;
        std::string GetLibrary() const;
        int GetNoOfArguments() const;
        std::vector<std::string> GetArgumentNames() const;
        std::vector<std::string> GetArgumentDescriptions() const;
        std::vector<std::string> GetArgumentTypes() const;

        bool GetVolatile() const
        {
            return Volatile;
        }

        bool GetThreadsafe() const
        {
            return Threadsafe;
        }

        bool GetAsynchronous() const
        {
            return Asynchronous;
        }

        bool GetMacroSheetEquivalent() const
        {
            return MacroSheetEquivalent;
        }

        bool GetClusterSafe() const
        {
            return ClusterSafe;
        }

        std::string GetReturnTypeCode() const;
        std::string GetHelpID() const;
    private:

        std::string FunctionName;
        std::string ExcelFunctionName;
        std::string FunctionDescription;
        std::string Library;
        int NoOfArguments;
        std::vector<std::string> ArgumentNames;
        std::vector<std::string> ArgumentDescriptions;
        std::vector<std::string> ArgumentTypes;
        bool Volatile;
        bool Threadsafe;
        std::string ReturnTypeCode;
        std::string helpID;
        bool Asynchronous;
        bool MacroSheetEquivalent;
        bool ClusterSafe;
    };

    class XLFunctionRegistrationHelper
    {
    public:

        XLFunctionRegistrationHelper(const std::string& FunctionName,
                         const std::string& ExcelFunctionName,
                         const std::string& FunctionDescription,
                         const std::string& Library,
                         const Arg Args[] = 0,
                         int NoOfArguments = 0,
                         bool Volatile = false,
                         bool Threadsafe = false,
                         const std::string &ReturnTypeCode_ = "",
                         const std::string &HelpID = "",
                         bool Asynchronous = false,
                         bool MacroSheetEquivalent = false,
                         bool ClusterSafe = false);
    private:

    };


    // singleton pattern, cf the Factory
    class ExcelFunctionRegistrationRegistry: public singleton<ExcelFunctionRegistrationRegistry>
    {
	friend class singleton<ExcelFunctionRegistrationRegistry>;
    public:
        void DoTheRegistrations() const;
        void DoTheDeregistrations() const;
        void AddFunction(const XLFunctionRegistrationData&);

    private:
        ExcelFunctionRegistrationRegistry(){}
        std::list<XLFunctionRegistrationData> RegistrationData;

    };

    }

}

#endif

