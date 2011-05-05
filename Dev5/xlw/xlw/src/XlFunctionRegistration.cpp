/*
Copyright (C) 2006 Mark Joshi
Copyright (C) 2007, 2008 Eric Ehlers
Copyright (C) 2011 Narinder S Claire
Copyright (C) 2011 John Adcock

This file is part of XLW, a free-software/open-source C++ wrapper of the
Excel C API - http://xlw.sourceforge.net/

XLW is free software: you can redistribute it and/or modify it under the
terms of the XLW license.  You should have received a copy of the
license along with this program; if not, please email xlw-users@lists.sf.net

This program is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
FOR A PARTICULAR PURPOSE.  See the license for more details.
*/


#include <xlw/XlFunctionRegistration.h>
#include <xlw/xlfFuncDesc.h>
#include <xlw/xlfCmdDesc.h>
#include <xlw/xlfArgDescList.h>

#include <stdio.h>

using namespace xlw;
using namespace XLRegistration;

XLFunctionRegistrationData::XLFunctionRegistrationData(const std::string& FunctionName_,
    const std::string& ExcelFunctionName_,
    const std::string& FunctionDescription_,
    const std::string& Library_,
    const Arg Arguments[],
    int NoOfArguments_,
    bool Volatile_,
    bool Threadsafe_,
    const std::string& ReturnTypeCode_,
    const std::string& HelpID_,
    bool Asynchronous_,
    bool MacroSheetEquivalent_,
    bool ClusterSafe_)
    :                FunctionName(FunctionName_),
    ExcelFunctionName(ExcelFunctionName_),
    FunctionDescription(FunctionDescription_),
    Library(Library_),
    NoOfArguments(NoOfArguments_),
    Volatile(Volatile_),
    Threadsafe(Threadsafe_),
    ReturnTypeCode(ReturnTypeCode_),
    helpID(HelpID_),
    Asynchronous(Asynchronous_),
    MacroSheetEquivalent(MacroSheetEquivalent_),
    ClusterSafe(ClusterSafe_)
{

    ArgumentNames.reserve(NoOfArguments);
    ArgumentDescriptions.reserve(NoOfArguments);

    for (int i=0; i < NoOfArguments; i++)
    {
        ArgumentNames.push_back(Arguments[i].ArgumentName);
        ArgumentDescriptions.push_back(Arguments[i].ArgumentDescription);
        ArgumentTypes.push_back(Arguments[i].ArgumentType);
    }
}




const std::string & XLFunctionRegistrationData::GetFunctionName() const
{
    return FunctionName;
}
const std::string & XLFunctionRegistrationData::GetExcelFunctionName() const
{
    return ExcelFunctionName;
}
const std::string & XLFunctionRegistrationData::GetFunctionDescription() const
{
    return FunctionDescription;
}

const std::string &  XLFunctionRegistrationData::GetReturnTypeCode() const
{
    return ReturnTypeCode;
}

const std::string & XLFunctionRegistrationData::GetHelpID() const
{
    return helpID;
}

const std::string & XLFunctionRegistrationData::GetLibrary() const
{
    return Library;
}

int XLFunctionRegistrationData::GetNoOfArguments() const{
    return NoOfArguments;
}

const std::vector<std::string> & XLFunctionRegistrationData::GetArgumentNames() const
{
    return ArgumentNames;
}

const std::vector<std::string> & XLFunctionRegistrationData::GetArgumentDescriptions() const
{
    return ArgumentDescriptions;
}

const std::vector<std::string> & XLFunctionRegistrationData::GetArgumentTypes() const
{
    return ArgumentTypes;
}



const std::string& XLCommandRegistrationData::GetCommandName() const
{
    return CommandName ; 
}


const std::string& XLCommandRegistrationData::GetExcelCommandName() const
{
    return ExcelCommandName ; 
}


const std::string& XLCommandRegistrationData::GetCommandComment() const
{
    return Comment ; 
}


const std::string& XLCommandRegistrationData::GetMenu() const
{
    return  Menu; 
}

const std::string& XLCommandRegistrationData::GetMenuText() const
{
    return  MenuText; 
}


XLFunctionRegistrationHelper::XLFunctionRegistrationHelper(const std::string& FunctionName,
    const std::string& ExcelFunctionName,
    const std::string& FunctionDescription,
    const std::string& Library,
    const Arg Args[],
    int NoOfArguments,
    bool Volatile,
    bool Threadsafe,
    const std::string& returnTypeCode,
    const std::string& helpID,
    bool Asynchronous,
    bool MacroSheetEquivalent,
    bool ClusterSafe)
{
    XLFunctionRegistrationData tmp(FunctionName,
        ExcelFunctionName,
        FunctionDescription,
        Library,
        Args,
        NoOfArguments,
        Volatile,
        Threadsafe,
        returnTypeCode,
        helpID,
        Asynchronous,
        MacroSheetEquivalent,
        ClusterSafe);

    ExcelFunctionRegistrationRegistry::Instance().AddFunction(tmp);
}

XLCommandRegistrationHelper::XLCommandRegistrationHelper(const std::string& CommandName,
    const std::string& ExcelCommandName,
    const std::string& Comment,
    const std::string& Menu,
    const std::string& MenuText)
{
    XLCommandRegistrationData tmp(CommandName,
        ExcelCommandName,
        Comment,
        Menu,
        MenuText);

    ExcelFunctionRegistrationRegistry::Instance().AddCommand(tmp);
}

void ExcelFunctionRegistrationRegistry::DoTheRegistrations() const
{

    for (functionCache::const_iterator it = Functions.begin(); it !=  Functions.end(); ++it)
    {
        it->second->Register();
    }


    for (commandCache::const_iterator it = Commands.begin(); it !=  Commands.end(); ++it)
    {
        it->second->Register();
        it->second->AddToMenuBar();
    }
}

void ExcelFunctionRegistrationRegistry::DoTheDeregistrations() const
{
    for (functionCache::const_iterator it = Functions.begin(); it !=  Functions.end(); ++it)
    {
        it->second->Unregister();
    }


    for (commandCache::const_iterator it = Commands.begin(); it !=  Commands.end(); ++it)
    {
        it->second->RemoveFromMenuBar();
        it->second->Unregister();
    }


}

void ExcelFunctionRegistrationRegistry::AddFunction(const XLFunctionRegistrationData& data)
{
    
    XlfFuncDesc::RecalcPolicy policy = data.GetVolatile() ? XlfFuncDesc::Volatile : XlfFuncDesc::NotVolatile;
    xlw_tr1::shared_ptr<XlfFuncDesc>  xlFunction( new XlfFuncDesc(
            data.GetFunctionName(),
            data.GetExcelFunctionName(),
            data.GetFunctionDescription(),
            data.GetLibrary(),
            policy,
            data.GetThreadsafe(),
            data.GetReturnTypeCode(),
            data.GetHelpID(),
            data.GetAsynchronous(),
            data.GetMacroSheetEquivalent(),
            data.GetClusterSafe()));
        XlfArgDescList xlFunctionArgs;

        for (int i=0; i < data.GetNoOfArguments(); ++i)
        {
            XlfArgDesc ThisArgumentDescription(data.GetArgumentNames()[i],
                data.GetArgumentDescriptions()[i],
                data.GetArgumentTypes()[i]);
            xlFunctionArgs + ThisArgumentDescription;
            
        }
 
    xlFunction->SetArguments(xlFunctionArgs);
    Functions[data.GetExcelFunctionName()] = xlFunction;
}
void ExcelFunctionRegistrationRegistry::AddCommand(const XLCommandRegistrationData& data)
{
    
        xlw_tr1::shared_ptr<XlfCmdDesc>  
            theCommand(new XlfCmdDesc(data.GetCommandName(),
            data.GetExcelCommandName(),
            data.GetCommandComment(),
            data.GetMenu(),
            data.GetMenuText(),
            !data.GetMenu().empty()));

        Commands[data.GetExcelCommandName()] = theCommand;

}
