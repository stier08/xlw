//
//
//                            ArgList.h
//
//
/*
 Copyright (C) 2006 Mark Joshi
 Copyright (C) 2009 Narinder S Claire
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

#ifndef ARG_LIST_H
#define ARG_LIST_H

#include "CellMatrix.h"
#include "MyContainers.h"
#include <map>
#include <string>
#include <vector>

namespace xlw {


    /*! isSameType<type1, type2>::value returns true if types are the same
        used to allow optimization if no conversion is required.
        compiler should remove unused branch in release mode
    */
    template<typename, typename>
    struct isSameType
    {
        static const bool value = false;
    };

    template<typename _Tp>
    struct isSameType<_Tp, _Tp>
    {
        static const bool value = true;
    };


    void MakeLowerCase(std::string& input);

    class ArgumentList
    {
    public:

        ArgumentList(CellMatrix cells,
                          std::string ErrorIdentifier);

        ArgumentList(std::string name);


        enum ArgumentType
        {
            string, number, vector, matrix,
            boolean, list, cells
        };

        std::string GetStructureName() const;

        const std::vector<std::pair<std::string, ArgumentType> >& GetArgumentNamesAndTypes() const;

        std::string GetStringArgumentValue(const std::string& ArgumentName);
        unsigned long GetULArgumentValue(const std::string& ArgumentName);
        double GetDoubleArgumentValue(const std::string& ArgumentName);
        inline MyArray GetArrayArgumentValue(const std::string& ArgumentName);
        inline MyMatrix GetMatrixArgumentValue(const std::string& ArgumentName);
        bool GetBoolArgumentValue(const std::string& ArgumentName);
        CellMatrix GetCellsArgumentValue(const std::string& ArgumentName);
        ArgumentList GetArgumentListArgumentValue(const std::string& ArgumentName);


        // bool indicates whether the argument was found
        bool GetIfPresent(const std::string& ArgumentName,
            unsigned long& ArgumentValue);
        bool GetIfPresent(const std::string& ArgumentName,
            double& ArgumentValue);
        inline bool GetIfPresent(const std::string& ArgumentName,
            MyArray& ArgumentValue);
        inline bool GetIfPresent(const std::string& ArgumentName,
            MyMatrix& ArgumentValue);
        bool GetIfPresent(const std::string& ArgumentName,
            bool& ArgumentValue);
        bool GetIfPresent(const std::string& ArgumentName,
            CellMatrix& ArgumentValue);
        bool GetIfPresent(const std::string& ArgumentName,
            ArgumentList& ArgumentValue);


        bool IsArgumentPresent(const std::string& ArgumentName) const;

        void CheckAllUsed(const std::string& ErrorId) const;

        CellMatrix AllData() const; // makes data into a cell matrix that could be used for
                                    // creating the same argument list
                                    // useful for checking the class works!

        // data insertions

        void add(const std::string& ArgumentName, const std::string& value);
        void add(const std::string& ArgumentName, double value);
        inline void add(const std::string& ArgumentName, const MyArray& value);
        inline void add(const std::string& ArgumentName, const MyMatrix& value);
        void add(const std::string& ArgumentName, bool value);
        void add(const std::string& ArgumentName, const CellMatrix& values);
        void addList(const std::string& ArgumentName, const CellMatrix& values);
        void add(const std::string& ArgumentName, const ArgumentList& values);

    private:
        inline std::vector<double> GetArrayArgumentValueInternal(const std::string& ArgumentName);
        inline NCMatrix GetMatrixArgumentValueInternal(const std::string& ArgumentName);
        void addArray(const std::string& ArgumentName, const std::vector<double>& value);
        void addMatrix(const std::string& ArgumentName, const NCMatrix& value);
        bool GetIfPresentInternal(const std::string& ArgumentName,
            std::vector<double>& ArgumentValue);
        bool GetIfPresentInternal(const std::string& ArgumentName,
            NCMatrix& ArgumentValue);
        std::string StructureName;

        std::vector<std::pair<std::string, ArgumentType> > ArgumentNames;
        std::map<std::string,double> DoubleArguments;
        std::map<std::string,std::vector<double> > ArrayArguments;
        std::map<std::string,NCMatrix> MatrixArguments;
        std::map<std::string,std::string> StringArguments;
        std::map<std::string,CellMatrix> ListArguments;


        std::map<std::string,CellMatrix> CellArguments;
        std::map<std::string,bool> BoolArguments;

        std::map<std::string,ArgumentType> Names;

        std::map<std::string,bool> ArgumentsUsed;

        void GenerateThrow(std::string message, size_t row, size_t column);
        void UseArgumentName(const std::string& ArgumentName); // private as no error checking performed
        void RegisterName(const std::string& ArgumentName, ArgumentType type);
    };
}

inline void xlw::ArgumentList::add(const std::string& ArgumentName, const MyArray& value)
{
    if(isSameType<MyArray, std::vector<double> >::value)
    {
        addArray(ArgumentName, value);
    }
    else
    {
        throw("Not implemented");
    }
}

inline void xlw::ArgumentList::add(const std::string& ArgumentName, const MyMatrix& value)
{
    if(isSameType<MyMatrix, NCMatrix>::value)
    {
        addMatrix(ArgumentName, value);
    }
    else
    {
        throw("Not implemented");
    }
}
inline xlw::MyArray xlw::ArgumentList::GetArrayArgumentValue(const std::string& ArgumentName)
{
    if(isSameType<MyArray, std::vector<double> >::value)
    {
        return GetArrayArgumentValueInternal(ArgumentName);
    }
    else
    {
        throw("Not implemented");
    }
}
inline xlw::MyMatrix xlw::ArgumentList::GetMatrixArgumentValue(const std::string& ArgumentName)
{
    if(isSameType<MyMatrix, NCMatrix>::value)
    {
        return GetMatrixArgumentValueInternal(ArgumentName);
    }
    else
    {
        throw("Not implemented");
    }
}
inline bool xlw::ArgumentList::GetIfPresent(const std::string& ArgumentName, MyArray& ArgumentValue)
{
    if(isSameType<MyArray, std::vector<double> >::value)
    {
        return GetIfPresentInternal(ArgumentName, ArgumentValue);
    }
    else
    {
        throw("Not implemented");
    }
}
inline bool xlw::ArgumentList::GetIfPresent(const std::string& ArgumentName, MyMatrix& ArgumentValue)
{
    if(isSameType<MyMatrix, NCMatrix>::value)
    {
        return GetIfPresentInternal(ArgumentName, ArgumentValue);
    }
    else
    {
        throw("Not implemented");
    }
}

#endif
