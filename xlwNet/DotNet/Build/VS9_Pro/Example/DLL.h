#ifndef DLL_H
#define DLL_H

#include <xlwDotNet.h>

//<xlw:libraryname=DLL


std::wstring //tests empty args 
EmptyArgFunction();

short //echoes a short 
EchoShort(short x // number to be echoed
);

MyMatrix //echoes a matrix
EchoMat(const MyMatrix&  Echoee // argument to be echoed
);

MyArray //echoes an array
EchoArray(const MyArray&  Echoee //  argument to be echoed
);

CellMatrix //echoes a  CellMatrix
EchoCells(const CellMatrix&  Echoee //  argument to be echoed
);

double //computes the circumference of a circle 
Circ(double Diameter // the circle's diameter
);

std::wstring //Concatenates two strings
Concat(const std::wstring&  str1, // first string
const std::wstring&  str2 // second string
);

MyArray //computes mean and variance of a range 
stats(const MyArray&  data // input for computation
);

double //throws an exception
Exception();

std::wstring //says hello name
HelloWorldAgain(const std::wstring&  name // name to be echoed
);

CellMatrix //echoes arg list
EchoArgList(const ArgumentList&  args //  arguments to echo
);

MyArray //takes double[]
CastToCSArray(const MyArray&  csarray //  double Array
);

MyArray //takes double[]
CastToCSArrayTwice(const MyArray&  csarray //  double Array
);

MyMatrix //takes double[,]
CastToCSMatrix(const MyMatrix&  csmatrix //  double Array
);

MyMatrix //takes double[,]
CastToCSMatrixTwice(const MyMatrix&  csmatrix //  double Array
);

double //throws an exception of type ArgumentNullException
throwString(const std::wstring&  err // Just any random string 
);

double //throws an exception of type cellMatrixException
throwCellMatrix(const std::wstring&  err // Just any random string 
);

double //makes the C Runtime throw an exception
throwCError(const std::wstring&  err // Just any random string 
);

#endif 
