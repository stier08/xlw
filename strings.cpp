#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>

#define DLLEXPORT extern "C" __declspec(dllexport)

DLLEXPORT int xlAutoOpen() {
    static XLOPER12 xDll;
    Excel12f(xlGetName, &xDll, 0);

    Excel12f(xlfRegister, 0, 7, &xDll,
        TempStr12(L"str_len1"),     // function code name
        TempStr12(L"JQ"),           // parameter codes
        TempStr12(L"str_len1"),     // function display name
        TempStr12(L""),             // comma-delimited list of parameters
        TempStr12(L"1"),            // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
        TempStr12(L"Strings"));     // function category

    Excel12f(xlfRegister, 0, 7, &xDll,
        TempStr12(L"str_len2"),     // function code name
        TempStr12(L"JQ"),           // parameter codes
        TempStr12(L"str_len2"),     // function display name
        TempStr12(L""),             // comma-delimited list of parameters
        TempStr12(L"1"),            // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
        TempStr12(L"Strings"));     // function category

    Excel12f(xlfRegister, 0, 7, &xDll,
        TempStr12(L"str_len3"),     // function code name
        TempStr12(L"JU"),           // parameter codes
        TempStr12(L"str_len3"),     // function display name
        TempStr12(L""),             // comma-delimited list of parameters
        TempStr12(L"1"),            // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
        TempStr12(L"Strings"));     // function category

    Excel12f(xlfRegister, 0, 7, &xDll,
        TempStr12(L"str_len4"),     // function code name
        TempStr12(L"JU"),           // parameter codes
        TempStr12(L"str_len4"),     // function display name
        TempStr12(L""),             // comma-delimited list of parameters
        TempStr12(L"1"),            // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
        TempStr12(L"Strings"));     // function category

    Excel12f(xlfRegister, 0, 7, &xDll,
        TempStr12(L"str_len5"),     // function code name
        TempStr12(L"JC%"),          // parameter codes
        TempStr12(L"str_len5"),     // function display name
        TempStr12(L""),             // comma-delimited list of parameters
        TempStr12(L"1"),            // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
        TempStr12(L"Strings"));     // function category
                                
    Excel12f(xlFree, 0, 1, &xDll);
    return 1;
}

DLLEXPORT int str_len1(LPXLOPER12 in) {
    if (xltypeStr != in->xltype)
        return -1;
    unsigned short ret = in->val.str[0];
    return ret;
}

DLLEXPORT int str_len2(LPXLOPER12 in) {
    XLOPER12 xStr;
    if (xlretSuccess != Excel12f(xlCoerce, &xStr, 2, in, TempInt12(xltypeStr)))
        return -1;
    unsigned short ret = xStr.val.str[0];
    Excel12f(xlFree, 0, 1, &xStr);
    return ret;
}

DLLEXPORT int str_len3(LPXLOPER12 in) {
    if (xltypeStr != in->xltype)
        return -1;
    unsigned short ret = in->val.str[0];
    return ret;
}

DLLEXPORT int str_len4(LPXLOPER12 in) {
    XLOPER12 xStr;
    if (xlretSuccess != Excel12f(xlCoerce, &xStr, 2, in, TempInt12(xltypeStr)))
        return -1;
    unsigned short ret = xStr.val.str[0];
    Excel12f(xlFree, 0, 1, &xStr);
    return ret;
}

DLLEXPORT int str_len5(wchar_t *in) {
    return static_cast<int>(wcslen(in));
}
