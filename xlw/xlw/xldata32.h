/*!
\file xldata32.h
\brief Data types from the header file shipped with Excel 97

Declares the datatypes that are used to communicate with Excel.
*/

// $Id$

#ifndef INC_xldata32_H
#define INC_xldata32_H

// include the guards in the MIDL output
#if defined(__midl)
#pragma midl_echo("#ifndef INC_xldata32_H") 
#pragma midl_echo("#define INC_xldata32_H") 
#endif


/*!
* XLREF structure
*
* Describes a single rectangular reference
*/

typedef struct xlref
{
    WORD rwFirst;       //!< first row.
    WORD rwLast;        //!< last row.
    BYTE colFirst;      //!< first column.
    BYTE colLast;       //!< last column.
} XLREF, *LPXLREF;

/*!
 * XLMREF structure
 *
 * Describes multiple rectangular references.
 * This is a variable size structure, default
 * size is 1 reference.
 */

typedef struct xlmref
{
    WORD count;			/*!< \brief Nb of XLREF. */
    XLREF reftbl[1];	/*!< \brief Array of XLREF, actually actually reftbl[count] */
} XLMREF, *LPXLMREF;

/*!
 * XLOPER structure
 *
 * Excel's fundamental data type: can hold data
 * of any type. Use "R" as the argument type in the
 * REGISTER function.
 */

typedef struct xloper
{
    union
    {
        double num;                     /*!< \brief xltypeNum */
        LPSTR str;                      /*!< \brief xltypeStr */
        WORD bool_;                      /*!< \brief xltypeBool */
        WORD err;                       /*!< \brief xltypeErr */
        short int w;                    /*!< \brief xltypeInt */
        struct
        {
            WORD count;                 /*!< \brief always = 1 */
            XLREF ref;
        } sref;                         /*!< \brief  xltypeSRef */
        struct
        {
            XLMREF *lpmref;
            DWORD idSheet;
        } mref;                         /*!< \brief  xltypeRef */
        struct
        {
            struct xloper *lparray;
            WORD rows;
            WORD columns;
        } array;                        /*!< \brief  xltypeMulti */
        struct
        {
            union
            {
                short int level;        /*!< \brief  xlflowRestart */
                short int tbctrl;       /*!< \brief  xlflowPause */
                DWORD idSheet;          /*!< \brief  xlflowGoto */
            } valflow;
            WORD rw;                    /*!< \brief  xlflowGoto */
            BYTE col;                   /*!< \brief  xlflowGoto */
            BYTE xlflow;
        } flow;                         /*!< \brief  xltypeFlow */
        struct
        {
            union
            {
                BYTE *lpbData;      /*!< \brief  data passed to XL */
                HANDLE hdata;           /*!< \brief  data returned from XL */
            } h;
            long cbData;
        } bigdata;                      /*!< \brief  xltypeBigData */
    } val;                              /*!< \brief data bits*/
    WORD xltype;                        /*!< \brief type flag */
} XLOPER, *LPXLOPER;


/*!
 * FP structure 
 *
 * Excel's array data type
 * Use "O" as the argument type in the REGISTER function.
 */

typedef struct _FP
{
	unsigned short int rows;
	unsigned short int columns;
	double array[1];			/*!< \brief Actually, array[rows][columns] */
} FP, *LPFP;



#if defined(__midl)
#pragma midl_echo("#endif // INC_xldata32_H") 
#endif

#endif // INC_xldata32_H