/*!
\file xlcall32.h
\brief Header file shipped with Excel 97

Declares the functions exported by the file \c xlcall32.dll
shipped with Excel as well. Microsoft is supposed to maintain
backward compatibility with this API in the future versions
of Ms Excel.

This file is documented here for information only if you want
to extend the wrapper with new functionalities. But this does
not really constitute a part of the wrapper but rather the
basis of it.
*/

// $Id$

#ifndef INC_excel32_H
#define INC_excel32_H

// jT: I wonder why we not include <WINDEF.H>
#ifndef STRICT
#define STRICT 1
#endif
#include "windows.h"

/*
*  Microsoft Excel Developer's Toolkit
*  Version 5.0
*
*  File:           INCLUDE\XLCALL.H
*  Description:    Header file for for Microsoft Excel callbacks
*  Platform:       Microsoft Windows
*
*  This file defines the constants and
*  data types which are used in the
*  Microsoft Excel C API. Include
*  <windows.h> before you include this.
*
*/

#include <xlw/xldata32.h>

#ifdef __cplusplus
extern "C" {
#endif	/*! __cplusplus */

/*!
* XLOPER data types
*
* Used for xltype field of XLOPER structure
*/
static const int xltypeNum        = 0x0001;
static const int xltypeStr        = 0x0002;
static const int xltypeBool       = 0x0004;
static const int xltypeRef        = 0x0008;
static const int xltypeErr        = 0x0010;
static const int xltypeFlow       = 0x0020;
static const int xltypeMulti      = 0x0040;
static const int xltypeMissing    = 0x0080;
static const int xltypeNil        = 0x0100;
static const int xltypeSRef       = 0x0400;
static const int xltypeInt        = 0x0800;
static const int xlbitXLFree      = 0x1000;
static const int xlbitDLLFree     = 0x4000;

static const int xltypeBigData    = (xltypeStr | xltypeInt);


/*!
* Error codes
*
* Used for val.err field of XLOPER structure
* when constructing error XLOPERs
*/

static const int xlerrNull    = 0;
static const int xlerrDiv0    = 7;
static const int xlerrValue   = 15;
static const int xlerrRef     = 23;
static const int xlerrName    = 29;
static const int xlerrNum     = 36;
static const int xlerrNA      = 42;


/*!
* Flow data types
*
* Used for val.flow.xlflow field of XLOPER structure
* when constructing flow-control XLOPERs
*/

static const int xlflowHalt       = 1;
static const int xlflowGoto       = 2;
static const int xlflowRestart    = 8;
static const int xlflowPause      = 16;
static const int xlflowResume     = 64;


/*!
* Function prototypes
*/

/*!  These functions prototypes were modified by jl
*   from function declaration
*   to pointer to function declaration
*   to use the xlcall32.dll instead of the .lib
*/

extern int (__cdecl *Excel4)(int xlfn, LPXLOPER operRes, int count,... );

/*! followed by count LPXLOPERs */

extern int (__stdcall *Excel4v)(int xlfn, LPXLOPER operRes, int count, LPXLOPER far opers[]);
extern int (__stdcall *XLCallVer)(void);


/*!
* Return codes
*
* These values can be returned from Excel4() or Excel4v().
*/

static const int xlretSuccess    = 0;     /*! success */ 
static const int xlretAbort      = 1;     /*! macro halted */
static const int xlretInvXlfn    = 2;     /*! invalid function number */ 
static const int xlretInvCount   = 4;     /*! invalid number of arguments */ 
static const int xlretInvXloper  = 8;     /*! invalid OPER structure */  
static const int xlretStackOvfl  = 16;    /*! stack overflow */  
static const int xlretFailed     = 32;    /*! command failed */  
static const int xlretUncalced   = 64;    /*! uncalced cell */


/*!
* Function number bits
*/

static const int xlCommand    = 0x8000;
static const int xlSpecial    = 0x4000;
static const int xlIntl       = 0x2000;
static const int xlPrompt     = 0x1000;


/*!
* Auxiliary function numbers
*
* These functions are available only from the C API,
* not from the Excel macro language.
*/

static const int xlFree             = (0  | xlSpecial);
static const int xlStack            = (1  | xlSpecial);
static const int xlCoerce           = (2  | xlSpecial);
static const int xlSet              = (3  | xlSpecial);
static const int xlSheetId          = (4  | xlSpecial);
static const int xlSheetNm          = (5  | xlSpecial);
static const int xlAbort            = (6  | xlSpecial);
static const int xlGetInst          = (7  | xlSpecial);
static const int xlGetHwnd          = (8  | xlSpecial);
static const int xlGetName          = (9  | xlSpecial);
static const int xlEnableXLMsgs     = (10 | xlSpecial);
static const int xlDisableXLMsgs    = (11 | xlSpecial);
static const int xlDefineBinaryName = (12 | xlSpecial);
static const int xlGetBinaryName    = (13 | xlSpecial);


/*!
* User defined function
*
* First argument should be a function reference.
*/

static const int xlUDF      = 255;


/*!
* Built-in Functions and Command Equivalents
*/


/*! Excel function numbers */

static const int xlfCount = 0;
static const int xlfIsna = 2;
static const int xlfIserror = 3;
static const int xlfSum = 4;
static const int xlfAverage = 5;
static const int xlfMin = 6;
static const int xlfMax = 7;
static const int xlfRow = 8;
static const int xlfColumn = 9;
static const int xlfNa = 10;
static const int xlfNpv = 11;
static const int xlfStdev = 12;
static const int xlfDollar = 13;
static const int xlfFixed = 14;
static const int xlfSin = 15;
static const int xlfCos = 16;
static const int xlfTan = 17;
static const int xlfAtan = 18;
static const int xlfPi = 19;
static const int xlfSqrt = 20;
static const int xlfExp = 21;
static const int xlfLn = 22;
static const int xlfLog10 = 23;
static const int xlfAbs = 24;
static const int xlfInt = 25;
static const int xlfSign = 26;
static const int xlfRound = 27;
static const int xlfLookup = 28;
static const int xlfIndex = 29;
static const int xlfRept = 30;
static const int xlfMid = 31;
static const int xlfLen = 32;
static const int xlfValue = 33;
static const int xlfTrue = 34;
static const int xlfFalse = 35;
static const int xlfAnd = 36;
static const int xlfOr = 37;
static const int xlfNot = 38;
static const int xlfMod = 39;
static const int xlfDcount = 40;
static const int xlfDsum = 41;
static const int xlfDaverage = 42;
static const int xlfDmin = 43;
static const int xlfDmax = 44;
static const int xlfDstdev = 45;
static const int xlfVar = 46;
static const int xlfDvar = 47;
static const int xlfText = 48;
static const int xlfLinest = 49;
static const int xlfTrend = 50;
static const int xlfLogest = 51;
static const int xlfGrowth = 52;
static const int xlfGoto = 53;
static const int xlfHalt = 54;
static const int xlfPv = 56;
static const int xlfFv = 57;
static const int xlfNper = 58;
static const int xlfPmt = 59;
static const int xlfRate = 60;
static const int xlfMirr = 61;
static const int xlfIrr = 62;
static const int xlfRand = 63;
static const int xlfMatch = 64;
static const int xlfDate = 65;
static const int xlfTime = 66;
static const int xlfDay = 67;
static const int xlfMonth = 68;
static const int xlfYear = 69;
static const int xlfWeekday = 70;
static const int xlfHour = 71;
static const int xlfMinute = 72;
static const int xlfSecond = 73;
static const int xlfNow = 74;
static const int xlfAreas = 75;
static const int xlfRows = 76;
static const int xlfColumns = 77;
static const int xlfOffset = 78;
static const int xlfAbsref = 79;
static const int xlfRelref = 80;
static const int xlfArgument = 81;
static const int xlfSearch = 82;
static const int xlfTranspose = 83;
static const int xlfError = 84;
static const int xlfStep = 85;
static const int xlfType = 86;
static const int xlfEcho = 87;
static const int xlfSetName = 88;
static const int xlfCaller = 89;
static const int xlfDeref = 90;
static const int xlfWindows = 91;
static const int xlfSeries = 92;
static const int xlfDocuments = 93;
static const int xlfActiveCell = 94;
static const int xlfSelection = 95;
static const int xlfResult = 96;
static const int xlfAtan2 = 97;
static const int xlfAsin = 98;
static const int xlfAcos = 99;
static const int xlfChoose = 100;
static const int xlfHlookup = 101;
static const int xlfVlookup = 102;
static const int xlfLinks = 103;
static const int xlfInput = 104;
static const int xlfIsref = 105;
static const int xlfGetFormula = 106;
static const int xlfGetName = 107;
static const int xlfSetValue = 108;
static const int xlfLog = 109;
static const int xlfExec = 110;
static const int xlfChar = 111;
static const int xlfLower = 112;
static const int xlfUpper = 113;
static const int xlfProper = 114;
static const int xlfLeft = 115;
static const int xlfRight = 116;
static const int xlfExact = 117;
static const int xlfTrim = 118;
static const int xlfReplace = 119;
static const int xlfSubstitute = 120;
static const int xlfCode = 121;
static const int xlfNames = 122;
static const int xlfDirectory = 123;
static const int xlfFind = 124;
static const int xlfCell = 125;
static const int xlfIserr = 126;
static const int xlfIstext = 127;
static const int xlfIsnumber = 128;
static const int xlfIsblank = 129;
static const int xlfT = 130;
static const int xlfN = 131;
static const int xlfFopen = 132;
static const int xlfFclose = 133;
static const int xlfFsize = 134;
static const int xlfFreadln = 135;
static const int xlfFread = 136;
static const int xlfFwriteln = 137;
static const int xlfFwrite = 138;
static const int xlfFpos = 139;
static const int xlfDatevalue = 140;
static const int xlfTimevalue = 141;
static const int xlfSln = 142;
static const int xlfSyd = 143;
static const int xlfDdb = 144;
static const int xlfGetDef = 145;
static const int xlfReftext = 146;
static const int xlfTextref = 147;
static const int xlfIndirect = 148;
static const int xlfRegister = 149;
static const int xlfCall = 150;
static const int xlfAddBar = 151;
static const int xlfAddMenu = 152;
static const int xlfAddCommand = 153;
static const int xlfEnableCommand = 154;
static const int xlfCheckCommand = 155;
static const int xlfRenameCommand = 156;
static const int xlfShowBar = 157;
static const int xlfDeleteMenu = 158;
static const int xlfDeleteCommand = 159;
static const int xlfGetChartItem = 160;
static const int xlfDialogBox = 161;
static const int xlfClean = 162;
static const int xlfMdeterm = 163;
static const int xlfMinverse = 164;
static const int xlfMmult = 165;
static const int xlfFiles = 166;
static const int xlfIpmt = 167;
static const int xlfPpmt = 168;
static const int xlfCounta = 169;
static const int xlfCancelKey = 170;
static const int xlfInitiate = 175;
static const int xlfRequest = 176;
static const int xlfPoke = 177;
static const int xlfExecute = 178;
static const int xlfTerminate = 179;
static const int xlfRestart = 180;
static const int xlfHelp = 181;
static const int xlfGetBar = 182;
static const int xlfProduct = 183;
static const int xlfFact = 184;
static const int xlfGetCell = 185;
static const int xlfGetWorkspace = 186;
static const int xlfGetWindow = 187;
static const int xlfGetDocument = 188;
static const int xlfDproduct = 189;
static const int xlfIsnontext = 190;
static const int xlfGetNote = 191;
static const int xlfNote = 192;
static const int xlfStdevp = 193;
static const int xlfVarp = 194;
static const int xlfDstdevp = 195;
static const int xlfDvarp = 196;
static const int xlfTrunc = 197;
static const int xlfIslogical = 198;
static const int xlfDcounta = 199;
static const int xlfDeleteBar = 200;
static const int xlfUnregister = 201;
static const int xlfUsdollar = 204;
static const int xlfFindb = 205;
static const int xlfSearchb = 206;
static const int xlfReplaceb = 207;
static const int xlfLeftb = 208;
static const int xlfRightb = 209;
static const int xlfMidb = 210;
static const int xlfLenb = 211;
static const int xlfRoundup = 212;
static const int xlfRounddown = 213;
static const int xlfAsc = 214;
static const int xlfDbcs = 215;
static const int xlfRank = 216;
static const int xlfAddress = 219;
static const int xlfDays360 = 220;
static const int xlfToday = 221;
static const int xlfVdb = 222;
static const int xlfMedian = 227;
static const int xlfSumproduct = 228;
static const int xlfSinh = 229;
static const int xlfCosh = 230;
static const int xlfTanh = 231;
static const int xlfAsinh = 232;
static const int xlfAcosh = 233;
static const int xlfAtanh = 234;
static const int xlfDget = 235;
static const int xlfCreateObject = 236;
static const int xlfVolatile = 237;
static const int xlfLastError = 238;
static const int xlfCustomUndo = 239;
static const int xlfCustomRepeat = 240;
static const int xlfFormulaConvert = 241;
static const int xlfGetLinkInfo = 242;
static const int xlfTextBox = 243;
static const int xlfInfo = 244;
static const int xlfGroup = 245;
static const int xlfGetObject = 246;
static const int xlfDb = 247;
static const int xlfPause = 248;
static const int xlfResume = 251;
static const int xlfFrequency = 252;
static const int xlfAddToolbar = 253;
static const int xlfDeleteToolbar = 254;
static const int xlfResetToolbar = 256;
static const int xlfEvaluate = 257;
static const int xlfGetToolbar = 258;
static const int xlfGetTool = 259;
static const int xlfSpellingCheck = 260;
static const int xlfErrorType = 261;
static const int xlfAppTitle = 262;
static const int xlfWindowTitle = 263;
static const int xlfSaveToolbar = 264;
static const int xlfEnableTool = 265;
static const int xlfPressTool = 266;
static const int xlfRegisterId = 267;
static const int xlfGetWorkbook = 268;
static const int xlfAvedev = 269;
static const int xlfBetadist = 270;
static const int xlfGammaln = 271;
static const int xlfBetainv = 272;
static const int xlfBinomdist = 273;
static const int xlfChidist = 274;
static const int xlfChiinv = 275;
static const int xlfCombin = 276;
static const int xlfConfidence = 277;
static const int xlfCritbinom = 278;
static const int xlfEven = 279;
static const int xlfExpondist = 280;
static const int xlfFdist = 281;
static const int xlfFinv = 282;
static const int xlfFisher = 283;
static const int xlfFisherinv = 284;
static const int xlfFloor = 285;
static const int xlfGammadist = 286;
static const int xlfGammainv = 287;
static const int xlfCeiling = 288;
static const int xlfHypgeomdist = 289;
static const int xlfLognormdist = 290;
static const int xlfLoginv = 291;
static const int xlfNegbinomdist = 292;
static const int xlfNormdist = 293;
static const int xlfNormsdist = 294;
static const int xlfNorminv = 295;
static const int xlfNormsinv = 296;
static const int xlfStandardize = 297;
static const int xlfOdd = 298;
static const int xlfPermut = 299;
static const int xlfPoisson = 300;
static const int xlfTdist = 301;
static const int xlfWeibull = 302;
static const int xlfSumxmy2 = 303;
static const int xlfSumx2my2 = 304;
static const int xlfSumx2py2 = 305;
static const int xlfChitest = 306;
static const int xlfCorrel = 307;
static const int xlfCovar = 308;
static const int xlfForecast = 309;
static const int xlfFtest = 310;
static const int xlfIntercept = 311;
static const int xlfPearson = 312;
static const int xlfRsq = 313;
static const int xlfSteyx = 314;
static const int xlfSlope = 315;
static const int xlfTtest = 316;
static const int xlfProb = 317;
static const int xlfDevsq = 318;
static const int xlfGeomean = 319;
static const int xlfHarmean = 320;
static const int xlfSumsq = 321;
static const int xlfKurt = 322;
static const int xlfSkew = 323;
static const int xlfZtest = 324;
static const int xlfLarge = 325;
static const int xlfSmall = 326;
static const int xlfQuartile = 327;
static const int xlfPercentile = 328;
static const int xlfPercentrank = 329;
static const int xlfMode = 330;
static const int xlfTrimmean = 331;
static const int xlfTinv = 332;
static const int xlfMovieCommand = 334;
static const int xlfGetMovie = 335;
static const int xlfConcatenate = 336;
static const int xlfPower = 337;
static const int xlfPivotAddData = 338;
static const int xlfGetPivotTable = 339;
static const int xlfGetPivotField = 340;
static const int xlfGetPivotItem = 341;
static const int xlfRadians = 342;
static const int xlfDegrees = 343;
static const int xlfSubtotal = 344;
static const int xlfSumif = 345;
static const int xlfCountif = 346;
static const int xlfCountblank = 347;
static const int xlfScenarioGet = 348;
static const int xlfOptionsListsGet = 349;
static const int xlfIspmt = 350;
static const int xlfDatedif = 351;
static const int xlfDatestring = 352;
static const int xlfNumberstring = 353;
static const int xlfRoman = 354;
static const int xlfOpenDialog = 355;
static const int xlfSaveDialog = 356;
static const int xlfViewGet = 357;
static const int xlfGetPivotData = 358;
static const int xlfHyperlink = 359;
static const int xlfPhonetic = 360;
static const int xlfAverageA = 361;
static const int xlfMaxA = 362;
static const int xlfMinA = 363;
static const int xlfStDevPA = 364;
static const int xlfVarPA = 365;
static const int xlfStDevA = 366;
static const int xlfVarA = 367;


/*! Excel command numbers */
static const int xlcBeep = (0 | xlCommand);
static const int xlcOpen = (1 | xlCommand);
static const int xlcOpenLinks = (2 | xlCommand);
static const int xlcCloseAll = (3 | xlCommand);
static const int xlcSave = (4 | xlCommand);
static const int xlcSaveAs = (5 | xlCommand);
static const int xlcFileDelete = (6 | xlCommand);
static const int xlcPageSetup = (7 | xlCommand);
static const int xlcPrint = (8 | xlCommand);
static const int xlcPrinterSetup = (9 | xlCommand);
static const int xlcQuit = (10 | xlCommand);
static const int xlcNewWindow = (11 | xlCommand);
static const int xlcArrangeAll = (12 | xlCommand);
static const int xlcWindowSize = (13 | xlCommand);
static const int xlcWindowMove = (14 | xlCommand);
static const int xlcFull = (15 | xlCommand);
static const int xlcClose = (16 | xlCommand);
static const int xlcRun = (17 | xlCommand);
static const int xlcSetPrintArea = (22 | xlCommand);
static const int xlcSetPrintTitles = (23 | xlCommand);
static const int xlcSetPageBreak = (24 | xlCommand);
static const int xlcRemovePageBreak = (25 | xlCommand);
static const int xlcFont = (26 | xlCommand);
static const int xlcDisplay = (27 | xlCommand);
static const int xlcProtectDocument = (28 | xlCommand);
static const int xlcPrecision = (29 | xlCommand);
static const int xlcA1R1c1 = (30 | xlCommand);
static const int xlcCalculateNow = (31 | xlCommand);
static const int xlcCalculation = (32 | xlCommand);
static const int xlcDataFind = (34 | xlCommand);
static const int xlcExtract = (35 | xlCommand);
static const int xlcDataDelete = (36 | xlCommand);
static const int xlcSetDatabase = (37 | xlCommand);
static const int xlcSetCriteria = (38 | xlCommand);
static const int xlcSort = (39 | xlCommand);
static const int xlcDataSeries = (40 | xlCommand);
static const int xlcTable = (41 | xlCommand);
static const int xlcFormatNumber = (42 | xlCommand);
static const int xlcAlignment = (43 | xlCommand);
static const int xlcStyle = (44 | xlCommand);
static const int xlcBorder = (45 | xlCommand);
static const int xlcCellProtection = (46 | xlCommand);
static const int xlcColumnWidth = (47 | xlCommand);
static const int xlcUndo = (48 | xlCommand);
static const int xlcCut = (49 | xlCommand);
static const int xlcCopy = (50 | xlCommand);
static const int xlcPaste = (51 | xlCommand);
static const int xlcClear = (52 | xlCommand);
static const int xlcPasteSpecial = (53 | xlCommand);
static const int xlcEditDelete = (54 | xlCommand);
static const int xlcInsert = (55 | xlCommand);
static const int xlcFillRight = (56 | xlCommand);
static const int xlcFillDown = (57 | xlCommand);
static const int xlcDefineName = (61 | xlCommand);
static const int xlcCreateNames = (62 | xlCommand);
static const int xlcFormulaGoto = (63 | xlCommand);
static const int xlcFormulaFind = (64 | xlCommand);
static const int xlcSelectLastCell = (65 | xlCommand);
static const int xlcShowActiveCell = (66 | xlCommand);
static const int xlcGalleryArea = (67 | xlCommand);
static const int xlcGalleryBar = (68 | xlCommand);
static const int xlcGalleryColumn = (69 | xlCommand);
static const int xlcGalleryLine = (70 | xlCommand);
static const int xlcGalleryPie = (71 | xlCommand);
static const int xlcGalleryScatter = (72 | xlCommand);
static const int xlcCombination = (73 | xlCommand);
static const int xlcPreferred = (74 | xlCommand);
static const int xlcAddOverlay = (75 | xlCommand);
static const int xlcGridlines = (76 | xlCommand);
static const int xlcSetPreferred = (77 | xlCommand);
static const int xlcAxes = (78 | xlCommand);
static const int xlcLegend = (79 | xlCommand);
static const int xlcAttachText = (80 | xlCommand);
static const int xlcAddArrow = (81 | xlCommand);
static const int xlcSelectChart = (82 | xlCommand);
static const int xlcSelectPlotArea = (83 | xlCommand);
static const int xlcPatterns = (84 | xlCommand);
static const int xlcMainChart = (85 | xlCommand);
static const int xlcOverlay = (86 | xlCommand);
static const int xlcScale = (87 | xlCommand);
static const int xlcFormatLegend = (88 | xlCommand);
static const int xlcFormatText = (89 | xlCommand);
static const int xlcEditRepeat = (90 | xlCommand);
static const int xlcParse = (91 | xlCommand);
static const int xlcJustify = (92 | xlCommand);
static const int xlcHide = (93 | xlCommand);
static const int xlcUnhide = (94 | xlCommand);
static const int xlcWorkspace = (95 | xlCommand);
static const int xlcFormula = (96 | xlCommand);
static const int xlcFormulaFill = (97 | xlCommand);
static const int xlcFormulaArray = (98 | xlCommand);
static const int xlcDataFindNext = (99 | xlCommand);
static const int xlcDataFindPrev = (100 | xlCommand);
static const int xlcFormulaFindNext = (101 | xlCommand);
static const int xlcFormulaFindPrev = (102 | xlCommand);
static const int xlcActivate = (103 | xlCommand);
static const int xlcActivateNext = (104 | xlCommand);
static const int xlcActivatePrev = (105 | xlCommand);
static const int xlcUnlockedNext = (106 | xlCommand);
static const int xlcUnlockedPrev = (107 | xlCommand);
static const int xlcCopyPicture = (108 | xlCommand);
static const int xlcSelect = (109 | xlCommand);
static const int xlcDeleteName = (110 | xlCommand);
static const int xlcDeleteFormat = (111 | xlCommand);
static const int xlcVline = (112 | xlCommand);
static const int xlcHline = (113 | xlCommand);
static const int xlcVpage = (114 | xlCommand);
static const int xlcHpage = (115 | xlCommand);
static const int xlcVscroll = (116 | xlCommand);
static const int xlcHscroll = (117 | xlCommand);
static const int xlcAlert = (118 | xlCommand);
static const int xlcNew = (119 | xlCommand);
static const int xlcCancelCopy = (120 | xlCommand);
static const int xlcShowClipboard = (121 | xlCommand);
static const int xlcMessage = (122 | xlCommand);
static const int xlcPasteLink = (124 | xlCommand);
static const int xlcAppActivate = (125 | xlCommand);
static const int xlcDeleteArrow = (126 | xlCommand);
static const int xlcRowHeight = (127 | xlCommand);
static const int xlcFormatMove = (128 | xlCommand);
static const int xlcFormatSize = (129 | xlCommand);
static const int xlcFormulaReplace = (130 | xlCommand);
static const int xlcSendKeys = (131 | xlCommand);
static const int xlcSelectSpecial = (132 | xlCommand);
static const int xlcApplyNames = (133 | xlCommand);
static const int xlcReplaceFont = (134 | xlCommand);
static const int xlcFreezePanes = (135 | xlCommand);
static const int xlcShowInfo = (136 | xlCommand);
static const int xlcSplit = (137 | xlCommand);
static const int xlcOnWindow = (138 | xlCommand);
static const int xlcOnData = (139 | xlCommand);
static const int xlcDisableInput = (140 | xlCommand);
static const int xlcEcho = (141 | xlCommand);
static const int xlcOutline = (142 | xlCommand);
static const int xlcListNames = (143 | xlCommand);
static const int xlcFileClose = (144 | xlCommand);
static const int xlcSaveWorkbook = (145 | xlCommand);
static const int xlcDataForm = (146 | xlCommand);
static const int xlcCopyChart = (147 | xlCommand);
static const int xlcOnTime = (148 | xlCommand);
static const int xlcWait = (149 | xlCommand);
static const int xlcFormatFont = (150 | xlCommand);
static const int xlcFillUp = (151 | xlCommand);
static const int xlcFillLeft = (152 | xlCommand);
static const int xlcDeleteOverlay = (153 | xlCommand);
static const int xlcNote = (154 | xlCommand);
static const int xlcShortMenus = (155 | xlCommand);
static const int xlcSetUpdateStatus = (159 | xlCommand);
static const int xlcColorPalette = (161 | xlCommand);
static const int xlcDeleteStyle = (162 | xlCommand);
static const int xlcWindowRestore = (163 | xlCommand);
static const int xlcWindowMaximize = (164 | xlCommand);
static const int xlcError = (165 | xlCommand);
static const int xlcChangeLink = (166 | xlCommand);
static const int xlcCalculateDocument = (167 | xlCommand);
static const int xlcOnKey = (168 | xlCommand);
static const int xlcAppRestore = (169 | xlCommand);
static const int xlcAppMove = (170 | xlCommand);
static const int xlcAppSize = (171 | xlCommand);
static const int xlcAppMinimize = (172 | xlCommand);
static const int xlcAppMaximize = (173 | xlCommand);
static const int xlcBringToFront = (174 | xlCommand);
static const int xlcSendToBack = (175 | xlCommand);
static const int xlcMainChartType = (185 | xlCommand);
static const int xlcOverlayChartType = (186 | xlCommand);
static const int xlcSelectEnd = (187 | xlCommand);
static const int xlcOpenMail = (188 | xlCommand);
static const int xlcSendMail = (189 | xlCommand);
static const int xlcStandardFont = (190 | xlCommand);
static const int xlcConsolidate = (191 | xlCommand);
static const int xlcSortSpecial = (192 | xlCommand);
static const int xlcGallery3dArea = (193 | xlCommand);
static const int xlcGallery3dColumn = (194 | xlCommand);
static const int xlcGallery3dLine = (195 | xlCommand);
static const int xlcGallery3dPie = (196 | xlCommand);
static const int xlcView3d = (197 | xlCommand);
static const int xlcGoalSeek = (198 | xlCommand);
static const int xlcWorkgroup = (199 | xlCommand);
static const int xlcFillGroup = (200 | xlCommand);
static const int xlcUpdateLink = (201 | xlCommand);
static const int xlcPromote = (202 | xlCommand);
static const int xlcDemote = (203 | xlCommand);
static const int xlcShowDetail = (204 | xlCommand);
static const int xlcUngroup = (206 | xlCommand);
static const int xlcObjectProperties = (207 | xlCommand);
static const int xlcSaveNewObject = (208 | xlCommand);
static const int xlcShare = (209 | xlCommand);
static const int xlcShareName = (210 | xlCommand);
static const int xlcDuplicate = (211 | xlCommand);
static const int xlcApplyStyle = (212 | xlCommand);
static const int xlcAssignToObject = (213 | xlCommand);
static const int xlcObjectProtection = (214 | xlCommand);
static const int xlcHideObject = (215 | xlCommand);
static const int xlcSetExtract = (216 | xlCommand);
static const int xlcCreatePublisher = (217 | xlCommand);
static const int xlcSubscribeTo = (218 | xlCommand);
static const int xlcAttributes = (219 | xlCommand);
static const int xlcShowToolbar = (220 | xlCommand);
static const int xlcPrintPreview = (222 | xlCommand);
static const int xlcEditColor = (223 | xlCommand);
static const int xlcShowLevels = (224 | xlCommand);
static const int xlcFormatMain = (225 | xlCommand);
static const int xlcFormatOverlay = (226 | xlCommand);
static const int xlcOnRecalc = (227 | xlCommand);
static const int xlcEditSeries = (228 | xlCommand);
static const int xlcDefineStyle = (229 | xlCommand);
static const int xlcLinePrint = (240 | xlCommand);
static const int xlcEnterData = (243 | xlCommand);
static const int xlcGalleryRadar = (249 | xlCommand);
static const int xlcMergeStyles = (250 | xlCommand);
static const int xlcEditionOptions = (251 | xlCommand);
static const int xlcPastePicture = (252 | xlCommand);
static const int xlcPastePictureLink = (253 | xlCommand);
static const int xlcSpelling = (254 | xlCommand);
static const int xlcZoom = (256 | xlCommand);
static const int xlcResume = (258 | xlCommand);
static const int xlcInsertObject = (259 | xlCommand);
static const int xlcWindowMinimize = (260 | xlCommand);
static const int xlcSize = (261 | xlCommand);
static const int xlcMove = (262 | xlCommand);
static const int xlcSoundNote = (265 | xlCommand);
static const int xlcSoundPlay = (266 | xlCommand);
static const int xlcFormatShape = (267 | xlCommand);
static const int xlcExtendPolygon = (268 | xlCommand);
static const int xlcFormatAuto = (269 | xlCommand);
static const int xlcGallery3dBar = (272 | xlCommand);
static const int xlcGallery3dSurface = (273 | xlCommand);
static const int xlcFillAuto = (274 | xlCommand);
static const int xlcCustomizeToolbar = (276 | xlCommand);
static const int xlcAddTool = (277 | xlCommand);
static const int xlcEditObject = (278 | xlCommand);
static const int xlcOnDoubleclick = (279 | xlCommand);
static const int xlcOnEntry = (280 | xlCommand);
static const int xlcWorkbookAdd = (281 | xlCommand);
static const int xlcWorkbookMove = (282 | xlCommand);
static const int xlcWorkbookCopy = (283 | xlCommand);
static const int xlcWorkbookOptions = (284 | xlCommand);
static const int xlcSaveWorkspace = (285 | xlCommand);
static const int xlcChartWizard = (288 | xlCommand);
static const int xlcDeleteTool = (289 | xlCommand);
static const int xlcMoveTool = (290 | xlCommand);
static const int xlcWorkbookSelect = (291 | xlCommand);
static const int xlcWorkbookActivate = (292 | xlCommand);
static const int xlcAssignToTool = (293 | xlCommand);
static const int xlcCopyTool = (295 | xlCommand);
static const int xlcResetTool = (296 | xlCommand);
static const int xlcConstrainNumeric = (297 | xlCommand);
static const int xlcPasteTool = (298 | xlCommand);
static const int xlcPlacement = (300 | xlCommand);
static const int xlcFillWorkgroup = (301 | xlCommand);
static const int xlcWorkbookNew = (302 | xlCommand);
static const int xlcScenarioCells = (305 | xlCommand);
static const int xlcScenarioDelete = (306 | xlCommand);
static const int xlcScenarioAdd = (307 | xlCommand);
static const int xlcScenarioEdit = (308 | xlCommand);
static const int xlcScenarioShow = (309 | xlCommand);
static const int xlcScenarioShowNext = (310 | xlCommand);
static const int xlcScenarioSummary = (311 | xlCommand);
static const int xlcPivotTableWizard = (312 | xlCommand);
static const int xlcPivotFieldProperties = (313 | xlCommand);
static const int xlcPivotField = (314 | xlCommand);
static const int xlcPivotItem = (315 | xlCommand);
static const int xlcPivotAddFields = (316 | xlCommand);
static const int xlcOptionsCalculation = (318 | xlCommand);
static const int xlcOptionsEdit = (319 | xlCommand);
static const int xlcOptionsView = (320 | xlCommand);
static const int xlcAddinManager = (321 | xlCommand);
static const int xlcMenuEditor = (322 | xlCommand);
static const int xlcAttachToolbars = (323 | xlCommand);
static const int xlcVbaReset = (324 | xlCommand);
static const int xlcOptionsChart = (325 | xlCommand);
static const int xlcStart = (326 | xlCommand);
static const int xlcVbaEnd = (327 | xlCommand);
static const int xlcVbaInsertFile = (328 | xlCommand);
static const int xlcVbaProcedureDefinition = (330 | xlCommand);
static const int xlcVbaReferences = (331 | xlCommand);
static const int xlcVbaStepInto = (332 | xlCommand);
static const int xlcVbaStepOver = (333 | xlCommand);
static const int xlcVbaToggleBreakpoint = (334 | xlCommand);
static const int xlcVbaClearBreakpoints = (335 | xlCommand);
static const int xlcRoutingSlip = (336 | xlCommand);
static const int xlcRouteDocument = (338 | xlCommand);
static const int xlcMailLogon = (339 | xlCommand);
static const int xlcInsertPicture = (342 | xlCommand);
static const int xlcEditTool = (343 | xlCommand);
static const int xlcGalleryDoughnut = (344 | xlCommand);
static const int xlcVbaObjectBrowser = (345 | xlCommand);
static const int xlcVbaDebugWindow = (346 | xlCommand);
static const int xlcVbaAddWatch = (347 | xlCommand);
static const int xlcVbaEditWatch = (348 | xlCommand);
static const int xlcVbaInstantWatch = (349 | xlCommand);
static const int xlcChartTrend = (350 | xlCommand);
static const int xlcPivotItemProperties = (352 | xlCommand);
static const int xlcWorkbookInsert = (354 | xlCommand);
static const int xlcOptionsTransition = (355 | xlCommand);
static const int xlcOptionsGeneral = (356 | xlCommand);
static const int xlcFilterAdvanced = (370 | xlCommand);
static const int xlcMailAddMailer = (373 | xlCommand);
static const int xlcMailDeleteMailer = (374 | xlCommand);
static const int xlcMailReply = (375 | xlCommand);
static const int xlcMailReplyAll = (376 | xlCommand);
static const int xlcMailForward = (377 | xlCommand);
static const int xlcMailNextLetter = (378 | xlCommand);
static const int xlcDataLabel = (379 | xlCommand);
static const int xlcInsertTitle = (380 | xlCommand);
static const int xlcFontProperties = (381 | xlCommand);
static const int xlcMacroOptions = (382 | xlCommand);
static const int xlcWorkbookHide = (383 | xlCommand);
static const int xlcWorkbookUnhide = (384 | xlCommand);
static const int xlcWorkbookDelete = (385 | xlCommand);
static const int xlcWorkbookName = (386 | xlCommand);
static const int xlcGalleryCustom = (388 | xlCommand);
static const int xlcAddChartAutoformat = (390 | xlCommand);
static const int xlcDeleteChartAutoformat = (391 | xlCommand);
static const int xlcChartAddData = (392 | xlCommand);
static const int xlcAutoOutline = (393 | xlCommand);
static const int xlcTabOrder = (394 | xlCommand);
static const int xlcShowDialog = (395 | xlCommand);
static const int xlcSelectAll = (396 | xlCommand);
static const int xlcUngroupSheets = (397 | xlCommand);
static const int xlcSubtotalCreate = (398 | xlCommand);
static const int xlcSubtotalRemove = (399 | xlCommand);
static const int xlcRenameObject = (400 | xlCommand);
static const int xlcWorkbookScroll = (412 | xlCommand);
static const int xlcWorkbookNext = (413 | xlCommand);
static const int xlcWorkbookPrev = (414 | xlCommand);
static const int xlcWorkbookTabSplit = (415 | xlCommand);
static const int xlcFullScreen = (416 | xlCommand);
static const int xlcWorkbookProtect = (417 | xlCommand);
static const int xlcScrollbarProperties = (420 | xlCommand);
static const int xlcPivotShowPages = (421 | xlCommand);
static const int xlcTextToColumns = (422 | xlCommand);
static const int xlcFormatCharttype = (423 | xlCommand);
static const int xlcLinkFormat = (424 | xlCommand);
static const int xlcTracerDisplay = (425 | xlCommand);
static const int xlcTracerNavigate = (430 | xlCommand);
static const int xlcTracerClear = (431 | xlCommand);
static const int xlcTracerError = (432 | xlCommand);
static const int xlcPivotFieldGroup = (433 | xlCommand);
static const int xlcPivotFieldUngroup = (434 | xlCommand);
static const int xlcCheckboxProperties = (435 | xlCommand);
static const int xlcLabelProperties = (436 | xlCommand);
static const int xlcListboxProperties = (437 | xlCommand);
static const int xlcEditboxProperties = (438 | xlCommand);
static const int xlcPivotRefresh = (439 | xlCommand);
static const int xlcLinkCombo = (440 | xlCommand);
static const int xlcOpenText = (441 | xlCommand);
static const int xlcHideDialog = (442 | xlCommand);
static const int xlcSetDialogFocus = (443 | xlCommand);
static const int xlcEnableObject = (444 | xlCommand);
static const int xlcPushbuttonProperties = (445 | xlCommand);
static const int xlcSetDialogDefault = (446 | xlCommand);
static const int xlcFilter = (447 | xlCommand);
static const int xlcFilterShowAll = (448 | xlCommand);
static const int xlcClearOutline = (449 | xlCommand);
static const int xlcFunctionWizard = (450 | xlCommand);
static const int xlcAddListItem = (451 | xlCommand);
static const int xlcSetListItem = (452 | xlCommand);
static const int xlcRemoveListItem = (453 | xlCommand);
static const int xlcSelectListItem = (454 | xlCommand);
static const int xlcSetControlValue = (455 | xlCommand);
static const int xlcSaveCopyAs = (456 | xlCommand);
static const int xlcOptionsListsAdd = (458 | xlCommand);
static const int xlcOptionsListsDelete = (459 | xlCommand);
static const int xlcSeriesAxes = (460 | xlCommand);
static const int xlcSeriesX = (461 | xlCommand);
static const int xlcSeriesY = (462 | xlCommand);
static const int xlcErrorbarX = (463 | xlCommand);
static const int xlcErrorbarY = (464 | xlCommand);
static const int xlcFormatChart = (465 | xlCommand);
static const int xlcSeriesOrder = (466 | xlCommand);
static const int xlcMailLogoff = (467 | xlCommand);
static const int xlcClearRoutingSlip = (468 | xlCommand);
static const int xlcAppActivateMicrosoft = (469 | xlCommand);
static const int xlcMailEditMailer = (470 | xlCommand);
static const int xlcOnSheet = (471 | xlCommand);
static const int xlcStandardWidth = (472 | xlCommand);
static const int xlcScenarioMerge = (473 | xlCommand);
static const int xlcSummaryInfo = (474 | xlCommand);
static const int xlcFindFile = (475 | xlCommand);
static const int xlcActiveCellFont = (476 | xlCommand);
static const int xlcEnableTipwizard = (477 | xlCommand);
static const int xlcVbaMakeAddin = (478 | xlCommand);
static const int xlcMailSendMailer = (482 | xlCommand);

#ifdef __cplusplus
}			/*! End of extern "C" { */
#endif	/*! __cplusplus */


#endif