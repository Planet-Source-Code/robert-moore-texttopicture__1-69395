VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'General Declarations
Option Explicit
Private Type BROWSEINFO  ' Folder Dialog
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Private Type OPENFILENAME 'Open & Save Dialog
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type CHOOSECOLOR 'Color Dialog
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Const LF_FACESIZE = 32 'Font Dialog
Private Type LOGFONT 'Font Dialog
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type ChooseFont 'Font Dialog
    lStructSize As Long
    hWndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
' extra font constant
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const SCREEN_FONTTYPE = &H2000
Const BOLD_FONTTYPE = &H100
Const FW_BOLD = 700
Const LOGPIXELSY = 90

Private Type PrintDlg 'PrintDialog
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Const CCHDEVICENAME = 32 'PrintDialog
Const CCHFORMNAME = 32 'PrintDialog
Private Type DEVMODE 'PrintDialog
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Type DEVNAMES 'PrintDialog
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
'extra printer constants - for Printer Dialog
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
' memory management constants - for Printer Dialog
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40


' ------------- Dialog calling functions
' -------------- Standard
'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
'Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
'Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hOwner As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Declare Function SHAbout Lib "Shell32" Alias "ShellAboutA" (ByVal hOwner As Long, ByVal sAppName As String, ByVal sPrompt As String, ByVal hicon As Long) As Long

' GDI functions
' For Font Dialog
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long


' user32 functions
' For Font and Printer Dialog
Private Declare Function GetDC Lib "user32" (ByVal hOwner As Long) As Long

' kernel32 functions
' For Font Dialog
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (p1 As Any, p2 As Any) As Long
' For Printer Dialog
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)


' common dialog action types

Public Enum CdlgExt_Actions
       cdlgOpen = 1
       cdlgSave = 2
       cdlgColor = 3
       cdlgFont = 4
       cdlgPrinter = 5
       cdlgHelp = 6
       cdlgAbout = 7
       cdlgFolder = 8
       cdlgFormat = 9
       cdlgIcon = 10
   '    cdlgObjectProp = 11
       cdlgRestart = 12
       cdlgRun = 13
       cdlgShutDown = 14
 End Enum
 
 Public Enum CdlgEx_IconSize
       IconSizeSmall = 16
       IconSizeLarge = 32
 End Enum

' --------------- Enum Flags r 10-2-05
Public Enum CdlgExt_Flags
 ' Open & Save Dialog
 cdlOFNAllowMultiselect = &H200
 cdlOFNCreatePrompt = &H2000
 cdlOFNExplorer = &H80000
 cdlOFNExtensionDifferent = &H400
 cdlOFNFileMustExist = &H1000
 cdlOFNHelpButton = &H10
 cdlOFNHideReadOnly = &H4
 cdlOFNLongNames = &H200000
 cdlOFNNoChangeDir = &H8
 cdlOFNNoDereferenceLinks = &H100000
 cdlOFNNoLongNames = &H40000
 cdlOFNNoReadOnlyReturn = &H8000
 cdlOFNNoValidate = &H100
 cdlOFNOverwritePrompt = &H2
 cdlOFNPathMustExist = &H800
 cdlOFNReadOnly = &H1
 cdlOFNShareAware = &H4000
 'Color Dialog
 cdlCCFullOpen = &H2
 cdlCCHelpButton = &H8
 cdlCCPreventFullOpen = &H4
 cdlCCRGBInit = &H1
' Printer Dialog
 cdlPDAllPages = &H0
 cdlPDCollate = &H10
 cdlPDDisablePrintToFile = &H80000
 cdlPDHelpButton = &H800
 cdlPDHidePrintToFile = &H100000
 cdlPDNoPageNums = &H8
 cdlPDNoSelection = &H4
 cdlPDNoWarning = &H80
 cdlPDPageNums = &H2
 cdlPDPrintSetup = &H40
 cdlPDPrintToFile = &H20
 cdlPDReturnDC = &H100
 cdlPDReturnDefault = &H400
 cdlPDReturnIC = &H200
 cdlPDSelection = &H1
 cdlPDUseDevModeCopies = &H40000
' Font Dialog
 cdlCFANSIOnly = &H400
 cdlCFApply = &H200
 cdlCFBoth = &H3
 cdlCFEffects = &H100
 cdlCFFixedPitchOnly = &H4000
 cdlCFForceFontExist = &H10000
 cdlCFHelpButton = &H4
 cdlCFLimitSize = &H2000
 cdlCFNoFaceSel = &H80000
 cdlCFNoSimulations = &H1000
 cdlCFNoSizeSel = &H200000
 cdlCFNoStyleSel = &H100000
 cdlCFNoVectorFonts = &H800
 cdlCFPrinterFonts = &H2
 cdlCFScalableOnly = &H20000
 cdlCFScreenFonts = &H1
 cdlCFTTOnly = &H40000
 cdlCFWYSIWYG = &H8000
' Other Dialog
 'Restart Dialog
 Restart_Logoff = &H0
 Restart_ShutDown = &H1
 Restart_Reboot = &H2
 Restart_Force = &H4
 ' Run Dialog
 Run_NoBrowse = &H10
 Run_NoDefault = &H20
 Run_CalcDir = &H40
 Run_NoLable = &H80
 ' Properties Dialog
 ObjProp_System = &H0
 ObjProp_Printer = &H100
 ObjProp_File = &H200
 ObjProp_Mouse = &H300
 ObjProp_Locale = &H400
 ObjProp_MMedia = &H500
 ObjProp_TimeDate = &H600
 ObjProp_Network = &H700
 ObjProp_Screen = &H800
 ObjProp_Internet = &H900
 ' Browse for Folder Dialog
 Folder_COMPUTER = &H1000
 Folder_PRINTER = &H2000
 Folder_INCLUDEFILES = &H4001
End Enum
'Enum Help Commands
Public Enum CdlgExt_HelpCommand
 HelpCommandHelp = &H102&
 HelpContents = &H3&
 HelpContext = &H1
 HelpContextPOPUP = &H8&
 HelpForceFile = &H9&
 HelpHelpOnHelp = &H4
 HelpIndex = &H3
 HelpKeyHelp = &H101
 HelpPartialKey = &H105&
 HelpQuit = &H2
 HelpSetContents = &H5&
 HelpSetIndex = &H5
 HelpMultiKey = &H201&
 HelpSetWinPos = &H203&
End Enum

Private retvalue As Long 'General
Const MAX_PATH = 260 'General
Private OFN As OPENFILENAME ' Open & Save Dialog

'Inner variables for properties
Private mFileName As String
Private mFileTitle As String
Private mhOwner As Long
Private mDialogTitle As String
Private mFilter As String
Private mInitDir As String
Private mDefaultExt As String
Private mFilterIndex As Long
Private mHelpFile As String
Private mHelpCommand As CdlgExt_HelpCommand
Private mHelpKey As Long
Private mRGBResult As Long
Private mItalic As Boolean
Private mUnderline As Boolean
Private mStrikethru As Boolean
Private mFontName As String
Private mFontSize As Long
Private mBold As Boolean
Private mDialogPrompt As String
Private mFlags As CdlgExt_Flags
Private mCancelError As Boolean
Private mhIcon As Long
Private mAppName As String
Private mIconSize As CdlgEx_IconSize

Public Property Let CancelError(ByVal vData As Boolean)
   mCancelError = vData
End Property

Public Property Get CancelError() As Boolean
  CancelError = mCancelError
End Property

Public Property Get hOwner() As Long
    hOwner = mhOwner
End Property

Public Property Let hOwner(ByVal New_hOwner As Long)
    mhOwner = New_hOwner
End Property

Public Property Get Flags() As CdlgExt_Flags
    Flags = mFlags
End Property

Public Property Let Flags(ByVal New_Flags As CdlgExt_Flags)
    mFlags = New_Flags
End Property

Public Property Get DialogTitle() As String
   DialogTitle = mDialogTitle
End Property

Public Property Let DialogTitle(sTitle As String)
   mDialogTitle = sTitle
End Property

Public Property Get DialogPrompt() As String
    DialogPrompt = mDialogPrompt
End Property

Public Property Let DialogPrompt(ByVal New_Prompt As String)
    mDialogPrompt = New_Prompt
End Property

Public Property Get AppName() As String
    AppName = mAppName
End Property

Public Property Let AppName(ByVal New_AppName As String)
    mAppName = New_AppName
End Property

Public Property Let hicon(ByVal vData As Long)
    mhIcon = vData
End Property

Public Property Get hicon() As Long
   hicon = mhIcon
End Property

' Font Properties
Public Property Get Bold() As Boolean
  Bold = mBold
End Property

Public Property Let Bold(bBold As Boolean)
   mBold = bBold
End Property

Public Property Get FontName() As String
   FontName = mFontName
End Property

Public Property Let FontName(sName As String)
   mFontName = sName
End Property

Public Property Get FontSize() As Long
  FontSize = mFontSize
End Property

Public Property Let FontSize(lSize As Long)
   mFontSize = lSize
End Property

Public Property Get Italic() As Boolean
  Italic = mItalic
End Property

Public Property Let Italic(BItalic As Boolean)
   mItalic = BItalic
End Property

Public Property Get StrikeThru() As Boolean
   StrikeThru = mStrikethru
End Property

Public Property Let StrikeThru(bStrikethru As Boolean)
   mStrikethru = bStrikethru
End Property

Public Property Get Underline() As Boolean
   Underline = mUnderline
End Property

Public Property Let Underline(bUnderline As Boolean)
   mUnderline = bUnderline
End Property

' Open , Save, Folder, Icon

Public Property Get DefaultExt() As String
   DefaultExt = mDefaultExt
End Property

Public Property Let DefaultExt(sDefExt As String)
   mDefaultExt = sDefExt
End Property

Public Property Get filename() As String
   filename = mFileName
End Property

Public Property Let filename(sFileName As String)
   mFileName = sFileName
End Property

Public Property Get FileTitle() As String
   FileTitle = mFileTitle
End Property

Public Property Let FileTitle(sTitle As String)
   mFileTitle = sTitle
End Property

Public Property Get Filter() As String
   Filter = mFilter
End Property

Public Property Let Filter(sFilter As String)
   mFilter = sFilter
End Property

Public Property Get FilterIndex() As Long
   FilterIndex = mFilterIndex
End Property

Public Property Let FilterIndex(lIndex As Long)
    mFilterIndex = lIndex
End Property

Public Property Get InitDir() As String
   InitDir = mInitDir
End Property

Public Property Let InitDir(sDir As String)
    mInitDir = sDir
End Property

Public Property Get IconSize() As CdlgEx_IconSize
  If mIconSize <> IconSizeLarge And mIconSize <> IconSizeSmall Then mIconSize = IconSizeLarge
  IconSize = mIconSize
End Property

Public Property Let IconSize(nSize As CdlgEx_IconSize)
  If nSize <> IconSizeLarge And nSize <> IconSizeSmall Then nSize = IconSizeLarge
  mIconSize = nSize
End Property
' Help Properties
Public Property Get HelpCommand() As CdlgExt_HelpCommand
   HelpCommand = mHelpCommand
End Property

Public Property Let HelpCommand(lCommand As CdlgExt_HelpCommand)
   mHelpCommand = lCommand
End Property

Public Property Get HelpFile() As String
   HelpFile = mHelpFile
End Property

Public Property Let HelpFile(sFile As String)
   mHelpFile = sFile
End Property

Public Property Get HelpKey() As Long
   HelpKey = mHelpKey
End Property

Public Property Let HelpKey(sKey As Long)
   mHelpKey = sKey
End Property

'Color Dialog
Public Property Get rgbResult() As Long
   rgbResult = mRGBResult
End Property

Public Property Let rgbResult(lValue As Long)
   mRGBResult = lValue
End Property



'About Dialog
Public Function ShowAbout()
    If mAppName = "" Then mAppName = Chr$(0)
    SHAbout mhOwner, mAppName, mDialogPrompt, mhIcon
End Function

Private Sub InitOFN()
  Dim sTemp As String, i As Integer
  Dim uFlag As Long
  uFlag = mFlags And (&H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
  With OFN
       .lStructSize = Len(OFN)
       .hWndOwner = mhOwner
       .Flags = uFlag
       .lpstrDefExt = mDefaultExt
       sTemp = mInitDir
       If sTemp = "" Then sTemp = App.Path
       .lpstrInitialDir = sTemp
       sTemp = mFileName
       .lpstrFile = sTemp & String$(255 - Len(sTemp), 0)
       .nMaxFile = 255
       .lpstrFileTitle = String$(255, 0)
       .nMaxFileTitle = 255
        sTemp = mFilter
        For i = 1 To Len(sTemp)
            If Mid(sTemp, i, 1) = "|" Then
               Mid(sTemp, i, 1) = vbNullChar
            End If
        Next
        sTemp = sTemp & String$(2, 0)
        .lpstrFilter = sTemp
        .nFilterIndex = mFilterIndex
        .lpstrTitle = mDialogTitle
        .hInstance = App.hInstance
 End With
End Sub
'Public Sub ShowHelp() 07-03-05
' 'mHelpKey = &H101
' RetValue = WinHelp(mhOwner, mHelpFile, mHelpCommand, mHelpKey)
'End Sub
'Public Sub ShowColor()
'  Dim CC As CHOOSECOLOR
'  Dim CustomColors() As Byte
'  Dim uFlag As Long
'  ReDim CustomColors(0 To 16 * 4 - 1) As Byte
'  For i = LBound(CustomColors) To UBound(CustomColors)
'     CustomColors(i) = 255 ' white
'  Next i
'  uFlag = mFlags And (&H1 Or &H2 Or &H4 Or &H8)
'  With CC
'       .lStructSize = Len(CC)
'       .hwndOwner = mhOwner
'       .hInstance = App.hInstance
'       .lpCustColors = StrConv(CustomColors, vbUnicode)
'       .Flags = uFlag
'       .rgbResult = mRGBResult
'       RetValue = ChooseColorAPI(CC)
'       If RetValue = 0 Then
'          If mCancelError Then err.Raise (RetValue)
'       Else
'          CustomColors = StrConv(.lpCustColors, vbFromUnicode)
'          mRGBResult = .rgbResult
'       End If
'  End With
'End Sub
Public Sub ShowFont()
  Dim CF As ChooseFont
  Dim LF As LOGFONT
  Dim TempByteArray() As Byte
  Dim ByteArrayLimit As Long
  Dim OldhDC As Long
  Dim FontToUse As Long
  Dim tbuf As String * 80
  Dim x As Long
  Dim uFlag As Long
  uFlag = mFlags And (&H1 Or &H2 Or &H3 Or &H4 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H10000 Or &H20000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
  TempByteArray = StrConv(mFontName & vbNullChar, vbFromUnicode)
  ByteArrayLimit = UBound(TempByteArray)
  With LF
     For x = 0 To ByteArrayLimit
        .lfFaceName(x) = TempByteArray(x)
     Next
    .lfHeight = mFontSize / 72 * GetDeviceCaps(GetDC(mhOwner), LOGPIXELSY)
    .lfItalic = mItalic * -1
    .lfUnderline = mUnderline * -1
    .lfStrikeOut = mStrikethru * -1
    If mBold Then .lfWeight = FW_BOLD
  End With
  With CF
      .lStructSize = Len(CF)
      .hWndOwner = mhOwner
      .hdc = GetDC(mhOwner)
      .lpLogFont = lstrcpy(LF, LF)
      If Not uFlag Then
         .Flags = cdlCFScreenFonts
      Else
         .Flags = uFlag Or cdlCFWYSIWYG
      End If
     .Flags = .Flags Or cdlCFEffects Or CF_INITTOLOGFONTSTRUCT
     .rgbColors = mRGBResult
     .lCustData = 0
     .lpfnHook = 0
     .lpTemplateName = 0
     .hInstance = 0
     .lpszStyle = 0
     .nFontType = SCREEN_FONTTYPE
     .nSizeMin = 0
     .nSizeMax = 0
     .iPointSize = mFontSize * 10
    End With
    retvalue = ChooseFont(CF)
    If retvalue = 0 Then
       If mCancelError Then err.Raise (retvalue)
    Else
       With LF
            mItalic = .lfItalic * -1
            mUnderline = .lfUnderline * -1
            mStrikethru = .lfStrikeOut * -1
       End With
       With CF
            mFontSize = .iPointSize \ 10
            mRGBResult = .rgbColors
            If .nFontType And BOLD_FONTTYPE Then
                mBold = True
            Else
                mBold = False
            End If
       End With
       FontToUse = CreateFontIndirect(LF)
       If FontToUse = 0 Then Exit Sub
          OldhDC = SelectObject(CF.hdc, FontToUse)
          retvalue = GetTextFace(CF.hdc, 79, tbuf)
          mFontName = Mid$(tbuf, 1, retvalue)
       End If
End Sub
