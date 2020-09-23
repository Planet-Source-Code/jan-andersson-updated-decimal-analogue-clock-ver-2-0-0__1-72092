Attribute VB_Name = "modDecimalClock"
Option Explicit

' =========================================================================================
' Decimal Clock Version 2.0.0
' Code written by Pappsegull Sweden, pappsegull@yahoo.se
' Copyright © 2009 - Pappsegull - All rights reserved.
'
' Disclaimer:
' This example program is provided "as is" with no warranty of any kind. It is
' intended for demonstration purposes only. You can use the code in any form,
' but please mention the author ;)
'
' Decimal time is the representation of the time of day using units
' which are decimally related. This term is often used to refer
' specifically to French Revolutionary Time, which divides the day into
' 10 decimal hours, each decimal hour into 100 decimal minutes and
' each decimal minute into 100 decimal seconds.
' =========================================================================================


'///Registry, File association & Tray icon
'Private Const C_Typ = "Decimal Clock"
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ As Long = 1
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_DWORD As Long = 4
Private Const KEY_WRITE = &H20006

Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'///Create an Icon in System Tray
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203 'Double-click
Private Const WM_LBUTTONDOWN = &H201   'Button down
Private Const WM_LBUTTONUP = &H202     'Button up
Private Const WM_RBUTTONDBLCLK = &H206 'Double-click
Private Const WM_RBUTTONDOWN = &H204   'Button down
Private Const WM_RBUTTONUP = &H205     'Button up

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim nID As NOTIFYICONDATA ' Trayicon variable

'/// Add & Remove Font

Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'///Picture stuff

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function TransparentBlt Lib "MSIMG32.DLL" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function FoxAlphaBlend Lib "FoxCBmp3.dl" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Alpha As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long
Private Declare Function FoxHSL Lib "FoxCBmp3.dl" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Hue As Single, ByVal Saturation As Single, ByVal Lightness As Single, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long
Private Declare Function FoxAlphaMask Lib "FoxCBmp3.dl" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal hMaskDC As Long, ByVal xMask As Long, ByVal yMask As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long
Private Declare Function FoxRotate Lib "FoxCBmp3.dl" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Angle As Double, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long
'Use this nice functions from Florian Egel to get good speed:) Thanks!
'The FoxCBmp3.dl (64kb) file is in the resource file and will be written to SysDir if not exists.
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=21470&lngWId=1

Global Const COLORONCOLOR As Long = 3
Global Const HALFTONE     As Long = 4

'///Timezone

Private Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte  'unicode (0-based)
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte  'unicode (0-based)
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'///Common dialog

'ShowOpen/ShowSave flags:
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2 ''''
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const OFS_MAXPATHNAME = 128

'ChooseColor flags:
Private Const CC_ANYCOLOR = &H100
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
Private Const CC_SHOWHELP = &H8
Private Const CC_SOLIDCOLOR = &H80

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetFileTitleAPI Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private OFName As OPENFILENAME, CC As CHOOSECOLOR, DefDialogColors&(), CustomColors() As Byte, sTmpMsg$

'///Play sound

'Play wave sound from Resource or disk file
Private Const SND_APPLICATION = &H80    'Look for application specific association
Private Const SND_ALIAS = &H10000       'Name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000   'Name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1           'Play asynchronously
Private Const SND_FILENAME = &H20000    'Name is a file name
Private Const SND_LOOP = &H8            'Loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4          'lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2       'Silence not default, if sound not found
Private Const SND_NOSTOP = &H10         'Don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000       'Don't wait if the driver is busy
Private Const SND_PURGE = &H40          'Purge non-static events for task
Private Const SND_RESOURCE = &H40004    'Name is a resource name or atom
Private Const SND_SYNC = &H0            'Play synchronously (default)
'Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function SNDPLAYSOUND Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

'///Mixed stuff
Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    
Const LWA_COLORKEY = 1
Const LWA_ALPHA = 2
Const LWA_BOTH = 3

Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const BM_SETSTATE = &HF3

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40


'///Clock stuff

'Clock hands
Enum Hands
    DecHour
    DecMinute
    DecSecond
    ClockFace
End Enum
'Indexed Horizontal scrollbar controls
Enum HscValues
    HandHourSize
    HandMinuteSize
    HandSecondSize
    SizeFactor
    BorderPos
    NumbersPos
    HourSpotsPos
    MinuteSpotsPos
    BackgroundPos
    BorderSize
    NumbersSize
    HourSpotsSize
    MinuteSpotsSize
    BackgroundSize
    DigiClockX
    WeekdayX
    DateX
    MyTextX
    DigiClockY
    WeekdayY
    DateY
    MyTextY
    DigiClockSize
    WeekdaySize
    DateSize
    MyTextSize
    HandHourX
    HandMinuteX
    HandSecondX
    HandHourY
    HandMinuteY
    HandSecondY
    HandHourStyle
    HandMinuteStyle
    HandSecondStyle
    Translucent
    FaceHue
    FaceSaturation
    FaceBrightness
    HHue
    HSaturation
    HBrightness
    MHue
    MSaturation
    MBrightness
    SHue
    SSaturation
    SBrightness
    HAlfa
    MAlfa
    SAlfa
End Enum
'Indexed Checkbox controls
Enum ChkValues
    ShowDigital
    ShowDate
    ShowSmooth
    UsePictureFile
    Topmost
    SettingsSave
    DrawOnPicture
    ShowSecondHand
    ShowWeekday
    DrawBorder
    DrawNumbers
    DrawHourSpots
    DrawMinuteSpots
    DrawBackground
    DrawMyText
    RunWinStartUp
    PictureHands
    SecondTick
    AlarmOn
    AlarmSound
    MinimizeToTray
    ShowMinuteHand
    ShowHourHand
    StopClock
    ShowDateMonth
    AntiAlias
    DrawNumbersMinutes
    AlphaHands
    AlphaFace
    CenterCircle
    DrawHourLines
    DrawMinuteLines
    SaveFontToFile
End Enum
'Indexed Command button controls to select colors and more
Enum CmdColors
    Border
    Numbers
    HourSpots
    MinuteSpots
    Background
    DigiClock
    Weekday
    Date
    MyText
    HourHand
    MinuteHand
    SecondHand
    TransparentBack
'Sound
    SoundPlay
    SoundSelect
'Hand gradient & outline colors
    HourHand2
    MinuteHand2
    SecondHand2
    HourHandBorder
    MinuteHandBorder
    SecondHandBorder
    CenterC1
    CenterC2
    Border2
    HourSpots2
    MinuteSpots2
    Background2
    Numbers2 'Not in us as of now
End Enum
'Some stuff for Undo/Redo
Enum ControlType
    DoURhsc
    DoURcmd
    DoURchk
    DoURpic
    DoURcmb
    DoURopt
End Enum
Enum UndoRedoAction
    DoReset
    DoUndo
    DoRedo
    DoSave
    DoKillTmp
End Enum
Type UndoRedoType
    CtrlTyp As Integer
    CtrlIndex As Integer
    CtrlValueOld As Variant
    CtrlValueNew As Variant
    InfoText As String
End Type
Type UndoFiles
    Ubounds As Integer
    UndoPos As Integer
    Undos() As UndoRedoType
End Type

'Type to save all clockdata to disk
Type Settings
    Version As String
    ChkValue() As Integer
    CmdColors() As Long
    HscValue() As Single
    ClockRadiusPxOrg As Integer
    DecimalClock As Boolean
    MyText As String
    FontName As String
    FileName As String
    Top As Integer
    Left As Integer
    TopS As Integer
    LeftS As Integer
    TimeZoneOffset As Integer
    TimeZoneSaved As Integer
    TimeZoneName As String
    PictureBack() As Byte
    PictureHourHand() As Byte
    PictureMinuteHand() As Byte
    PictureSecondHand() As Byte
    AlarmHour As Integer
    AlarmMinute As Integer
    AlarmCommand As String
    AlarmShowMsgBox As Boolean
    AlarmWavFile() As Byte
    Reserved() As Variant
    Comments As String
    FileTest As String * 5
    FontFile() As Byte
    FontFileName As String
End Type
'Global constants
Global Const c_Sec24h& = 86400, c_DecH& = c_Sec24h& / 10, c_DecM! = c_DecH& / 100, _
  c_NewFile = "New Clock", c_F = "##00", c_100 = 100, c_DefautRadius = 100, c_ClcDir = "\Clocks", _
  c_Def = "\Default.clc", c_Cust = "CUSTOM", c_BltErr = "Transparent Blitt error.", _
  c_URL = "www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72092&lngWId=1", _
  c_Mail = "pappsegull@yahoo.se", Pi = 3.14159265358979, c_FileTest = "®ø©¥¿" 'Don't change this

'Global variables
Global Tm!, Gs As New clsLineGS, bErr As Boolean, nDll%, sVerNow$, bUndoRedo As Boolean, _
  HH%, MM%, SS%, AngH!, AngM!, AngS!, AngPrevS!, sFile$, CR%, bRedraw As Boolean, sTmpUndo$, _
  Clock As Settings, bDeci As Boolean, bNotNow As Boolean, bStop As Boolean, Tmp() As Byte, _
  IsDirty As Boolean, bCancelSave As Boolean, CX%, CY%, FacePicRadius%, UndoPos%, _
  bAlarmIsOn As Boolean, PrevMM%, bpicsHorM As Boolean, bpic(2) As Boolean, sTmp$, _
  ByteSound() As Byte, bColorPick As Boolean, lColorpick&, UndoRedos() As UndoRedoType, _
  CPX%, CPY%, PrevX!, PrevY!, TransC&(3), UndoFile As UndoFiles
  

Sub Main()
Dim Cl As Settings, X%, Y%, s$, bCmd As Boolean
On Error GoTo ErrMain
ErrFile:
    sVerNow$ = App.Major & "." & App.Minor & "." & App.Revision
    frmC.tmr.Enabled = False
    If bErr Then Err.Raise 7 'Error when did load settings
    With Clock
        If Command <> "" Then 'Opening a associated setting file
            s$ = Command: bNotNow = True: bCmd = True
            'Remove '' chars
            s$ = Left(s$, Len(s$) - 1): s$ = Right(s$, Len(s$) - 1)
            'Convert to long filename
            s$ = LongFileName(s$): sFile$ = s$: s$ = ""
        End If
        bNotNow = True: IsDirty = False: Unload frmS: bNotNow = False
        s$ = App.Path & c_Def: If sFile$ = "" Then sFile$ = s$
OpenClockFile:
        If Dir(sFile$) <> "" Then 'Load settings
            X% = FreeFile: Open sFile$ For Binary As X%
            Get X%, , Cl: Close X%
            If Cl.FileTest <> c_FileTest Then 'Test if file is valid
                MsgBox sFile$ & vbLf & "The file is corrupt or not a valid " & _
                  App.Title & "-File.", 16 + vbSystemModal
                  If bCmd Then End
                  sFile$ = IIf(sTmp$ = c_NewFile, s$, sTmp$)
                  sTmp$ = "": GoTo OpenClockFile
            End If
            Clock = Cl: bDeci = .DecimalClock
            'Default settings exits so open as new file
            If sFile$ = s$ Then sFile$ = c_NewFile
            'Check if this computor have another TimeZone then the one file was saved on
            X% = 0: s$ = GetTimeZone(X%)
            If .TimeZoneSaved <> X% And .TimeZoneOffset <> 0 Then
                Y% = .TimeZoneOffset - ((X% - .TimeZoneSaved) / 60)
                MsgBox "This file '" & .FileName & _
                  "' was saved on a computer with time zone: " & _
                  .TimeZoneName & vbLf & vbLf & "Will compensate that to time zone '" & _
                  s$ & "'." & vbLf & "Have change time zone offset from " & _
                  .TimeZoneOffset & " to " & Y% & ".", 64 + vbSystemModal
                .TimeZoneOffset = Y%: .TimeZoneName = s$
            End If
        Else 'Get default settings from Resourcefile
            Dim ba() As Byte
            ba() = LoadResData("DEFAULT.CLC", c_Cust)
            X% = FreeFile: Open sFile$ For Binary As X%
            Put X%, , ba: Close X%: Erase ba()
            X% = FreeFile: Open sFile$ For Binary As X%
            Get X%, , Clock: Close X%
            bDeci = .DecimalClock
            .Top = Screen.Height / 2 - (.ClockRadiusPxOrg * Screen.TwipsPerPixelY)
            .Left = Screen.Width / 2 - (.ClockRadiusPxOrg * Screen.TwipsPerPixelY)
            frmC.Move .Left, .Top
            If sFile$ = s$ Then 'Create a default setting file if is missing
                Call CreateAssociation: CheckDLL 'Create file association and check dll
                MsgBox "Right click on the clock to show the popup menu, drag & drop to move it.", 64
                Call SaveSettings: sFile$ = ""
            End If
        End If
    End With
    With frmC
        If Len(Clock.FontName) Then frmC.FontName = Clock.FontName
        Call LoadClockPictures: Call InitPolygon: Call SetStyle
        .Top = Clock.Top: .Left = Clock.Left
        Call SetTopMost(.hwnd)
        If sFile$ = c_NewFile Then frmS.Show 'Start application new clockfile
        If IsSelected(MinimizeToTray) Then
            Call TrayMinimizeAppTo 'Minimize to tray
        Else: .Show: End If
        IsDirty = False: SetCaption
        .tmr.Enabled = True: bNotNow = False: bRedraw = True
    End With
    Exit Sub
ErrMain:
    If Err = 380 Then 'Font error
        If Clock.FontFileName <> "" Then
            s$ = Left$(sFile$, Len(sFile$) - 4) & " " & Clock.FontFileName
            'Write the font file to disk if stored in clockfile
            X% = FreeFile: Open s$ For Binary As X%
            Put X%, , Clock.FontFile: Close X%: Err.Clear
            If AddFont&(s$) Then 'Sucess to add font
                On Local Error Resume Next
                frmC.FontName = Clock.FontName 'Try one more if could use the font
                If Err = 0 Then
                    MsgBox "The font '" & Clock.FontName & _
                      "', where missing on your system. But where added from the clock file '" _
                      & Clock.FileName & "', to the font file '" & s$ & "'.", 64 + vbSystemModal
                    Resume Next
                Else: GoTo FontErr: End If
            Else: GoTo FontErr: End If
        Else
FontErr:
            MsgBox "Coulden't use/add the font " & Clock.FontName & _
              ", will use " & frmC.FontName & " instead.", 48 + vbSystemModal
            Clock.FontName = frmC.FontName
            Resume Next
        End If
    End If
    If Err = 7 Then
        If Len(sFile$) Then
        If sFile$ = c_NewFile Then sFile$ = App.Path & c_Def
        X% = MsgBox("This version of " & App.Title & " is '" & _
          sVerNow$ & "' and the file was created in version '" & _
          Clock.Version & "'." & vbLf & vbLf & "Do you want to delete '" & sFile$ & _
          "'?", vbCritical + vbYesNoCancel + vbDefaultButton2 + vbSystemModal)
        If X% = vbCancel Then End
        If X% = vbYes Then
            Clock = Cl: On Error Resume Next
            Close X%: If Dir(sFile$) <> "" Then Kill sFile$
        Else: sFile$ = "": End If
        End If
        bNotNow = False: bErr = False
        Resume ErrFile
    End If
    MsgBox Err.Description, vbCritical + vbSystemModal
    Exit Sub
    Resume
End Sub

Function ConvToDecTime!(Hand As Hands, Optional RetValAngle!, Optional ByVal TmD! = -1)
Dim T!, A!
    If TmD! = -1 Then TmD! = IIf(Tm! = 0, Timer, Tm!)
    If Hand = DecHour Then 'Hour hand
        T! = c_Sec24h& / 360                            'Number of seconds in one degree on hour hand
        A! = TmD! / T!                                  'Seconds elapsed since midnight, angle on hourhand
        T! = Fix(A! / 36)                               'Calculate decimal hours
    Else                   'Minute/Second
        T! = IIf(Hand = DecMinute, c_DecH&, c_DecM!)    'Select value if minute or second
        T! = TmD! / T!: T! = T! - Fix(T!)               'Remove whole laps (the integer)
        A! = 360 * T!                                   'Calculate the angle on minutehand
        T! = Fix(A! / 3.6)                              'Calculate decimal minutes/Seconds
    End If
    ConvToDecTime! = T!: RetValAngle! = A!
End Function

Function CheckTime() As Boolean 'Check time and print date on clock
Dim L!, n!, T As Date, s$, DH%, DM%, DS%, NH%, NM%, NS%: Static BeenHereB4 As Boolean

    T = Now 'Time, prevents calling Time multiple times
    T = DateAdd("h", Clock.TimeZoneOffset, T)
'Alarm
    With Clock
        If IsSelected(AlarmOn) And BeenHereB4 Then
            If .AlarmHour = HH% And .AlarmMinute = MM% Then
                'If Not bAlarmIsOn And SS% < 2 Then
                If Not bAlarmIsOn Then
                    bAlarmIsOn = True
                    s$ = "The time is now " & Format(.AlarmHour, c_F) & ":" & _
                      Format(.AlarmMinute, c_F)
                    If IsSelected(AlarmSound) Then
                        frmC.tmr.Enabled = False: s$ = s$ & ", turn off alarm sound."
                        PlayWav , , True 'Stop current sound if any...
                        PlayWav , True 'Loop alarm sound
                    Else: s$ = s$ & ".": End If
                    If Len(.AlarmCommand) Then
                        If .AlarmShowMsgBox Then
                            MsgBox .AlarmCommand & vbLf & vbLf & _
                              s$, 64 + vbSystemModal: s$ = ""
                        Else
                            RunCommand .AlarmCommand
                        End If
                    End If
                    If Len(s$) Then MsgBox s$, 64 + vbSystemModal
                    PlayWav , , True  'Stop sound if any running
                    frmC.tmr.Enabled = True
                End If
            Else: bAlarmIsOn = False: End If
        End If
    End With
'Get decimal Time & Angle
    DH% = ConvToDecTime!(DecHour, AngH!)
    DM% = ConvToDecTime!(DecMinute, AngM!)
    DS% = ConvToDecTime!(DecSecond, AngS!)
'Get standard Time
    NH% = Hour(T): NM% = Minute(T): NS% = Second(T)
'Decimal clock selected
    If bDeci Then
        HH% = DH%: MM% = DM%: SS% = DS%
    Else
'Standard clock selected
        HH% = NH%: MM% = NM%: SS% = NS%
        'Convert time components to degrees °
        AngM! = MM% * 6
        AngS! = (SS% + (Tm! - Fix(Tm!))) * 6
        'Adjust hour hand to include minute angle \12
        AngH! = (HH% Mod 12) * 30 + AngM! \ 12
    End If
    'Digital clocks (Decimal & Standard) in form
    s$ = IIf(DH% < 10, "0", "") & DH% & ":" & _
      IIf(DM% < 10, "0", "") & DM% & ":" & IIf(DS% < 10, "0", "") & DS%
    frmS.lbl(18) = Format(T, "Long Time"): frmS.lbl(17) = s$
'Stop Selected, redraw only if have edit clock
    If IsSelected(StopClock) Then
        AngH! = 0: AngM! = 0: AngS! = 0: Tm! = 0
        If Not bRedraw Then CheckTime = True: Exit Function
    End If
'Only if smooth mode not selected...
    If Not IsSelected(ShowSmooth) Then
        'Not redraw clock if none second hand every second
        If Not bRedraw And Not bStop Then
            If Not IsSelected(ShowSecondHand) Then
                If MM% = PrevMM% Then Exit Function
            End If
            AngS! = SS% * IIf(bDeci, 3.6, 6)
        End If
        If AngS! = AngPrevS! Then Exit Function
        If AngS! < AngPrevS! Then AngPrevS! = 0
        AngPrevS! = AngS!: PrevMM% = MM%
    End If
'Print text
    With frmC
        CheckTime = True: .Cls
        If Tm! = 0 Then
            T = Format(T, "yyyy-mm-dd") & " 00:00:00"
        End If
    'Show Digital Clock
        If IsSelected(ShowDigital) Then
            If bDeci Then
                s$ = IIf(HH% < 10, "0", "") & HH% & ":" & _
                  IIf(MM% < 10, "0", "") & MM%
                If IsSelected(ShowSecondHand) Then
                    s$ = s$ & ":" & IIf(SS% < 10, "0", "") & SS%
                End If
            Else
                s$ = IIf(IsSelected(ShowSecondHand), "Long", "Short") & " Time"
                s$ = Trim$(Format(T, s$))
            End If
            PrintText s$, GetHscValue(DigiClockY), _
              GetHscValue(DigiClockSize), GetColor(DigiClock), _
              , , GetHscValue(DigiClockX)
        End If
    'Show Date & mounth
        If IsSelected(ShowDate) Or IsSelected(ShowDateMonth) Then
            If IsSelected(ShowDate) And _
              IsSelected(ShowDateMonth) Then
                s$ = "d mmm"
            ElseIf IsSelected(ShowDate) Then
                s$ = "d"
            Else: s$ = "mmm": End If
            PrintText UCase(Format(T, s$)), GetHscValue(DateY), _
            GetHscValue(DateSize), GetColor(Date), , , GetHscValue(DateX)
        End If
    'Show Weekday
        If IsSelected(ShowWeekday) Then
            PrintText UCase(Format(T, "dddd")), GetHscValue(WeekdayY), _
              GetHscValue(WeekdaySize), GetColor(Weekday), , , GetHscValue(WeekdayX)
        End If
    'Show Custom text or App.Title if decimal
        s$ = IIf(IsSelected(DrawMyText), Clock.MyText, _
          IIf(bDeci, "Decimal", "Standard"))
        PrintText s$, GetHscValue(MyTextY), _
          GetHscValue(MyTextSize), GetColor(MyText), , , GetHscValue(MyTextX)
    End With
    BeenHereB4 = True
End Function

Sub ResetAll(Optional Reload As Boolean)
'Reset all settings to default on current clock or reload it
Dim s$, T%, L%, CX%, CY%, CS As Settings
    T% = frmS.Top: L% = frmS.Left
    Clock = CS
    With frmC
        CX% = .Left + (.Width / 2): CY% = .Top + (.Height / 2)
        bNotNow = True: Unload frmS
        If Dir(sFile$) <> "" And Not Reload Then Kill sFile$
        Call Main
        .Left = CX% - (.Width / 2): .Top = CY% - (.Height / 2)
    End With
    frmS.Left = L%: frmS.Top = T%: frmS.Show
End Sub

Sub DrawHands()
Dim X%, Y%, G&(), Hand As Hands, s$, C&: Static bSnd As Boolean

'Draw a transparent cross on clock to make it more easy to center hand pictures
    C& = GetColor(TransparentBack)
    If bStop And bRedraw Then
        With frmC
            frmS.pic(3) = LoadPicture(""): frmC.Cls
            .DrawWidth = IIf(CR% > 150, 3, 1)
            Gs.CircleGP .hDC, CR%, CR%, CR% / 2, CR% / 2, C&
            frmC.Line (CR%, CR% - CR%)-(CR%, CR% + CR%), C&
            frmC.Line (CR% - CR%, CR%)-(CR% + CR%, CR%), C&
            .DrawWidth = 1: frmC.PSet (CR%, CR%), InvertColor(C&)
            .Picture = .Image
        End With
    End If
With frmS
    If IsSelected(PictureHands) Then
    'Redraw hour & minute hand if seconds = 00 or bRedraw = True
        If SS% = 0 Or bRedraw Then
            bpicsHorM = False
            For X% = DecHour To DecSecond 'Check if handstyle = -1 (Picture)
                bpic(X%) = IIf((GetHscValue(Choose(X% + 1, HandHourStyle, _
                  HandMinuteStyle, HandSecondStyle))) = -1, True, False)
                'If bStop And bpic(X%) Then PrepairPicture X%
            Next
            .pic(3) = LoadPicture(""): .pic(3).BackColor = C&
            .pic(3).Height = frmC.Height: .pic(3).Width = frmC.Width
            'Check if can draw hour hand picture
            If Not IsArrayEmpty(Clock.PictureHourHand) And bpic(DecHour) Then
                DrawHand DecHour: bpicsHorM = True
                .pic(3) = .pic(3).Image
            Else: bpic(DecHour) = False: End If
            'Check if can draw minute hand picture
            If Not IsArrayEmpty(Clock.PictureMinuteHand) And bpic(DecMinute) Then
                DrawHand DecMinute: bpicsHorM = True
                .pic(3) = .pic(3).Image
            Else: bpic(DecMinute) = False: End If
            'Check if can draw second hand picture
            If Not IsArrayEmpty(Clock.PictureSecondHand) And bpic(DecSecond) Then
            Else: bpic(DecSecond) = False: End If
        End If
        .pic(3).Cls
        If bpicsHorM And Not bpic(DecSecond) Then
            If TransparentBlt( _
              frmC.hDC, 0, 0, .pic(3).ScaleWidth, .pic(3).ScaleHeight, _
              .pic(3).hDC, 0, 0, .pic(3).ScaleWidth, .pic(3).ScaleHeight, _
              GetColor(TransparentBack)) = False Then _
                MsgBox c_BltErr, 16 + vbSystemModal
        End If
        'Check if draw lines or polygones if can't draw picture hands
        If Not bpic(DecHour) Then DrawHandNonePic DecHour
        If Not bpic(DecMinute) Then DrawHandNonePic DecMinute
        'Draw second hand
        If bpic(DecSecond) Then DrawHand DecSecond Else DrawHandNonePic DecSecond
    Else 'Picture hands not selected so draw using clsLineGS
        If Not bpicsHorM Then
            For Hand = DecHour To DecSecond
                DrawHandNonePic Hand
            Next
        End If
        bpicsHorM = False
    End If
End With
'Draw a gradient center circle on top of hands
    If IsSelected(CenterCircle) Then
        Y% = 6 * Clock.HscValue(SizeFactor)
        BlendColors GetColor(CenterC2), GetColor(CenterC1), Y%, G&()
        For X% = 1 To Y%
            Gs.CircleGP frmC.hDC, CR%, CR%, X%, X%, G&(X% - 1) ', Thin
        Next
        SetPixel frmC.hDC, CR%, CR%, G&(0) 'Fill the last pixel in center
        Erase G&()
    End If
    bRedraw = False: frmC.Refresh
'Play second tick minute/hour change sound if selected
    If IsSelected(SecondTick) And Not bAlarmIsOn And Not bStop Then
        If MM% = 0 And SS% = 0 Then 'Hour strike
            s$ = "H"
        ElseIf SS% = 0 Then         'Minute change
            s$ = "M"
        Else                        'Second tick
            If Not IsSelected(ShowSmooth) Then s$ = "S"
        End If
        If MM% <> 0 And SS% = 1 Then bSnd = True
        If Len(s$) And bSnd Then PlayWav "CHANGE_" & s$ & ".WAV"
        'To avoid multiple play if smooth mode, or play whole hour strike
        bSnd = IIf(SS% = 0 Or (MM% = 0 And SS% < 3), False, True)
    End If
    'Stop clock if selected
    If IsSelected(StopClock) Then bStop = True
End Sub

Sub DrawHandNonePic(ByVal Hand As Hands)
Dim Y%
    Y% = GetHscValue(Choose(Hand + 1, HandHourStyle, _
      HandMinuteStyle, HandSecondStyle))
    If Y% = 0 Then  'Line hand
        DrawHand Hand
    Else            'Draw polygon hand
        If IsSelected(Choose(Hand + 1, ShowHourHand, _
          ShowMinuteHand, ShowSecondHand)) Then DrawPolygonHand Hand
    End If
End Sub

Sub DrawHand(Hand As Hands) 'Draw a Hand picture or line
Dim X%, C&, R&, Angle!, SzX!, SzY!, L!, X2%, Y2%
On Error GoTo ErrDrawHand
    If nDll% > 1 Then Exit Sub
    If bStop And Not bRedraw Then Exit Sub
    If Not IsSelected(ShowHourHand) And Hand = DecHour Then Exit Sub
    If Not IsSelected(ShowMinuteHand) And Hand = DecMinute Then Exit Sub
    If Not IsSelected(ShowSecondHand) And Hand = DecSecond Then Exit Sub
'Set values depending on hand  L! = Choose(Hand + 1, 5, 7, 8) / 10 '%-Length on hand
    Select Case Hand
        Case DecHour: SzX! = HandHourX: SzY! = HandHourY: Angle! = AngH!: C& = HourHandBorder: L! = 0.5
        Case DecMinute: SzX! = HandMinuteX: SzY! = HandMinuteY: Angle! = AngM!: C& = MinuteHandBorder: L! = 0.7
        Case DecSecond: SzX! = HandSecondX: SzY! = HandSecondY: Angle! = AngS!: C& = SecondHandBorder: L! = 0.8
    End Select
    SzX! = GetHscValue(SzX!): SzY! = GetHscValue(SzY!): C& = GetColor(C&)
    If IsSelected(StopClock) Then Angle! = 0 'Clock is stopped
'Blit picture
    If bpic(Hand) And IsSelected(PictureHands) Then
        With frmS
            C& = TransC&(ClockFace) 'GetColor(TransparentBack)
        'Rotate hand
            FoxRotate .pic(3).hDC, CR%, CR%, _
              .pic(Hand).ScaleWidth, .pic(Hand).ScaleHeight, _
              .pic(Hand).hDC, 0, 0, Angle!, TransC&(Hand), _
              IIf(IsSelected(AntiAlias), &H3, &H1) ' Mask = 1, AntiAlias = 2
        'If time to transparent blitt frmS.pic(3) to frmC (The Clock)
            If Hand = DecSecond Or _
              (Hand = DecMinute And Not (IsSelected(ShowSecondHand) Or _
                Not bpic(DecSecond))) Or _
              (Hand = DecHour And (Not IsSelected(ShowMinuteHand) Or _
                (Not bpic(DecMinute) Or Not bpic(DecMinute)))) Then
                .pic(3).Refresh
                If IsSelected(AlphaHands) Then  'Use Alpha Mask
                    FoxAlphaMask frmC.hDC, 0, 0, _
                    .pic(3).ScaleWidth, .pic(3).ScaleHeight, _
                    .pic(3).hDC, 0, 0, .pic(3).hDC, 0, 0, C&, &H1
                Else                            'Use Transparent blitt
                    If TransparentBlt( _
                      frmC.hDC, 0, 0, .pic(3).ScaleWidth, .pic(3).ScaleHeight, _
                      .pic(3).hDC, 0, 0, .pic(3).ScaleWidth, .pic(3).ScaleHeight, _
                      C&) = False Then _
                        MsgBox c_BltErr, 16 + vbSystemModal
                End If
            End If
        End With
    Else
'Draw Hand line using clsLineGS if hand picture is missing or not selected
        L! = L! * CR% * GetHscValue(Hand) 'Size on hand
    'Calculate X & Y for the line
        X2% = L! * Cos(Pi / 180 * (Angle! - 90)) + (CR% * SzX!)
        Y2% = L! * Sin(Pi / 180 * (Angle! - 90)) + (CR% * SzY!)
        CX% = CR% * SzX!: CY% = CR% * SzY!
        If IsSelected(AntiAlias) Then 'Anti Alias selected
            Gs.LineGP frmC.hDC, CX%, CY%, X2%, Y2%, C&
        Else
            frmC.Line (CX%, CY%)-(X2%, Y2%), C&
        End If
    End If
    Exit Sub
ErrDrawHand:
    If Err = 53 Then 'Can't find the dll-file
        nDll% = nDll% + 1
        If nDll% = 1 Then CheckDLL: DoEvents: Resume
        If nDll% > 2 Then Err.Clear: Exit Sub
    End If
    MsgBox Err.Description, 16 + vbSystemModal: Err.Clear
End Sub

Sub SetStyle() 'Redraw clock
Dim dx%, dy%
    If bUndoRedo And Len(sTmpUndo$) = 0 Then
        IsDirty = True: CtrlEnabled: Exit Sub
    End If
    With frmC
        .BackColor = GetColor(TransparentBack)
        If IsSelected(UsePictureFile) Then 'Use background picture
            PrepairPicture ClockFace
            If IsSelected(DrawOnPicture) Then Call DrawClock
        Else
            If FacePicRadius% <> c_DefautRadius And FacePicRadius% <> 0 Then
                CR% = FacePicRadius%
            Else: CR% = c_DefautRadius: End If
            CR% = CR% * Clock.HscValue(SizeFactor): .Picture = LoadPicture("")
        'Resize clock form & compensate for borders
            dx% = (.Width - (.ScaleWidth * Screen.TwipsPerPixelX))
            dy% = (.Height - (.ScaleHeight * Screen.TwipsPerPixelY))
            .Width = (CR% + dx%) * 2 * Screen.TwipsPerPixelX
            .Height = (CR% + dy%) * 2 * Screen.TwipsPerPixelY
            Call DrawClock
        End If
    'Save background picture to img(0) to be used if save backpicture
        frmS.pic(5) = .Picture: frmS.pic(5).BackColor = .BackColor
        frmS.pic(5).Width = .Width: frmS.pic(5).Height = .Height
        frmS.Img(0) = frmS.pic(5).Image: frmS.pic(5).Picture = LoadPicture()
        Call SetTranslucent
    End With
    IsDirty = True: CtrlEnabled: bRedraw = True
End Sub

Sub SetTranslucent()
Dim R&, L&
    With frmC
        'Make form transparent and set translucent value
        R& = GetWindowLong(.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetWindowLong .hwnd, GWL_EXSTYLE, R&
        L& = GetHscValue(Translucent): R& = IIf(L& = 255, LWA_COLORKEY, LWA_BOTH)
        SetLayeredWindowAttributes .hwnd, GetColor(TransparentBack), L&, R&
    End With
End Sub

Sub DrawClock() 'Draw the clock picture to use in frmS (The Clock)
Dim X%, Y%, Z%, L!, A!, n%, M%, i%, C&, C2&, CG&(), Sz!, MSz!, b As Boolean
    With frmC
        .DrawWidth = 2: .FillStyle = 1: .Cls
        .BackColor = GetColor(TransparentBack): Screen.MousePointer = 11
    'Draw circle inner area
        If IsSelected(DrawBackground) Then
            X% = CR% * GetHscValue(BackgroundPos): Y% = CR% * GetHscValue(BackgroundSize)
            n% = Y% - X% + 1: i% = 0: BlendColors GetColor(Background2), GetColor(Background), n%, CG&()
            For Z% = X% To Y%
                If Z% < X% + 2 Or Z% > Y% - 2 Then 'Just around edges
                    Gs.CircleGP .hDC, CR%, CR%, Z%, Z%, CG&(i%), _
                      IIf(Z% = X% Or Z% = Y%, Thin, Thick)
                Else
                    frmC.Circle (CR%, CR%), Z%, CG&(i%) 'Standard circle to speed up
                End If
                i% = i% + 1
            Next
        End If
    'Draw outer clock border
        If IsSelected(DrawBorder) Then
            Y% = CR% * GetHscValue(BorderPos): X% = Y% * (1 - GetHscValue(BorderSize))
            n% = Y% - X% + 1: i% = 0: BlendColors GetColor(Border2), GetColor(Border), n%, CG&()
            For Z% = X% To Y%
                If Z% < X% + 2 Or Z% > Y% - 2 Then 'Just around edges
                    Gs.CircleGP .hDC, CR%, CR%, Z%, Z%, CG&(i%), _
                      IIf(Z% = X% Or Z% = Y%, Thin, Thick)
                Else
                    frmC.Circle (CR%, CR%), Z%, CG&(i%) 'Standard circle to speed up
                End If
                i% = i% + 1
            Next
        End If
    'Draw Hour spots or lines
        If IsSelected(DrawHourSpots) Then
            L! = (.ScaleHeight / 2) * GetHscValue(HourSpotsPos)
            M% = GetHscValue(HourSpotsSize): C& = GetColor(HourSpots)
            b = IsSelected(DrawHourLines): M% = M% / IIf(b, 1, 3)
            n% = IIf(bDeci, 10, 12): A! = 360 / n%: .DrawWidth = 1
            Sz! = L! - (M% / 100 * L!): M% = M% * (GetHscValue(SizeFactor) / 2)
            If Not b Then
                BlendColors GetColor(HourSpots2), C&, M%, CG&()
            End If
            For Z% = 1 To n%
                'Calculate X & Y for the Spot/Line
                X% = L! * Cos(Pi / 180 * ((A! * Z%) - 90)) + CR%
                Y% = L! * Sin(Pi / 180 * ((A! * Z%) - 90)) + CR%
                If b Then 'Line
                    Gs.LineGP .hDC, X%, Y%, _
                      Sz! * Cos(Pi / 180 * ((A! * Z%) - 90)) + CR%, _
                      Sz! * Sin(Pi / 180 * ((A! * Z%) - 90)) + CR%, C&
                Else     'Circel
                    If M% > 0 Then
                        For i% = 1 To M% 'Draw circles gradiet filled if diff HourSpots & HourSpots2
                            Gs.CircleGP .hDC, X%, Y%, i%, i%, CG&(i% - 1), _
                              IIf(i% = 1 Or i% = M%, Thin, Thick)
                        Next
                    End If
                    SetPixel frmC.hDC, X%, Y%, GetColor(HourSpots2)   'Just set one pixel in center
                End If
            Next
        End If
    'Draw minute spots or lines
        If IsSelected(DrawMinuteSpots) Then
            L! = (.ScaleHeight / 2) * GetHscValue(MinuteSpotsPos)
            M% = GetHscValue(MinuteSpotsSize): C& = GetColor(MinuteSpots)
            b = IsSelected(DrawMinuteLines): M% = M% / IIf(b, 1, 10): .DrawWidth = 1
            n% = IIf(bDeci, 100, 60): A! = 360 / n%: Sz! = L! - (M% / 100 * L!)
            If Not b Then
                M% = M% * (GetHscValue(SizeFactor) / 2)
                BlendColors GetColor(MinuteSpots2), GetColor(MinuteSpots), M%, CG&()
            End If
            For Z% = 1 To n%
                'Calculate X & Y for the Spot/Line
                X% = L! * Cos(Pi / 180 * ((A! * Z%) - 90)) + CR%
                Y% = L! * Sin(Pi / 180 * ((A! * Z%) - 90)) + CR%
                If b Then 'Line
                    'Make minute lines 25% longer every 5 minutes
                    MSz! = IIf(Z% Mod 5, L! - (M% / 100 * L!) * 0.75, Sz!)
                    Gs.LineGP .hDC, X%, Y%, _
                      MSz! * Cos(Pi / 180 * ((A! * Z%) - 90)) + CR%, _
                      MSz! * Sin(Pi / 180 * ((A! * Z%) - 90)) + CR%, C&
                Else     'Circel
                    If M% > 0 Then
                        For i% = 1 To M% 'Draw circles gradiet filled if diff MinuteSpots & MinuteSpots2
                            Gs.CircleGP .hDC, X%, Y%, i%, i%, CG&(i% - 1), _
                              IIf(i% = 1 Or i% = M%, Thin, Thick)
                        Next
                    End If
                    SetPixel frmC.hDC, X%, Y%, GetColor(MinuteSpots2) 'Just set one pixel in cener
                End If
            Next
        End If
    'Print numbers
        If IsSelected(DrawNumbers) Then
            'If draw hour or minute numbers
            If IsSelected(DrawNumbersMinutes) Then
                n% = IIf(bDeci, 100, 60)
            Else: n% = IIf(bDeci, 10, 12): End If
            A! = 360 / n%: Sz! = GetHscValue(NumbersSize)
            L! = (.ScaleHeight / 2) * GetHscValue(NumbersPos)
            C& = GetColor(Numbers): .ForeColor = C&
            For Z% = 1 To n%
                'Calculate X & Y for the text
                X% = L! * Cos(Pi / 180 * ((A! * Z%) - 90)) + CR%
                Y% = L! * Sin(Pi / 180 * ((A! * Z%) - 90)) + CR%
                If n% = Z% Then
                    If IsSelected(DrawNumbersMinutes) Then M% = 0 Else M% = Z%
                Else: M% = Z%: End If
                PrintText Trim$(M%), 0, 14 * Sz!, _
                  GetColor(Numbers), X%, Y%
            Next
        End If
        'Save as clock bakground picture
        Erase CG&(): .Picture = .Image: .Refresh: Screen.MousePointer = 0
    End With
    
End Sub

Sub PrintText(sText$, FactorY!, FntSz!, Optional Color& = vbBlack, _
  Optional CurX%, Optional CurY%, Optional FactorX!)
Dim X%, Y%
On Local Error GoTo ErrPrintText ' Print a text string to frmC (The clock)
TryAgain:
    With frmC
        .FontSize = FntSz! * Clock.HscValue(SizeFactor)
        .FontName = Clock.FontName
        .ForeColor = Color&
        If FactorY! <> 0 Then
            X% = CR% * FactorX! - .TextWidth(sText$) / 2
            Y% = CR% * FactorY! - .TextHeight(sText$) / 2
        Else
            X% = CurX% - .TextWidth(sText$) / 2
            Y% = CurY% - .TextHeight(sText$) / 2
        End If
        .CurrentX = X%: .CurrentY = Y%
        frmC.FontItalic = False: frmC.FontBold = False
        frmC.Print sText$
    End With
    Exit Sub
ErrPrintText:
    bErr = True: Call Main: Resume TryAgain
End Sub

Sub SetTopMost(hwnd&, Optional ForceTopmost As Boolean) 'Set to Top most or not
    If IsSelected(Topmost) Or ForceTopmost Then
        SetWindowPos hwnd&, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos hwnd&, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Function IsSelected(Value As ChkValues) As Boolean
On Local Error Resume Next 'Get value check boxes
    IsSelected = Clock.ChkValue(Value)
    If Err = 9 Then Err.Clear: Exit Function
    If Err.Number <> 0 Then bErr = True: Call Main
End Function

Function GetColor(ByVal Value As CmdColors) As Long
On Local Error Resume Next 'Get colors to command buttons
    GetColor = Clock.CmdColors(Value)
    If Err.Number <> 0 Then bErr = True: Call Main
End Function

Function GetHscValue(ByVal Value As HscValues) As Single
On Local Error Resume Next 'Get value from horizintal scrollbars
    GetHscValue = Clock.HscValue(Value)
    If Err.Number <> 0 Then bErr = True: Call Main
End Function

Function FileToByteArray(ByVal FileName$) As Byte()
Dim X%, b() As Byte  'Load font, picture or sound file to byte array
    On Error GoTo LoadFileErr
    If Dir(FileName$, vbNormal Or vbArchive) = "" Then Exit Function
    X% = FreeFile
    Open FileName$ For Binary Access Read As #X%
    ReDim b(0 To LOF(X%) - 1)
    Get #X%, , b: Close #X%
    FileToByteArray = b: Erase b()
    Exit Function
LoadFileErr:
    MsgBox Err.Description, 16 + vbSystemModal, "FileToByteArray"
End Function

Function PicFromByteArray(ByteArray() As Byte, Optional OpenedFile$) As IPicture
Dim LB&, ByteCount&, hMem&, lpMem&, IID_IPicture(15), istM As stdole.IUnknown, b As Boolean
'Read stored picture from byte array and return the picture
On Error GoTo Err_PicFromByteArray
    If UBound(ByteArray, 1) < 0 Then Exit Function
    LB = LBound(ByteArray)
    ByteCount = (UBound(ByteArray) - LB) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, ByteArray(LB), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istM) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), _
                  IID_IPicture(0)) = 0 Then
                  Call OleLoadPicture(ByVal ObjPtr(istM), _
                    ByteCount, 0, IID_IPicture(0), PicFromByteArray)
                End If
            End If
        End If
    End If
    If Len(OpenedFile$) Then
        frmS.pic(5) = PicFromByteArray: If frmS.pic(5) = 0 Then b = True
        If b Then MsgBox "Invalid graphic file." & vbLf & OpenedFile$, vbSystemModal + 16: OpenedFile$ = ""
    End If
    Exit Function
Err_PicFromByteArray:
    If Err.Number = 9 Then
        'Uninitialized array
       '' MsgBox "You must pass a non-empty byte array to this function!", 64 + vbSystemModal
    Else
        MsgBox Err.Description + vbSystemModal, 16
    End If
End Function

Function SaveQ() As Boolean
Dim R%
    bCancelSave = False: If Not IsDirty Then Exit Function
    R% = MsgBox("Do you like to save changes to '" & _
      sFile$ & "'?", vbYesNoCancel + vbQuestion + vbSystemModal)
    If R% = vbCancel Then SaveQ = True: bCancelSave = True: Exit Function
    If R% = vbYes Then frmS.mnuFile_Click 2  'Save file
    If R% = vbNo Then UndoRedo DoKillTmp     'Remove Undo Tempfile
End Function

Sub SaveSettings() 'Save clock settings to disk
Dim X%, s$
    With Clock
        .Top = frmC.Top: .Left = frmC.Left: .TopS = frmS.Top: .LeftS = frmS.Left
        .ClockRadiusPxOrg = frmC.ScaleWidth / 2: .Version = sVerNow$: .FileTest = c_FileTest
        .TimeZoneName = GetTimeZone(X%): .TimeZoneSaved = X%  'Save TimeZone
        If IsSelected(SaveFontToFile) Then 'Save selected font in clock file
            .FontFileName = GetFontFileName(.FontName, True)
            s$ = GetFontFileName(.FontName)
            If Len(s$) Then .FontFile() = FileToByteArray(s$)
        Else: Erase .FontFile(): .FontFileName = "": End If
    End With
    Call SaveStartUpWindows
    If Dir(sFile$) <> "" Then Kill sFile$ 'Remove first might get it smaller
    X% = FreeFile: Open sFile$ For Binary As X%
    Put X%, , Clock: Close X%: UndoRedo DoKillTmp: UndoRedo 'Remove Undo Tempfile & Reset
    bCancelSave = False: IsDirty = False: SetCaption: CtrlEnabled
End Sub

Function CheckDLL() As Boolean
Dim X%, s$, ArrDLL() As Byte: Const C = "FoxCBmp3.dl"
'Check if Dll exist in Sys-path, if not get it from resource file and put it to disk.
    s$ = Space$(260)
    GetSystemDirectory s$, Len(s$)
    s$ = Left$(s$, InStr(s$ & vbNullChar, vbNullChar) - 1)
    s$ = s$ & "\" & C
    If Dir(s$) = "" Then
        ArrDLL = LoadResData(C, c_Cust)
        X% = FreeFile: Open s$ For Binary As X%
        Put X%, , ArrDLL: Close X%
        CheckDLL = True: Erase ArrDLL: DoEvents
        'Is compiled so need to restart so VB understand that the dll now exists
        If App.LogMode <> 0 Then
            If nDll% > 1 Then MsgBox "Need to restart coz, '" & s$ & "', was missing.", 48
            s$ = App.Path & "\" & App.Title & ".exe"
            If Dir(s$) <> "" Then
                RunCommand s$: DoEvents: Unload frmC: End
            End If
        End If
    End If
End Function

Function SetCaption$()
Dim s$: s$ = sFile$ 'Change caption in settings form
    If sFile$ = "" Then sFile$ = c_NewFile
    If sFile$ <> c_NewFile Then
        s$ = GetFileTitle(s$): s$ = Left$(s$, Len(s$) - 4)
    End If
    Clock.FileName = s$
    SetCaption$ = App.Title & "  [" & s$ & "]" & _
      IIf(IsDirty, "  *", "") ' & GetTimeZone
    With Clock
        If IsSelected(AlarmOn) Then
            s$ = "  Alarm is on - " & Format(.AlarmHour, c_F) & _
              ":" & Format(.AlarmMinute, c_F)
        Else: s$ = "": End If
    End With
    SetCaption$ = SetCaption$ & s$: frmS.Caption = SetCaption$
End Function

Private Function GetTimeZone(Optional RetValMinutesOffSet%) As String
Dim TZI As TIME_ZONE_INFORMATION, s$, T%, D%
'Get timezone and minutes offset from GMT
   Select Case GetTimeZoneInformation(TZI)
      Case 0:  s$ = "Cannot determine current time zone."
      Case 1:  s$ = TZI.StandardName
      Case 2:  s$ = TZI.DaylightName: D% = 60
   End Select
   T% = TZI.Bias * -1
   RetValMinutesOffSet% = TZI.Bias + RetValMinutesOffSet% * -1
   GetTimeZone = StripTerminator(s$) & _
     " (GMT " & IIf(T% > 0, "+", "") & T% / 60 & " h)"
   RetValMinutesOffSet% = D% + T%
End Function

Sub CtrlEnabled()
Dim X%, s$, b As Boolean, b2 As Boolean: Static IsHere As Boolean
'Enable/Disable controls in form frmS
    If bNotNow Or IsHere Then Exit Sub
    If Len(sTmpUndo$) = 0 Then Exit Sub
    IsHere = True
    On Error GoTo ErrCtrlEnabled
    With frmS
    'Undo/Redo
        b = IIf(UndoPos% > 0, True, False)
        .mnuOpt(10).Enabled = b  'Undo
        If b Then
            s$ = "&Undo " & UndoRedos(UndoPos%).InfoText & _
              "  (Max " & UndoPos% & " Step" & _
              IIf(UndoPos% = 1, "", "s") & ")"
        Else: s$ = "Can't undo": End If
        .mnuOpt(10).Caption = s$
        b = IIf(UndoPos% < UBound(UndoRedos), True, False)
        .mnuOpt(11).Enabled = b 'Redo
        If b Then
            X% = UBound(UndoRedos) - UndoPos%
            s$ = "&Redo " & UndoRedos(UndoPos% + 1).InfoText & _
              "  (Max " & X% & " Step" & IIf(X% = 1, "", "s") & ")"
        Else: s$ = "Can't Redo": End If
        .mnuOpt(11).Caption = s$
    'Backpicture menu
        b2 = False
        b = Not IsArrayEmpty(Clock.PictureBack)
        .mnuFile(8).Checked = b         'Select Back Picture
        .mnuFile(9).Enabled = b         'Remove Back Picture
        b = IIf(b And .chk(UsePictureFile).Value = 1, True, False)
        .chk(AlphaFace).Enabled = b     'Checkbox Use Alpha mask on clock face
        'Scrollbars Clock Face, Hue, Saturation, Brightness...
        .hsc(FaceHue).Enabled = b: .hsc(FaceSaturation).Enabled = b: .hsc(FaceBrightness).Enabled = b
    'Hands picture menu
        b = Not IsArrayEmpty(Clock.PictureHourHand)
        If b Then b2 = True
        .mnuFile(10).Checked = b    'Select Hour hand Picture
        .mnuFile(5).Enabled = b     'Save Hour hand Picture
        'Scrollbars Hour hand for, Hue, Saturation, Brightness, Alpha...
        b = IIf(.chk(ShowHourHand) And b, True, False)
        .hsc(HHue).Enabled = b: .hsc(HAlfa).Enabled = b
        .hsc(HSaturation).Enabled = b: .hsc(HBrightness).Enabled = b
        b = Not IsArrayEmpty(Clock.PictureMinuteHand)
        If b Then b2 = True
        .mnuFile(11).Checked = b    'Select Minute hand Picture
        .mnuFile(6).Enabled = b     'Save Minute hand Picture
        'Scrollbars Minute hand for, Hue, Saturation, Brightness, Alpha...
        b = IIf(.chk(ShowMinuteHand) And b, True, False)
        .hsc(MHue).Enabled = b: .hsc(MAlfa).Enabled = b
        .hsc(MSaturation).Enabled = b: .hsc(MBrightness).Enabled = b
        b = Not IsArrayEmpty(Clock.PictureSecondHand)
        If b Then b2 = True
        .mnuFile(12).Checked = b    'Select Second hand Picture
        .mnuFile(7).Enabled = b     'Save Second hand Picture
        .mnuFile(13).Enabled = b2       'Remove hand Pictures
        'Scrollbars Second hand for, Hue, Saturation, Brightness, Alpha...
        b = IIf(.chk(ShowSecondHand) And b, True, False)
        .hsc(SHue).Enabled = b: .hsc(SAlfa).Enabled = b
        .hsc(SSaturation).Enabled = b: .hsc(SBrightness).Enabled = b
        .chk(PictureHands).Enabled = b2 'Checkbox Use Picturefiles as hands
        b = IIf(b2 And .chk(PictureHands).Value = 1, True, False)
        .chk(AlphaHands).Enabled = b     'Checkbox Use Alpha mask on hands
        If Not b Then 'Disable scroll adjustments for hand pictures
            For X% = HHue To SAlfa: .hsc(X%).Enabled = False: Next
        End If
    'Option menu items
        'Reset pictures to orginal size
        .mnuOpt(1).Enabled = IIf(.hsc(SizeFactor) <> c_100 Or .hsc(HandHourSize) <> c_100 Or _
          .hsc(HandMinuteSize) <> c_100 Or .hsc(HandSecondSize) <> c_100, True, False)
        'Reset alarm sound to default from resourcefile
        .mnuOpt(2).Enabled = IIf(IsArrayEmpty(Clock.AlarmWavFile), False, True)
        'Center Hour hand
        .mnuOpt(3).Enabled = IIf(.hsc(HandHourX) <> c_100 Or .hsc(HandHourY) <> c_100, True, False)
        'Center Minute hand
        .mnuOpt(4).Enabled = IIf(.hsc(HandMinuteX) <> c_100 Or .hsc(HandMinuteY) <> c_100, True, False)
        'Center Second hand
        .mnuOpt(5).Enabled = IIf(.hsc(HandSecondX) <> c_100 Or .hsc(HandSecondY) <> c_100, True, False)
        'Reload clockfile
        .mnuOpt(7).Enabled = IIf(sFile$ <> c_NewFile, True, False)
        .mnuOpt(7).Caption = "Rel&oad " & IIf(sFile$ <> c_NewFile, Clock.FileName, "")
    'Second Hand
        b = .chk(ShowSecondHand).Value
        .chk(ShowSmooth).Enabled = b
        If Not b Then .chk(ShowSmooth).Value = b
        b = .chk(ShowSmooth).Value: '.chk(SecondTick).Enabled = Not B
        'If B Then .chk(SecondTick).Value = Not B
    'Alarm
        b = IsSelected(AlarmOn)
        .cmd(SoundPlay).Enabled = b: .cmd(SoundSelect).Enabled = b
        .chk(AlarmSound).Enabled = b: .txt(1).Enabled = b
        .cmb(2).Enabled = b: .cmb(3).Enabled = b
        .opt(2).Enabled = b: .opt(3).Enabled = b
        .lbl(2).Enabled = b
    'Minimize to tray
        '.chk(MinimizeToTray).Enabled = .chk(RunWinStartUp).Value
    'Draw on picture...
        b = CBool(.chk(UsePictureFile).Value)
        .chk(DrawOnPicture).Enabled = b
        .chk(DrawNumbersMinutes).Enabled = CBool(.chk(DrawNumbers).Value)
        b = IIf(Not b Or (.chk(DrawOnPicture) And b), True, False)
        For X% = 9 To 13: .chk(X%).Enabled = b: Next
        For X% = 0 To 4: .cmd(X%).Enabled = b: Next
        For X% = 21 To 27: .cmd(X%).Enabled = b: Next
        If Not b Then .chk(DrawNumbersMinutes).Enabled = b
        .chk(CenterCircle).Enabled = b
        .txt(0).Enabled = CBool(.chk(DrawMyText).Value)
        b = IIf(.chk(DrawHourSpots).Enabled And .chk(DrawHourSpots), True, False)
        .chk(DrawHourLines).Enabled = b
        b = IIf(.chk(DrawMinuteSpots).Enabled And .chk(DrawMinuteSpots), True, False)
        .chk(DrawMinuteLines).Enabled = b
    'Horizontal scrollbars
        b = .chk(ShowHourHand): .hsc(HandHourSize).Enabled = b
        .hsc(HandHourX).Enabled = b: .hsc(HandHourY).Enabled = b: .hsc(HandHourStyle).Enabled = b
        b = .chk(ShowMinuteHand): .hsc(HandMinuteSize).Enabled = b
        .hsc(HandMinuteX).Enabled = b: .hsc(HandMinuteY).Enabled = b: .hsc(HandMinuteStyle).Enabled = b
        b = .chk(ShowSecondHand): .hsc(HandSecondSize).Enabled = b
        .hsc(HandSecondX).Enabled = b: .hsc(HandSecondY).Enabled = b: .hsc(HandSecondStyle).Enabled = b
        b2 = IIf(.chk(DrawOnPicture) = 1 Or .chk(DrawOnPicture).Enabled = False, True, False)
        b = .chk(DrawBorder) And b2: .hsc(BorderPos).Enabled = b: .hsc(BorderSize).Enabled = b
        b = .chk(DrawHourSpots) And b2: .hsc(HourSpotsPos).Enabled = b: .hsc(HourSpotsSize).Enabled = b
        b = .chk(DrawMinuteSpots) And b2: .hsc(MinuteSpotsPos).Enabled = b: .hsc(MinuteSpotsSize).Enabled = b
        b = .chk(DrawBackground) And b2: .hsc(BackgroundPos).Enabled = b: .hsc(BackgroundSize).Enabled = b
        b = .chk(DrawNumbers) And b2: .hsc(NumbersPos).Enabled = b: .hsc(NumbersSize).Enabled = b
        b = .chk(ShowDigital): .hsc(DigiClockX).Enabled = b
        .hsc(DigiClockY).Enabled = b: .hsc(DigiClockSize).Enabled = b
        b = .chk(ShowWeekday): .hsc(WeekdayX).Enabled = b
        .hsc(WeekdayY).Enabled = b: .hsc(WeekdaySize).Enabled = b
        b = .chk(ShowDate) Or .chk(ShowDateMonth): .hsc(DateX).Enabled = b
        .hsc(DateY).Enabled = b: .hsc(DateSize).Enabled = b
        'b = .chk(DrawMyText): .hsc(MyTextSize).Enabled = b
        '.hsc(MyTextY).Enabled = b: .hsc(MyTextX).Enabled = b
        bNotNow = False: SetCaption: IsHere = False
    End With
    Exit Sub
ErrCtrlEnabled:
    Err.Clear: IsHere = False
End Sub

Sub FillAlarmCombos(Optional Changed As Boolean)
Dim X%, Y%, H%, M%, AH&, AM%, T!, L&, b As Boolean: Const C = "00"
    With frmS
        AH& = Clock.AlarmHour: AM% = Clock.AlarmMinute
        If Changed And (AH& <> 0 Or AM% <> 0) Then b = True 'Convert alarm time
        If Clock.DecimalClock Then
            If b Then 'Convert from standard to decimal time
                T! = (AH& * 3600) + (AM% * 60) 'Seconds past since midnight
                AH& = ConvToDecTime!(DecHour, , T!)
                AM% = ConvToDecTime!(DecMinute, , T!)
            End If
            H% = 9: M% = 99
        Else
            If b Then 'Convert from decimal to standard time
                'Decimal seconds past since midnight
                T! = ((AH& * 10000) + (AM% * 100)) / 100000 '% time past since  00:00
                T! = c_Sec24h& * T!: L& = Fix(T! / 3600): AH& = L&
                L& = (T! - (L& * 3600)) / 60: AM% = L&
                L& = T! - (AH& * 3600) - (AM% * 60) 'l& = Seconds diff here
            End If
            H% = 23: M% = 59
        End If
        For X% = 2 To 3: .cmb(X%).Clear: Next
    'Alarm hour time
        For X% = 0 To H%
            .cmb(2).AddItem Format(X%, c_F)
        Next
        .cmb(2).Text = Format(AH&, c_F)
        'If Val(.cmb(2).Text) > H% Then .cmb(2).Text = C
    'Alarm minute time
        For X% = 0 To M%
            .cmb(3).AddItem Format(X%, c_F)
        Next
        .cmb(3).Text = Format(AM%, c_F)
        'If Val(.cmb(3).Text) > M% Then .cmb(3).Text = C
        Clock.AlarmHour = AH&: Clock.AlarmMinute = AM%
    End With
End Sub

Sub LoadSettings()
Dim X%, Y%, s$
TryAgain: 'Load settings from the type "Clock"
    On Error GoTo ErrLoadSettings
    If bNotNow Then Exit Sub
    With Clock
        bNotNow = True: bUndoRedo = True
    'Color buttons
      'Have added a command button control since last version
        If frmS.cmd.UBound > UBound(.CmdColors) Then
            Y% = UBound(.CmdColors)
            ReDim Preserve .CmdColors(frmS.cmd.UBound)
            'Apply new default values from command button
            For X% = Y% + 1 To frmS.cmd.UBound
                .CmdColors(X%) = frmS.cmd(X%).BackColor
            Next
        End If
        For X% = 0 To frmS.cmd.UBound
            frmS.cmd(X%).BackColor = .CmdColors(X%)
            If frmS.cmd(X%).Tag <> "" Then _
              frmS.cmd(X%).ToolTipText = "Select " & _
                frmS.cmd(X%).Tag & " color..."
        Next
    'Checkboxes
      'Have added a checkbox control since last version
        If frmS.chk.UBound > UBound(.ChkValue) Then
            Y% = UBound(.ChkValue)
            ReDim Preserve .ChkValue(frmS.chk.UBound)
            'Apply new default values from checkbox
            For X% = Y% + 1 To frmS.chk.UBound
                .ChkValue(X%) = frmS.chk(X%)
            Next
        End If
        For X% = 0 To frmS.chk.UBound
            frmS.chk(X%) = .ChkValue(X%)
        Next
    'Set maxvalue on Horizontal scrollbars depending-
    'of avalible number of polygon Hand-styles
        frmS.hsc(HandHourStyle).Max = c_NoOffPolyStyles * c_100
        frmS.hsc(HandMinuteStyle).Max = c_NoOffPolyStyles * c_100
        frmS.hsc(HandSecondStyle).Max = c_NoOffPolyStyles * c_100
    'H-Scroll values
      'Have added a horizontal scrollbar control since last version
        If frmS.hsc.UBound > UBound(.HscValue) Then
            Y% = UBound(.HscValue)
            ReDim Preserve .HscValue(frmS.hsc.UBound)
            'Apply new default values from horizontal scrollbar
            For X% = Y% + 1 To frmS.hsc.UBound
                .HscValue(X%) = frmS.hsc(X%) / c_100
            Next
        End If
        For X% = 0 To frmS.hsc.UBound
            frmS.hsc(X%) = .HscValue(X%) * c_100
        Next
    'Textboxes
        frmS.txt(0) = .MyText           'Custom clock text
        frmS.txt(1) = .AlarmCommand     'Alarm Command
        frmS.txt(2) = .Comments         'Clock Comments
    'Get settings form position
        frmS.Top = .TopS: frmS.Left = .LeftS
    'Clear all combo's
        For X% = 0 To frmS.cmb.UBound: frmS.cmb(X%).Clear: Next
    'Font
        For X% = 0 To Screen.FontCount - 1
            s$ = Screen.Fonts(X%) 'Only fonts found in Registry
            If Len(GetFontFileName(s$, True)) Then
                frmS.cmb(0).AddItem s$
            End If
        Next
        frmS.cmb(0).Text = .FontName
    'Timezones
        For X% = 11 To -11 Step -1
            frmS.cmb(1).AddItem Trim$(X%)
        Next
        frmS.cmb(1).Text = .TimeZoneOffset
    'Alarm message or run command
        frmS.opt(IIf(.AlarmShowMsgBox, 3, 2)).Value = True
    'Decimal or Standard clock
        frmS.opt(IIf(.DecimalClock, 0, 1)).Value = True
    'Alarm combos HH:MM
        Call FillAlarmCombos: bDeci = .DecimalClock
    End With
    bNotNow = False: IsDirty = False: Call UndoRedo 'Reset type to keep track on changes
    bUndoRedo = False: Call CtrlEnabled: SetCaption
    Exit Sub
ErrLoadSettings:
    If Err = 9 Then Resume Next 'Have added indexed control
    bErr = True: Call Main: Resume TryAgain
End Sub

Sub UndoRedo(Optional Action As UndoRedoAction = DoReset, Optional ByVal Ctrl As ControlType, _
  Optional ByVal CtrlIdx%, Optional ByVal PrevValue, Optional ByVal NewValue)
Dim Val, X%, Y%, Z%, s$ 'Undo/Redo handeling

    On Local Error GoTo ErrUndoRedo
    If Action = DoUndo Or Action = DoRedo Then 'Undo or Redo
        bUndoRedo = True: If Action = DoRedo Then UndoPos% = UndoPos% + 1
        With UndoRedos(UndoPos%)
            Val = IIf(Action = DoUndo, .CtrlValueOld, .CtrlValueNew)
            Select Case .CtrlTyp
                Case DoURhsc 'Horizontal scrollbar
                    frmS.hsc(.CtrlIndex).Value = Val
                Case DoURcmd 'Command button (Colors)
                    frmS.cmd(.CtrlIndex).BackColor = Val
                    Clock.CmdColors(.CtrlIndex) = Val
                    If .CtrlIndex = TransparentBack Then LoadClockPictures
                    SetStyle
                Case DoURchk 'Checkbox
                    frmS.chk(.CtrlIndex).Value = Val
                Case DoURpic 'Changed picture
                    X% = Choose(.CtrlIndex + 1, 10, 11, 12, 8)
                    Tmp() = Val: frmS.mnuFile_Click X%
                Case DoURcmb 'Combo box
                    Y% = .CtrlIndex
                    With frmS.cmb(.CtrlIndex)
                        Z% = .ListCount - 1
                        For X% = 0 To Z%
                            If Y% = 0 Then 'Font (String value)
                                If .List(X%) = Val Then .ListIndex = X%: Exit For
                            Else           'Timezone & Alarm HH & MM
                                If CInt(.List(X%)) = CInt(Val) Then .ListIndex = X%: Exit For
                            End If
                        Next
                    End With
                Case DoURopt 'Option button
                    X% = Val: Y% = IIf(Val = .CtrlValueNew, .CtrlValueOld, .CtrlValueNew)
                    With frmS
                        .opt(Y%).Value = False: .opt(X%).Value = True
                    End With
            End Select
        End With
        If Action = DoUndo Then UndoPos% = UndoPos% - 1
        bUndoRedo = False: CtrlEnabled
    ElseIf Action = DoSave Then           'Save old and new values
        UndoPos% = UndoPos% + 1: ReDim Preserve UndoRedos(UndoPos%)
        With UndoRedos(UndoPos%)
            If Ctrl = DoURpic Then 'Changing picture
                s$ = Choose(CtrlIdx% + 1, "hour hand", "minute hand", _
                  "second hand", "clock face") & " picture"
            End If
            .CtrlTyp = Ctrl: .CtrlIndex = CtrlIdx%
            .CtrlValueNew = NewValue: .CtrlValueOld = PrevValue
            .InfoText = Choose(Ctrl + 1, "scrollbar value", _
              "color", "click on checkbox", s$, "combo box", "option button")
        End With
    ElseIf Action = DoReset Then 'Reset saved Undos...
        UndoPos% = 0: ReDim UndoRedos(0): s$ = ".undo"
        If Dir(sFile$) = "" Or sFile$ = c_NewFile Then
            If sFile$ = c_NewFile Then Clock.FileName = sFile$
            sTmpUndo$ = App.Path & "\" & c_NewFile & s$
        Else
            sTmpUndo$ = Left$(sFile$, Len(sFile$) - 4) & s$
        End If
        If Dir(sTmpUndo$) <> "" Then  'Temporary Undo file exists for the clock file
            X% = FreeFile: Open sTmpUndo$ For Binary As X%
            Get X%, , UndoFile: Close X%
            ReDim UndoRedos(UndoFile.Ubounds) 'Dump the file back into the UDT UndoRedos.
            UndoRedos() = UndoFile.Undos(): UndoPos% = UndoFile.UndoPos
            If UBound(UndoRedos) > 0 Then
                s$ = App.Title & " did not exit properly last time you did work with '" & Clock.FileName & "'." & _
                  vbLf & vbLf & "Do you like to get back to previous undo point where it did terminate?"
                If MsgBox(s$, 48 + vbYesNo + vbSystemModal) = vbYes Then
                    frmC.Visible = False: frmC.tmr.Enabled = False
                    Y% = UndoPos% - 1: UndoPos% = 0: s$ = sTmpUndo$: sTmpUndo$ = ""
                    Do Until UndoPos% = Y%  'Move to the position just before it crached
                        UndoRedo DoRedo
                    Loop
                    sTmpUndo$ = s$: SetStyle
                    frmC.Visible = True: frmC.tmr.Enabled = True
                    frmS.Visible = True: IsDirty = True: CtrlEnabled
                Else
                    UndoRedo DoKillTmp: UndoRedo: CtrlEnabled
                End If
            Else
                UndoRedo DoKillTmp: UndoRedo: CtrlEnabled
            End If
        End If
    ElseIf Action = DoKillTmp Then 'Remove temporary undo file
        If Dir(sTmpUndo$) <> "" Then Kill sTmpUndo$
    End If
    If Action <> DoKillTmp And Action <> DoReset And sTmpUndo$ <> "" Then
        'Save undos temporary to disk if crach...:(
         If Dir(sTmpUndo$) <> "" Then Kill sTmpUndo$ 'Remove first might get it smaller
         X% = FreeFile: UndoFile.Ubounds = UBound(UndoRedos)
         UndoFile.Undos() = UndoRedos(): UndoFile.UndoPos = UndoPos%
         Open sTmpUndo$ For Binary As X%
         Put X%, , UndoFile: Close X%
    End If
    Exit Sub
ErrUndoRedo:
    If Action = DoReset And (Err = 13 Or Err = 7) Then 'Not a valid Undo Temp-File
        MsgBox "The undofile '" & sTmpUndo$ & "' was damage.", 48 + vbSystemModal
        Err.Clear: On Local Error Resume Next
        Close #X%: If Err <> 0 Then Err.Clear
        UndoRedo DoKillTmp: UndoRedo DoReset: Exit Sub
    End If
    MsgBox Err.Description & vbLf & sTmpUndo$, 16 + vbSystemModal
    bUndoRedo = False: Err.Clear
End Sub
Sub LoadClockPictures()
Dim X%, b As Boolean
'Load and resize back & hand pictures
    With Clock
     'Backpicture
        If IsSelected(UsePictureFile) Then PrepairPicture ClockFace
    'Hour, minute and second Hand pictures
        For X% = DecHour To DecSecond
            b = Not IsArrayEmpty(Choose(X% + 1, .PictureHourHand(), _
              .PictureMinuteHand(), .PictureSecondHand()))
            If b Then PrepairPicture X% Else BuildPolygon X%
        Next
    End With
End Sub

Sub PrepairPicture(ByVal Hand As Hands)
Dim X%, Y%, Sz!, W&, H&, dx&, dy&, C&, s$, SzX!, SzY!, Hue%, Sat%, Lig%, Abl%
On Error GoTo ErrPrepairPicture
'Resize pictures, if hands so can blit them direct
    If Not IsSelected(PictureHands) And Hand <> ClockFace Then
        Exit Sub
    End If
    If Not IsSelected(UsePictureFile) And Hand = ClockFace Then Exit Sub
    With frmS
        'Clear temporary pictureboxes
        .pic(4) = LoadPicture(""): .pic(5) = LoadPicture("")
        Select Case Hand 'Load Orginal size picture from byte array
            Case DecHour: .pic(5) = PicFromByteArray(Clock.PictureHourHand)
                SzX! = HandHourX: SzY! = HandHourY
                Hue% = HHue: Sat% = HSaturation: Lig% = HBrightness: Abl% = HAlfa
            Case DecMinute: .pic(5) = PicFromByteArray(Clock.PictureMinuteHand)
                SzX! = HandMinuteX: SzY! = HandMinuteY
                Hue% = MHue: Sat% = MSaturation: Lig% = MBrightness: Abl% = MAlfa
            Case DecSecond: .pic(5) = PicFromByteArray(Clock.PictureSecondHand)
                SzX! = HandSecondX: SzY! = HandSecondY
                Hue% = SHue: Sat% = SSaturation: Lig% = SBrightness: Abl% = SAlfa
            Case ClockFace: .pic(5) = PicFromByteArray(Clock.PictureBack)
                frmC.Picture = LoadPicture()
                If .pic(5) = 0 Then Exit Sub
                SzX! = 1: SzY! = 1
        End Select
        If Hand <> ClockFace Then SzX! = GetHscValue(SzX!): SzY! = GetHscValue(SzY!)
        C& = GetColor(TransparentBack)
        For Y% = 0 To 5: .pic(Y%).BackColor = C&: Next
        'Calculate Scale factor
        Sz! = GetHscValue(SizeFactor)
        If Hand <> ClockFace Then 'Hands
            Sz! = GetHscValue(Choose(Hand + 1, HandHourSize, _
              HandMinuteSize, HandSecondSize)) * Sz!
        End If
        W& = .pic(5).ScaleWidth: H& = .pic(5).ScaleHeight
        If W& <> H& Then 'Not a squared picture
            If W& < H& Then
                dx& = H& / 2 - W& / 2: W& = H&
            Else: dy& = W& / 2 - H& / 2: H& = W&: End If
        End If
        'Make picture squared on .pic(4) and-
        'use selected transparent backcolor
        .pic(4).Width = W& * Screen.TwipsPerPixelX
        .pic(4).Height = H& * Screen.TwipsPerPixelY
        'Get transparent color from, source pictures Top-Left Pixel
        If TransparentBlt( _
          .pic(4).hDC, dx&, dy&, .pic(5).ScaleWidth, .pic(5).ScaleHeight, _
          .pic(5).hDC, 0, 0, .pic(5).ScaleWidth, .pic(5).ScaleHeight, _
          GetPixel(.pic(5).hDC, 0, 0)) = False Then _
            MsgBox c_BltErr, 16 + vbSystemModal
        .pic(4).Refresh
        .pic(5).Picture = .pic(4).Image: .pic(4).Cls
        .pic(4).Width = .pic(5).Width * Sz!
        .pic(4).Height = .pic(5).Height * Sz!
        If Sz! <> 1 Then 'Resize picture if scaled
            SetStretchBltMode .pic(4).hDC, COLORONCOLOR
            StretchBlt .pic(4).hDC, 0&, 0&, W& * Sz!, H& * Sz!, _
              .pic(5).hDC, 0, 0, W&, H&, vbSrcCopy
            .pic(4).Refresh
        Else
            .pic(4) = .pic(5)
        End If
    'Blitt hand in correct position so can use sourcepicture direct when rotate
        If Hand <> ClockFace And (SzX! <> 1 Or SzY! <> 1) Then
            SzX! = CR% - .pic(4).ScaleWidth / 2 * SzX!
            SzY! = CR% - .pic(4).ScaleHeight / 2 * SzY!
            .pic(5).Height = frmC.Height: .pic(5).Width = frmC.Width
            .pic(5).Picture = LoadPicture
            If TransparentBlt( _
              .pic(5).hDC, SzX!, SzY!, .pic(4).ScaleWidth, .pic(4).ScaleHeight, _
              .pic(4).hDC, 0, 0, .pic(4).ScaleWidth, .pic(4).ScaleHeight, _
              C&) = False Then _
                MsgBox c_BltErr, 16 + vbSystemModal
              .pic(5).Refresh
        End If
    'Make adjustments to pictures
        If Hand = ClockFace Then 'Clock face
            With frmS.pic(4)
                frmC.Move frmC.Left, frmC.Top, W& * Screen.TwipsPerPixelX * Sz!, _
                  H& * Screen.TwipsPerPixelY * Sz!
                'Adjust Hue, Saturation & Brightness
                FoxHSL .hDC, 0, 0, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, _
                  GetHscValue(FaceHue) * 10, GetHscValue(FaceSaturation), GetHscValue(FaceBrightness), C&, 1&
                If IsSelected(AlphaFace) Then  'Alpha mask clock face
                    frmC.Picture = LoadPicture()
                    FoxAlphaMask frmC.hDC, 0, 0, .ScaleWidth, .ScaleHeight, _
                    .hDC, 0, 0, .hDC, 0, 0, C&, &H1
                    frmC.Picture = frmC.Image
                Else: frmC.Picture = .Image: End If
                TransC(Hand) = GetPixel(frmC.hDC, 0, 0)
                'Ready to use Clock face picture saved to frmC.Picture
                FacePicRadius% = W& / 2: CR% = W& / 2 * Sz!
            End With
        Else    'Clock hands
            'Adjust Alpha blend on the clock hand
            .pic(4) = .pic(5).Image: .pic(4).Picture = LoadPicture()
             FoxAlphaBlend .pic(4).hDC, 0, 0, _
               .pic(5).ScaleWidth, .pic(5).ScaleHeight, _
               .pic(5).hDC, 0, 0, GetHscValue(Abl%), C&, 1&
            .pic(4).Refresh: .pic(5).Picture = LoadPicture()
            'Adjust Hue, Saturation & Brightness on the clock hand
            FoxHSL .pic(5).hDC, 0, 0, _
              .pic(5).ScaleWidth, .pic(5).ScaleHeight, _
              .pic(4).hDC, 0, 0, _
              GetHscValue(Hue%) * 10, GetHscValue(Sat%), GetHscValue(Lig%), C&, 1&
            'Ready to use hands saved to .pic(0-2)
            .pic(5).Refresh: .pic(Hand) = .pic(5).Image
            TransC(Hand) = GetPixel(.pic(5).hDC, 0, 0) 'Save backcolor until rotate
        End If
    End With
    Exit Sub
ErrPrepairPicture:
    If Err = 53 Then 'Can't find the dll-file
        nDll% = nDll% + 1
        If nDll% = 1 Then CheckDLL: DoEvents: Resume
        If nDll% > 2 Then Err.Clear: Exit Sub
    End If
    MsgBox Err.Description, 16 + vbSystemModal: Err.Clear
End Sub

Sub RunCommand(ByVal sCmd$) 'Run a command
    ShellExecute frmC.hwnd, "Open", sCmd$, vbNullString, "C:\", 1
End Sub

Sub PlayWav(Optional ByVal ResID$, _
  Optional PlayLoop As Boolean, Optional StopPlay As Boolean)
Dim L&, b As Boolean: Const Alarm = "CHANGE_H.WAV"
'Play a sound from resourcefile or disk via frmC.PlayAlarmWaveFile
    L& = SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
'Workaround to stop playing the alarm without a default sound
    If StopPlay Then
        ByteSound = LoadResData("CHANGE_S.WAV", c_Cust)
        SNDPLAYSOUND ByteSound(0), SND_MEMORY Or SND_MEMORY
        Erase ByteSound: DoEvents: Exit Sub
    End If
    If (ResID$) = "" Then 'Alarm sound
        'Check if external file selected and exsists in Clock.AlarmWavFile()
        If Not IsArrayEmpty(Clock.AlarmWavFile) Then
            ByteSound = Clock.AlarmWavFile: b = True
        Else
            ResID$ = Alarm 'Play default alarm sound from resource file
        End If
    End If
 'Play sound from Resource fil or Clock.AlarmWavFile()
    If PlayLoop Then L& = L& Or SND_LOOP
    If ResID$ <> Alarm Then
        L& = L& Or SND_NOSTOP 'Dont let clock sounds interupt alarm sound
    End If
    If Not b Then ByteSound = LoadResData(ResID$, c_Cust)
    SNDPLAYSOUND ByteSound(0), L&
End Sub

'/////Register - Autostart & File-association stuff

Sub SaveStartUpWindows()
Dim s$, Ret&: Const c_P = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN"
'Add or remove key in registry to make the clock start when Windows starts
    If Dir(sFile$) = "" Or Clock.FileName = "" Then Exit Sub
    s$ = App.Title & " - " & Clock.FileName
    If Not IsSelected(RunWinStartUp) Then 'Delete value
        'Open the key, exit if not found
        If RegOpenKeyEx(HKEY_LOCAL_MACHINE, c_P, 0, KEY_WRITE, Ret&) Then Exit Sub
        'Delete the value (returns 0 if success)
        RegDeleteValue Ret&, s$
    Else
        'Create a new key
        RegCreateKey HKEY_LOCAL_MACHINE, c_P, Ret&
        'Save a string to the key
        RegSetValueEx Ret&, s$, 0&, REG_SZ, ByVal sFile$, Len(sFile$)
    End If
    'Close the key
    RegCloseKey Ret&
End Sub

Private Sub CreateAssociation(Optional RemoveAssociation As Boolean)
'Create a file-association (*.clc) so can open it if double click
'The exe-file need to have the same name as the App.Title to make this sub work
Const c_X = ".clc", c_P = "\SHELL\OPEN\COMMAND"
Dim i%, Ret&, s$, EP$: s$ = App.Title: EP$ = App.Path & "\" & s$ & ".exe" & " '%1'"
    s$ = Replace(s$, " ", ".")
'Remove association
    If RemoveAssociation Then
        For i% = 0 To 1
            'Open the key
            RegOpenKeyEx HKEY_CLASSES_ROOT, s$, 0, KEY_ALL_ACCESS, Ret&
            If Ret& Then                'Delete key if exists
                RegDeleteKey Ret&, ""   'Delete the key
                RegCloseKey Ret&        'Close the handle
            End If
            s$ = c_X
        Next
'Create association
    Else
        CreateNewKey c_X, HKEY_CLASSES_ROOT
        SetKeyValue c_X, "", s$, REG_SZ
        CreateNewKey s$ & c_P, HKEY_CLASSES_ROOT
        SetKeyValue s$ & c_P, "", EP$, REG_SZ
    End If
End Sub

Private Sub CreateNewKey(sNewKeyName$, lPredefinedKey&)
Dim L&, k&
    L& = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, _
      REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, k&, L&)
    Call RegCloseKey(k&)
End Sub
Private Function SetValueEx(ByVal hKey&, sValueName As String, lType&, vValue As Variant) As Long
Dim nValue&, sValue$
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, _
              0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            nValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, _
              0&, lType, nValue, 4)
    End Select
End Function
Private Sub SetKeyValue(sKeyName$, sValueName$, vValueSetting As Variant, lValueType&)
Dim R&, hKey&
    R& = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    R& = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    Call RegCloseKey(hKey)
End Sub

'///Minimize to tray stuff

Sub TrayMinimizeAppTo() 'Create tray icon
    frmC.Hide: frmS.Hide
    With nID
        .cbSize = Len(nID)
        .hwnd = frmS.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = frmS.Icon
        .szTip = frmS.Caption & vbNullChar
        Shell_NotifyIcon NIM_ADD, nID
        frmS.mnuPop(3).Caption = "Show " & frmS.Caption
        frmS.mnuPop(3).Visible = True  'View menu show digiclock
        frmS.mnuPop(1).Visible = False 'Hide menu minimize to tray
    End With
End Sub

Sub TraySendPosX(PosX As Single)
'Tray icon actions when mouse click/move on it
    Select Case PosX / Screen.TwipsPerPixelX
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK: frmC.Show: TrayRemoveIcon   'Show clock
        Case WM_RBUTTONDOWN
        Case WM_RBUTTONUP: frmS.PopupMenu frmS.mnuClockPop 'Show PopUp menu
        Case WM_RBUTTONDBLCLK
    End Select
End Sub

Sub TrayRemoveIcon() 'Remove tray icon
    frmS.mnuPop(3).Visible = False      'Hide menu show digiclock
    frmS.mnuPop(1).Visible = True       'View menu minimize to tray
    Shell_NotifyIcon NIM_DELETE, nID    'Delete tray icon
End Sub

Public Function LongFileName$(ByVal ShortName$)
Dim LongName$, P%, s$ 'Convert from short to long filename

    'Start after the drive letter if any.
    If Mid$(ShortName$, 2, 1) = ":" Then
        s$ = Left$(ShortName$, 2): P% = 3
    Else: P% = 1: s$ = "": End If
    'Consider each section in the file name.
    Do While P% > 0
        ' Find the next \.
        P% = InStr(P% + 1, ShortName$, "\")
        ' Get the next piece of the path.
        If P% = 0 Then
            LongName$ = Dir$(ShortName$, vbNormal + _
                vbHidden + vbSystem + vbDirectory)
        Else
            LongName$ = Dir$(Left$(ShortName$, P% - 1), _
                vbNormal + vbHidden + vbSystem + _
                vbDirectory)
        End If
        s$ = s$ & "\" & LongName$
    Loop
    LongFileName = s$
End Function

Sub Long2RGB(ByVal lColor As Long, RetValR&, _
  RetValG&, RetValB&) 'Return RGB values from long...
Const C& = 65536, b& = 256
   RetValB = lColor \ C&
   RetValG = (lColor - (RetValB * C&)) \ b&
   RetValR = lColor - (RetValB * C&) - (RetValG * b&)
End Sub

Function InvertColor(ByVal lColor As Long) As Long
Dim R&, G&, b& 'Invert a color
On Local Error Resume Next
    Long2RGB lColor, R&, G&, b&
    R& = Abs(R& - 255): G& = Abs(G& - 255): b& = Abs(b& - 255)
    InvertColor = RGB(R&, G&, b&)
    If Err <> 0 Then Err.Clear
End Function

Public Function SetCaptionColorDialog(ByVal hwnd As Long, ByVal uMsg As Long, _
  ByVal wParam As Long, ByVal lParam As Long) As Long
'Set caption in select color dialog
   If uMsg = &H110 Then Call SetWindowText(hwnd, sTmpMsg$)
End Function

Public Function SetColor(ByVal Value As CmdColors) As Long
Dim s$, X%, Y%, Z%, b As Boolean
'Create default colors in Color dialog

    If bColorPick Then 'Have picked a color from the clock
        SetColor = lColorpick&: GoTo SetPickColor
    End If
    ReDim DefDialogColors&(16)
    For X% = 0 To frmS.cmd.UBound
        b = True
        For Y% = 0 To Z%
            If DefDialogColors&(Y%) = GetColor(X%) Then b = False: Exit For
        Next
        If b Then
            Z% = Z% + 1: DefDialogColors&(Z%) = GetColor(X%)
        End If
    Next
'Set caption in color dialog and show it
    s$ = App.Title & " - Select " & frmS.cmd(Value).Tag & " color"
    SetColor = ShowColor(frmS.hwnd, CC_FULLOPEN, Clock.CmdColors(Value), s$)
    If SetColor = -1 Then Exit Function 'No color selected
SetPickColor:
    UndoRedo DoSave, DoURcmd, Value, Clock.CmdColors(Value), SetColor  'Save to undo
'Return the selected color and update clock ...
    Clock.CmdColors(Value) = SetColor: frmS.cmd(Value).BackColor = SetColor
'Redraw pictures if change transparent color
    IsDirty = True: If Value = TransparentBack Then LoadClockPictures
    Select Case Value
        Case Border, Numbers, HourSpots, MinuteSpots, Background, TransparentBack, Border2, HourSpots2, MinuteSpots2, Background2, Numbers2
            SetStyle 'Need to redraw clock
        Case Else: CtrlEnabled: bRedraw = True
    End Select
    
End Function

Public Function ShowColor(hwndOwner As Long, Optional nFlags As Long, _
  Optional cInitColor As Long, Optional DlgCaption$) As Long
Dim Custcolor(16) As Long, lReturn As Long, i As Integer, X%, C&
'Show select color dialog
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
'Set usercolors in dialog
    For i = LBound(CustomColors) To UBound(CustomColors) Step 4
        X% = X% + 1 'Split color to RGB
        C& = DefDialogColors&(X%)
        If C < 0 Then C = C * -1
        CustomColors(i) = C& Mod &H100: C& = C& \ &H100
        CustomColors(i + 1) = C& Mod &H100: C& = C& \ &H100
        CustomColors(i + 2) = C& Mod &H100
    Next
    With CC
        .lStructSize = Len(CC)
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .lpCustColors = StrConv(CustomColors, vbUnicode)
        .Flags = nFlags Or CC_ANYCOLOR Or CC_RGBINIT Or CC_ENABLEHOOK
        .rgbResult = cInitColor
        .lpfnHook = FarProc(AddressOf SetCaptionColorDialog)
        sTmpMsg$ = DlgCaption$
        If CHOOSECOLOR(CC) <> 0 Then
            ShowColor = .rgbResult: X% = 0 'Save custom colors
            CustomColors = StrConv(.lpCustColors, vbFromUnicode)
            For i = LBound(CustomColors) To UBound(CustomColors) Step 4
                X% = X% + 1
                DefDialogColors&(X%) = RGB(CustomColors(i), _
                  CustomColors(i + 1), CustomColors(i + 2))
            Next i
        Else: ShowColor = -1: End If
    End With
End Function

Public Function ShowOpen(hwndOwner As Long, sFilter As String, sTitle As String, _
  Optional sIntDir, Optional nFlags As Long = OFN_EXPLORER) As String
Static PrevDir$ 'Open file dialog
    If IsMissing(sIntDir) Then
        sIntDir = IIf(Len(PrevDir$), PrevDir$, App.Path & c_ClcDir)
    End If
    With OFName
        .lStructSize = Len(OFName)
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .lpstrFilter = sFilter
        .lpstrInitialDir = sIntDir
        .lpstrFile = String(254, vbNullChar)
        .nMaxFile = 255
        .lpstrFileTitle = String(254, vbNullChar)
        .nMaxFileTitle = 255
        .lpstrTitle = sTitle
        .Flags = nFlags
        If GetOpenFileName(OFName) Then
            ShowOpen = StripTerminator(.lpstrFile): PrevDir$ = ShowOpen
        End If
    End With
End Function

Public Function ShowSave(hwndOwner As Long, sFilter As String, sTitle As String, _
  Optional sIntDir, Optional nFlags As Long = OFN_EXPLORER Or OFN_OVERWRITEPROMPT, _
  Optional DefName$, Optional Extension$) As String
Static PrevDir$ 'Save file dialog
    With OFName
        If IsMissing(sIntDir) Then
            sIntDir = IIf(Len(PrevDir$), PrevDir$, App.Path & c_ClcDir)
        End If
        .lStructSize = Len(OFName)
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .lpstrFilter = sFilter
        .lpstrInitialDir = sIntDir
        .lpstrFile = String(254, vbNullChar)
        .nMaxFile = 255
        .lpstrFileTitle = String(254, vbNullChar)
        .nMaxFileTitle = 255
        .lpstrTitle = sTitle
        .Flags = nFlags
        .lpstrFile = DefName$ & "." & Extension$ & _
          String(255 - Len(DefName$), vbNullChar)
        If GetSaveFileName(OFName) Then
            ShowSave = StripTerminator(.lpstrFile)
            If LCase(Right$(ShowSave, 4)) <> "." & Extension$ Then
                ShowSave = ShowSave & "." & Extension$
                PrevDir$ = ShowSave
            End If
        End If
    End With
End Function

Function BlendColors&(ByVal C1&, ByVal C2&, ByVal Steps&, RetValColors&())
'Creates an array of colors blending from C1& to C2& in Steps& number of steps.
'Returns the count and fills the RetValColors() array.
Dim lIdx&, R&, G&, b&, Rs!, Gs!, Bs!

'Stop possible error
    If Steps& < 2 Then Steps& = 2
'Extract Red, Blue and Green values from the start and end colors.
    R& = (C1& And &HFF&)
    G& = (C1& And &HFF00&) / &H100
    b& = (C1& And &HFF0000) / &H10000
'Find the amount of change for each color element per color change.
    Rs! = Div(CSng((C2& And &HFF&) - R&), CSng(Steps&))
    Gs! = Div(CSng(((C2& And &HFF00&) / &H100&) - G&), CSng(Steps&))
    Bs! = Div(CSng(((C2& And &HFF0000) / &H10000) - b&), CSng(Steps&))
'Create the colors
    ReDim RetValColors(Steps& - 1)
    RetValColors(0) = C1&            'First Color
    RetValColors(Steps& - 1) = C2&   'Last Color
    For lIdx = 1 To Steps& - 2       'All Colors between
        RetValColors(lIdx) = CLng(R& + (Rs! * CSng(lIdx))) + _
            (CLng(G& + (Gs! * CSng(lIdx))) * &H100&) + _
            (CLng(b& + (Bs! * CSng(lIdx))) * &H10000)
    Next lIdx
'Return number of colors in array
    BlendColors = Steps&
End Function

Private Function Div#(ByVal dN#, ByVal dD#)
'Divides dN# by dD# if dD# <> 0, eliminates 'Division By Zero' error.
    If dD# <> 0 Then
        Div = dN# / dD#
    Else: Div = 0: End If
End Function

Function GetFileTitle$(sFile$) 'Get filename from path
    GetFileTitle$ = String(255, vbNullChar)
    GetFileTitleAPI sFile$, GetFileTitle$, 255
    GetFileTitle$ = StripTerminator(GetFileTitle$)
End Function

Function StripTerminator$(sInput$)
Dim X% 'Remove null chars from string
    X% = InStr(1, sInput, vbNullChar)
    If X% > 0 Then
        StripTerminator = Left$(sInput, X% - 1)
    Else: StripTerminator = sInput: End If
End Function

Function IsArrayEmpty(ByVal Arr As Variant) As Boolean
On Local Error Resume Next  'Check if array is empty
    Arr = Arr(LBound(Arr)): If Err <> 0 Then IsArrayEmpty = True: Err.Clear
End Function

Private Function FarProc(ByVal pfn&): FarProc = pfn: End Function 'Workaround...

'///Font handeling

Private Function GetFontFileName(sFontName$, Optional NotPath As Boolean)
Dim X%, s$, sRet$, R&, lRet& ': Const c_Sz% = 1024
'Get the fonts fil name
    R& = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", _
      0, KEY_ALL_ACCESS, lRet&) ' Open Registry Key
    'Only works with "Windows Me" or higher if lower use: "SOFTWARE\Microsoft\Windows\CurrentVersion\Fonts"
    If R& = 0 Then  'Did find that key
        For X% = 1 To 3
            sRet$ = String$(256, 0) 'Loop around to find the font...
            s$ = sFontName$ & Choose(X%, "", " (TrueType)", " (All res)")
            R& = RegQueryValueEx(lRet&, s$, 0, REG_SZ, sRet$, 256) ' Get the Key Value
            If R& = 0 Then Exit For 'Did find the font name
        Next
        sRet$ = StripTerminator$(sRet$)
        If Len(sRet$) Then 'Have found the fonts filname
            If NotPath Then
                GetFontFileName = sRet$ 'Just return file name without path
            Else
                Dim oShlApp, oFolder, oFItem 'Find the fonts path
                On Local Error Resume Next
                Set oShlApp = CreateObject("Shell.Application")
                Set oFolder = oShlApp.Namespace(&H14)
                If Not oFolder Is Nothing Then
                   Set oFItem = oFolder.Self: s$ = oFItem.Path
                   s$ = s$ & "\" & sRet$
                   If Dir(s$) <> "" And Err = 0 Then GetFontFileName = s$
                End If
                Set oShlApp = Nothing: Set oFolder = Nothing: Set oFItem = Nothing
            End If
        End If
    End If
    If lRet& > 0 Then RegCloseKey lRet& 'Close the key
End Function

Function AddFont&(ByVal sFontFile$) 'To add the font
Const HWND_BROADCAST = &HFFFF&, WM_FONTCHANGE = &H1D
'i.e. AddFont& = AddFontResource("c:\myFont.ttf")
    AddFont& = AddFontResource(sFontFile$)
    If AddFont& > 0 Then 'Alert all windows that a font was added
        SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0&, 0&
    End If
End Function

Function RemoveFont&(ByVal sFontFile$) 'To remove the font
'i.e. RemoveFont& = RemoveFontResource("c:\myFont.ttf")
    RemoveFont& = RemoveFontResource(sFontFile$)
End Function

