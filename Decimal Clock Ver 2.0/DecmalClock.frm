VERSION 5.00
Begin VB.Form frmC 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   990
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   1560
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   15  'Size All
   ScaleHeight     =   66
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   60
   End
End
Attribute VB_Name = "frmC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub tmr_Timer()
'Timer returns a Single representing the number of seconds elapsed since midnight
    Tm! = Timer: DoEvents: Tm! = Tm! + ((c_Sec24h& / 24) * Clock.TimeZoneOffset)
    If Tm! < 0 Then Tm! = Tm! + c_Sec24h&
    If bStop Then Tm! = 0
    If CheckTime Or bRedraw Then DrawHands
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bColorPick Then
        PrevX! = X: PrevY! = Y: Magnify 'Redraw when in colorpick mode
    Else
        If Button = 1 Then 'Move Clock
            Left = Left + (X - PrevX!): Top = Top + (Y - PrevY!)
        End If
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PrevX! = X: PrevY! = Y 'Save previous X & Y
    If Button = 2 Then PopupMenu frmS.mnuClockPop
    If Button = 1 And bColorPick Then 'Save the picked color
        lColorpick& = GetPixel(hDC, X, Y)
        frmS.lbl(11).BackColor = lColorpick&
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next 'Just to keep focus on the setting form
    If frmS.Visible Then frmS.SetFocus
    If Err <> 0 Then Err.Clear
    If Button = 1 And Shift Then Unload Me 'Exit if hold Shift-Key down and left click clock
End Sub

Sub Magnify(Optional ByVal Sz!) 'Magnify when using colorpicker
Dim C&, C2&, SW&, SH&: Static SzS!: Const cSz = 50
    With frmS.pic(7)
        If Sz! = 0 Then Sz! = SzS! Else SzS! = Sz!
        Sz! = Sz! / 100: SW& = ScaleWidth: SH& = ScaleHeight
        If PrevX! = True And PrevY! = True Then
            PrevX! = SW& / 2: PrevY! = SH& / 2
        End If
        PrevX! = PrevX! - (CPX% / Sz!): PrevY! = PrevY! - (CPY% / Sz!)
    'Stretch blit to .pic(7)
        SetStretchBltMode .hDC, COLORONCOLOR: .Cls
        StretchBlt .hDC, 0&, 0&, SW& * Sz!, SH& * Sz!, _
          hDC, PrevX!, PrevY!, SW&, SH&, vbSrcCopy
    'Draw center cross
        C2& = GetPixel(.hDC, CPX%, CPY%)
        frmS.lbl(12).BackColor = C2&: C& = InvertColor(C2&)
        Gs.CircleGP .hDC, CPX%, CPY%, cSz - (cSz / 10), cSz - (cSz / 10), C&
        Gs.LineGP .hDC, CPX% - cSz, CPY%, CPX% + cSz, CPY%, C&
        Gs.LineGP .hDC, CPX%, CPY% - cSz, CPX%, CPY% + cSz, C&
        frmS.pic(7).PSet (CPX%, CPY%), C2&: .Refresh
    End With
End Sub

Private Sub Form_DblClick()
    TrayMinimizeAppTo 'Minimize to tray
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Question if like to save before End...
    If IsSelected(SettingsSave) Then 'Auto save is selected
        frmS.mnuFile_Click 2         'Save settings, even not Isdirty so save clock's position
    Else: SaveQ: End If
    If bCancelSave And IsDirty Then
        Cancel = True: Exit Sub
    End If
'Clean up and End Decimal Clock!
    Set Gs = Nothing
    Erase Rotate!(): Erase PrevTime%(): Erase ByteSound(): Erase UndoFile.Undos()
    Erase ptHour(): Erase ptMinute(): Erase ptSecond(): Erase UndoRedos()
    Erase ptNewHour(): Erase ptNewMinute(): Erase ptNewSecond()
    Set frmS = Nothing: Set frmC = Nothing: Set frmAbout = Nothing: End
End Sub
