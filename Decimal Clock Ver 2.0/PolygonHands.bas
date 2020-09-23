Attribute VB_Name = "modPolygonHands"
Option Explicit

' =========================================================================================
' Decimal Clock Version 2.0.0
' Code written by Pappsegull Sweden, pappsegull@yahoo.se
' Copyright © 2009 - Pappsegull - All rights reserved.
'
' Thanks to Dana Seaman I have used and modify the code from then code project:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=38582&lngWId=1
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



Type POINTAPI
   X As Long
   Y As Long
End Type
Private Type TRIVERTEX
   X     As Long
   Y     As Long
   Red   As Integer
   Green As Integer
   Blue  As Integer
   Alpha As Integer
End Type
Private Type RGB
   Red   As Integer
   Green As Integer
   Blue  As Integer
End Type
Private Type GradientTRIANGLE
   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long
End Type

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

'Style on hands
Enum Handstyles
    NotSelected = 0
    Arrow3DTwisted = 100
    ArrowTwistedFilled = 1001
    ArrowTwistedTransparent = 10
    CompassFilled = 2001
    CompassTransparent = 20
    NeedleFilled = 3001
    NeedleTransparent = 30
    RectFilled = 4001
    RectTransparent = 40
    RectX2Filled = 5001
    RectX2Transparent = 50
    ArrowFilled = 6001
    ArrowTransparent = 60
End Enum
Global Const c_NoOffPolyStyles = 13
Dim HandStyle As Handstyles
Global Rotate!(3599, 1) 'Sin/Cos look-up table
Global PrevTime%(DecHour To DecSecond)
Dim HandGradientEnd As OLE_COLOR, HandGradientStart As OLE_COLOR
'Polygons point to 3 O'clock
Global ptHour() As POINTAPI, ptMinute() As POINTAPI, ptSecond() As POINTAPI
'Rotated polygons
Global ptNewHour() As POINTAPI, ptNewMinute() As POINTAPI, ptNewSecond() As POINTAPI
'For Gradients
Dim RGB1 As RGB, RGB2 As RGB

Sub DrawPolygonHand(Hand As Hands)
Dim A!, C&, b As Boolean
    SetHandStyle Hand
    If HandStyle < 1 Then Exit Sub
    If bRedraw Then BuildPolygon Hand
    If HandStyle > 0 Then
        C& = Choose(Hand + 1, HH%, MM%, SS%)
        A! = Format(Choose(Hand + 1, AngH!, AngM!, AngS!), "##0.0")
        'If smooth mode selected on second hand
        If IsSelected(ShowSmooth) And Hand = DecSecond Then b = True
        If Hand = DecMinute Then 'Adjust minute hand
            A! = MM% * IIf(bDeci, 3.6, 6)
        End If
        If bStop Then A! = 0: C& = 0
        'Rotate hand only when change is detected
        If C& <> PrevTime%(Hand) Or bRedraw Or b Then
            PrevTime%(Hand) = C&
            Select Case Hand
                Case DecHour: RotatePoints ptHour, ptNewHour, A!, Hand
                Case DecMinute: RotatePoints ptMinute, ptNewMinute, A!, Hand
                Case DecSecond: RotatePoints ptSecond, ptNewSecond, A!, Hand
            End Select
        End If
        'Set border & gradient colors
        C& = GetColor(Choose(Hand + 1, HourHand, MinuteHand, SecondHand))
        HandGradientStart = C&: RGB1 = GetRGB(C&)
        C& = GetColor(Choose(Hand + 1, HourHand2, MinuteHand2, SecondHand2))
        HandGradientEnd = C&: RGB2 = GetRGB(C&)
        C& = GetColor(Choose(Hand + 1, HourHandBorder, MinuteHandBorder, SecondHandBorder))
        'Draw the polygon
        Select Case Hand
            Case DecHour: DrawPolygon ptNewHour, C&, HandGradientStart
            Case DecMinute: DrawPolygon ptNewMinute, C&, HandGradientStart
            Case DecSecond: DrawPolygon ptNewSecond, C&, HandGradientStart
        End Select
    End If
End Sub

Sub BuildPolygon(ByVal Hand As Hands)
Dim n%, PT() As POINTAPI, Sz!, T!, X!, Y!, Z!

    If HandStyle = 0 Then 'Check what hand to build if any selected...
        SetHandStyle Hand
    End If
    If HandStyle = 0 Then Exit Sub
    Select Case HandStyle
        Case Arrow3DTwisted, ArrowTwistedFilled, ArrowTwistedTransparent: n% = 18
        Case ArrowFilled, ArrowTransparent: n% = 7
        Case RectX2Filled, RectX2Transparent: n% = 8
        Case Else: n% = 4
    End Select
    Select Case Hand
        Case DecHour: ReDim ptHour(n%): ReDim ptNewHour(n%): Sz! = GetHscValue(HandHourSize)
        Case DecMinute: ReDim ptMinute(n%): ReDim ptNewMinute(n%): Sz! = GetHscValue(HandMinuteSize)
        Case DecSecond: ReDim ptSecond(n%): ReDim ptNewSecond(n%): Sz! = GetHscValue(HandSecondSize)
    End Select
    ReDim PT(n%)

    Select Case HandStyle
    'Twisted arrow
        Case Arrow3DTwisted, ArrowTwistedFilled, ArrowTwistedTransparent
            For n% = 0 To 9
                'Define upper half of hand
                PT(n).X = CR% * Choose(n + 1, -0.15, -0.1, _
                  -0.05, -0.05, 0.05, 0.05, 0.425, 0.8, 0.8, 0.95)
                PT(n).Y = CR% * Choose(n + 1, 0, 0.025, _
                  0.025, 0.01, 0.01, 0.025, 0.05, 0.025, 0.07, 0)
                Select Case Hand
                    Case DecHour
                        ptHour(n).X = PT(n).X * Sz!: ptHour(n).Y = PT(n).Y * Sz!
                    Case DecMinute
                        ptMinute(n).X = PT(n).X * Sz!: ptMinute(n).Y = PT(n).Y * Sz!
                    Case DecSecond
                        ptSecond(n).X = PT(n).X * Sz!: ptSecond(n).Y = PT(n).Y * Sz!
                End Select
            Next
            'Adjust hands to get differet styles between them,5,7,8
            Select Case Hand
                Case DecHour   'Hour hand
                    ptHour(6).X = CR% * 0.2 * Sz!: ptHour(7).X = CR% * 0.45 * Sz!
                    ptHour(8).X = CR% * 0.45 * Sz!: ptHour(8).Y = CR% * 0.09 * Sz!
                    ptHour(9).X = CR% * 0.6 * Sz!
                Case DecSecond 'Second hand
                    ptSecond(6).X = CR% * 0.2 * Sz!: ptSecond(7).X = CR% * 0.6 * Sz!
                    ptSecond(8).X = CR% * 0.6 * Sz!: ptSecond(8).Y = CR% * 0.06 * Sz!
                    ptSecond(9).X = CR% * 0.95 * Sz!
            End Select
            MirrorVerticals Hand, 10, 18, 18: Erase PT
    'Arrow
        Case ArrowFilled, ArrowTransparent
            Select Case Hand
                Case DecHour: BuildArrowHand ptHour(), Hand, Sz!
                Case DecMinute: BuildArrowHand ptMinute(), Hand, Sz!
                Case DecSecond: BuildArrowHand ptSecond(), Hand, Sz!
            End Select
    'Rectangular hand style
        Case RectFilled, RectTransparent
            Select Case Hand
                Case DecHour: BuildRectHand ptHour(), Hand, Sz!
                Case DecMinute: BuildRectHand ptMinute(), Hand, Sz!
                Case DecSecond: BuildRectHand ptSecond(), Hand, Sz!
            End Select
    'Rectangular X2 hand style
        Case RectX2Filled, RectX2Transparent
            Select Case Hand
                Case DecHour: Build2xRectHand ptHour(), Hand, Sz!
                Case DecMinute: Build2xRectHand ptMinute(), Hand, Sz!
                Case DecSecond: Build2xRectHand ptSecond(), Hand, Sz!
            End Select
    'Needle Triangle style
        Case NeedleFilled, NeedleTransparent
            Select Case Hand '(0).X = Base from center, (1).Y = Width, (2).X = Length
                Case DecHour
                    ptHour(0).X = -CR% * 0.01 * Sz!: ptHour(1).Y = CR% * 0.04 * Sz!: ptHour(2).X = CR% * 0.5 * Sz!
                Case DecMinute
                    ptMinute(0).X = -CR% * 0.01 * Sz!: ptMinute(1).Y = CR% * 0.02 * Sz!: ptMinute(2).X = CR% * 0.7 * Sz!
                Case DecSecond
                    ptSecond(0).X = -CR% * 0.01 * Sz!: ptSecond(1).Y = CR% * 0.01 * Sz!: ptSecond(2).X = CR% * 0.8 * Sz!
            End Select
            MirrorVerticals Hand, 3, 4, 4 'From, To, Index
    'Compass style
        Case CompassFilled, CompassTransparent
            Select Case Hand '(0).X = Base from center, (1).Y = Width, (2).X = Length
                Case DecHour
                    ptHour(0).X = -CR% * 0.2 * Sz!: ptHour(1).Y = CR% * 0.075 * Sz!: ptHour(2).X = CR% * 0.5 * Sz!
                Case DecMinute
                    ptMinute(0).X = -CR% * 0.2 * Sz!: ptMinute(1).Y = CR% * 0.05 * Sz!: ptMinute(2).X = CR% * 0.7 * Sz!
                Case DecSecond
                    ptSecond(0).X = -CR% * 0.1 * Sz!: ptSecond(1).Y = CR% * 0.04 * Sz!: ptSecond(2).X = CR% * 0.8 * Sz!
            End Select
            MirrorVerticals Hand, 3, 4, 4 'From, To, Index
    End Select
End Sub

Private Sub DrawPolygon(ptNew() As POINTAPI, OutlineColor As Long, FillColor As Long)
Dim i%, n%, hDC&, C&, C2&

    With frmC
        hDC& = .hDC
        If HandStyle = Arrow3DTwisted Then
            .FillStyle = 1
            'Fill all the triangles with Gradients, create 3D twisted effect
            DrawTG hDC&, ptNew(), 0, 1, 17: DrawTG hDC, ptNew(), 16, 1, 17
            DrawTG hDC&, ptNew(), 1, 2, 16: DrawTG hDC, ptNew(), 15, 3, 14
            DrawTG hDC&, ptNew(), 4, 3, 14: DrawTG hDC, ptNew(), 5, 6, 12
            DrawTG hDC&, ptNew(), 12, 5, 13: DrawTG hDC, ptNew(), 6, 7, 11
            DrawTG hDC&, ptNew(), 11, 6, 12: DrawTG hDC, ptNew(), 9, 8, 10
            If IsSelected(AntiAlias) Then  'Only if we are anti-aliasing
                'This line fixes glitch where the Gradients meet, it occurs only at some angles.
                Gs.LineGP hDC&, ptNew(7).X, ptNew(7).Y, ptNew(11).X, ptNew(11).Y, HandGradientStart
                'These lines accentuate the twist effect as well as hiding the jaggies where gradients meet .
                Gs.LineGP hDC&, ptNew(1).X, ptNew(1).Y, ptNew(16).X, ptNew(16).Y, HandGradientStart
                Gs.LineGP hDC&, ptNew(5).X, ptNew(5).Y, ptNew(12).X, ptNew(12).Y, HandGradientStart
                Gs.LineGP hDC&, ptNew(6).X, ptNew(6).Y, ptNew(11).X, ptNew(11).Y, HandGradientStart
            End If
        Else 'Draw API polygon
            C& = .FillColor: C2& = .ForeColor: .ForeColor = OutlineColor
            'If HandStyle = NeedleFilled Or HandStyle = ArrowTwistedFilled Or HandStyle = CompassFilled Then
            If HandStyle >= 1000 Then
                 If HandStyle <> ArrowTwistedFilled Then   'Gradient fill
                    .FillStyle = 1
                    If HandStyle = RectX2Filled Then
                        DrawTG hDC&, ptNew(), 2, 5, 4: DrawTG hDC, ptNew(), 2, 3, 4
                        DrawTG hDC&, ptNew(), 1, 7, 0: DrawTG hDC, ptNew(), 1, 6, 7
                    ElseIf HandStyle = ArrowFilled Then
                        DrawTG hDC&, ptNew(), 6, 5, 0: DrawTG hDC, ptNew(), 1, 5, 0
                        DrawTG hDC&, ptNew(), 2, 4, 3
                    Else
                        DrawTG hDC&, ptNew(), 2, 3, 1: DrawTG hDC, ptNew(), 0, 3, 1
                    End If
                 Else                               'Fill polygon
                    .FillColor = FillColor: .FillStyle = 0
                    Polygon .hDC, ptNew(0), UBound(ptNew)
                End If
            ElseIf IsSelected(AntiAlias) = False Then 'Transparent
                .FillStyle = 1: Polygon hDC&, ptNew(0), UBound(ptNew)
            End If
        End If
        'Anti-alias outline, overwrites API polygon.
        If IsSelected(AntiAlias) Then
            n% = UBound(ptNew) - 1
            For i% = 0 To n%
                Gs.LineGP hDC&, ptNew(i%).X, ptNew(i%).Y, _
                  ptNew(i% + 1).X, ptNew(i% + 1).Y, OutlineColor
            Next
        End If
        .FillColor = C&: .ForeColor = C2&
   End With
End Sub

Private Sub RotatePoints(Points() As POINTAPI, NewPoints() As POINTAPI, _
  ByVal Angle As Single, Hand As Hands)
Dim i%, j%, L!, DiffX!, DiffY!
    
    On Local Error Resume Next: j% = UBound(Points)
    If Err <> 0 Then 'Polygon not yet built
        For i% = 0 To 2: BuildPolygon (i%): Next
        j% = UBound(Points): Err.Clear
    End If
'Set values depending on hand
    Select Case Hand
        Case DecHour: DiffX! = HandHourX: DiffY! = HandHourY
        Case DecMinute: DiffX! = HandMinuteX: DiffY! = HandMinuteY
        Case DecSecond: DiffX! = HandSecondX: DiffY! = HandSecondY
    End Select
    DiffX! = GetHscValue(DiffX!): DiffY! = GetHscValue(DiffY!)
    CX% = CR% * DiffX!: CY% = CR% * DiffY!
    Angle = Angle * 10: Angle = Angle Mod 3600
   'Use Sin/Cos lookup table Rotate() for speed
   For i% = 0 To j%
      NewPoints(i%).X = Points(i%).X * Rotate(Angle, 0) + _
        Points(i%).Y * Rotate(Angle, 1) + CR% * DiffX!
      NewPoints(i%).Y = -Points(i%).X * Rotate(Angle, 1) + _
        Points(i%).Y * Rotate(Angle, 0) + CR% * DiffY!
   Next
End Sub

Private Sub MirrorVerticals(ByVal Hand As Hands, ByVal Start As Integer, ByVal Finish As Integer, ByVal Idx As Integer)
   Dim n As Integer
    For n = Start To Finish
        Select Case Hand
            Case DecHour: ptHour(n).X = ptHour(Idx - n).X: ptHour(n).Y = -ptHour(Idx - n).Y
            Case DecMinute: ptMinute(n).X = ptMinute(Idx - n).X: ptMinute(n).Y = -ptMinute(Idx - n).Y
            Case DecSecond: ptSecond(n).X = ptSecond(Idx - n).X: ptSecond(n).Y = -ptSecond(Idx - n).Y
        End Select
   Next
End Sub

Private Function GetRGB(lColour As Long) As RGB
Dim HexColour As String 'Split long color to RGB pices
   TranslateColor lColour, 0, lColour
   HexColour = String(6 - Len(Hex$(lColour)), "0") & Hex$(lColour)
   GetRGB.Red = "&H" & Mid$(HexColour, 5, 2) & "00"
   GetRGB.Green = "&H" & Mid$(HexColour, 3, 2) & "00"
   GetRGB.Blue = "&H" & Mid$(HexColour, 1, 2) & "00"
End Function

Private Function DrawTG(hDC As Long, ptNew() As POINTAPI, A As Integer, b As Integer, C As Integer) As Long
Dim Triangle As GradientTRIANGLE, V(2) As TRIVERTEX
'Draw triangle gradient
On Local Error Resume Next
   V(0).X = ptNew(A).X: V(0).Y = ptNew(A).Y
   V(0).Red = RGB2.Red: V(0).Green = RGB2.Green: V(0).Blue = RGB2.Blue
   
   V(1).X = ptNew(b).X: V(1).Y = ptNew(b).Y
   V(1).Red = RGB1.Red: V(1).Green = RGB1.Green: V(1).Blue = RGB1.Blue
   
   
   V(2).X = ptNew(C).X: V(2).Y = ptNew(C).Y
   V(2).Red = RGB1.Red: V(2).Green = RGB1.Green: V(2).Blue = RGB1.Blue
   
   
   Triangle.Vertex1 = 0: Triangle.Vertex2 = 1: Triangle.Vertex3 = 2
   DrawTG = GradientFillRect(hDC, V(0), 3, Triangle, 1, &H2)
    If Err <> 0 Then Err.Clear
End Function

Sub InitPolygon()
Dim n%, R!, A!: R! = Pi / 180
    For n% = 0 To 3599 'Sin/Cos look-up array(singles) for polygon
        A! = n% / 10: Rotate(n%, 0) = Sin(A! * R!): Rotate(n%, 1) = Cos(A! * R!)
    Next
End Sub

Sub BuildRectHand(PT() As POINTAPI, Hand As Hands, Size!)
Dim i%, T!, X!, Y!, Z! 'X! = Y Start pos. | Y! = % of rad. Height | Z! = Width
    Select Case Hand
        Case DecHour: X! = -0.1: Y! = 0.5: Z! = 0.03
        Case DecMinute: X! = -0.1: Y! = 0.7: Z! = 0.02
        Case DecSecond: X! = -0.1: Y! = 0.8: Z! = 0.01
    End Select
    T! = CR% * Size!: X! = X! * T!: Y! = Y! * T!: Z! = Z! * T!
    For i% = 0 To 4
        PT(i%).X = Choose(i% + 1, Y!, Y!, X!, X!, Y!)
        PT(i%).Y = Choose(i% + 1, -Z!, Z!, Z!, -Z!, -Z!)
    Next
End Sub

Sub Build2xRectHand(PT() As POINTAPI, Hand As Hands, Size!)
Dim i%, T!, X!, Y!, Y2!, Z!, Z2!
'X! = Y Start pos | Y! = % of rad. Height | Z! = Width Thick | Y2! = Y-Pos. Where Thin | Z2! = Width Thin
    Select Case Hand
        Case DecHour: X! = -0.3: Y! = 0.5: Z! = 0.045: Z2! = 0.015: Y2! = -0.1
        Case DecMinute: X! = -0.35: Y! = 0.7: Z! = 0.035: Z2! = 0.01: Y2! = -0.1
        Case DecSecond: X! = -0.4: Y! = 0.8: Z! = 0.02: Z2! = 0.005: Y2! = -0.1
    End Select
    T! = CR% * Size!: X! = X! * T!: Y! = Y! * T!: Z! = Z! * T!: Y2! = Y2! * T!: Z2! = Z2! * T!
    For i% = 0 To 8
        PT(i%).X = Choose(i% + 1, X!, Y2!, Y2!, Y!, Y!, Y2!, Y2!, X!, X!)
        PT(i%).Y = Choose(i% + 1, -Z!, -Z!, -Z2!, -Z2!, Z2!, Z2!, Z!, Z!, -Z!)
    Next
End Sub

Sub BuildArrowHand(PT() As POINTAPI, Hand As Hands, Size!)
Dim i%, T!, X!, Y!, Y2!, Z!, Z2!
'X! = Y Start | Y! = % of rad. Height | Z! = Width Thick | Y2! = Height Where Thin | Z2! = Width Thin
    Select Case Hand
        Case DecHour: X! = -0.15: Y! = 0.5: Z! = 0.05: Z2! = 0.02: Y2! = 0.35
        Case DecMinute: X! = -0.15: Y! = 0.7: Z! = 0.04: Z2! = 0.02: Y2! = 0.55
        Case DecSecond: X! = -0.15: Y! = 0.8: Z! = 0.02: Z2! = 0.01: Y2! = 0.65
    End Select
    T! = CR% * Size!: X! = X! * T!: Y! = Y! * T!: Z! = Z! * T!: Y2! = Y2! * T!: Z2! = Z2! * T!
    For i% = 0 To 7
        PT(i%).X = Choose(i% + 1, X!, Y2!, Y2!, Y!, Y2!, Y2!, X!, X!)
        PT(i%).Y = Choose(i% + 1, -Z2!, -Z2!, -Z!, 0, Z!, Z2!, Z2!, -Z2!)
    Next
End Sub

Sub SetHandStyle(Hand As Hands)
Dim X% 'Select hand style depending on scrollbar value
    X% = GetHscValue(Choose(Hand + 1, HandHourStyle, _
      HandMinuteStyle, HandSecondStyle))
    Select Case X%
        Case 0: X% = NotSelected
        Case 1: X% = Arrow3DTwisted
        Case 2: X% = ArrowTwistedFilled
        Case 3: X% = ArrowTwistedTransparent
        Case 4: X% = ArrowFilled
        Case 5: X% = ArrowTransparent
        Case 6: X% = CompassFilled
        Case 7: X% = CompassTransparent
        Case 8: X% = NeedleFilled
        Case 9: X% = NeedleTransparent
        Case 10: X% = RectFilled
        Case 11: X% = RectTransparent
        Case 12: X% = RectX2Filled
        Case 13: X% = RectX2Transparent
    End Select
    HandStyle = X%
End Sub


