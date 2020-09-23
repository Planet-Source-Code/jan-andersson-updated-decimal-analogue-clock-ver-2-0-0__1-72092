VERSION 5.00
Begin VB.Form frmS 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DecimalSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7305
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   10320
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   125
      Top             =   2880
      Width           =   315
      Visible         =   0   'False
      Begin VB.VScrollBar vsc 
         Height          =   375
         Index           =   0
         LargeChange     =   100
         Left            =   0
         Max             =   4000
         Min             =   200
         SmallChange     =   10
         TabIndex        =   159
         Top             =   0
         Value           =   500
         Width           =   135
         Visible         =   0   'False
      End
   End
   Begin VB.TextBox txt 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   10200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   126
      Text            =   "DecimalSettings.frx":08CA
      Top             =   3540
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   10560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   127
      Top             =   3540
      Width           =   330
      Visible         =   0   'False
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   93
      Top             =   4440
      Width           =   675
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   0
      Left            =   660
      Sorted          =   -1  'True
      TabIndex        =   91
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CheckBox chk 
      Caption         =   "Save font to file"
      Height          =   240
      Index           =   32
      Left            =   120
      TabIndex        =   92
      ToolTipText     =   "Will save selected font in current clock file"
      Top             =   4290
      Width           =   1995
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   4
      Left            =   7380
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   151
      Top             =   2580
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   27
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   86
      Tag             =   "numbers 2nd"
      Top             =   3660
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   26
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   78
      Tag             =   "background 2nd"
      Top             =   3420
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   25
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   72
      Tag             =   "minute spots 2nd"
      Top             =   3180
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   24
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   66
      Tag             =   "hour spots 2nd"
      Top             =   2940
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   23
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   61
      Tag             =   "outer border 2nd"
      Top             =   2700
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   22
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   82
      Tag             =   "end center circle"
      Top             =   3420
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   21
      Left            =   2130
      Style           =   1  'Graphical
      TabIndex        =   81
      Tag             =   "start center circle"
      Top             =   3420
      Width           =   195
   End
   Begin VB.CheckBox chk 
      Height          =   240
      Index           =   31
      Left            =   2340
      TabIndex        =   74
      Top             =   3180
      Width           =   195
   End
   Begin VB.CheckBox chk 
      Height          =   240
      Index           =   30
      Left            =   2340
      TabIndex        =   68
      Top             =   2940
      Width           =   195
   End
   Begin VB.CheckBox chk 
      Caption         =   "Audible"
      Height          =   240
      Index           =   17
      Left            =   2760
      TabIndex        =   8
      Top             =   420
      Width           =   855
   End
   Begin VB.CheckBox chk 
      Caption         =   "Hour"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   22
      Left            =   3900
      TabIndex        =   24
      Top             =   600
      Value           =   1  'Checked
      Width           =   795
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   50
      LargeChange     =   1000
      Left            =   6180
      Max             =   25500
      Min             =   100
      SmallChange     =   100
      TabIndex        =   59
      Top             =   2340
      Value           =   25500
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   49
      LargeChange     =   1000
      Left            =   5040
      Max             =   25500
      Min             =   100
      SmallChange     =   100
      TabIndex        =   47
      Top             =   2340
      Value           =   25500
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   48
      LargeChange     =   1000
      Left            =   3900
      Max             =   25500
      Min             =   100
      SmallChange     =   100
      TabIndex        =   35
      Top             =   2340
      Value           =   25500
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   47
      LargeChange     =   1000
      Left            =   6180
      Max             =   25500
      Min             =   -25500
      SmallChange     =   100
      TabIndex        =   56
      Top             =   1800
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   46
      LargeChange     =   100
      Left            =   6180
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   57
      Top             =   1980
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   45
      LargeChange     =   100
      Left            =   6180
      Max             =   3600
      SmallChange     =   10
      TabIndex        =   58
      Top             =   2160
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   44
      LargeChange     =   1000
      Left            =   5040
      Max             =   25500
      Min             =   -25500
      SmallChange     =   100
      TabIndex        =   44
      Top             =   1800
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   43
      LargeChange     =   100
      Left            =   5040
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   45
      Top             =   1980
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   42
      LargeChange     =   100
      Left            =   5040
      Max             =   3600
      SmallChange     =   10
      TabIndex        =   46
      Top             =   2160
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   41
      LargeChange     =   1000
      Left            =   3900
      Max             =   25500
      Min             =   -25500
      SmallChange     =   100
      TabIndex        =   32
      Top             =   1800
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   40
      LargeChange     =   100
      Left            =   3900
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   33
      Top             =   1980
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   39
      LargeChange     =   100
      Left            =   3900
      Max             =   3600
      SmallChange     =   10
      TabIndex        =   34
      Top             =   2160
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   36
      LargeChange     =   100
      Left            =   3900
      Max             =   3600
      SmallChange     =   10
      TabIndex        =   4
      Top             =   240
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   37
      LargeChange     =   100
      Left            =   6180
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   5
      Top             =   240
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   38
      LargeChange     =   1000
      Left            =   3900
      Max             =   25500
      Min             =   -25500
      SmallChange     =   100
      TabIndex        =   2
      Top             =   60
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   35
      LargeChange     =   1000
      Left            =   6180
      Max             =   25500
      Min             =   100
      SmallChange     =   100
      TabIndex        =   3
      Top             =   60
      Value           =   25500
      Width           =   1050
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   51
      Tag             =   "second hand line/border"
      Top             =   840
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "minute hand line/border"
      Top             =   840
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   4740
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "hour hand line/border"
      Top             =   840
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   50
      Tag             =   "second hand 2nd"
      Top             =   840
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "minute hand 2nd"
      Top             =   840
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "hour hand 2nd"
      Top             =   840
      Width           =   195
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   34
      LargeChange     =   100
      Left            =   6180
      Max             =   500
      Min             =   -100
      SmallChange     =   100
      TabIndex        =   55
      Top             =   1620
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   33
      LargeChange     =   100
      Left            =   5040
      Max             =   500
      Min             =   -100
      SmallChange     =   100
      TabIndex        =   43
      Top             =   1620
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   32
      LargeChange     =   100
      Left            =   3900
      Max             =   500
      Min             =   -100
      SmallChange     =   100
      TabIndex        =   31
      Top             =   1620
      Width           =   1050
   End
   Begin VB.CheckBox chk 
      Caption         =   "C-Circle"
      Height          =   240
      Index           =   29
      Left            =   1260
      TabIndex        =   80
      ToolTipText     =   "Center circle above hands"
      Top             =   3420
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CheckBox chk 
      Caption         =   "Alpha mask"
      Height          =   240
      Index           =   28
      Left            =   1560
      TabIndex        =   14
      ToolTipText     =   "Use Alpha Mask on clock face picture"
      Top             =   1140
      Width           =   1155
   End
   Begin VB.CheckBox chk 
      Caption         =   "Alpha mask"
      Height          =   240
      Index           =   27
      Left            =   1560
      TabIndex        =   20
      ToolTipText     =   "Use Alpha Mask on picture hands"
      Top             =   1620
      Width           =   1155
   End
   Begin VB.CheckBox chk 
      Caption         =   "Minutes"
      Height          =   240
      Index           =   26
      Left            =   1620
      TabIndex        =   88
      ToolTipText     =   "Draw minutes instead of hours"
      Top             =   3660
      Width           =   915
   End
   Begin VB.CheckBox chk 
      Caption         =   "Use Anti Alias on hands"
      Height          =   240
      Index           =   25
      Left            =   120
      TabIndex        =   21
      Top             =   1860
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   3
      Left            =   7680
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   156
      Top             =   5820
      Width           =   1035
   End
   Begin VB.CheckBox chk 
      Height          =   240
      Index           =   24
      Left            =   1260
      TabIndex        =   106
      Top             =   5280
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      Height          =   315
      Index           =   14
      Left            =   6840
      Picture         =   "DecimalSettings.frx":147C
      Style           =   1  'Graphical
      TabIndex        =   124
      ToolTipText     =   "Select wave file..."
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Height          =   315
      Index           =   13
      Left            =   6435
      Picture         =   "DecimalSettings.frx":172E
      Style           =   1  'Graphical
      TabIndex        =   123
      ToolTipText     =   "Play alarm sound"
      Top             =   4320
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Caption         =   "Stop"
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   23
      Left            =   2760
      TabIndex        =   11
      Top             =   660
      Width           =   735
   End
   Begin VB.CheckBox chk 
      Caption         =   "Minute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   21
      Left            =   5040
      TabIndex        =   36
      Top             =   600
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "transparent "
      Top             =   1395
      Width           =   195
   End
   Begin VB.CheckBox chk 
      Caption         =   "Minimized to tray"
      Height          =   240
      Index           =   20
      Left            =   120
      TabIndex        =   9
      Top             =   660
      Width           =   1575
   End
   Begin VB.CheckBox chk 
      Caption         =   "Play sound"
      Height          =   240
      Index           =   19
      Left            =   5340
      TabIndex        =   122
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   6
      Left            =   5340
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   152
      Top             =   3900
      Width           =   1815
      Begin VB.OptionButton opt 
         Caption         =   "Show message"
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   119
         Top             =   0
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton opt 
         Caption         =   "Run command string"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   120
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.TextBox txt 
      Height          =   255
      Index           =   1
      Left            =   3420
      TabIndex        =   121
      Top             =   4320
      Width           =   1815
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   3
      Left            =   4140
      TabIndex        =   118
      Top             =   3960
      Width           =   615
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   2
      Left            =   3420
      TabIndex        =   117
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "Alarm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   2580
      TabIndex        =   116
      Top             =   4020
      Width           =   840
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   5
      Left            =   7440
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   147
      Top             =   60
      Width           =   1035
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   2
      Left            =   4860
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   155
      Top             =   5940
      Width           =   1035
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   2460
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   154
      Top             =   5940
      Width           =   735
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   153
      Top             =   5940
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "TopMost"
      Height          =   240
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   660
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   31
      LargeChange     =   10
      Left            =   6180
      Max             =   200
      Min             =   1
      TabIndex        =   54
      Top             =   1440
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   30
      LargeChange     =   10
      Left            =   5040
      Max             =   200
      Min             =   1
      TabIndex        =   42
      Top             =   1440
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   29
      LargeChange     =   10
      Left            =   3900
      Max             =   200
      Min             =   1
      TabIndex        =   30
      Top             =   1440
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   28
      LargeChange     =   10
      Left            =   6180
      Max             =   200
      Min             =   1
      TabIndex        =   53
      Top             =   1260
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   27
      LargeChange     =   10
      Left            =   5040
      Max             =   200
      Min             =   1
      TabIndex        =   41
      Top             =   1260
      Value           =   100
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   26
      LargeChange     =   10
      Left            =   3900
      Max             =   200
      Min             =   1
      TabIndex        =   29
      Top             =   1260
      Value           =   100
      Width           =   1050
   End
   Begin VB.CheckBox chk 
      Caption         =   "Smooth second hand moves"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   2100
      Width           =   2415
   End
   Begin VB.CheckBox chk 
      Caption         =   "Picture hands"
      Height          =   240
      Index           =   16
      Left            =   120
      TabIndex        =   18
      Top             =   1620
      Width           =   1575
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   25
      LargeChange     =   500
      Left            =   5700
      Max             =   3000
      Min             =   100
      SmallChange     =   100
      TabIndex        =   115
      Top             =   5580
      Value           =   600
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   24
      LargeChange     =   500
      Left            =   5700
      Max             =   3000
      Min             =   100
      SmallChange     =   100
      TabIndex        =   109
      Top             =   5340
      Value           =   1000
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   23
      LargeChange     =   500
      Left            =   5700
      Max             =   3000
      Min             =   100
      SmallChange     =   100
      TabIndex        =   103
      Top             =   5100
      Value           =   1000
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   22
      LargeChange     =   500
      Left            =   5700
      Max             =   3000
      Min             =   100
      SmallChange     =   100
      TabIndex        =   98
      Top             =   4860
      Value           =   1100
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   21
      LargeChange     =   10
      Left            =   4140
      Max             =   200
      Min             =   1
      TabIndex        =   114
      Top             =   5580
      Value           =   65
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   20
      LargeChange     =   10
      Left            =   4140
      Max             =   200
      Min             =   1
      TabIndex        =   108
      Top             =   5340
      Value           =   132
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   19
      LargeChange     =   10
      Left            =   4140
      Max             =   200
      Min             =   1
      TabIndex        =   102
      Top             =   5100
      Value           =   115
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   18
      LargeChange     =   10
      Left            =   4140
      Max             =   200
      Min             =   1
      TabIndex        =   97
      Top             =   4860
      Value           =   85
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   17
      LargeChange     =   10
      Left            =   2580
      Max             =   200
      Min             =   1
      TabIndex        =   113
      Top             =   5580
      Value           =   100
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   16
      LargeChange     =   10
      Left            =   2580
      Max             =   200
      Min             =   1
      TabIndex        =   107
      Top             =   5340
      Value           =   100
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   15
      LargeChange     =   10
      Left            =   2580
      Max             =   200
      Min             =   1
      TabIndex        =   101
      Top             =   5100
      Value           =   100
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   14
      LargeChange     =   10
      Left            =   2580
      Max             =   200
      Min             =   1
      TabIndex        =   96
      Top             =   4860
      Value           =   100
      Width           =   1500
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   13
      LargeChange     =   10
      Left            =   5160
      Max             =   100
      Min             =   1
      TabIndex        =   84
      Top             =   3480
      Value           =   97
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   12
      LargeChange     =   1000
      Left            =   5160
      Max             =   10000
      Min             =   100
      SmallChange     =   100
      TabIndex        =   76
      Top             =   3240
      Value           =   100
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   11
      LargeChange     =   1000
      Left            =   5160
      Max             =   10000
      Min             =   100
      SmallChange     =   100
      TabIndex        =   70
      Top             =   3000
      Value           =   2000
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   10
      LargeChange     =   40
      Left            =   5160
      Max             =   400
      Min             =   20
      SmallChange     =   4
      TabIndex        =   90
      Top             =   3720
      Value           =   100
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   9
      LargeChange     =   10
      Left            =   5160
      Max             =   100
      Min             =   1
      TabIndex        =   64
      Top             =   2760
      Value           =   10
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   8
      LargeChange     =   10
      Left            =   2940
      Max             =   100
      Min             =   1
      TabIndex        =   83
      Top             =   3480
      Value           =   1
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   7
      LargeChange     =   10
      Left            =   2940
      Max             =   100
      Min             =   1
      TabIndex        =   75
      Top             =   3240
      Value           =   80
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   6
      LargeChange     =   10
      Left            =   2940
      Max             =   100
      Min             =   20
      TabIndex        =   69
      Top             =   3000
      Value           =   80
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   5
      LargeChange     =   10
      Left            =   2940
      Max             =   100
      Min             =   20
      TabIndex        =   89
      Top             =   3720
      Value           =   60
      Width           =   2050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   4
      LargeChange     =   10
      Left            =   2940
      Max             =   100
      Min             =   1
      TabIndex        =   63
      Top             =   2760
      Value           =   97
      Width           =   2050
   End
   Begin VB.CheckBox chk 
      Caption         =   "Run clock when Windows start"
      Height          =   240
      Index           =   15
      Left            =   120
      TabIndex        =   7
      Top             =   420
      Width           =   2835
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   49
      Tag             =   "second hand 1st"
      Top             =   840
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "minute hand 1st"
      Top             =   840
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "hour hand 1st"
      Top             =   840
      Width           =   195
   End
   Begin VB.TextBox txt 
      Height          =   255
      Index           =   0
      Left            =   1500
      TabIndex        =   112
      Top             =   5460
      Width           =   975
   End
   Begin VB.CheckBox chk 
      Caption         =   "My Text"
      Height          =   240
      Index           =   14
      Left            =   420
      TabIndex        =   111
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   110
      Tag             =   "my text"
      Top             =   5520
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   104
      Tag             =   "date"
      Top             =   5280
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   99
      Tag             =   "weekday"
      Top             =   5040
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   94
      Tag             =   "digital clock"
      Top             =   4800
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   77
      Tag             =   "background 1st"
      Top             =   3420
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   71
      Tag             =   "minute spots/lines 1st"
      Top             =   3180
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   65
      Tag             =   "hour spots/lines 1st"
      Top             =   2940
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   85
      Tag             =   "numbers 1st"
      Top             =   3660
      Width           =   195
   End
   Begin VB.CheckBox chk 
      Caption         =   "Back"
      Height          =   240
      Index           =   13
      Left            =   600
      TabIndex        =   79
      ToolTipText     =   "Filled circle as backcolor"
      Top             =   3420
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.CheckBox chk 
      Caption         =   "Minute Spots/Lines"
      Height          =   240
      Index           =   12
      Left            =   600
      TabIndex        =   73
      Top             =   3180
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Hour Spots/Lines"
      Height          =   240
      Index           =   11
      Left            =   600
      TabIndex        =   67
      Top             =   2940
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Numbers"
      Height          =   240
      Index           =   10
      Left            =   600
      TabIndex        =   87
      Top             =   3660
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox chk 
      Caption         =   "Draw outer border"
      Height          =   240
      Index           =   9
      Left            =   600
      TabIndex        =   62
      Top             =   2700
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   3
      LargeChange     =   10
      Left            =   6180
      Max             =   200
      Min             =   20
      TabIndex        =   6
      Top             =   420
      Value           =   100
      Width           =   1050
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   60
      Tag             =   "outer border 1st"
      Top             =   2700
      Width           =   195
   End
   Begin VB.OptionButton opt 
      Caption         =   "Standard Clock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      Top             =   60
      Width           =   1755
   End
   Begin VB.CheckBox chk 
      Caption         =   "Weekday"
      Height          =   240
      Index           =   8
      Left            =   420
      TabIndex        =   100
      Top             =   5040
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   2
      LargeChange     =   10
      Left            =   6180
      Max             =   200
      TabIndex        =   52
      Top             =   1080
      Value           =   80
      Width           =   1050
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   1
      LargeChange     =   10
      Left            =   5040
      Max             =   200
      TabIndex        =   40
      Top             =   1080
      Value           =   75
      Width           =   1050
   End
   Begin VB.OptionButton opt 
      Caption         =   "Decimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.HScrollBar hsc 
      Height          =   135
      Index           =   0
      LargeChange     =   10
      Left            =   3900
      Max             =   200
      TabIndex        =   28
      Top             =   1080
      Value           =   50
      Width           =   1050
   End
   Begin VB.CheckBox chk 
      Caption         =   "Draw on clock face picture"
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CheckBox chk 
      Caption         =   "Save Settings when exit"
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   900
      Width           =   2175
   End
   Begin VB.CheckBox chk 
      Caption         =   "Picture as face"
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1140
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chk 
      Caption         =   "Day/M"
      Height          =   240
      Index           =   1
      Left            =   420
      TabIndex        =   105
      Top             =   5280
      Value           =   1  'Checked
      Width           =   795
   End
   Begin VB.CheckBox chk 
      Caption         =   " Digital Clock"
      Height          =   240
      Index           =   0
      Left            =   420
      TabIndex        =   95
      Top             =   4800
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chk 
      Caption         =   "Second"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   6180
      TabIndex        =   48
      Top             =   600
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.Label lbl 
      Caption         =   "Translucency"
      Height          =   195
      Index           =   20
      Left            =   5100
      TabIndex        =   142
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   2460
      TabIndex        =   158
      ToolTipText     =   "Color under mousepointer"
      Top             =   1380
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   2160
      TabIndex        =   157
      ToolTipText     =   "Picked color"
      Top             =   1380
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Label lbl 
      Caption         =   "Alpha Blend"
      Height          =   195
      Index           =   27
      Left            =   2940
      TabIndex        =   150
      Top             =   2280
      Width           =   930
   End
   Begin VB.Label lbl 
      Caption         =   "Brightness"
      Height          =   195
      Index           =   22
      Left            =   3060
      TabIndex        =   144
      Top             =   30
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Style"
      Height          =   195
      Index           =   19
      Left            =   2940
      TabIndex        =   141
      Top             =   1560
      Width           =   930
   End
   Begin VB.Label lbl 
      Caption         =   "Brightness"
      Height          =   195
      Index           =   21
      Left            =   2940
      TabIndex        =   143
      Top             =   1740
      Width           =   930
   End
   Begin VB.Label lbl 
      Caption         =   "Hue"
      Height          =   195
      Index           =   26
      Left            =   2940
      TabIndex        =   149
      Top             =   2100
      Width           =   930
   End
   Begin VB.Label lbl 
      Caption         =   "Saturation"
      Height          =   195
      Index           =   25
      Left            =   2940
      TabIndex        =   148
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label lbl 
      Caption         =   "Hue"
      Height          =   195
      Index           =   24
      Left            =   3060
      TabIndex        =   146
      Top             =   195
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Saturation"
      Height          =   195
      Index           =   23
      Left            =   5100
      TabIndex        =   145
      Top             =   195
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   18
      Left            =   1440
      TabIndex        =   140
      Top             =   240
      Width           =   45
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   17
      Left            =   360
      TabIndex        =   139
      Top             =   240
      Width           =   45
   End
   Begin VB.Label lbl 
      Caption         =   "Transparent Backcolor"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   1395
      Width           =   1695
   End
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   405
      Index           =   1
      Left            =   4860
      Picture         =   "DecimalSettings.frx":1800
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   435
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   0
      Left            =   1860
      Stretch         =   -1  'True
      ToolTipText     =   "Current clock face"
      Top             =   4815
      Width           =   600
   End
   Begin VB.Label lbl 
      Caption         =   "Message:"
      Height          =   195
      Index           =   2
      Left            =   2580
      TabIndex        =   19
      Top             =   4320
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   ":"
      Height          =   195
      Index           =   1
      Left            =   4020
      TabIndex        =   17
      Top             =   4020
      Width           =   135
   End
   Begin VB.Label lbl 
      Caption         =   "Y- Position"
      Height          =   195
      Index           =   16
      Left            =   2940
      TabIndex        =   138
      Top             =   1380
      Width           =   990
   End
   Begin VB.Label lbl 
      Caption         =   "Hand Size"
      Height          =   195
      Index           =   15
      Left            =   2940
      TabIndex        =   137
      Top             =   1035
      Width           =   840
   End
   Begin VB.Label lbl 
      Caption         =   "X- Position"
      Height          =   195
      Index           =   14
      Left            =   2940
      TabIndex        =   136
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Y - Position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   4140
      TabIndex        =   135
      Top             =   4650
      Width           =   1500
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "X- Position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   2580
      TabIndex        =   134
      Top             =   4650
      Width           =   1500
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Font Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   5700
      TabIndex        =   133
      Top             =   4650
      Width           =   1500
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   5160
      TabIndex        =   132
      Top             =   2550
      Width           =   2055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Radius position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   2940
      TabIndex        =   131
      Top             =   2550
      Width           =   2050
   End
   Begin VB.Label lbl 
      Caption         =   "Font:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   130
      Top             =   4020
      Width           =   555
   End
   Begin VB.Label lbl 
      Caption         =   "Time Zone Offset:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   129
      Top             =   4530
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Clock Size"
      Height          =   195
      Index           =   3
      Left            =   5100
      TabIndex        =   128
      Top             =   360
      Width           =   855
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New clockfile"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &as..."
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &clock face picture as..."
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save  H&our Hand picture as..."
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save  M&inute Hand picture as..."
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save  S&econd  Hand picture as..."
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Select  cl&ock face picture..."
         Index           =   8
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Remove cloc&k face picture..."
         Index           =   9
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Select  &Hour Hand picture..."
         Index           =   10
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Select  &Minute Hand picture..."
         Index           =   11
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Select  S&econd Hand picture..."
         Index           =   12
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Remove Hand &pictures"
         Index           =   13
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close this form"
         Index           =   14
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Close this form and minimi&ze clock to tray"
         Index           =   15
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   16
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuO 
      Caption         =   "&Options"
      Begin VB.Menu mnuOpt 
         Caption         =   "&Reset all to default settings"
         Index           =   0
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Reset all to original si&ze"
         Index           =   1
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Reset to default &alarm sound"
         Index           =   2
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Center &Hour hand"
         Index           =   3
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Center &Minute hand"
         Index           =   4
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Center &Second hand"
         Index           =   5
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Center Cl&ock on screen"
         Index           =   6
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Rel&oad..."
         Index           =   7
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Re&fresh"
         Index           =   8
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Color p&icker"
         Index           =   9
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Undo"
         Index           =   10
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Redo"
         Index           =   11
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Show comments about this clock"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "S&hort info about this application"
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "The World Clock - Time &Zones"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Vote for this code :-)"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Donate :-)"
         Index           =   4
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Comments and suggestions ..."
         Index           =   5
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   6
      End
   End
   Begin VB.Menu mnuClockPop 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnuPop 
         Caption         =   "Settings"
         Index           =   0
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Minimize to tray"
         Index           =   1
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Exit"
         Index           =   2
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Show Clock"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop 
         Caption         =   "About"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmS"
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

Private Sub Form_Load()
    Call LoadSettings: bColorPick = False: frmC.MousePointer = 15
'Compensate for titlebar & borders
    Height = 5805 + (Height - ScaleHeight): Width = 7350 + (Width - ScaleWidth)
    txt(2).Move 60, 60, ScaleWidth - 120, ScaleHeight - 120         'Comments textbox
    txt(3).Move 60, 60, ScaleWidth - 120, ScaleHeight - 120         'Short info textbox
    pic(7).Move 2580, 2580, 4650, 3140: pic(7).DrawWidth = 4        'Color picker picturebox
    vsc(0).Move pic(7).ScaleWidth - 10, 0, 10, pic(7).ScaleHeight   'Vertical Scrollbar Zoom
    CPX% = pic(7).ScaleWidth / 2: CPY% = pic(7).ScaleHeight / 2     'To use when magnify (Zoom)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TraySendPosX X 'Message to tray icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If txt(2).Visible Then 'Just hide comments textbox
        bNotNow = True: mnuHelp_Click 0
        bNotNow = False: Cancel = True: Exit Sub
    End If
    If txt(3).Visible Then 'Just hide short info textbox
        mnuHelp_Click 1: Cancel = True: Exit Sub
    End If
    If bColorPick Then mnuOpt_Click 9 'Remove colorpick mode
    If Not bNotNow Then 'Hide only
        Cancel = True: Hide: Exit Sub
    Else
        If SaveQ Then Cancel = True: Show: Exit Sub
    End If
End Sub

Sub mnuFile_Click(Index As Integer) 'File menu
Dim X%, s$, s2$: Const CF = "Picture-Files (*.bmp;*.jpg;*.gif)" + vbNullChar + _
  "*.bmp;*.jpg;*.gif" & vbNullChar & "All Files (*.*)" + vbNullChar + "*.*"
    On Error GoTo mnuFile_ClickErr
    s$ = App.Title & "-Files (*.clc)" + vbNullChar + "*.clc"
    If Index > 3 Then s$ = "Bitmap-Files (*.bmp)" + vbNullChar + "*.bmp"
    With Clock
        Select Case Index
            Case 0 'New
                If Not SaveQ Then 'Open default settings file
                    sFile$ = "": Call Main
                End If
            Case 1 'Open
                If Not SaveQ Then 'Check if need to save current file first
                    s$ = ShowOpen(hwnd, s$, "Open Clockfile...")
                    If Len(s$) Then
                        frmC.Left = -20000: frmC.tmr.Enabled = False
                        Me.Hide: DoEvents: PrevMM% = -1
                        sTmp$ = sFile$: sFile$ = s$: bStop = False: Call Main
                        'If sFile$ = "" Then sFile$ = s2$ 'Not a valid Decimal Clock-file
                        frmC.Visible = True: frmC.tmr.Enabled = True: frmS.Visible = True
                    End If
                End If
            Case 2 'Save
                If sFile$ = c_NewFile Then
                    mnuFile_Click 3: Exit Sub 'Save as... if a new file
                Else: Call SaveSettings: End If
            Case 3 'Save as
                If .FileName <> c_NewFile Then
                    s2$ = .FileName
                Else: s2$ = "My Clock": End If
                s$ = ShowSave(hwnd, s$, "Save as...", , , s2$, "clc")
                If Len(s$) Then
                    sFile$ = s$: Call SetCaption: Call SaveSettings
                End If
            Case 4 'Save backpicture as
                s$ = ShowSave(hwnd, s$, "Save clock face picture as...", , , "My clock face", "bmp")
                'If Len(s$) Then SavePicture PicFromByteArray(.PictureBack()), s$
                If Len(s$) Then SavePicture img(0), s$
            Case 5 'Save hourpicture as
                s$ = ShowSave(hwnd, s$, "Save hour hand picture as...", , , "My Hour hand", "bmp")
                If Len(s$) Then SavePicture PicFromByteArray(.PictureHourHand()), s$
            Case 6 'Save minutepicture as
                s$ = ShowSave(hwnd, s$, "Save minute hand picture as...", , , "My Minute hand", "bmp")
                If Len(s$) Then SavePicture PicFromByteArray(.PictureMinuteHand()), s$
            Case 7 'Save secondpicture as
                s$ = ShowSave(hwnd, s$, "Save second hand picture as...", , , "My Second hand", "bmp")
                If Len(s$) Then SavePicture PicFromByteArray(.PictureSecondHand()), s$
            Case 8 'Select Back - Picturefile
                If bUndoRedo Then GoTo UndoClockFace
                s$ = ShowOpen(hwnd, CF, "Select background picture...")
                If Len(s$) Then
                    'Test if valid grapic file
                    Tmp = FileToByteArray(s$): PicFromByteArray Tmp, s$
                    If s$ = "" Then Erase Tmp: Exit Sub
UndoClockFace:
                    bNotNow = True: hsc(SizeFactor) = c_100
                    If Not bUndoRedo Then 'Save to undo
                        UndoRedo DoSave, DoURpic, ClockFace, .PictureBack, Tmp
                    End If
                    'Save to .PictureBack
                    .PictureBack = Tmp: Erase Tmp: PrepairPicture ClockFace
                    chk(UsePictureFile).Value = 1: chk(UsePictureFile).Enabled = True
                    bNotNow = False: Call SetStyle
                End If
            Case 9 'Remove Back - Picturefile
                UndoRedo DoSave, DoURpic, ClockFace, .PictureBack, Tmp 'Save to undo
                bNotNow = True: Erase .PictureBack()
                chk(UsePictureFile).Value = 0: FacePicRadius% = 0
                chk(UsePictureFile).Enabled = False: bNotNow = False: Call SetStyle
            Case 10 'Select Hourhand - Picturefile
                If bUndoRedo Then GoTo UndoHourhand
                s$ = ShowOpen(hwnd, CF, "Select Hourhand picture...")
                If Len(s$) Then
                    'Test if valid grapic file
                    Tmp = FileToByteArray(s$): PicFromByteArray Tmp, s$
                    If s$ = "" Then Erase Tmp: Exit Sub
UndoHourhand:
                    bNotNow = True: hsc(HandHourX) = c_100: hsc(HandHourY) = c_100
                    hsc(HandHourSize) = c_100: hsc(HandHourStyle) = c_100 * -1
                    If chk(PictureHands) = 0 Then chk(PictureHands) = 1
                    If Not bUndoRedo Then 'Save to undo
                        UndoRedo DoSave, DoURpic, DecHour, .PictureHourHand, Tmp
                    End If
                    'Save to .PictureHourHand
                    .PictureHourHand = Tmp: PrepairPicture DecHour: Erase Tmp
                    chk(ShowHourHand).Value = 1: bNotNow = False: Call SetStyle
                End If
            Case 11 'Select Minutehand - Picturefile
                If bUndoRedo Then GoTo UndoMinutehand
                s$ = ShowOpen(hwnd, CF, "Select Minutehand picture...")
                If Len(s$) Then
                    'Test if valid grapic file
                    Tmp = FileToByteArray(s$): PicFromByteArray Tmp, s$
                    If s$ = "" Then Erase Tmp: Exit Sub
UndoMinutehand:
                    bNotNow = True: hsc(HandMinuteX) = c_100: hsc(HandMinuteY) = c_100
                    hsc(HandMinuteSize) = c_100: hsc(HandMinuteStyle) = c_100 * -1
                    If chk(PictureHands) = 0 Then chk(PictureHands) = 1
                    If Not bUndoRedo Then 'Save to undo
                        UndoRedo DoSave, DoURpic, DecMinute, .PictureMinuteHand, Tmp
                    End If
                    'Save to .PictureMinuteHand
                    .PictureMinuteHand = Tmp: PrepairPicture DecMinute: Erase Tmp
                    chk(ShowMinuteHand).Value = 1: bNotNow = False: Call SetStyle
                End If
            Case 12 'Select Secondhand - Picturefile
                If bUndoRedo Then GoTo UndoSecondhand
                s$ = ShowOpen(hwnd, CF, "Select Secondhand picture...")
                If Len(s$) Then
                    'Test if valid grapic file
                    Tmp = FileToByteArray(s$): PicFromByteArray Tmp, s$
                    If s$ = "" Then Erase Tmp: Exit Sub
UndoSecondhand:
                    bNotNow = True: hsc(HandSecondX) = c_100: hsc(HandSecondY) = c_100
                    hsc(HandSecondSize) = c_100: hsc(HandSecondStyle) = c_100 * -1
                    If chk(PictureHands) = 0 Then chk(PictureHands) = 1
                    If Not bUndoRedo Then 'Save to undo
                        UndoRedo DoSave, DoURpic, DecSecond, .PictureSecondHand, Tmp
                    End If
                    'Save to .PictureSecondHand
                    .PictureSecondHand = Tmp: PrepairPicture DecSecond: Erase Tmp
                    chk(ShowSecondHand).Value = 1: bNotNow = False: Call SetStyle
                End If
            Case 13 'Remove Hand - Picturefile
                If MsgBox("Is it okey to remove hand pictures from this file (" & _
                  .FileName & ")?", 36 + vbSystemModal) = vbNo Then Exit Sub
                bNotNow = True: hsc(HandHourStyle) = 0: hsc(HandMinuteStyle) = 0: hsc(HandSecondStyle) = 0
                chk(PictureHands).Value = 0: chk(PictureHands).Enabled = False
                chk(ShowHourHand).Value = 1: chk(ShowMinuteHand).Value = 1: chk(ShowSecondHand).Value = 1
                mnuOpt_Click 3: mnuOpt_Click 4: mnuOpt_Click 5 'Center hands
                'Save to undo
                UndoRedo DoSave, DoURpic, DecHour, .PictureHourHand, Tmp
                UndoRedo DoSave, DoURpic, DecMinute, .PictureMinuteHand, Tmp
                UndoRedo DoSave, DoURpic, DecSecond, .PictureSecondHand, Tmp
                Erase .PictureHourHand(): Erase .PictureMinuteHand(): Erase .PictureSecondHand()
                bNotNow = False: Call SetStyle
            Case 14 'Hide this form
                Me.Hide
            Case 15 'Hide this form and minimize clock to tray
                Me.Hide: TrayMinimizeAppTo
            Case 16 'Exit application
                Unload frmC: Exit Sub
        End Select
    End With
    Me.SetFocus: nDll% = 0
    Exit Sub
mnuFile_ClickErr:
    'Set focus error
    s$ = IIf(Err = 76, vbLf & sFile$, "") 'Path error
    If Err <> 5 Then MsgBox Err.Description & s$, 16 + vbSystemModal
    Err.Clear
End Sub

Private Sub mnuOpt_Click(Index As Integer)
Static NoMsg As Boolean
'Option menu
    If IsSelected(PictureHands) And Index > 2 And Index < 6 Then
    If MsgBox("Can only center hand if it was centered in the original picture, continue?", _
      vbQuestion + vbDefaultButton2 + vbYesNo + vbSystemModal) = vbNo Then Exit Sub
    End If
    Select Case Index
        Case 0 'Reset to default settings in current clock file
            If MsgBox("Is it OK to reset all settings to default in file '" & Clock.FileName & _
              "'?", vbQuestion + vbDefaultButton2 + vbYesNo + vbSystemModal) = vbYes Then ResetAll
        Case 1 'Reset pictures to orginal size
            bNotNow = True: hsc(SizeFactor) = c_100: hsc(HandHourSize) = c_100
            hsc(HandMinuteSize) = c_100: hsc(HandSecondSize) = c_100
            bNotNow = False: LoadClockPictures: SetStyle: bRedraw = True
        Case 2 'Reset alarm sound to default from resourcefile
            Clock.AlarmWavFile = "": PlayWav: MsgBox "Done!", 64 + vbSystemModal: CtrlEnabled
        Case 3 'Center Hour hand
            hsc(HandHourX) = c_100: hsc(HandHourY) = c_100
        Case 4 'Center Minute hand
            hsc(HandMinuteX) = c_100: hsc(HandMinuteY) = c_100
        Case 5 'Center Second hand
            hsc(HandSecondX) = c_100: hsc(HandSecondY) = c_100
        Case 6 'Center clock on screen
            frmC.Top = Screen.Height / 2 - (CR% * Screen.TwipsPerPixelY)
            frmC.Left = Screen.Width / 2 - (CR% * Screen.TwipsPerPixelX)
        Case 7 'Reload clockfile
            ResetAll True
        Case 8 'Refresh clock
            Call LoadClockPictures: SetStyle ': bUndoRedo = False: UndoRedo
        Case 9 'Color Picker mode
            bColorPick = Not mnuOpt(8).Checked: mnuOpt(8).Checked = bColorPick
            lbl(11).Visible = bColorPick: lbl(12).Visible = bColorPick
            pic(7).Visible = bColorPick: vsc(0).Visible = bColorPick
            frmC.MousePointer = IIf(bColorPick, 2, 15)
            If bColorPick Then vsc_Change 0 'Draw a centered picture of the clock
            If bColorPick And Not NoMsg Then
                MsgBox "1.) Move the mouse on the clock." & vbLf & _
                "2.) Click on the clock to select a color." & vbLf & _
                "3.) Click on any colored button to change it's color." & vbLf & vbLf & _
                "You can adjust the magnification by moving the vertical scrollbar." & vbLf & _
                "Press the F6-Key to exit color picker mode.", _
                64 + vbSystemModal: NoMsg = True
            End If
        Case 10 'Undo
            UndoRedo DoUndo
        Case 11 'Redo
            UndoRedo DoRedo
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
Const Tnx = "Thanks! :-)"
    Select Case Index 'Help menu
        Case 0 'Show edit comments about this clock
            If txt(3).Visible Then mnuHelp_Click 1 'Hide short info
            mnuHelp(0).Checked = Not mnuHelp(0).Checked
            txt(2).Visible = mnuHelp(0).Checked
            If Clock.Comments = "" And Not bNotNow And mnuHelp(0).Checked Then _
              MsgBox "Here you can write your own comments about your clock.", 64 + vbSystemModal
        Case 1: 'Show short info
            If txt(2).Visible Then bNotNow = True: mnuHelp_Click 0 'Hide comments
            mnuHelp(1).Checked = Not mnuHelp(1).Checked
            txt(3).Visible = mnuHelp(1).Checked: bNotNow = False
        Case 2 'World -Clock
            RunCommand "http://www.timeanddate.com/worldclock/full.html"
        Case 3 'Vote for my code :-)
            MsgBox Tnx, 64 + vbSystemModal: RunCommand c_URL
        Case 4 'Make a donation if you feel 4 it :-)
            MsgBox Tnx, 64 + vbSystemModal: RunCommand "https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=pappsegull%40yahoo%2ese&item_name=Pappsegull%27s%20Decimal%20Clock.%20Thank%20you%20very%20much%20for%20your%20donation%3a%2d%29&no_shipping=0&no_note=1&tax=0&currency_code=USD&lc=US&bn=PP%2dDonationsBF&charset=UTF%2d8"
        Case 5 'About form
            frmAbout.Contact c_Mail
        Case 6 'About form
            frmAbout.Show , Me
    End Select
End Sub

Private Sub mnuPop_Click(Index As Integer)
'Show popup menu if right click clock form or icon tray
    Select Case Index
        Case 0: frmS.WindowState = 0: frmS.Show 'Show Settings Form
        Case 1: TrayMinimizeAppTo               'Minimize to tray
        Case 2
            If bColorPick Then _
              mnuOpt_Click 9: Exit Sub          'Remove colorpick mode
            Unload frmC                         'Exit Application
        Case 3: frmC.Show                       'Show the clock if minimized to tray
            Call TrayRemoveIcon                 'Remove tray icon and change popup menu
        Case 4: frmAbout.Show , frmC            'Show About Form
    End Select
End Sub

Private Sub chk_Click(Index As Integer)
Dim s$, X% 'Click on checkboxes
    If Not bUndoRedo Then 'Save to undo
        UndoRedo DoSave, DoURchk, Index, Clock.ChkValue(Index), chk(Index)
    End If
    Clock.ChkValue(Index) = chk(Index)
    If bNotNow Then Exit Sub
    'Check if have a compiled exe..
    If App.LogMode = 0 And Index = RunWinStartUp And chk(RunWinStartUp) Then
        s$ = App.Path & "\" & App.Title & ".exe"
        If Dir(s$) = "" Then
            MsgBox "You can't use '" & chk(RunWinStartUp).Caption & _
              "', if you don't have a compiled exe." & vbLf & s$, 64 + vbSystemModal
            bNotNow = True: chk(RunWinStartUp) = 0: bNotNow = False: Exit Sub
        End If
    End If
    bStop = IIf(IsSelected(StopClock), True, False)
    If Index = PictureHands Then 'Picture hands or not, change style on hands
        X% = IIf(chk(PictureHands), -1, 0) * 100: bNotNow = True: hsc(HandHourStyle) = X%
        hsc(HandMinuteStyle) = X%: hsc(HandSecondStyle) = X%: bNotNow = False
    End If
    Select Case Index
        Case ShowHourHand, ShowMinuteHand, ShowSecondHand
            If sTmpUndo$ <> "" Then 'To redraw clock if second hand not visible
                 bRedraw = True: PrevMM% = -1: CtrlEnabled
            End If
        Case StopClock, AlphaFace, DrawOnPicture, UsePictureFile, DrawBackground, DrawBorder, DrawHourLines, DrawHourSpots, DrawMinuteLines, DrawMinuteSpots, DrawNumbers, DrawNumbersMinutes
             Call SetStyle: bRedraw = True 'Redraw clock face
        Case Topmost 'Make the clock TopMost
            If frmC.Visible Then frmC.SetFocus
            SetTopMost frmC.hwnd: frmS.SetFocus
            Call TrayRemoveIcon 'Remove tray icon and change popup menu
    End Select
    IsDirty = True: CtrlEnabled: bRedraw = True
End Sub

Private Sub hsc_Change(Index As Integer)
'Click on horizontal scrollbars
    If Not bUndoRedo Then  'Save to Undo
        UndoRedo DoSave, DoURhsc, Index, Clock.HscValue(Index) * 100, hsc(Index)
    End If
    Clock.HscValue(Index) = hsc(Index) / 100 'Save value to Type "Clock"
    If bNotNow Then Exit Sub
    Select Case Index
    'Resize Clock
        Case ClockFace: LoadClockPictures: SetStyle
    'Redraw hands if changing style, size, Hue, Saturation, Brightness, Alfa, X or Y positions
      'Hour
        Case HandHourStyle, HandHourSize, HandHourX, HandHourY, HHue, HSaturation, HBrightness, HAlfa
            bRedraw = True: If GetHscValue(HandHourStyle) = -1 Then _
              PrepairPicture DecHour Else BuildPolygon DecHour
      'Minute
        Case HandMinuteStyle, HandMinuteSize, HandMinuteX, HandMinuteY, MHue, MSaturation, MBrightness, MAlfa
            bRedraw = True: If GetHscValue(HandMinuteStyle) = -1 Then _
              PrepairPicture DecMinute Else BuildPolygon DecMinute
      'Second
        Case HandSecondStyle, HandSecondSize, HandSecondX, HandSecondY, SHue, SSaturation, SBrightness, SAlfa
            bRedraw = True: If GetHscValue(HandSecondStyle) = -1 Then _
              PrepairPicture DecSecond Else BuildPolygon DecSecond
      'Clocktext no need redraw clock
        Case DigiClockSize, DigiClockX, DigiClockY, WeekdaySize, WeekdayX, WeekdayY, DateSize, DateX, DateY, MyTextSize, MyTextX, MyTextY
            bRedraw = True
    'Changing Translucency
        Case Translucent: SetTranslucent
        Case Else: SetStyle
    End Select
    IsDirty = True: CtrlEnabled: bRedraw = True
End Sub

Private Sub vsc_Change(Index As Integer)
'Change vertical scrollbars (Magnify Colorpicker)
    PrevX! = True: PrevY! = True: frmC.Magnify vsc(0)
End Sub

Sub opt_Click(Index As Integer)
Dim b As Boolean 'Click on option buttons
    If bNotNow Then Exit Sub
    If Not bUndoRedo Then UndoRedo DoSave, DoURopt, _
      Index, Choose(Index + 1, 1, 0, 3, 2), Index 'Save to undo
'Option show message or comand if Alarm selected
    b = IIf(opt(3), True, False)
    lbl(2) = IIf(b, "Message:", "Command:")
    Clock.AlarmShowMsgBox = b
'Option Decimal or standard clock
    Clock.DecimalClock = IIf(opt(0), True, False)
    'Change values to select in alarm time combos
    If bDeci <> Clock.DecimalClock Then FillAlarmCombos True
    bDeci = Clock.DecimalClock: IsDirty = True
    PrevMM% = -1: bRedraw = True: CtrlEnabled 'To redraw clock if second hand not visible
    If Index < 2 Then SetStyle                'Changing clock Decimal/Standard
End Sub

Private Sub cmb_Click(Index As Integer)
'Click on comboboxes
    IsDirty = True
    With Clock
        Select Case Index
            Case 0 'Font
                If Not bUndoRedo Then _
                  UndoRedo DoSave, DoURcmb, Index, .FontName, cmb(0).Text            'Save to undo
                .FontName = cmb(0).Text: SetStyle
            Case 1 'Timezone
                If Not bUndoRedo Then _
                  UndoRedo DoSave, DoURcmb, Index, .TimeZoneOffset, Val(cmb(1).Text) 'Save to undo
                .TimeZoneOffset = Val(cmb(1).Text): PrevMM = -1: CtrlEnabled: bRedraw = True
            Case 2 'Alarm hour
                If Not bUndoRedo Then _
                  UndoRedo DoSave, DoURcmb, Index, .AlarmHour, Val(cmb(2).Text)      'Save to undo
                .AlarmHour = Val(cmb(2).Text): CtrlEnabled
            Case 3 'Alarm minute
                If Not bUndoRedo Then _
                  UndoRedo DoSave, DoURcmb, Index, .AlarmMinute, Val(cmb(3).Text)    'Save to undo
                .AlarmMinute = Val(cmb(3).Text): CtrlEnabled
        End Select
    End With
End Sub

Private Sub cmd_Click(Index As Integer)
Dim s$
'Click on command buttons
    Select Case Index
        Case SoundPlay: PlayWav     'Play current alarm wave sound file
        Case SoundSelect            'Select a wave alarm sound file and store it in Clock.AlarmWavFile()
            s$ = "Wave sound-Files (*.wav)" + vbNullChar + "*.wav"
            s$ = ShowOpen(hwnd, s$, "Select wave file...", Clock.AlarmWavFile)
            If Len(s$) Then _
              Clock.AlarmWavFile = FileToByteArray(s$): PlayWav: IsDirty = True: CtrlEnabled
        Case Else: SetColor Index   'Select a color button
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
'Textboxes change, have skip undo for the textboxes...
    If bNotNow Then Exit Sub
    With Clock
        .MyText = txt(0)
        .AlarmCommand = txt(1)
        .Comments = txt(2)
    End With
    IsDirty = True: CtrlEnabled
End Sub

