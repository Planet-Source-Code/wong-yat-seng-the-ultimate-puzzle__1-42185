VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmKnight 
   Caption         =   "A Horrible Knight's Puzzle - By Yat Seng"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6255
   Icon            =   "frmKnight.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmOneSecond 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin MediaPlayerCtl.MediaPlayer WMP 
      Height          =   735
      Left            =   240
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   765
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -790
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Remain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   67
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblRemain 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "64 Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   66
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Elasped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   65
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblGameTime 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 Sec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   64
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   18
      Left            =   5280
      TabIndex        =   63
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   4560
      TabIndex        =   62
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   3840
      TabIndex        =   61
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   3120
      TabIndex        =   60
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   2400
      TabIndex        =   59
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   1680
      TabIndex        =   58
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   960
      TabIndex        =   57
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   240
      TabIndex        =   56
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   28
      Left            =   5280
      TabIndex        =   55
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   27
      Left            =   4560
      TabIndex        =   54
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   26
      Left            =   3840
      TabIndex        =   53
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   25
      Left            =   3120
      TabIndex        =   52
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   24
      Left            =   2400
      TabIndex        =   51
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   23
      Left            =   1680
      TabIndex        =   50
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   22
      Left            =   960
      TabIndex        =   49
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   21
      Left            =   240
      TabIndex        =   48
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   38
      Left            =   5280
      TabIndex        =   47
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   37
      Left            =   4560
      TabIndex        =   46
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   36
      Left            =   3840
      TabIndex        =   45
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   35
      Left            =   3120
      TabIndex        =   44
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   34
      Left            =   2400
      TabIndex        =   43
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   33
      Left            =   1680
      TabIndex        =   42
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   32
      Left            =   960
      TabIndex        =   41
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   31
      Left            =   240
      TabIndex        =   40
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   48
      Left            =   5280
      TabIndex        =   39
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   47
      Left            =   4560
      TabIndex        =   38
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   46
      Left            =   3840
      TabIndex        =   37
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   45
      Left            =   3120
      TabIndex        =   36
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   44
      Left            =   2400
      TabIndex        =   35
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   43
      Left            =   1680
      TabIndex        =   34
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   42
      Left            =   960
      TabIndex        =   33
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   41
      Left            =   240
      TabIndex        =   32
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   58
      Left            =   5280
      TabIndex        =   31
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   57
      Left            =   4560
      TabIndex        =   30
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   56
      Left            =   3840
      TabIndex        =   29
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   55
      Left            =   3120
      TabIndex        =   28
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   54
      Left            =   2400
      TabIndex        =   27
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   53
      Left            =   1680
      TabIndex        =   26
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   52
      Left            =   960
      TabIndex        =   25
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   51
      Left            =   240
      TabIndex        =   24
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   68
      Left            =   5280
      TabIndex        =   23
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   67
      Left            =   4560
      TabIndex        =   22
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   66
      Left            =   3840
      TabIndex        =   21
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   65
      Left            =   3120
      TabIndex        =   20
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   64
      Left            =   2400
      TabIndex        =   19
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   63
      Left            =   1680
      TabIndex        =   18
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   62
      Left            =   960
      TabIndex        =   17
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   61
      Left            =   240
      TabIndex        =   16
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   78
      Left            =   5280
      TabIndex        =   15
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   77
      Left            =   4560
      TabIndex        =   14
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   76
      Left            =   3840
      TabIndex        =   13
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   75
      Left            =   3120
      TabIndex        =   12
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   74
      Left            =   2400
      TabIndex        =   11
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   73
      Left            =   1680
      TabIndex        =   10
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   72
      Left            =   960
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   71
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   88
      Left            =   5280
      TabIndex        =   7
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   87
      Left            =   4560
      TabIndex        =   6
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   86
      Left            =   3840
      TabIndex        =   5
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   85
      Left            =   3120
      TabIndex        =   4
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   84
      Left            =   2400
      TabIndex        =   3
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   83
      Left            =   1680
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   82
      Left            =   960
      TabIndex        =   1
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   81
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   735
   End
   Begin VB.Menu mnGame 
      Caption         =   "&Game"
      Begin VB.Menu mnNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnScore 
         Caption         =   "&High Scores"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnExit 
         Caption         =   "E&xit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnMove 
      Caption         =   "&Move"
      Begin VB.Menu mnShow 
         Caption         =   "&Show Moves"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnUndo 
         Caption         =   "&Undo Last Move"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnInstruct 
         Caption         =   "&Instructions"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmKnight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Elasped As Long             'Total Elasped time
Dim Remain As Integer           'Remaining Boxes
Dim Box_Taken(88) As Boolean    'To determine whether the box (Index) is occupied
Dim Box_Avail(88) As Boolean    'To indicate boxes is within range of current position
Dim Moved(64) As Integer        'Used to save the Positions made by the user
Dim GameMode As Boolean         'Indicate whether game is in progress
Dim Hint As Boolean             'Show moves/Dont Show

Private Sub Box_Click(Index As Integer)
Dim i As Integer
Dim z As Integer
Dim Box_Index As Integer
    If GameMode = True Then     'If game is started...
        If Box_Avail(Index) = True Then
            
            For i = 1 To 88         'Reset all available boxes
                Box_Avail(i) = False
            Next i
            
            For i = 1 To 8
                For z = 1 To 8     'Reset Box Colors
                    Box_Index = Val(Trim(Str(i) & Str(z)))
                    If Box(Box_Index).Caption = "" Then Box(Box_Index).BackColor = &HC0FFFF
                Next z
            Next i
                
            WMP.Play
            mnUndo.Enabled = True   'Allow Undo
            
            '----- Mark the clicked Box -----
            
            Remain = Remain - 1             'Reduce remaining boxes
            lblRemain.Caption = Trim(Str(Remain)) & " Boxes"
            Box(Index).Caption = "K"        'Mark 'K' on the clicked box
            Box(Index).BackColor = &H8080FF 'Paint to Light Red
            Box_Taken(Index) = True         'Box is Filled
            Moved(64 - Remain) = Index      'Stores Previous Position
            
            '----- Calculates the next available boxes ------
            
            Dim Row As Integer
            Dim Col As Integer

            Row = Left(Index, 1)
            Col = Right(Index, 1)
            
            'Performs Knight's Move calculation
            'Up Left
            If Row - 2 > 0 And Col - 1 > 0 Then
                If Box_Taken(Val(Trim(Str(Row - 2) & Str(Col - 1)))) = False Then
                    Box_Avail(Val(Trim(Str(Row - 2) & Str(Col - 1)))) = True
                    If Hint = True Then Box(Val(Trim(Str(Row - 2) & Str(Col - 1)))).BackColor = &H80FF80
                End If
            End If
            'Up Right
            If Row - 2 > 0 And Col + 1 < 9 Then
                If Box_Taken(Val(Trim(Str(Row - 2) & Str(Col + 1)))) = False Then
                    Box_Avail(Val(Trim(Str(Row - 2) & Str(Col + 1)))) = True
                    If Hint = True Then Box(Val(Trim(Str(Row - 2) & Str(Col + 1)))).BackColor = &H80FF80
                End If
            End If
            'Down Left
            If Row + 2 < 9 And Col - 1 > 0 Then
                If Box_Taken(Val(Trim(Str(Row + 2) & Str(Col - 1)))) = False Then
                    Box_Avail(Val(Trim(Str(Row + 2) & Str(Col - 1)))) = True
                    If Hint = True Then Box(Val(Trim(Str(Row + 2) & Str(Col - 1)))).BackColor = &H80FF80
                End If
            End If
            'Down Right
            If Row + 2 < 9 And Col + 1 < 9 Then
                If Box_Taken(Val(Trim(Str(Row + 2) & Str(Col + 1)))) = False Then
                    Box_Avail(Val(Trim(Str(Row + 2) & Str(Col + 1)))) = True
                    If Hint = True Then Box(Val(Trim(Str(Row + 2) & Str(Col + 1)))).BackColor = &H80FF80
                End If
            End If
            'Left Up
            If Row - 1 > 0 And Col - 2 > 0 Then
                If Box_Taken(Val(Trim(Str(Row - 1) & Str(Col - 2)))) = False Then
                    Box_Avail(Val(Trim(Str(Row - 1) & Str(Col - 2)))) = True
                    If Hint = True Then Box(Val(Trim(Str(Row - 1) & Str(Col - 2)))).BackColor = &H80FF80
                End If
            End If
            'Left Down
            If Row + 1 < 9 And Col - 2 > 0 Then
                If Box_Taken(Val(Trim(Str(Row + 1) & Str(Col - 2)))) = False Then
                    Box_Avail(Val(Trim(Str(Row + 1) & Str(Col - 2)))) = True
                    If Hint = True Then Box(Val(Trim(Str(Row + 1) & Str(Col - 2)))).BackColor = &H80FF80
                End If
            End If
            'Right Up
            If Row - 1 > 0 And Col + 2 < 9 Then
                If Box_Taken(Val(Trim(Str(Row - 1) & Str(Col + 2)))) = False Then
                    Box_Avail(Val(Trim(Str(Row - 1) & Str(Col + 2)))) = True
                    If Hint = True Then Box(Val(Trim(Str(Row - 1) & Str(Col + 2)))).BackColor = &H80FF80
                End If
            End If
            'Right Down
            If Row + 1 < 9 And Col + 2 < 9 Then
                If Box_Taken(Val(Trim(Str(Row + 1) & Str(Col + 2)))) = False Then
                    Box_Avail(Val(Trim(Str(Row + 1) & Str(Col + 2)))) = True
                    If Hint = True Then Box(Val(Trim(Str(Row + 1) & Str(Col + 2)))).BackColor = &H80FF80
                End If
            End If
            
            '----- Check Whether there's any remaining Moves -----
            
            Dim GameOver As Boolean
            Dim WinGame As Boolean
            GameOver = True
            WinGame = True
            
            For i = 11 To 88     'Check for availability of Moves
                If Box_Avail(i) = True Then GameOver = False
            Next i
                     
            For i = 1 To 8
                For z = 1 To 8  'If there's still empty boxes
                    If Box(Val(Trim(Str(i) & Str(z)))).Caption = "" Then WinGame = False
                Next z
            Next i
            
            If WinGame = True Then  'If no more empty boxes
                MsgBox "Congratulations!" & vbCrLf & "You've COMPLETED the Puzzle!", vbInformation, "Completed"
                GameMode = False
                mnUndo.Enabled = False
                Call CheckScore
                Exit Sub
            End If
            
            If GameOver = True Then 'If no more available moves
                MsgBox "You've no more Moves.", vbInformation, "GAME OVER"
                GameMode = False
                mnUndo.Enabled = False
                Call CheckScore
                Exit Sub
            End If
            
        End If
    End If
End Sub

Private Sub Box_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)   'Short cut to UNDO
    If Button = 2 And mnUndo.Enabled = True Then Call mnUndo_Click
End Sub

Private Sub Form_Load()
    GameMode = False    'Game Not Started
    Hint = True         'Show moves
    WMP.Open (App.Path & "\move.wav")
End Sub

Private Sub mnExit_Click()  'Quit Game
    Set frmKnight = Nothing
    Unload Me
End Sub

Private Sub mnInstruct_Click()      'Help Instructions
    MsgBox "1." & vbTab & "Start by clicking on any empty boxes." & vbCrLf & "2." & vbTab & "Click on another box as you would move" & vbCrLf & vbTab & "a Knight in a game of Chess." & vbCrLf & "3." & vbTab & "Continue to fill up the boxes until no more" & vbCrLf & vbTab & "moves are available for you." & vbCrLf & vbCrLf & "Note: Right Click on any box to Undo Last Move.", vbInformation, "Knight's Puzzle"
End Sub

Private Sub mnNew_Click()
Dim i As Integer
Dim z As Integer
Dim Box_Index As Integer
   
    mnUndo.Enabled = False  'Disable Undo
    
    Elasped = 0             'Reset all values
    Remain = 64             'Remain boxes return to 64
    GameMode = True         'Game gets started
    
    For i = 1 To 88
        Box_Avail(i) = True     'All boxes are available
        Box_Taken(i) = False    'All boxes are empty
    Next i
        
    For i = 1 To 64         'Reset all moves
        Moved(i) = 0
    Next i
    
    For i = 1 To 8
        For z = 1 To 8     'Reset aLL values
            Box_Index = Val(Trim(Str(i) & Str(z)))
            Box(Box_Index).Caption = ""     'Empty all Captions
            Box(Box_Index).BackColor = &H80FF80
        Next z
    Next i
       
    MsgBox "Click on a Box to Start", vbOKOnly
End Sub

Private Sub CheckScore()        'Check Highscore
Dim i As Integer
Dim z As Integer
Dim Name As String
Dim Col_1(10) As String
Dim Col_2(10) As Integer
Dim Col_3(10) As Integer
    
    Open (App.Path & "\ScoreNA.dat") For Random As #1
    Open (App.Path & "\ScoreRE.dat") For Random As #2
    Open (App.Path & "\ScoreEL.dat") For Random As #3
    
        For i = 1 To 10
            Get #1, i, Col_1(i)     'Get Name from file
            Get #2, i, Col_2(i)     'Get Remain from file
            Get #3, i, Col_3(i)     'Get Seconds from file
        Next i
        For i = 1 To 10             'If current score is better than any highscore ...
            If Remain < Int(Col_2(i)) Or Remain = Int(Col_2(i)) And Elasped < Int(Col_3(i)) Then
                For z = 10 To i Step -1     'Shifts all HighScore below current score 1 level down
                    Put #1, z, Col_1(z - 1)
                    Put #2, z, Col_2(z - 1)
                    Put #3, z, Col_3(z - 1)
                Next z
                    Name = InputBox("Please Enter your Name:", "HIGH SCORE")
                    Put #1, i, Name         'Insert Current score
                    Put #2, i, Remain
                    Put #3, i, Elasped
                Exit For
            End If
        Next i
    Close #1            'Close all file after use (Always a good practice)
    Close #2
    Close #3
    
    '----- If puzzled is completed ----
    If Remain <= 0 Then
        Dim FreeNum
        FreeNum = FreeFile      'Write the steps taken into a file
        Open (App.Path & "\Solved.txt") For Output As FreeNum
            For i = 1 To 64
                Write #FreeNum, i, Moved(i)
            Next
        Close #1                'REMEMBER to send me this file
    End If
    
    Call mnScore_Click      'Display highscore
    
End Sub

Private Sub mnScore_Click()     'Display a msgbox of high score
Dim i As Integer
Dim Name As String
Dim Col_1(10) As String
Dim Col_2(10) As Integer
Dim Col_3(10) As Integer
Dim Display As String

'Display table Headers
Display = "Rank" & vbTab & "Name" & vbTab & vbTab & "Remain" & vbTab & "Time" & vbCrLf & vbCrLf

    'Open Highscore files
    Open (App.Path & "\ScoreRE.dat") For Random As #1
    Open (App.Path & "\ScoreEL.dat") For Random As #2
    Open (App.Path & "\ScoreNA.dat") For Random As #3
        
        For i = 1 To 10     'Extract High Scores
            
            Get #1, i, Col_2(i)
            Get #2, i, Col_3(i)
            Get #3, i, Col_1(i)
        
        '---FOR RESET SCORES ONLY---
        'If you wish to reset the Highscores, unComment the
        '4 lines below and Comment the 3 Get# statements above.
        'Run the game, view the Hiscore. Come back here and
        'reverse the process.
        '---------------------------
        
            'Put #1, i, 64
            'Put #2, i, 0
            'Name = "The Fool"
            'Put #3, i, Name
        
        Next i
    Close #1
    Close #2
    Close #3

    For i = 1 To 10     'Concatenate the Msgbox
        Display = Display & Trim(Str(i)) & vbTab & Col_1(i) & vbTab & vbTab & Col_2(i) & vbTab & Col_3(i) & vbCrLf
    Next i
                        'Show the Table
    MsgBox Display, vbOKOnly, "Top 10 HighScore"

End Sub

Private Sub mnShow_Click()      'Show Hints
Dim i As Integer
    If mnShow.Checked = True Then
        mnShow.Checked = False
        Hint = False    'Dont Show
        For i = 11 To 88
            If Box_Avail(i) = True Then Box(i).BackColor = &HC0FFFF
        Next i
    Else:
        mnShow.Checked = True
        Hint = True     'Show
        For i = 11 To 88
            If Box_Avail(i) = True Then Box(i).BackColor = &H80FF80
        Next i
    End If
End Sub

Private Sub mnUndo_Click()          'UNDO Moves
    If Remain < 63 Then
    
        'Return the last clicked box to unused state
        Box_Avail(Moved(64 - Remain)) = True
        Box(Moved(64 - Remain)).Caption = ""
        Box_Taken(Moved(64 - Remain)) = False
            
        'Make a call to the 2nd last box
        Remain = Remain + 1
        Box_Avail(Moved(64 - Remain)) = True
        Call Box_Click(Moved(64 - Remain))
    
        'Increase back the remain amount as the call
        'to the box click itself will reduce remain by 1.
        Remain = Remain + 1
        lblRemain.Caption = Trim(Str(Remain)) & " Boxes"
        
        'Cannot undo more than first step
        If Remain >= 63 Then mnUndo.Enabled = False
    End If
End Sub

Private Sub tmOneSecond_Timer()    'Seconds Timer
    If GameMode = True Then Elasped = Elasped + 1
    lblGameTime.Caption = Trim(Str(Elasped)) & " Sec"
End Sub

