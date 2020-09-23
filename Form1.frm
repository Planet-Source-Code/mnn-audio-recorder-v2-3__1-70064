VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Audio Recorder"
   ClientHeight    =   5265
   ClientLeft      =   9060
   ClientTop       =   645
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7695
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   48
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton btnUnhook 
      Caption         =   "Unhook"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   46
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btnHook 
      Caption         =   "Hook"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   45
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picLinear 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   840
      ScaleHeight     =   465
      ScaleWidth      =   6375
      TabIndex        =   39
      Top             =   3360
      Width           =   6375
      Begin VB.Label lblLinearScale 
         AutoSize        =   -1  'True
         Caption         =   "0.5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2950
         TabIndex        =   44
         Top             =   300
         Width           =   240
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   12
         X1              =   3082
         X2              =   3082
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   12
         X1              =   3067
         X2              =   3067
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   11
         X1              =   15
         X2              =   15
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   11
         X1              =   0
         X2              =   0
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   2
         X1              =   1548
         X2              =   1548
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   2
         X1              =   1533
         X2              =   1533
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   4616
         X2              =   4616
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   1
         X1              =   4601
         X2              =   4601
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   0
         X1              =   6240
         X2              =   6240
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   0
         X1              =   6225
         X2              =   6225
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Label lblLinearScale 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6195
         TabIndex        =   43
         Top             =   300
         Width           =   90
      End
      Begin VB.Label lblLinearScale 
         AutoSize        =   -1  'True
         Caption         =   "0.75"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   42
         Top             =   300
         Width           =   330
      End
      Begin VB.Label lblLinearScale 
         AutoSize        =   -1  'True
         Caption         =   "0.25"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1395
         TabIndex        =   41
         Top             =   300
         Width           =   330
      End
      Begin VB.Label lblLinearScale 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   40
         Top             =   300
         Width           =   90
      End
   End
   Begin VB.CheckBox chkMonitoring 
      Caption         =   "Peak monitoring"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   37
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CheckBox chkdB 
      Caption         =   "dB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   25
      Top             =   2805
      Width           =   495
   End
   Begin VB.PictureBox picPeak 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   840
      ScaleHeight     =   690
      ScaleWidth      =   6135
      TabIndex        =   17
      Top             =   2640
      Width           =   6135
      Begin VB.PictureBox picPeakRight 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3105
         TabIndex        =   19
         Tag             =   "3870"
         Top             =   360
         Width           =   3105
      End
      Begin VB.PictureBox picPeakLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3105
         TabIndex        =   18
         Tag             =   "3870"
         Top             =   0
         Width           =   3105
      End
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1
      Left            =   2880
      Top             =   0
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton btnStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton btnStart 
         Cancel          =   -1  'True
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton btnPause 
         Caption         =   "Pause"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblStateus 
         AutoSize        =   -1  'True
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2850
         TabIndex        =   16
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblTimeRecorded 
         AutoSize        =   -1  'True
         Caption         =   "Time recorded:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblBytesWritten 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 bytes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   13
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblBytes 
         AutoSize        =   -1  'True
         Caption         =   "Bytes written:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lblState 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "not recording"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   2280
         TabIndex        =   11
         Top             =   330
         Width           =   1260
      End
   End
   Begin VB.Frame fraRecSets 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton btnMP3Settings 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cmbOutMode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   2280
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   9
         Top             =   1920
         Width           =   230
      End
      Begin VB.TextBox txtFile 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Text            =   "c:\rec_%time%.wav"
         Top             =   1920
         Width           =   2415
      End
      Begin VB.ComboBox cmbBits 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0019
         Left            =   2280
         List            =   "Form1.frx":0023
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cmbFrequency 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":002E
         Left            =   2280
         List            =   "Form1.frx":0041
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbChannels 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0067
         Left            =   2280
         List            =   "Form1.frx":0071
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblOutputType 
         AutoSize        =   -1  'True
         Caption         =   "Output type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   1360
         Width           =   945
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         Caption         =   "File:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1965
         Width           =   300
      End
      Begin VB.Label lblBits 
         AutoSize        =   -1  'True
         Caption         =   "Bits per sec.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label lblFrequency 
         AutoSize        =   -1  'True
         Caption         =   "Frequency:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   645
         Width           =   825
      End
      Begin VB.Label lblChannels 
         AutoSize        =   -1  'True
         Caption         =   "Channels:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.TextBox txtDebug 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3960
      Width           =   7695
   End
   Begin VB.PictureBox picDecibels 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   840
      ScaleHeight     =   465
      ScaleWidth      =   6375
      TabIndex        =   26
      Top             =   3360
      Width           =   6375
      Begin VB.Label lblDBScale 
         AutoSize        =   -1  'True
         Caption         =   "-90"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   30
         TabIndex        =   34
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblDBScale 
         AutoSize        =   -1  'True
         Caption         =   "-70"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1245
         TabIndex        =   33
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblDBScale 
         AutoSize        =   -1  'True
         Caption         =   "-50"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2635
         TabIndex        =   32
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblDBScale 
         AutoSize        =   -1  'True
         Caption         =   "-30"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   4000
         TabIndex        =   31
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblDBScale 
         AutoSize        =   -1  'True
         Caption         =   "-20"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4680
         TabIndex        =   30
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblDBScale 
         AutoSize        =   -1  'True
         Caption         =   "-10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5350
         TabIndex        =   29
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblDBScale 
         AutoSize        =   -1  'True
         Caption         =   "-5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5720
         TabIndex        =   28
         Top             =   300
         Width           =   150
      End
      Begin VB.Label lblDBScale 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6200
         TabIndex        =   27
         Top             =   300
         Width           =   90
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   10
         X1              =   6225
         X2              =   6225
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   10
         X1              =   6240
         X2              =   6240
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   9
         X1              =   5794
         X2              =   5794
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   9
         X1              =   5810
         X2              =   5810
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   8
         X1              =   5453
         X2              =   5453
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   8
         X1              =   5468
         X2              =   5468
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   7
         X1              =   4771
         X2              =   4771
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   7
         X1              =   4786
         X2              =   4786
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   6
         X1              =   4090
         X2              =   4090
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   6
         X1              =   4105
         X2              =   4105
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   5
         X1              =   2726
         X2              =   2726
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   5
         X1              =   2740
         X2              =   2740
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   4
         X1              =   1363
         X2              =   1363
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   4
         X1              =   1378
         X2              =   1378
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linGray 
         BorderColor     =   &H8000000C&
         Index           =   3
         X1              =   20
         X2              =   20
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H80000005&
         Index           =   3
         X1              =   35
         X2              =   35
         Y1              =   240
         Y2              =   0
      End
   End
   Begin VB.Label lblLeft 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblRight 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Dim bDecibels As Boolean
Dim b_Recording As Boolean
Dim b_Monitoring As Boolean

Private Sub btnHelp_Click()
    MsgBox "Variables:" & vbCrLf & "======" & vbCrLf & _
        "%date%" & vbTab & "- current date in system format (e.g. 12/30/2006)" & vbCrLf & _
        "%time%" & vbTab & "- current time in system format (e.g. 2:30 PM)", vbInformation, Caption
End Sub

Private Sub btnHook_Click()
    Hook True
End Sub

Private Sub btnMP3Settings_Click()
    frmMP3Settings.Show vbModal, Me
End Sub

Private Sub btnPause_Click()
    b_Recording = Not b_Recording
    
    If b_Recording Then
        With lblState
            .Caption = "RECORDING"
            .ForeColor = &HFF&
            .FontBold = True
        End With
    Else
        With lblState
            .Caption = "PAUSED"
            .ForeColor = &H8000&
            .FontBold = True
        End With
    End If
    
    modWaveIn.Pause
End Sub

Private Sub btnReset_Click()
    modWaveIn.ResetMonitoring
    chkMonitoring.Value = vbUnchecked
    SetPeakmeter 0, 0
End Sub

Private Sub btnStart_Click()
    Dim File As String
    Dim channels As Integer
    
    EnableRecording
    
    txtDebug.Text = ""
    
    File = ParseFile(txtFile.Text)
    
    If cmbChannels.ListIndex = 0 Then
        channels = 2
    ElseIf cmbChannels.ListIndex = 1 Then
        channels = 1
    End If
    
    If cmbOutMode.ListIndex = 0 Then
        File = Replace(File, ".mp3", ".wav")
    Else
        File = Replace(File, ".wav", ".mp3")
    End If
    
    modWaveIn.PrepareRecording channels, _
                               CInt(cmbBits.Text), _
                               CLng(cmbFrequency.Text), _
                               cmbOutMode.ListIndex, _
                               CStr(File)
End Sub

Private Sub btnStop_Click()
    DisableRecording
    
    Hook False
    modWaveIn.StopRec
    Hook True
End Sub

Private Sub btnUnhook_Click()
    Hook False
End Sub

Private Sub cmbOutMode_Click()
    If cmbOutMode.ListIndex = 0 Then
        btnMP3Settings.Visible = False
    Else
        btnMP3Settings.Visible = True
    End If
End Sub

Private Sub Form_Load()
    Dim channels As Integer
    Dim File As String
    
    If Dir("C:\test.log") = "test.log" Then _
        Kill "C:\test.log"
    
    
    Caption = Caption & " v" & App.Major & "." & App.Minor
    
    INI_FILE = AppPath & "audio_recorder.ini"
    
    LoadSettings
    
    picPeakLeft.Width = 0
    picPeakRight.Width = 0
    
    cmbChannels.Text = cmbChannels.List(0)
    cmbBits.Text = cmbBits.List(0)
    cmbFrequency.Text = cmbFrequency.List(1)
    cmbOutMode.ListIndex = 1
    
    With lblState
        .Caption = "not recording"
        .ForeColor = &HFF0000
        .FontBold = False
    End With
    
    modWaveIn.msg = Space(Len(modWaveIn.msg))
    
    Hook True
    
    If Tag = "1" Then
        
        channels = 1
        If cmbChannels.ListIndex = 0 Then channels = 2
        
        File = ParseFile(txtFile.Text)
        
        EnableRecording
        
        modWaveIn.PrepareRecording channels, _
                                   CInt(cmbBits.Text), _
                                   CLng(cmbFrequency.Text), _
                                   File
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        txtDebug.Move 0, txtDebug.Top, ScaleWidth, ScaleHeight - txtDebug.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Hook False
    modWaveIn.StopRec
    modWaveIn.StopMonitoring
    SaveSettings
End Sub

Sub DebugIt(ByVal inString As String)
    txtDebug.Text = txtDebug.Text & inString & vbCrLf
End Sub

Private Sub chkdB_Click()
    If chkdB.Value = vbChecked Then
        bDecibels = True
        picDecibels.Visible = True
        picLinear.Visible = False
    Else
        bDecibels = False
        picDecibels.Visible = False
        picLinear.Visible = True
    End If
    
    If bMonitoring = False Then SetPeakmeter 0, 0
End Sub

Private Sub chkMonitoring_Click()
    If chkMonitoring.Value = vbChecked Then
        If hWaveIn_Monitor = 0 Then
            PrepareMonitoring
        Else
            PauseMonitoring
        End If
        
        b_Monitoring = True
    Else
        b_Monitoring = False
        PauseMonitoring
    End If
End Sub

Private Sub tmrTimer_Timer()
    Static sngLast As Single
    Static sLast As Single
    Static Max As Double
    Dim lLeft As Double
    Dim lRight As Double
    
    If Visible = False Then Exit Sub
    If WindowState = vbMinimized Then Exit Sub
    
    If bRecording = False And b_Recording = True Then
        DisableRecording
        b_Recording = False
    End If
    
    If bMonitoring = True And b_Monitoring = True Then
        
        If modWaveIn.GetCurPeak(lLeft, lRight) = True Then
            
            If (VBA.Timer - sLast) >= 0.05 Then
            
                SetPeakmeter lLeft, lRight
                
                sLast = VBA.Timer
            End If
            
        End If
    End If
    
    If b_Recording = True Then
    
        If (VBA.Timer - sngLast) >= 0.5 Then
            If lblState.Caption = "RECORDING" Then
                lblState.Caption = "                 "
            Else
                lblState.Caption = "RECORDING"
            End If
            
            sngLast = VBA.Timer
        End If
        
        lblTime.Caption = FormatTime(modWaveIn.GetTime)
        
        'lSample = (modWaveIn.SampleFrequency * 1000) \ (
    End If
    
    If Trim(modWaveIn.msg) <> "" Then
        DebugIt Trim(modWaveIn.msg)
        modWaveIn.msg = ""
    End If
    
    lblBytesWritten.Caption = CStr(modWaveIn.BytesWritten) & " bytes"
End Sub

Function FormatTime(ByVal lIn As Long, Optional ByVal bInMS As Boolean) As String
    Dim sec As String
    Dim min As String
    Dim tim As String
    Dim ms As String
    Dim hour As String
    
    ms = CStr(lIn)
    sec = CStr(ms)
    If bInMS Then sec = CStr(CLng(sec) \ 1000)
    
    
    hour = format(CStr(CLng(sec) \ 60 \ 60), "0#")
    min = format(CStr(CLng(sec) \ 60 - (CLng(hour) * 60)), "0#")
    sec = format(CStr(CLng(sec) - CLng(min) * 60 - (CLng(hour) * 60 * 60)), "0#")
    'ms = ms - (((CInt(min) * 60) + CInt(sec)) * 1000)
    
    tim = hour & ":" & min & ":" & sec
    
    FormatTime = tim
End Function

Sub EnableRecording()
    b_Recording = True
    
    btnPause.Enabled = True
    btnStop.Enabled = True
    btnStart.Enabled = False
    
    cmbChannels.Enabled = False
    cmbFrequency.Enabled = False
    cmbBits.Enabled = False
    txtFile.Enabled = False
    cmbOutMode.Enabled = False
    btnMP3Settings.Enabled = False
    
    With lblState
        .Caption = "RECORDING"
        .ForeColor = &HFF&
        .FontBold = True
    End With
End Sub

Sub DisableRecording()
    b_Recording = False
    
    btnPause.Enabled = False
    btnStop.Enabled = False
    btnStart.Enabled = True
    
    cmbChannels.Enabled = True
    cmbFrequency.Enabled = True
    cmbBits.Enabled = True
    txtFile.Enabled = True
    cmbOutMode.Enabled = True
    btnMP3Settings.Enabled = True
    
    With lblState
        .Caption = "not recording"
        .ForeColor = &HFF0000
        .FontBold = False
    End With
End Sub

Function ParseFile(ByVal File As String) As String
    ParseFile = Replace(File, "%date%", Replace(CStr(Date), ".", "-"))
    ParseFile = Replace(File, "%time%", Replace(CStr(Time), ":", "_"))
End Function

Sub LoadSettings()
    With MP3_Settings
        .MP3_Type = GetINILong("MP3", "MP3 Type", INI_FILE, 0)
        .VBR_MaxBitrate = GetINILong("MP3", "VBR Max Bitrate", INI_FILE, 320)
        .VBR_MinBitrate = GetINILong("MP3", "VBR Min Bitrate", INI_FILE, 256)
        .ABR_AvgBitrate = GetINILong("MP3", "ABR Average Bitrate", INI_FILE, 128)
        .CBR_Bitrate = GetINILong("MP3", "CBR Bitrate", INI_FILE, 192)
        .VBR_Quality = GetINILong("MP3", "VBR Quality", INI_FILE, 2)
        .VBR_Routine = GetINILong("MP3", "VBR Routine", INI_FILE, 0)
        
        .LAME = GetINIString("MP3", "Lame", INI_FILE)
    End With
End Sub

Sub SaveSettings()
    With MP3_Settings
        WriteINI "MP3", "MP3 Type", CStr(.MP3_Type), INI_FILE
        WriteINI "MP3", "VBR Max Bitrate", CStr(.VBR_MaxBitrate), INI_FILE
        WriteINI "MP3", "VBR Min Bitrate", CStr(.VBR_MinBitrate), INI_FILE
        WriteINI "MP3", "ABR Average Bitrate", CStr(.ABR_AvgBitrate), INI_FILE
        WriteINI "MP3", "VBR Quality", CStr(.VBR_Quality), INI_FILE
        WriteINI "MP3", "VBR Routine", CStr(.VBR_Routine), INI_FILE
        WriteINI "MP3", "CBR Bitrate", CStr(.CBR_Bitrate), INI_FILE
        
        WriteINI "MP3", "Lame", .LAME, INI_FILE
    End With
End Sub

Sub SetPeakmeter(ByVal lLeft As Double, ByVal lRight As Double)
    If bDecibels = False Then
        picPeakLeft.Width = (lLeft * CDbl(picPeak.ScaleWidth)) \ modWaveIn.PeakMax
        picPeakRight.Width = (lRight * CDbl(picPeak.ScaleWidth)) \ modWaveIn.PeakMax
        
        lblLeft.Caption = CStr(Round((lLeft / modWaveIn.PeakMax), 3))
        lblRight.Caption = CStr(Round((lRight / modWaveIn.PeakMax), 3))
    Else
        lLeft = modWaveIn.Calc_dB(lLeft)
        lRight = modWaveIn.Calc_dB(lRight)
        
        If lLeft < -100 Then lLeft = -100
        If lRight < -100 Then lRight = -100
        
        picPeakLeft.Width = (Abs(lLeft + 100) * CDbl(picPeak.ScaleWidth)) \ 90
        picPeakRight.Width = (Abs(lRight + 100) * CDbl(picPeak.ScaleWidth)) \ 90
        
        If lLeft > -100 Then
            lblLeft.Caption = CStr(Round(lLeft, 1))
            lblRight.Caption = CStr(Round(lRight, 1))
        Else
            lblLeft.Caption = "inf."
            lblRight.Caption = "inf."
        End If
    End If
End Sub
