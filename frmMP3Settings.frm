VERSION 5.00
Begin VB.Form frmMP3Settings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MP3 Encoding configuration"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optABR 
      Caption         =   "ABR"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton optVBR 
      Caption         =   "VBR"
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton optCBR 
      Caption         =   "CBR"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame fraABR 
      Height          =   1095
      Left            =   4320
      TabIndex        =   12
      Top             =   360
      Width           =   2175
      Begin VB.ComboBox cmbABRAvgBitrate 
         Height          =   315
         ItemData        =   "frmMP3Settings.frx":0000
         Left            =   1200
         List            =   "frmMP3Settings.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblABRMinBitrate 
         AutoSize        =   -1  'True
         Caption         =   "Average bitrate:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Frame fraLamePath 
      Caption         =   "Path to lame.exe"
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   6375
      Begin VB.OptionButton optAppPath 
         Caption         =   "Application Path"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optSpecific 
         Caption         =   "Specific:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   780
         Width           =   1815
      End
      Begin VB.TextBox txtLamePath 
         Height          =   350
         Left            =   2160
         TabIndex        =   11
         Top             =   720
         Width           =   3975
      End
   End
   Begin VB.Frame fraVBR 
      Height          =   1695
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2175
      Begin VB.ComboBox cmbVBRRoutine 
         Height          =   315
         ItemData        =   "frmMP3Settings.frx":0036
         Left            =   1200
         List            =   "frmMP3Settings.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbVBRMinBitrate 
         Height          =   315
         ItemData        =   "frmMP3Settings.frx":004E
         Left            =   1200
         List            =   "frmMP3Settings.frx":0064
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbVBRMaxBitrate 
         Height          =   315
         ItemData        =   "frmMP3Settings.frx":0084
         Left            =   1200
         List            =   "frmMP3Settings.frx":009A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbVBRQuality 
         Height          =   315
         ItemData        =   "frmMP3Settings.frx":00BA
         Left            =   1200
         List            =   "frmMP3Settings.frx":00BC
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblVBRRoutine 
         AutoSize        =   -1  'True
         Caption         =   "VBR Routine:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label lblMaxBitrate 
         Caption         =   "Max. bitrate:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblMinBitrate 
         AutoSize        =   -1  'True
         Caption         =   "Min. bitrate:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   885
      End
      Begin VB.Label lblQuality 
         AutoSize        =   -1  'True
         Caption         =   "Quality:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   570
      End
   End
   Begin VB.Frame fraCBR 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      Begin VB.ComboBox cmbCBRBitrate 
         Height          =   315
         ItemData        =   "frmMP3Settings.frx":00BE
         Left            =   840
         List            =   "frmMP3Settings.frx":00D4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblBitrate 
         AutoSize        =   -1  'True
         Caption         =   "Bitrate:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmMP3Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoaded As Boolean

Private Sub btnClose_Click()
    With MP3_Settings
        If optCBR.Value = True Then
            .MP3_Type = CBR
        ElseIf optVBR.Value = True Then
            .MP3_Type = vbr
        ElseIf optABR.Value = True Then
            .MP3_Type = ABR
        End If
        
        .VBR_Quality = cmbVBRQuality.ListIndex
        .VBR_Routine = cmbVBRRoutine.ListIndex
        .VBR_MaxBitrate = FindBitrate(cmbVBRMaxBitrate.ListIndex)
        .VBR_MinBitrate = FindBitrate(cmbVBRMinBitrate.ListIndex)
        .ABR_AvgBitrate = FindBitrate(cmbABRAvgBitrate.ListIndex)
        .CBR_Bitrate = FindBitrate(cmbCBRBitrate.ListIndex)
        
        If optSpecific.Value = True Then
            .LAME = txtLamePath.Text
        Else
            .LAME = ""
        End If
    End With
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To 9
        cmbVBRQuality.AddItem CStr(i), i
    Next i
    
    With MP3_Settings
        cmbCBRBitrate.ListIndex = FindIndex(.CBR_Bitrate)
        cmbABRAvgBitrate.ListIndex = FindIndex(.ABR_AvgBitrate)
        cmbVBRQuality.ListIndex = .VBR_Quality
        cmbVBRMinBitrate.ListIndex = FindIndex(.VBR_MinBitrate)
        cmbVBRMaxBitrate.ListIndex = FindIndex(.VBR_MaxBitrate)
        cmbVBRRoutine.ListIndex = .VBR_Routine
        
        If .MP3_Type = CBR Then
            optCBR.Value = True
        ElseIf .MP3_Type = ABR Then
            optABR.Value = True
        ElseIf .MP3_Type = vbr Then
            optVBR.Value = True
        End If
        
        If .LAME = "" Then
            optAppPath.Value = True
        Else
            optSpecific.Value = True
            txtLamePath.Text = .LAME
        End If
    End With
    
    bLoaded = True
End Sub

Private Sub optABR_Click()
    If optABR.Value = True Then
        ABR_Enable
        VBR_Disable
        CBR_Disable
    End If
End Sub

Private Sub optAppPath_Click()
    txtLamePath.Enabled = False
End Sub

Private Sub optSpecific_Click()
    txtLamePath.Enabled = True
    If bLoaded = True Then txtLamePath.SetFocus
End Sub

Private Sub optVBR_Click()
    If optVBR.Value = True Then
        VBR_Enable
        ABR_Disable
        CBR_Disable
    End If
End Sub

Private Sub optCBR_Click()
    If optCBR.Value = True Then
        CBR_Enable
        VBR_Disable
        ABR_Disable
    End If
End Sub

Sub CBR_Enable()
    fraCBR.Enabled = True
    cmbCBRBitrate.Enabled = True
End Sub

Sub CBR_Disable()
    fraCBR.Enabled = False
    cmbCBRBitrate.Enabled = False
End Sub

Sub VBR_Enable()
    fraVBR.Enabled = True
    cmbVBRMinBitrate.Enabled = True
    cmbVBRMaxBitrate.Enabled = True
    cmbVBRQuality.Enabled = True
    cmbVBRRoutine.Enabled = True
End Sub

Sub VBR_Disable()
    fraVBR.Enabled = False
    cmbVBRMinBitrate.Enabled = False
    cmbVBRMaxBitrate.Enabled = False
    cmbVBRQuality.Enabled = False
    cmbVBRRoutine.Enabled = False
End Sub

Sub ABR_Enable()
    fraABR.Enabled = True
    cmbABRAvgBitrate.Enabled = True
End Sub

Sub ABR_Disable()
    fraABR.Enabled = False
    cmbABRAvgBitrate.Enabled = False
End Sub

Function FindIndex(ByVal Bitrate As Integer) As Integer
    If Bitrate = 320 Then
        FindIndex = 0
    ElseIf Bitrate = 256 Then
        FindIndex = 1
    ElseIf Bitrate = 192 Then
        FindIndex = 2
    ElseIf Bitrate = 128 Then
        FindIndex = 3
    ElseIf Bitrate = 96 Then
        FindIndex = 4
    ElseIf Bitrate = 64 Then
        FindIndex = 5
    Else
        FindIndex = -1
    End If
End Function

Function FindBitrate(ByVal Index As Integer) As Integer
    If Index = 0 Then
        FindBitrate = 320
    ElseIf Index = 1 Then
        FindBitrate = 256
    ElseIf Index = 2 Then
        FindBitrate = 192
    ElseIf Index = 3 Then
        FindBitrate = 128
    ElseIf Index = 4 Then
        FindBitrate = 96
    ElseIf Index = 5 Then
        FindBitrate = 96
    Else
        FindBitrate = -1
    End If
End Function
