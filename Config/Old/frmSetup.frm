VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Frog3D"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Game Setup"
      Height          =   2520
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   3930
      Begin VB.PictureBox picPreview 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1560
         Left            =   135
         Picture         =   "frmSetup.frx":0E42
         ScaleHeight     =   1560
         ScaleWidth      =   1965
         TabIndex        =   11
         Top             =   345
         Width           =   1965
         Begin VB.PictureBox picDisplay 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1230
            Left            =   165
            ScaleHeight     =   82
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   109
            TabIndex        =   12
            Top             =   135
            Width           =   1635
         End
      End
      Begin VB.CheckBox chkJoystick 
         Caption         =   "Use Joystick"
         Height          =   195
         Left            =   510
         TabIndex        =   8
         Top             =   2175
         Width           =   1230
      End
      Begin VB.Frame Frame2 
         Caption         =   "Graphics"
         Height          =   1740
         Left            =   2235
         TabIndex        =   3
         Top             =   270
         Width           =   1575
         Begin VB.CheckBox chkFlat 
            Caption         =   "Render Flat"
            Height          =   210
            Left            =   135
            TabIndex        =   10
            Top             =   1200
            Width           =   1350
         End
         Begin VB.CheckBox chkFiltering 
            Caption         =   "Bilinear Filtering"
            Height          =   225
            Left            =   135
            TabIndex        =   9
            Top             =   975
            Width           =   1395
         End
         Begin VB.CheckBox chkFS 
            Caption         =   "Use Fullscreen"
            Height          =   330
            Left            =   135
            TabIndex        =   7
            Top             =   1365
            Width           =   1380
         End
         Begin VB.OptionButton optScreenRes 
            Caption         =   "320x240"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   255
            Width           =   1005
         End
         Begin VB.OptionButton optScreenRes 
            Caption         =   "512x384"
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   465
            Width           =   1080
         End
         Begin VB.OptionButton optScreenRes 
            Caption         =   "640x480"
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   675
            Width           =   1050
         End
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   360
         Left            =   3090
         TabIndex        =   2
         Top             =   2055
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   360
         Left            =   2325
         TabIndex        =   1
         Top             =   2055
         Width           =   705
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   195
         X2              =   2055
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   195
         X2              =   2055
         Y1              =   2055
         Y2              =   2055
      End
   End
   Begin VB.Image img320Full 
      Height          =   1230
      Left            =   3750
      Picture         =   "frmSetup.frx":10A1
      Top             =   4035
      Width           =   1635
   End
   Begin VB.Image img512Full 
      Height          =   1230
      Left            =   3750
      Picture         =   "frmSetup.frx":79F3
      Top             =   2625
      Width           =   1635
   End
   Begin VB.Image img640Full 
      Height          =   1230
      Left            =   1965
      Picture         =   "frmSetup.frx":E345
      Top             =   4050
      Width           =   1635
   End
   Begin VB.Image img320Window 
      Height          =   1230
      Left            =   180
      Picture         =   "frmSetup.frx":14C97
      Top             =   4035
      Width           =   1635
   End
   Begin VB.Image img512Window 
      Height          =   1230
      Left            =   1965
      Picture         =   "frmSetup.frx":1B5E9
      Top             =   2625
      Width           =   1635
   End
   Begin VB.Image img640Window 
      Height          =   1230
      Left            =   180
      Picture         =   "frmSetup.frx":21F3B
      Top             =   2625
      Width           =   1635
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ScreenX As String
Dim ScreenY As String

Dim UseFS As String
Dim UseJoystick As String
Dim UseFiltering As String
Dim UseFlat As String
Dim TickCount As String

Private Sub chkFS_Click()
    UpdatePreview
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    ChDir App.Path
    
    If chkFS.Value = 1 Then UseFS = "TRUE" Else UseFS = "FALSE"
    If optScreenRes(0).Value = True Then
        ScreenX = 320
        ScreenY = 240
    ElseIf optScreenRes(1).Value = True Then
        ScreenX = 512
        ScreenY = 384
    ElseIf optScreenRes(2).Value = True Then
        ScreenX = 640
        ScreenY = 480
    End If
    If chkJoystick.Value = 1 Then UseJoystick = "TRUE" Else UseJoystick = "FALSE"
    If chkFiltering.Value = 1 Then UseFiltering = "TRUE" Else UseFiltering = "FALSE"
    If chkFlat.Value = 1 Then UseFlat = "TRUE" Else UseFlat = "FALSE"
    On Error GoTo Oops
    
    Open "frog3d.cfg" For Output As #1
    Print #1, UseFS
    Print #1, ScreenX
    Print #1, ScreenY
    Print #1, UseJoystick
    Print #1, UseFiltering
    Print #1, UseFlat
    Print #1, TickCount
    
    Close #1
    
    GoTo Success
    
Oops:
    MsgBox "Sorry, config file not found", , "Error"
    End
Success:

End Sub

Private Sub Form_Load()
    ChDir App.Path
    
    On Error GoTo NoSuchFile
    Open "frog3d.cfg" For Input As #1
    
    Input #1, UseFS
    Input #1, ScreenX
    Input #1, ScreenY
    Input #1, UseJoystick
    Input #1, UseFiltering
    Input #1, UseFlat
    Input #1, TickCount
    
    If UseFS = "TRUE" Then chkFS.Value = 1 Else chkFS.Value = 0
    If UseJoystick = "TRUE" Then chkJoystick.Value = 1
    If UseFiltering = "TRUE" Then chkFiltering.Value = 1 Else chkFiltering.Value = 0
    If UseFlat = "TRUE" Then chkFlat.Value = 1
    
    If ScreenX = "320" Then optScreenRes(0).Value = True
    If ScreenX = "512" Then optScreenRes(1).Value = True
    If ScreenX = "640" Then optScreenRes(2).Value = True
    
    Close #1
    
    GoTo Success
    
NoSuchFile:
    MsgBox "Sorry, config file not found", , "Error"
    End
Success:
    UpdatePreview
End Sub

Private Sub optScreenRes_Click(Index As Integer)
    
    If Index = 0 Then
        ScreenX = 320
        ScreenY = 240
    ElseIf Index = 1 Then
        ScreenX = 512
        ScreenY = 384
    ElseIf Index = 2 Then
        ScreenX = 640
        ScreenY = 480
    End If
    UpdatePreview
    'If optScreenRes(Index).Value = True Then Exit Sub
End Sub

Private Sub UpdatePreview()
    If ScreenX = 320 Then
        If chkFS.Value = 1 Then picDisplay.Picture = img320Full Else picDisplay.Picture = img320Window
    End If
    
    If ScreenX = 512 Then
        If chkFS.Value = 1 Then picDisplay.Picture = img512Full Else picDisplay.Picture = img512Window
    End If
    
    If ScreenX = 640 Then
        If chkFS.Value = 1 Then picDisplay.Picture = img640Full Else picDisplay.Picture = img640Window
    End If
End Sub

