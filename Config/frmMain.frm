VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   0  'None
   Caption         =   "Config"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":0E42
   ScaleHeight     =   3750
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox lstDisplay 
      Height          =   315
      Left            =   2445
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   675
      Width           =   1380
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000C000&
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   90
      Width           =   255
   End
   Begin VB.PictureBox picPreview 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   240
      Picture         =   "frmMain.frx":1A09
      ScaleHeight     =   1560
      ScaleWidth      =   1965
      TabIndex        =   8
      Top             =   495
      Width           =   1965
      Begin VB.PictureBox picDisplay 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1230
         Left            =   165
         Picture         =   "frmMain.frx":1C68
         ScaleHeight     =   82
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
         TabIndex        =   9
         Top             =   135
         Width           =   1635
      End
   End
   Begin VB.CheckBox chkMusic 
      BackColor       =   &H00004000&
      Caption         =   "Play Music"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   615
      TabIndex        =   7
      Top             =   2265
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00004000&
      Caption         =   "Graphics"
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   2340
      TabIndex        =   3
      Top             =   405
      Width           =   1575
      Begin VB.CheckBox chkFlat 
         BackColor       =   &H00004000&
         Caption         =   "Render Flat"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   990
         Width           =   1350
      End
      Begin VB.CheckBox chkFiltering 
         BackColor       =   &H00004000&
         Caption         =   "Bilinear Filtering"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   690
         Width           =   1395
      End
      Begin VB.CheckBox chkFS 
         BackColor       =   &H00004000&
         Caption         =   "Fullscreen"
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   1455
         Y1              =   1290
         Y2              =   1290
      End
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3195
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2205
      Width           =   705
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000C000&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2205
      Width           =   705
   End
   Begin VB.PictureBox picMainSkin 
      Height          =   2715
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   0
      Width           =   4275
      Begin VB.Image imgIcon 
         Height          =   240
         Left            =   195
         Picture         =   "frmMain.frx":2A12
         Stretch         =   -1  'True
         Top             =   75
         Width           =   240
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Frog3D Configuration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   495
         TabIndex        =   10
         Top             =   90
         Width           =   1860
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   300
      X2              =   2160
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   300
      X2              =   2160
      Y1              =   2205
      Y2              =   2205
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ScreenX As String
Dim ScreenY As String

Dim UseFS As String
Dim UseMusic As String
Dim UseFiltering As String
Dim UseFlat As String
Dim TickCount As String

Private Sub chkFS_Click()
    'UpdatePreview
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    ChDir App.Path
    
    If chkFS.Value = 1 Then UseFS = "TRUE" Else UseFS = "FALSE"
    
    If lstDisplay.Text = "320x240" Then
        ScreenX = 320: ScreenY = 240
    ElseIf lstDisplay.Text = "512x384" Then
        ScreenX = 512: ScreenY = 384
    ElseIf lstDisplay.Text = "640x480" Then
        ScreenX = 640: ScreenY = 480
    ElseIf lstDisplay.Text = "800x600" Then
        ScreenX = 800: ScreenY = 600
    ElseIf lstDisplay.Text = "1024x768" Then
        ScreenX = 1024: ScreenY = 768
    End If
    
    If chkMusic.Value = 1 Then UseMusic = "TRUE" Else UseMusic = "FALSE"
    If chkFiltering.Value = 1 Then UseFiltering = "TRUE" Else UseFiltering = "FALSE"
    If chkFlat.Value = 1 Then UseFlat = "TRUE" Else UseFlat = "FALSE"
    On Error GoTo Oops
    
    Open "frog3d.cfg" For Output As #1
    Print #1, UseFS
    Print #1, ScreenX
    Print #1, ScreenY
    Print #1, UseMusic
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

Private Sub LoadConfig()
    ChDir App.Path
    
    lstDisplay.AddItem "320x240"
    lstDisplay.AddItem "512x384"
    lstDisplay.AddItem "640x480"
    lstDisplay.AddItem "800x600"
    lstDisplay.AddItem "1024x768"
    
    On Error GoTo NoSuchFile
    Open "frog3d.cfg" For Input As #1
    
    Input #1, UseFS
    Input #1, ScreenX
    Input #1, ScreenY
    Input #1, UseMusic
    Input #1, UseFiltering
    Input #1, UseFlat
    Input #1, TickCount
    
    If UseFS = "TRUE" Then chkFS.Value = 1 Else chkFS.Value = 0
    If UseMusic = "TRUE" Then chkMusic.Value = 1
    If UseFiltering = "TRUE" Then chkFiltering.Value = 1 Else chkFiltering.Value = 0
    If UseFlat = "TRUE" Then chkFlat.Value = 1
    
    If ScreenX = "320" Then lstDisplay.Text = "320x240"
    If ScreenX = "512" Then lstDisplay.Text = "512x384"
    If ScreenX = "640" Then lstDisplay.Text = "640x480"
    If ScreenX = "800" Then lstDisplay.Text = "800x600"
    If ScreenX = "1024" Then lstDisplay.Text = "1024x768"
    
    Close #1
    
    GoTo Success
    
NoSuchFile:
    MsgBox "Sorry, config file not found", , "Error"
    End
Success:
    'UpdatePreview
End Sub

Private Sub img640Full_Click()
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
    'UpdatePreview
    'If optScreenRes(Index).Value = True Then Exit Sub
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdWeb_Click()
    Shell "start http://members.truepath.com/marshallh", vbNormalFocus
End Sub

Private Sub Form_Load()
    Dim WindowRegion As Long
    
    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    picMainSkin.Top = 0
    picMainSkin.Left = 0
    Me.BorderStyle = vbBSNone
        
    'Set picMainSkin.Picture = LoadPicture(App.Path & "\main.bmp")
    picMainSkin.Picture = Me.Picture
    
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
    
    LoadConfig
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lstDisplay_Change()
    If Int(Right(lstDisplay.Text, 3)) > Screen.Height / Screen.TwipsPerPixelY And chkFS.Value = 0 Then
        'user chose a bigger window than their display mode can handle
        MsgBox "You chose a display size that, when windowed, will be bigger than your desktop. Either select a lower resolution or switch to full screen.", vbCritical + vbOKOnly
    End If
End Sub

Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub


