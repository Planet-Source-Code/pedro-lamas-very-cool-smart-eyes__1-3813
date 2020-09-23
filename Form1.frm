VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   3720
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      Height          =   735
      Left            =   1590
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   735
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   2280
      Width           =   1500
      Begin VB.Image Skin 
         Height          =   735
         Left            =   0
         Picture         =   "Form1.frx":0A63
         Top             =   0
         Width           =   1455
      End
      Begin VB.Image Pupil 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "Form1.frx":135F
         Top             =   240
         Width           =   375
      End
      Begin VB.Image Pupil 
         Height          =   375
         Index           =   1
         Left            =   960
         Picture         =   "Form1.frx":17A4
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   0
      Picture         =   "Form1.frx":1BE9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":266A
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":306E
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":39D8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Window 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   330
      TabIndex        =   3
      Top             =   1200
      Width           =   4005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Window under cursor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1875
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   1880
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   120
      Shape           =   2  'Oval
      Top             =   840
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "I'm watching you!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private XT(1) As Single, YT As Single, M As Single
Private XScreen As Single, YScreen As Single
Private I As Integer, II As Integer
Private Const MAX_DELOCATION = 250
Private Const PUPIL_DISTANCE = 30

Private Sub Form_Load()
    XT(0) = Skin.Left - Pupil(0).Width / 2 + MAX_DELOCATION + PUPIL_DISTANCE
    XT(1) = Skin.Left + Skin.Width / 2 - Pupil(1).Width / 2 + MAX_DELOCATION - PUPIL_DISTANCE
    YT = Skin.Top - Pupil(0).Height / 2 + MAX_DELOCATION
    M = Skin.Width / 2 - MAX_DELOCATION * 2
    XScreen = Screen.Width / Screen.TwipsPerPixelX
    YScreen = Screen.Height / Screen.TwipsPerPixelY
    II = 1
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Programmed by Pedro Lamas" & vbCrLf & "Copyright ©1997-1999 Underground Software" & vbCrLf & vbCrLf & "Home-Page (Dedicated to VB): www.terravista.pt/portosanto/3723/" & vbCrLf & "E-Mail: sniper@hotpop.com", 9 + vbInformation, "Credits!"
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub Timer1_Timer()
    Dim CP As POINTAPI, hWnd As Long, S As String
    GetCursorPos CP
    Pupil(0).Move XT(0) + CP.X * M / XScreen, YT + CP.Y * M / YScreen
    Pupil(1).Move XT(1) + CP.X * M / XScreen, YT + CP.Y * M / YScreen
    hWnd = WindowFromPoint(CP.X, CP.Y)
    S = Space(128)
    GetWindowText hWnd, S, 128
    If Asc(Left(S, 1)) = 0 Then GetClassName hWnd, S, 128
    Window.Caption = S
    'With Picture1
    '    .PaintPicture Image1(I).Picture, Skin.Left, Skin.Top
    '    .PaintPicture Pupil(0).Picture, Pupil(0).Left, Pupil(0).Top
    '    .PaintPicture Pupil(1).Picture, Pupil(1).Left, Pupil(1).Top
    'End With
    DoEvents
End Sub

Private Sub Timer2_Timer()
    If I + II < 0 Or I + II > 3 Then II = -II
    I = I + II
    Skin.Picture = Image1(I)
    If I = 0 Then
        Timer2.Interval = 3000
    Else
        Timer2.Interval = 100
    End If
End Sub
