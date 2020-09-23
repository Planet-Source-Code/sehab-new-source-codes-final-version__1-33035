VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "NSC"
   ClientHeight    =   5340
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6915
   FillColor       =   &H0000FF00&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "NSC.frx":0000
   ScaleHeight     =   5340
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Preview of Background"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Forward"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1800
      Top             =   720
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   3975
      Left            =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   6015
      ExtentX         =   10610
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label6 
      Caption         =   "Change Code Showers Language and Background Mask"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "__"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000000&
      Caption         =   "New Source Codes               Move Me"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   5985
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                      
Private Sub Command1_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub

Private Sub Command2_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub Command3_Click()
Form3.Image1.Picture = Form1.Picture
Form3.Show
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=-1"
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then
Me.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hWnd = Me.hWnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
Shell_NotifyIcon NIM_DELETE, nid
End If
 Dim Result
 Result = CreateRoundRectRgn(0, 0, Me.Width / 15, (Me.Height / 15), 120, 120)
 Result = SetWindowRgn(Me.hWnd, Result, True)
End Sub

Private Sub Label1_Click()
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub Label2_Click()
Me.WindowState = 1
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
Module1.ReleaseCapture
Module1.SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTMOVE, 0
End If
End Sub

Private Sub Label4_Click()
WebBrowser1.Refresh
End Sub

Private Sub Label5_Click()
Form2.Show
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Form1.PopupMenu Form2.mnuccl
End Sub

Private Sub Timer2_Timer()
If Me.WindowState = 0 Then Me.BorderStyle = 0
If Me.WindowState = 1 Then Me.BorderStyle = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Sys As Long
Sys = x / Screen.TwipsPerPixelX
Select Case Sys
Case WM_LBUTTONDOWN:
Me.PopupMenu Form2.mnust
End Select
End Sub
