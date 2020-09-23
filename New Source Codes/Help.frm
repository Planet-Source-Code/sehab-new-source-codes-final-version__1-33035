VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "..:: Help ::.."
   ClientHeight    =   2940
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Change Code Shower Language?"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "What does the Bo Forward do?"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2220
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "What does the Go Back do?"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2040
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   3000
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   2475
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "How do I minimize this form?"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "How do I Exit?"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "How do I use this?"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "What is Refresh?"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "What is Move Me?"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1350
   End
   Begin VB.Menu mnuccl 
      Caption         =   "SCCSL"
      Visible         =   0   'False
      Begin VB.Menu mnual 
         Caption         =   "All Languages"
      End
      Begin VB.Menu mnuvb 
         Caption         =   "Visual Basic"
      End
      Begin VB.Menu mnujj 
         Caption         =   "Java / Javascript"
      End
      Begin VB.Menu mnucc 
         Caption         =   "C / C++"
      End
      Begin VB.Menu mnuavs 
         Caption         =   "ASP / VbScript"
      End
      Begin VB.Menu mnusql 
         Caption         =   "SQL"
      End
      Begin VB.Menu mnup 
         Caption         =   "Perl"
      End
      Begin VB.Menu mnud 
         Caption         =   "Delphi"
      End
      Begin VB.Menu muphp 
         Caption         =   "PHP"
      End
      Begin VB.Menu mnucf 
         Caption         =   "Cold Fusion"
      End
      Begin VB.Menu mnunet 
         Caption         =   ".Net  ..:: New ::.."
      End
      Begin VB.Menu mnucm 
         Caption         =   "Change Mask/Background"
         Begin VB.Menu mnupd 
            Caption         =   "Poka-Dot"
         End
         Begin VB.Menu mnum 
            Caption         =   "Colorful"
         End
         Begin VB.Menu mnur 
            Caption         =   "Rain"
         End
         Begin VB.Menu mnubh 
            Caption         =   "Blue Hills"
         End
         Begin VB.Menu mnus 
            Caption         =   "Sunset"
         End
         Begin VB.Menu mnuwl 
            Caption         =   "Water Lilies"
         End
         Begin VB.Menu mnuw 
            Caption         =   "Winter"
         End
         Begin VB.Menu mnun 
            Caption         =   "None"
         End
      End
   End
   Begin VB.Menu mnust 
      Caption         =   "System Tray"
      Visible         =   0   'False
      Begin VB.Menu mnushow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form2
End Sub

Private Sub Label1_Click()
Label6.Caption = "If you click 'Move Me' on the 'New Source Code' form you will be able to move the form by moving your mouse as you hold the left button."
End Sub

Private Sub Label2_Click()
Label6.Caption = "If you click 'Refresh' on the 'New Source Code' form your browser will refresh."
End Sub

Private Sub Label3_Click()
Label6.Caption = "Well you really don't do anything, but you can move the form, refresh the form, and exit the form."
End Sub

Private Sub Label4_Click()
Label6.Caption = "You have to click the 'X' button on the 'New Source Code' form."
End Sub

Private Sub Label5_Click()
Label6.Caption = "You have to click the '_' button on the 'New Source Code' form."
End Sub

Private Sub Label7_Click()
Label6.Caption = "The 'Go Back' button in the 'New Source Codes' form makes the browser(Code Shower) go back like AOL's back button."
End Sub

Private Sub Label8_Click()
Label6.Caption = "The 'Go Forward' button in the 'New Source Codes' form makes the browser(Code Shower) go forward like AOL's forward button."
End Sub

Private Sub Label9_Click()
Label6.Caption = "This is a popup menu that allows you to change the language.  EX: Visual Basic"
End Sub

Private Sub mnual_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=-1"
End Sub

Private Sub mnuavs_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=4"
End Sub

Private Sub mnubh_Click()
Form1.Picture = LoadPicture(App.Path + "\blue hills.jpg")
End Sub

Private Sub mnucc_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=3"
End Sub

Private Sub mnucf_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=9"
End Sub

Private Sub mnud_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=7"
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnujj_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=2"
End Sub

Private Sub mnum_Click()
Form1.Picture = LoadPicture(App.Path + "\colorful.bmp")
End Sub

Private Sub mnun_Click()
Form1.Picture = Nothing
End Sub

Private Sub mnunet_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=10"
End Sub

Private Sub mnup_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=6"
End Sub

Private Sub mnupd_Click()
Form1.Picture = LoadPicture(App.Path + "\pd.bmp")
End Sub

Private Sub mnur_Click()
Form1.Picture = LoadPicture(App.Path + "\thing.ico")
End Sub

Private Sub mnus_Click()
Form1.Picture = LoadPicture(App.Path + "\sunset.jpg")
End Sub

Private Sub mnushow_Click()
Form1.WindowState = vbNormal
Form1.Show
End Sub

Private Sub mnusql_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=5"
End Sub

Private Sub mnuvb_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1"
End Sub

Private Sub mnuw_Click()
Form1.Picture = LoadPicture(App.Path + "\winter.jpg")
End Sub

Private Sub mnuwl_Click()
Form1.Picture = LoadPicture(App.Path + "\water lilies.jpg")
End Sub

Private Sub muphp_Click()
Form1.WebBrowser1.Navigate2 "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=8"
End Sub
