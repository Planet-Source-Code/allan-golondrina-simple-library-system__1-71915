VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Main Menu"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   15240
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   10515
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Donation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF0000&
      Height          =   735
      Left            =   8280
      Picture         =   "Menu.frx":240042
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9000
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   10440
      Width           =   1815
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   12
         Left            =   0
         TabIndex        =   13
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Catalogue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   2
         Left            =   3795
         TabIndex        =   12
         Top             =   1875
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password Maintenence"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   735
         Index           =   10
         Left            =   0
         TabIndex        =   11
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   8
         Left            =   0
         TabIndex        =   10
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Returning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   6
         Left            =   0
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Borrowing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Image Image12 
      Height          =   2370
      Left            =   3240
      Picture         =   "Menu.frx":240484
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   8790
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   11760
      Picture         =   "Menu.frx":25DD46
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image10 
      Height          =   720
      Left            =   11760
      Picture         =   "Menu.frx":261220
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   9840
      Picture         =   "Menu.frx":2646FA
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image8 
      Height          =   720
      Left            =   9840
      Picture         =   "Menu.frx":266B54
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   6480
      Picture         =   "Menu.frx":268FAE
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   6480
      Picture         =   "Menu.frx":26D5B0
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   4800
      Picture         =   "Menu.frx":271BB2
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   4800
      Picture         =   "Menu.frx":27409C
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   3120
      Picture         =   "Menu.frx":276586
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   8040
      TabIndex        =   5
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   11880
      TabIndex        =   4
      Top             =   9840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password Maintenence"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   9720
      TabIndex        =   3
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Returning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6360
      TabIndex        =   2
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Borrowing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4560
      TabIndex        =   1
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Catalogue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   0
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   3120
      Picture         =   "Menu.frx":2795C8
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   10650
      Left            =   0
      Picture         =   "Menu.frx":27C60A
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   15330
   End
   Begin VB.Menu MNUabout 
      Caption         =   "&About"
      Begin VB.Menu SUBauthor 
         Caption         =   "Author"
      End
      Begin VB.Menu SUBhelp 
         Caption         =   "&Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub Command1_Click()
DONATION.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1(1).ForeColor = vbBlack
Label1(2).ForeColor = vbYellow
Label1(3).ForeColor = vbBlack
Label1(4).ForeColor = vbGreen
Label1(5).ForeColor = vbBlack
Label1(6).ForeColor = vbBlue
Label1(7).ForeColor = vbBlack
Label1(8).ForeColor = vbRed
Label1(9).ForeColor = vbBlack
Label1(10).ForeColor = &HFF00FF
Label1(11).ForeColor = vbBlack
Label1(12).ForeColor = vbWhite

End Sub
    


Private Sub Command3_Click()

'Unload Me
End Sub
Private Sub Command4_Click()
Form7.Show

End Sub
Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1(8).ForeColor = vbBlack
Label1(7).ForeColor = vbRed
End Sub


Private Sub Command5_Click()
End Sub
Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

End Sub

Private Sub Command6_Click()

End Sub
Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1(12).ForeColor = vbBlack
Label1(11).ForeColor = vbWhite
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1(1).ForeColor = vbBlack
Label1(2).ForeColor = vbYellow
Label1(3).ForeColor = vbBlack
Label1(4).ForeColor = vbGreen
Label1(5).ForeColor = vbBlack
Label1(6).ForeColor = vbBlue
Label1(7).ForeColor = vbBlack
Label1(8).ForeColor = vbRed
Label1(9).ForeColor = vbBlack
Label1(10).ForeColor = &HFF00FF
Label1(11).ForeColor = vbBlack
Label1(12).ForeColor = vbWhite

End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image11.Visible = False
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image11.Visible = True
End
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Visible = False
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1(2).ForeColor = vbBlack
Label1(1).ForeColor = vbYellow
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Visible = True
Form5.Show
'Unload Me
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Visible = False
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1(4).ForeColor = vbBlack
Label1(3).ForeColor = vbGreen
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Visible = True
Form2.Show
'Unload Me
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image7.Visible = False
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1(6).ForeColor = vbBlack
Label1(5).ForeColor = vbBlue
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image7.Visible = True
Form3.Show

End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image9.Visible = False
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Label1(10).ForeColor = vbBlack
Label1(9).ForeColor = &HFF00FF
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image9.Visible = True
Form4.Show
'Unload Me

End Sub
