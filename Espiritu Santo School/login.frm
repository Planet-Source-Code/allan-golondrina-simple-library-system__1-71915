VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2325
   ClientLeft      =   5115
   ClientTop       =   2280
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1200
      TabIndex        =   0
      Text            =   "a"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "a"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image10 
      Height          =   1095
      Left            =   3480
      Picture         =   "login.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   960
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3360
      Picture         =   "login.frx":330EA
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   960
   End
   Begin VB.Image Image6 
      Height          =   315
      Left            =   3360
      Picture         =   "login.frx":346EC
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   960
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2400
      Picture         =   "login.frx":35CEE
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   960
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   2400
      Picture         =   "login.frx":372F0
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   -240
      Picture         =   "login.frx":388F2
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   9570
   End
   Begin VB.Image Image3 
      Height          =   435
      Left            =   0
      Picture         =   "login.frx":3DF80
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9570
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Password:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   -120
      Picture         =   "login.frx":4360E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5850
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As Database
Dim tn As Recordset
Option Compare Text
Dim C As Integer

Private Sub Data2_Validate(Action As Integer, Save As Integer)
End Sub

Private Sub Form_Load()

Set dn = OpenDatabase(App.Path + "\SADconvert.mdb")
Set tn = dn.OpenRecordset("login", dbOpenDynaset)

End Sub


Private Sub CMDok_Click()

End Sub

Private Sub CMDclose_Click()

End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Visible = False
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Visible = True

 'If tn!UserName = Text1.Text And tn!Password = Text2.Text Then
  'Form1.Enabled = True
  'Form1.Show
  'Unload Me
  
 'Else
  'MsgBox "Invalid User Name or Password !!", vbOKOnly + vbInformation, "Login"
  'Text1.SetFocus
  'SendKeys "{Home}+{End}"
  ' Form1.Enabled = True
 'End If
 If C <> 1000 Then
 C = C + 1
 
 tn.FindFirst " password like '" + Text2.Text + "'"
   If tn.Fields("Username") = Text1.Text And tn.Fields("password") = Text2.Text Then
   d = MsgBox("Access Granted", vbInformation, "Try again")
   Form1.Show
   Unload Me
   Form1.Enabled = True
 
  
  ElseIf tn.Fields("UserName") <> Text1.Text And tn.Fields("password") = Text2.Text Then
  d = MsgBox("Your user name is incorrect", vbCritical, "Try again")
  
  ElseIf tn.Fields("UserName") = Text1.Text And tn.Fields("password") <> Text2.Text Then
  d = MsgBox("Your user password is incorrect", vbCritical, "Try again")
  
  ElseIf tn.Fields("UserName") <> Text1.Text And tn.Fields("password") <> Text2.Text Then
  d = MsgBox("Access denied", vbCritical, "Try again")
  End If
  End If



End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image7.Visible = False
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image7.Visible = True

End
End Sub

