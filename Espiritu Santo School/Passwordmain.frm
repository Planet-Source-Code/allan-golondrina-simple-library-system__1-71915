VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Password Maintenance"
   ClientHeight    =   2340
   ClientLeft      =   3765
   ClientTop       =   2340
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   2340
   ScaleWidth      =   5685
   Begin VB.CommandButton NEWOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8160
      TabIndex        =   17
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   8160
      TabIndex        =   16
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9240
      TabIndex        =   15
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CMDclose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CMDcancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CMDok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      TabIndex        =   7
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Image Image11 
      Height          =   1095
      Left            =   4680
      Picture         =   "Passwordmain.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   960
   End
   Begin VB.Image Image10 
      Height          =   1095
      Left            =   4680
      Picture         =   "Passwordmain.frx":330EA
      Stretch         =   -1  'True
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName:"
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password:"
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   18
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image Image8 
      Height          =   435
      Left            =   5760
      Picture         =   "Passwordmain.frx":661D4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5850
   End
   Begin VB.Image Image7 
      Height          =   555
      Left            =   5760
      Picture         =   "Passwordmain.frx":6B862
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   5850
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1440
      X2              =   3120
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Image Image5 
      Height          =   555
      Left            =   0
      Picture         =   "Passwordmain.frx":70EF0
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   5850
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   0
      Picture         =   "Passwordmain.frx":7657E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   5850
   End
   Begin VB.Image Image3 
      Height          =   435
      Left            =   0
      Picture         =   "Passwordmain.frx":7BC0C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5850
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New UserName:"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName:"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   0
      Picture         =   "Passwordmain.frx":8129A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5850
   End
   Begin VB.Image Image4 
      Height          =   435
      Left            =   0
      Picture         =   "Passwordmain.frx":D5EF4
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   5850
   End
   Begin VB.Image Image6 
      Height          =   2295
      Left            =   0
      Picture         =   "Passwordmain.frx":DB582
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   5850
   End
   Begin VB.Image Image9 
      Height          =   2295
      Left            =   5760
      Picture         =   "Passwordmain.frx":1301DC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5850
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As Database
Dim tn As Recordset
Dim C As Integer
Option Compare Text



Private Sub Command1_Click()
a = MsgBox("Do you want to add this user", vbYesNo + vbQuestion, "save")
Dim sw As Boolean
If a = vbYes Then
tn.FindFirst "username like'" + Text6.Text + "'"
If tn.NoMatch = False Then
MsgBox "This Username  already exist"
Text6.Text = ""
Else
If Text6.Text <> "" And Text7.Text <> "" Then

tn.AddNew
tn!UserName = Text6.Text
tn!Password = Text7.Text
tn.Update
Else
a = MsgBox("you must complete the information", vbExclamation + vbOK, "save")
End If
End If
End If
End Sub

Private Sub Command2_Click()
Me.Height = 2865
Me.Width = 11265
CMDok.Top = 1920
 CMDcancel.Top = 1920
 CMDclose.Top = 1920
 NEWOK.Visible = False
 CMDok.Visible = True
End Sub

Private Sub Command3_Click()
Me.Width = 5510
End Sub

Private Sub Form_Load()

Set dn = OpenDatabase(App.Path + "\SADconvert.mdb")
Set tn = dn.OpenRecordset("login", dbOpenDynaset)
End Sub


Private Sub CMDok_Click()
'Select Case CMDok.Caption
' Case "&Ok"
 ' If tn!UserName = Text1.Text And tn!Password = Text2.Text Then
  ' Frame1.Visible = False
   'CMDok.Caption = "C&hage"
  ' Form4.Height = 5370
  ' CMDok.Top = 4440
  ' CMDcancel.Top = 4440
  ' CMDclose.Top = 4440
  ' Text3.SetFocus
  ' SendKeys "{Home}+{End}"
'Else
'   MsgBox "Invalid User Name or Password !!", vbOKOnly + vbInformation, "Login"
'   Text1.SetFocus
'   SendKeys "{Home}+{End}"
 '  Text2.Text = ""
 '
 ' End If
 'Case "C&hage"
 ' If Text4.Text = text5.Text Then
  ' tn.Edit
   ' tn!UserName = Text3.Text
    'tn!Password = Text4.Text
   'tn.Update
   ' MsgBox "Login password change!!", vbOKOnly + vbInformation, "Login"
   ' Frame1.Visible = True
    'Text1.Text = ""
   ' Text2.Text = ""
   ' Text3.Text = ""
   ' Text4.Text = ""
   ' text5.Text = ""
   ' Form4.Height = 3075
   ' CMDok.Top = 1680
   ' CMDcancel.Top = 1680
    'CMDclose.Top = 1680

    
  'Else
   ' MsgBox "Password dont mach  !!", vbOKOnly + vbInformation, "Login"
   ' text5.Text = ""
   ' Text4.SetFocus
   ' SendKeys "{Home}+{End}"
    
  'End If

'End Select


''''''''''''

If C <> 1000 Then
 C = C + 1
tn.FindFirst " password like '" + Text2.Text + "'"
   If tn.Fields("Username") = Text1.Text And tn.Fields("password") = Text2.Text Then
   Form4.Height = 5370
   Command2.Visible = True
    CMDok.Top = 4440
    CMDok.Visible = False
    CMDcancel.Top = 4440
    CMDclose.Top = 4440
    NEWOK.Visible = True
    '''Command1.Visible = False
    '''Command2.Visible = False
   
 ElseIf tn.Fields("UserName") <> Text1.Text And tn.Fields("password") = Text2.Text Then
  d = MsgBox("Your user name is incorrect", vbCritical, "Try again")
  
  ElseIf tn.Fields("UserName") = Text1.Text And tn.Fields("password") <> Text2.Text Then
  d = MsgBox("Your user password is incorrect", vbCritical, "Try again")
  
  ElseIf tn.Fields("UserName") <> Text1.Text And tn.Fields("password") <> Text2.Text Then
  d = MsgBox("Access denied", vbCritical, "Try again")
  End If
  End If
End Sub

Private Sub CMDclose_Click()
End
End Sub

Private Sub CMDcancel_Click()
Form1.Show
Unload Me

End Sub

Private Sub NEWOK_Click()
If Text4.Text = text5.Text Then
   tn.Edit
    tn!UserName = Text3.Text
    tn!Password = Text4.Text
   tn.Update
    MsgBox "Login password change!!", vbOKOnly + vbInformation, "Login"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    text5.Text = ""
    Me.Height = 2835
    Me.Top = 1680
    Me.Top = 1680
    Me.Top = 1680
    CMDok.Top = 1920
    CMDcancel.Top = 1920
    CMDclose.Top = 1920
    CMDok.Visible = True
    Command1.Visible = True
    Command2.Visible = True
    
  Else
    MsgBox "Password dont mach  !!", vbOKOnly + vbInformation, "Login"
    text5.Text = ""
    Text4.SetFocus
    SendKeys "{Home}+{End}"
    
  End If
End Sub
