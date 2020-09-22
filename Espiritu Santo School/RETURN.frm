VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Returning Section"
   ClientHeight    =   6390
   ClientLeft      =   3510
   ClientTop       =   2010
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   6390
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8400
      Top             =   120
   End
   Begin VB.CommandButton CMDclose 
      Caption         =   "&Close"
      Height          =   855
      Left            =   6240
      Picture         =   "RETURN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CMDview 
      Caption         =   "&View"
      Height          =   855
      Left            =   5280
      Picture         =   "RETURN.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CMDdamage 
      Caption         =   "&Damage"
      Height          =   855
      Left            =   4320
      Picture         =   "RETURN.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CDMrefresh 
      Caption         =   "&Refresh"
      Height          =   855
      Left            =   3360
      Picture         =   "RETURN.frx":3026
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CMDsearch 
      Caption         =   "&Search"
      Height          =   855
      Left            =   2400
      Picture         =   "RETURN.frx":3170
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CMDreturn 
      Caption         =   "Re&turn"
      Height          =   855
      Left            =   1440
      Picture         =   "RETURN.frx":35B2
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Borrower"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   3975
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Aquisition No:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image Image6 
         Height          =   1020
         Left            =   0
         Picture         =   "RETURN.frx":39F4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   8115
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5520
         TabIndex        =   23
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5520
         TabIndex        =   22
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5520
         TabIndex        =   21
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5520
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5520
         TabIndex        =   19
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date:"
         Height          =   255
         Index           =   11
         Left            =   4650
         TabIndex        =   10
         Top             =   2040
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Edition Date:"
         Height          =   255
         Index           =   10
         Left            =   4425
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   255
         Index           =   9
         Left            =   4785
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Call No:"
         Height          =   375
         Index           =   8
         Left            =   4785
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Title of the book:"
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   6
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tell:"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   5
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   4
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yr / Section:"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "I.D. No:"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.Image Image5 
         Height          =   2820
         Left            =   0
         Picture         =   "RETURN.frx":4AF0E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8115
      End
   End
   Begin VB.Image Image8 
      Height          =   705
      Left            =   0
      Picture         =   "RETURN.frx":92428
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2385
   End
   Begin VB.Image Image7 
      Height          =   735
      Left            =   120
      Picture         =   "RETURN.frx":98CA2
      Top             =   840
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   0
      Picture         =   "RETURN.frx":9C650
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   9570
   End
   Begin VB.Image Image3 
      Height          =   795
      Left            =   0
      Picture         =   "RETURN.frx":A1CDE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9570
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5640
      Picture         =   "RETURN.frx":A736C
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   31
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   30
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   5760
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image4 
      Height          =   6255
      Left            =   0
      Picture         =   "RETURN.frx":A77AE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   9450
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As Database
Dim tn, tn1, tn2, tn3 As Recordset

Option Compare Text




Private Sub Form_Load()
Set dn = OpenDatabase(App.Path + "\SADconvert.mdb")
Set tn = dn.OpenRecordset("borrowed material", dbOpenDynaset)
Set tn1 = dn.OpenRecordset("student table", dbOpenDynaset)
Set tn2 = dn.OpenRecordset("book table", dbOpenDynaset)
Set tn3 = dn.OpenRecordset("damage bok", dbOpenDynaset)
End Sub
'==================================================command
Private Sub CMDclose_Click()
Form1.Show
Unload Me
End Sub
Private Sub CMDview_Click()
Form7.Show
Unload Me
End Sub


Private Sub CDMrefresh_Click()
Text1.SetFocus
SendKeys "{Home}+{End}"
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
End Sub

Private Sub CMDreturn_Click()
Dim sw3, sw4 As Boolean
With tn2
 .MoveFirst
   Do Until .EOF
    If Label7.Caption = !TITLEOFTHEBOOK Then
     sw3 = True
   Exit Do
    End If
  .MoveNext
    Loop
      
      
      If sw3 = True Then
        .Edit
         !Status = "Available"
        .Update
       End If
End With



With tn

 .MoveFirst
   Do Until .EOF
    If Label2.Caption = !STUDENTNAME Then
     sw4 = True
   Exit Do
    End If
  .MoveNext
    Loop
   
      If sw4 = True Then
        .Edit
        .delete
        MsgBox "The book was returned!!", vbInformation, "Returned"
       End If
   
End With

End Sub

Private Sub CMDdamage_Click()
If Label2.Caption <> "" Then
    ans = MsgBox("This will save in the damage book inventory? Do you want to proceed?", vbOKCancel + vbExclamation, "DAMAGE BOOK")
     If ans = vbOK Then
       With tn3
         .AddNew
         !STUDENTNAME = Label2.Caption
         !IDNO = Label3.Caption
         !YRSEC = Label4.Caption
         !ADDRESS = Label5.Caption
         !TELL = Label6.Caption
         !TITLEOFTHEBOOK = Label7.Caption
         !AQUISITIONNO = Label8.Caption
         !AUTHOR = Label9.Caption
        .Update
       MsgBox "Save"
     End With
      
      tn.Edit
      tn.delete
      
     
     
    End If
 Else
  MsgBox "No Student Name Found!!!"
 End If
End Sub

Private Sub CMDsearch_Click()
Dim sw, sw1, sw2 As Boolean
CDMrefresh.Value = True

tn.MoveFirst
 Do Until tn.EOF
  If Text1.Text = tn!IDNO Then
    sw = True
    Exit Do
   End If
  tn.MoveNext
 Loop
    
    If sw = True Then
     Label7.Caption = tn!TITLEOFTHEBOOK
     Label11.Caption = tn!duedate
     Label8.Caption = tn!AQUISITIONNO
     
     tn1.MoveFirst
      Do Until tn1.EOF
        If Text1.Text = tn1!IDNO Then
         sw1 = True
      Exit Do
        End If
     tn1.MoveNext
      Loop
       If sw1 = True Then
         Label2.Caption = tn1!STUDENTNAME
         Label3.Caption = tn1!IDNO
         Label4.Caption = tn1!YRSEC
         Label5.Caption = tn1!ADDRESS
         Label6.Caption = tn1!TELL
       End If
     Else
      MsgBox "No current Record"
     End If
     
     tn2.MoveFirst
      Do Until tn2.EOF
       If Label7.Caption = tn2!TITLEOFTHEBOOK Then
        sw2 = True
      Exit Do
       End If
     tn2.MoveNext
      Loop
        If sw2 = True Then
         Label8.Caption = tn2!AQUISITIONNO
         Label9.Caption = tn2!AUTHOR
         Label10.Caption = tn2!EDITIONDATE
        End If
        
  
End Sub

Private Sub Timer1_Timer()
Label12.Caption = Format(Date, "mm/dd/yy")
Label13.Caption = Time

End Sub
