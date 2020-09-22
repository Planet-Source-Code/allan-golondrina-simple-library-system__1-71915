VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00C00000&
   Caption         =   "Inventory"
   ClientHeight    =   6975
   ClientLeft      =   2355
   ClientTop       =   1560
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   6975
   ScaleWidth      =   10680
   Begin TabDlg.SSTab Tab 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11880
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Borrowed Books / Materials"
      TabPicture(0)   =   "inventory.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FG1"
      Tab(0).Control(1)=   "Image5"
      Tab(0).Control(2)=   "Image2"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Damage Books / Materials"
      TabPicture(1)   =   "inventory.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FG2"
      Tab(1).Control(1)=   "Image6"
      Tab(1).Control(2)=   "Image1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "All Books / Materials"
      TabPicture(2)   =   "inventory.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Image4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Image7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "FG3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   1
         Top             =   1680
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid FG2 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   2
         Top             =   1680
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid FG3 
         Height          =   3735
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image7 
         Height          =   1095
         Left            =   0
         Picture         =   "inventory.frx":0054
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   10770
      End
      Begin VB.Image Image6 
         Height          =   1095
         Left            =   -75000
         Picture         =   "inventory.frx":56E2
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   10770
      End
      Begin VB.Image Image5 
         Height          =   1095
         Left            =   -75000
         Picture         =   "inventory.frx":AD70
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   10770
      End
      Begin VB.Image Image2 
         Height          =   6375
         Left            =   -75000
         Picture         =   "inventory.frx":103FE
         Stretch         =   -1  'True
         Top             =   360
         Width           =   10650
      End
      Begin VB.Image Image1 
         Height          =   6375
         Left            =   -75000
         Picture         =   "inventory.frx":65058
         Stretch         =   -1  'True
         Top             =   360
         Width           =   10650
      End
      Begin VB.Image Image4 
         Height          =   6375
         Left            =   0
         Picture         =   "inventory.frx":B9CB2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   10650
      End
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      Picture         =   "inventory.frx":10E90C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10770
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "&File"
      Begin VB.Menu MenScren 
         Caption         =   "&Menu Screen"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Borrow 
         Caption         =   "&Borrowing "
         Shortcut        =   {F2}
      End
      Begin VB.Menu Return 
         Caption         =   "&Returning"
         Shortcut        =   {F3}
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As Database
Dim tn, tn1, tn2 As Recordset
Dim a




Private Sub Form_Load()
On Error Resume Next
Dim a As Integer
Set dn = OpenDatabase(App.Path + "\SADconvert.mdb")
Set tn = dn.OpenRecordset("book table", dbOpenDynaset)
Set tn1 = dn.OpenRecordset("damage bok", dbOpenDynaset)
Set tn2 = dn.OpenRecordset("borrowed material", dbOpenDynaset)



FG3.TextMatrix(0, 1) = "Title of the book"
FG3.TextMatrix(0, 2) = "Aqui. No"
FG3.TextMatrix(0, 3) = "author"
FG3.TextMatrix(0, 4) = "Publisher"
FG3.TextMatrix(0, 5) = "Status"
FG3.TextMatrix(0, 6) = "Edit.Date"
FG3.TextMatrix(0, 7) = "Due Date"

FG3.ColWidth(0) = 200
FG3.ColWidth(1) = 2000
FG3.ColWidth(2) = 1000
FG3.ColWidth(3) = 1700
FG3.ColWidth(4) = 2000
FG3.ColWidth(5) = 1300
FG3.ColWidth(6) = 1000
FG3.ColWidth(7) = 1000
FG3.ColAlignment(2) = 4


tn.MoveFirst
   FG3.TextMatrix(1, 1) = tn!TITLEOFTHEBOOK
   FG3.TextMatrix(1, 2) = tn!AQUISITIONNO
   FG3.TextMatrix(1, 3) = tn!AUTHOR
   FG3.TextMatrix(1, 4) = tn!PUBLISHER
   FG3.TextMatrix(1, 5) = tn!Status
   FG3.TextMatrix(1, 6) = tn!EDITIONDATE
   FG3.TextMatrix(1, 7) = tn!duedate
 tn.MoveNext
 
 Do Until tn.EOF
    FG3.AddItem vbTab & tn!TITLEOFTHEBOOK & vbTab & tn!AQUISITIONNO & vbTab & _
    tn!AUTHOR & vbTab & tn!PUBLISHER & vbTab & tn!Status & vbTab & _
    tn!EDITIONDATE & vbTab & tn!duedate
tn.MoveNext
 Loop
Form7.Caption = tn1.EOF







FG2.TextMatrix(0, 1) = "Student Name"
FG2.TextMatrix(0, 2) = "I.D. No."
FG2.TextMatrix(0, 3) = "YR &Sec"
FG2.TextMatrix(0, 4) = "Address"
FG2.TextMatrix(0, 5) = "Tell"
FG2.TextMatrix(0, 6) = "Title of the Book"
FG2.TextMatrix(0, 7) = "Aquisition"
FG2.TextMatrix(0, 8) = "Author"
FG2.TextMatrix(0, 9) = "Publisher"

FG2.ColWidth(0) = 200
FG2.ColWidth(1) = 2000
FG2.ColWidth(2) = 1000
FG2.ColWidth(3) = 2000
FG2.ColWidth(4) = 3000
FG2.ColWidth(5) = 1000
FG2.ColWidth(6) = 2000
FG2.ColWidth(7) = 1000
FG2.ColWidth(8) = 1000
FG2.ColWidth(9) = 1000
FG2.ColAlignment(2) = 4


tn1.MoveFirst
   FG2.TextMatrix(1, 1) = tn1!STUDENTNAME
   FG2.TextMatrix(1, 2) = tn1!IDNO
   FG2.TextMatrix(1, 3) = tn1!YRSEC
   FG2.TextMatrix(1, 4) = tn1!ADDRESS
   FG2.TextMatrix(1, 5) = tn1!TELL
   FG2.TextMatrix(1, 6) = tn1!TITLEOFTHEBOOK
   FG2.TextMatrix(1, 7) = tn1!AQUISITIONNO
   FG2.TextMatrix(1, 8) = tn1!AUTHOR
   FG2.TextMatrix(1, 9) = tn1!PUBLISHER
   
 tn1.MoveNext
 
 Do Until tn1.EOF
    FG2.AddItem vbTab & tn1!STUDENTNAME & vbTab & tn1!IDNO & vbTab & _
    tn1!YRSEC & vbTab & tn1!ADDRESS & vbTab & tn1!TELL & vbTab & _
    tn1!TITLEOFTHEBOOK & vbTab & tn1!AQUISITIONNO & vbTab & tn1!AUTHOR & tn1!PUBLISHER
tn1.MoveNext
 Loop





FG1.TextMatrix(0, 1) = "Student name"
FG1.TextMatrix(0, 2) = "I.D No"
FG1.TextMatrix(0, 3) = "Title of the Book"
FG1.TextMatrix(0, 4) = "Aquisition No"
FG1.TextMatrix(0, 5) = "Due Date"

FG1.ColWidth(0) = 200
FG1.ColWidth(1) = 2000
FG1.ColWidth(2) = 1000
FG1.ColWidth(3) = 2500
FG1.ColWidth(4) = 2500
FG1.ColWidth(5) = 2000
FG1.ColAlignment(2) = 4


tn2.MoveFirst
   FG1.TextMatrix(1, 1) = tn2!STUDENTNAME
   FG1.TextMatrix(1, 2) = tn2!IDNO
   FG1.TextMatrix(1, 3) = tn2!TITLEOFTHEBOOK
   FG1.TextMatrix(1, 4) = tn2!AQUISITIONNO
   FG1.TextMatrix(1, 5) = tn2!duedate
 tn2.MoveNext
 
 Do Until tn2.EOF
    FG1.AddItem vbTab & tn2!STUDENTNAME & vbTab & tn2!IDNO & vbTab & _
    tn2!TITLEOFTHEBOOK & vbTab & tn2!AQUISITIONNO & vbTab & tn2!duedate
tn2.MoveNext
 Loop

End Sub


'===========================================================Menus


Private Sub MenScren_Click()
Form1.Show
Unload Me
End Sub
Private Sub Borrow_Click()
Form2.Show
Unload Me
End Sub

Private Sub Return_Click()
Form3.Show
Unload Me
End Sub
