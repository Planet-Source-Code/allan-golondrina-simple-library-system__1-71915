VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Catalogue"
   ClientHeight    =   8145
   ClientLeft      =   2160
   ClientTop       =   810
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   8145
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13361
      _Version        =   393216
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
      TabCaption(0)   =   "Inventory Selection"
      TabPicture(0)   =   "Catalogue.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(11)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(10)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(9)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Image6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CMDedit"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CMDsave"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "FG3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "text5"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CMDclose"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CMDrefresh"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "Catalogue.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image3"
      Tab(1).Control(1)=   "Label1(6)"
      Tab(1).Control(2)=   "Image1"
      Tab(1).Control(3)=   "Text7"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Command6"
      Tab(1).Control(6)=   "FG2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Student Section"
      TabPicture(2)   =   "Catalogue.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image4"
      Tab(2).Control(1)=   "Label1(4)"
      Tab(2).Control(2)=   "Label1(3)"
      Tab(2).Control(3)=   "Label1(2)"
      Tab(2).Control(4)=   "Label1(1)"
      Tab(2).Control(5)=   "Label1(0)"
      Tab(2).Control(6)=   "Image2"
      Tab(2).Control(7)=   "FG1"
      Tab(2).Control(8)=   "Text12"
      Tab(2).Control(9)=   "Text11"
      Tab(2).Control(10)=   "Text10"
      Tab(2).Control(11)=   "Text9"
      Tab(2).Control(12)=   "Text8"
      Tab(2).Control(13)=   "CMDedit1"
      Tab(2).Control(14)=   "CMDsave1"
      Tab(2).Control(15)=   "CMDdelete1"
      Tab(2).Control(16)=   "CMDclose1"
      Tab(2).ControlCount=   17
      Begin VB.CommandButton CMDrefresh 
         Caption         =   "&Refresh"
         Height          =   855
         Left            =   2880
         Picture         =   "Catalogue.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton CMDclose1 
         Caption         =   "&Close"
         Height          =   855
         Left            =   -68640
         Picture         =   "Catalogue.frx":019E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   6480
         Width           =   855
      End
      Begin VB.CommandButton CMDdelete1 
         Caption         =   "&Delete"
         Height          =   855
         Left            =   -69480
         Picture         =   "Catalogue.frx":05E0
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6480
         Width           =   855
      End
      Begin VB.CommandButton CMDsave1 
         Caption         =   "S&ave"
         Height          =   855
         Left            =   -70320
         Picture         =   "Catalogue.frx":0A22
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   6480
         Width           =   855
      End
      Begin VB.CommandButton CMDedit1 
         Caption         =   "&Edit"
         Height          =   855
         Left            =   -71160
         Picture         =   "Catalogue.frx":0E64
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   6480
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid FG2 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   36
         Top             =   2880
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
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
      Begin VB.CommandButton Command6 
         Caption         =   "&Ok"
         Height          =   975
         Left            =   -72720
         Picture         =   "Catalogue.frx":0FAE
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search By:"
         Height          =   1455
         Left            =   -70320
         TabIndex        =   30
         Top             =   840
         Width           =   3015
         Begin VB.OptionButton Option3 
            Caption         =   "Author"
            Height          =   240
            Left            =   600
            TabIndex        =   33
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Aquisition No."
            Height          =   240
            Left            =   600
            TabIndex        =   32
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Title of the book"
            Height          =   240
            Left            =   600
            TabIndex        =   31
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -73080
         TabIndex        =   29
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -72960
         TabIndex        =   23
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -72960
         TabIndex        =   22
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -72960
         TabIndex        =   21
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -68880
         TabIndex        =   20
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -68880
         TabIndex        =   19
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton CMDclose 
         Caption         =   "&Close"
         Height          =   855
         Left            =   6360
         Picture         =   "Catalogue.frx":13F0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6480
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Delete"
         Height          =   855
         Left            =   5520
         Picture         =   "Catalogue.frx":1832
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6480
         Width           =   855
      End
      Begin VB.ComboBox text5 
         Height          =   360
         ItemData        =   "Catalogue.frx":1C74
         Left            =   6120
         List            =   "Catalogue.frx":1C7E
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1080
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid FG3 
         Height          =   3735
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   10095
         _ExtentX        =   17806
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
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6120
         TabIndex        =   12
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton CMDsave 
         Caption         =   "S&ave"
         Height          =   855
         Left            =   4680
         Picture         =   "Catalogue.frx":1C9C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6480
         Width           =   855
      End
      Begin VB.CommandButton CMDedit 
         Caption         =   "&Edit"
         Height          =   855
         Left            =   3840
         Picture         =   "Catalogue.frx":20DE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6120
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   18
         Top             =   2400
         Width           =   10095
         _ExtentX        =   17806
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
      Begin VB.Image Image6 
         Height          =   1515
         Left            =   0
         Picture         =   "Catalogue.frx":2228
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   10890
      End
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   -75120
         Picture         =   "Catalogue.frx":78B6
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   10890
      End
      Begin VB.Image Image2 
         Height          =   1515
         Left            =   -75120
         Picture         =   "Catalogue.frx":CF44
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   10890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
         Height          =   255
         Index           =   6
         Left            =   -73920
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "I.D. No."
         Height          =   255
         Index           =   0
         Left            =   -73800
         TabIndex        =   28
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name:"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   27
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yr / Section:"
         Height          =   255
         Index           =   2
         Left            =   -74160
         TabIndex        =   26
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Index           =   3
         Left            =   -69720
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tell:"
         Height          =   255
         Index           =   4
         Left            =   -69360
         TabIndex        =   24
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Edition Date:"
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   14
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Title of the book:"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Call No:"
         Height          =   375
         Index           =   8
         Left            =   1305
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   255
         Index           =   9
         Left            =   1305
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher:"
         Height          =   255
         Index           =   10
         Left            =   5145
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   255
         Index           =   11
         Left            =   5505
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   5820
         Left            =   -75000
         Picture         =   "Catalogue.frx":125D2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   10635
      End
      Begin VB.Image Image3 
         Height          =   6300
         Left            =   -75000
         Picture         =   "Catalogue.frx":59AEC
         Stretch         =   -1  'True
         Top             =   360
         Width           =   10635
      End
      Begin VB.Image Image5 
         Height          =   5820
         Left            =   0
         Picture         =   "Catalogue.frx":A1006
         Stretch         =   -1  'True
         Top             =   360
         Width           =   10635
      End
   End
   Begin VB.Image Image8 
      Height          =   690
      Left            =   120
      Picture         =   "Catalogue.frx":E8520
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image7 
      Height          =   675
      Left            =   0
      Picture         =   "Catalogue.frx":EC6CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10650
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As Database
Dim tn As Recordset
Dim tn1 As Recordset
Dim tn2 As Recordset
Dim a, b, C
Dim category


Option Compare Text

Private Sub CMDclose1_Click()
Form1.Show
Unload Me
End Sub

Private Sub CMDrefresh_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text1.Refresh
Text2.Refresh
Text3.Refresh
Text4.Refresh
Text6.Refresh

End Sub

Private Sub Form_Load()
Set dn = OpenDatabase(App.Path + "\SADconvert.mdb")
Set tn = dn.OpenRecordset("book table", dbOpenDynaset)
Set tn1 = dn.OpenRecordset("student table", dbOpenDynaset)

Refreshable
ParasaStudent
End Sub
'==========================================tools
Public Sub Refreshable()
FG2.TextMatrix(0, 1) = "Title of the book"
FG2.TextMatrix(0, 2) = "Aqui. No"
FG2.TextMatrix(0, 3) = "author"
FG2.TextMatrix(0, 4) = "Publisher"
FG2.TextMatrix(0, 5) = "Status"
FG2.TextMatrix(0, 6) = "Edit.Date"
FG2.TextMatrix(0, 7) = "Due Date"

FG2.ColWidth(0) = 200
FG2.ColWidth(1) = 2000
FG2.ColWidth(2) = 1000
FG2.ColWidth(3) = 2000
FG2.ColWidth(4) = 2000
FG2.ColWidth(5) = 2000
FG2.ColWidth(6) = 1000
FG2.ColWidth(7) = 1000
FG2.ColAlignment(2) = 4

On Error Resume Next
FG3.Clear

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
FG3.ColWidth(3) = 2000
FG3.ColWidth(4) = 2000
FG3.ColWidth(5) = 2000
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
  a = a + 1
  
   FG3.AddItem vbTab & tn!TITLEOFTHEBOOK & vbTab & tn!AQUISITIONNO & vbTab & _
    tn!AUTHOR & vbTab & tn!PUBLISHER & vbTab & tn!Status & vbTab & _
    tn!EDITIONDATE & vbTab & tn!duedate
tn.MoveNext
 Loop

End Sub



Public Sub Clean()
FG3.Clear
Do Until a = 0
a = a - 1
FG3.RemoveItem FG3.Row
Loop
Refreshable

End Sub



'========================================================comand


Private Sub CMDclose_Click()
Form1.Show
Form1.Enabled = True
Unload Me
End Sub

Private Sub CMDedit_Click()
Dim sw As Boolean

With tn
 .MoveFirst
  Do Until .EOF
   If Text2.Text = !AQUISITIONNO Then
     sw = True
  Exit Do
   End If
 .MoveNext
  Loop
   
       If sw = True Then
        reply = MsgBox(" Do you want to edit this record", vbYesNo + vbQuestion, "SAVE")
          If reply = vbYes Then
          
            .Edit
               !TITLEOFTHEBOOK = Text1.Text
               !AQUISITIONNO = Text2.Text
               !AUTHOR = Text3.Text
               !PUBLISHER = Text4.Text
               !Status = text5.Text
               !EDITIONDATE = Text6.Text
            .Update
           Clean
            End If

         Else
         ans = MsgBox("There is no record found in the data base!! Do you want to save this as a new record", vbYesNo + vbQuestion, "SAVE")
         
           If ans = vbYes Then
             .AddNew
                !TITLEOFTHEBOOK = Text1.Text
                !AQUISITIONNO = Text2.Text
                !AUTHOR = Text3.Text
                !PUBLISHER = Text4.Text
                !Status = text5.Text
                !EDITIONDATE = Text6.Text
              .Update
              Clean
             End If
           End If
End With
End Sub

Private Sub CMDsave_Click()
Dim sw1 As Boolean
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text6.Text <> "" Then

With tn
 .MoveFirst
  Do Until .EOF
   If Text2.Text = !AQUISITIONNO Then
     sw1 = True
  Exit Do
   End If
 .MoveNext
  Loop



       If sw1 = True Then
        reply1 = MsgBox(" There is a current record in the data base! Do you want to edit this entry??", vbYesNo + vbQuestion, "SAVE")
          If reply1 = vbYes Then
          
            .Edit
               !TITLEOFTHEBOOK = Text1.Text
               !AQUISITIONNO = Text2.Text
               !AUTHOR = Text3.Text
               !PUBLISHER = Text4.Text
               !Status = text5.Text
               !EDITIONDATE = Text6.Text
            .Update
             Clean
            End If

         Else
         ans1 = MsgBox("Do you want to save this as a new record", vbYesNo + vbQuestion, "SAVE")
         
           If ans1 = vbYes Then
             .AddNew
                !TITLEOFTHEBOOK = Text1.Text
                !AQUISITIONNO = Text2.Text
                !AUTHOR = Text3.Text
                !PUBLISHER = Text4.Text
                !Status = text5.Text
                !EDITIONDATE = Text6.Text
              .Update
              Clean
             End If
           End If
 End With
Else
 MsgBox "You have to complete the information to save these book!", vbInformation + vbOKOnly, "Saving"
End If
End Sub


Private Sub Command5_Click()

With tn
 .MoveFirst
  Do Until .EOF
   If Text1.Text = !TITLEOFTHEBOOK Then
     sw1 = True
  Exit Do
   End If
 .MoveNext
  Loop

       If sw1 = True Then
        reply3 = MsgBox("  Do you want to delete this record??", vbYesNo + vbQuestion, "SAVE")
          If reply3 = vbYes Then
          
            .Edit
            .delete
                Clean
            End If
         Else
            MsgBox "There is no record found !!", vbOKOnly + vbExclamation, "Deletion"
        End If
        
        
End With

End Sub



'=================================================text

Private Sub FG3_Click()
Dim sw5 As Boolean
Text1.Text = FG3.Text
On Error Resume Next



With tn
 .MoveFirst
  Do Until .EOF
   If Text1.Text = !TITLEOFTHEBOOK Then
     sw5 = True
  Exit Do
   End If
 .MoveNext
  Loop
   
   
   
   If Text1.Text = !TITLEOFTHEBOOK Then
    Text2.Text = !AQUISITIONNO
    Text3.Text = !AUTHOR
    Text4.Text = !PUBLISHER
    text5.Text = !Status
    Text6.Text = !EDITIONDATE

   End If
End With

End Sub

'====================================================================================tab2
Private Sub Command6_Click()
Dim sw As Boolean

FG2.Clear

With tn
  Select Case category
    Case 1
      .MoveFirst
        Do Until .EOF
          If Text7.Text = !TITLEOFTHEBOOK Then
            sw = True
         Exit Do
           End If
           
      .MoveNext
        Loop
               
               
     Case 2
      .MoveFirst
        Do Until .EOF
          If Text7.Text = !AQUISITIONNO Then
            sw = True
         Exit Do
           End If
           
      .MoveNext
        Loop
        
     Case 3
      .MoveFirst
        Do Until .EOF
          If Text7.Text = !AUTHOR Then
            sw = True
         Exit Do
           End If
           
      .MoveNext
        Loop

End Select
       
       If sw = True Then
            FG2.AddItem vbTab & tn!TITLEOFTHEBOOK & vbTab & tn!AQUISITIONNO & vbTab & _
            tn!AUTHOR & vbTab & tn!PUBLISHER & vbTab & tn!Status & vbTab & _
            tn!EDITIONDATE & vbTab & tn!duedate
          
       Else
           MsgBox "Record not found!!", vbOKOnly + vbInformation, "Searching"
       End If
End With

End Sub


Private Sub Option1_Click()
category = 1

End Sub
Private Sub Option2_Click()
category = 2
End Sub
Private Sub Option3_Click()
category = 3
End Sub





'==============================================================Student Sectiom







Public Sub ParasaStudent()
FG1.Clear

FG1.TextMatrix(0, 1) = "ID No"
FG1.TextMatrix(0, 2) = "Student Name"
FG1.TextMatrix(0, 3) = "Year / Sec"
FG1.TextMatrix(0, 4) = "Address"
FG1.TextMatrix(0, 5) = "Tell No"

FG1.ColWidth(0) = 200
FG1.ColWidth(1) = 1000
FG1.ColWidth(2) = 2000
FG1.ColWidth(3) = 1000
FG1.ColWidth(4) = 3000
FG1.ColWidth(5) = 1000
FG1.ColAlignment(2) = 4


 tn1.MoveFirst
   FG1.TextMatrix(1, 1) = tn1!IDNO
   FG1.TextMatrix(1, 2) = tn1!STUDENTNAME
   FG1.TextMatrix(1, 3) = tn1!YRSEC
   FG1.TextMatrix(1, 4) = tn1!ADDRESS
   FG1.TextMatrix(1, 5) = tn1!TELL

 tn1.MoveNext
 
 Do Until tn1.EOF
  b = b + 1
  
   FG1.AddItem vbTab & tn1!IDNO & vbTab & tn1!STUDENTNAME & vbTab & _
    tn1!YRSEC & vbTab & tn1!ADDRESS & vbTab & tn1!TELL
tn1.MoveNext
 Loop
End Sub

Private Sub FG1_Click()
Dim sw4 As Boolean
Text8.Text = FG1.Text
On Error Resume Next



With tn1
 .MoveFirst
  Do Until .EOF
   If Text8.Text = !IDNO Then
     sw4 = True
  Exit Do
   End If
 .MoveNext
  Loop
   
   
   
   If Text8.Text = !IDNO Then
    Text9.Text = !STUDENTNAME
    Text10.Text = !YRSEC
    Text12.Text = !ADDRESS
    Text11.Text = !TELL
   End If
End With

End Sub



Private Sub CMDedit1_Click()
Dim sw2 As Boolean

With tn1
 .MoveFirst
  Do Until .EOF
   If Text8.Text = !IDNO Then
     sw2 = True
  Exit Do
   End If
 .MoveNext
  Loop
   
       If sw2 = True Then
        reply = MsgBox(" Do you want to edit this record", vbYesNo + vbQuestion, "SAVE")
          If reply = vbYes Then
          
            .Edit
               !IDNO = Text8.Text
               !STUDENTNAME = Text9.Text
               !YRSEC = Text10.Text
               !ADDRESS = Text12.Text
               !TELL = Text11.Text
               
            .Update
            cleanforstud
           
            End If

         Else
         ans = MsgBox("There is no record found in the data base!! Do you want to save this as a new record", vbYesNo + vbQuestion, "SAVE")
         
           If ans = vbYes Then
             .AddNew
               !IDNO = Text8.Text
               !STUDENTNAME = Text9.Text
               !YRSEC = Text10.Text
               !ADDRESS = Text12.Text
               !TELL = Text11.Text
              .Update
               
             End If
           End If
End With

End Sub



Private Sub CMDsave1_Click()
Dim sw2 As Boolean
If Text8.Text <> "" And Text9.Text <> "" And Text10.Text <> "" And Text11.Text <> "" Then

With tn1
 .MoveFirst
  Do Until .EOF
   If Text8.Text = !IDNO Then
     sw2 = True
  Exit Do
   End If
 .MoveNext
  Loop



       If sw2 = True Then
        reply1 = MsgBox(" There is a current record in the data base! Do you want to edit this entry??", vbYesNo + vbQuestion, "SAVE")
          If reply1 = vbYes Then
          
            .Edit
               !IDNO = Text8.Text
               !STUDENTNAME = Text9.Text
               !YRSEC = Text10.Text
               !ADDRESS = Text12.Text
               !TELL = Text11.Text
            .Update
            cleanforstud
             
            End If

         Else
         ans1 = MsgBox("Do you want to save this as a new record", vbYesNo + vbQuestion, "SAVE")
         
           If ans1 = vbYes Then
             .AddNew
               !IDNO = Text8.Text
               !STUDENTNAME = Text9.Text
               !YRSEC = Text10.Text
               !ADDRESS = Text12.Text
               !TELL = Text11.Text
              .Update
              cleanforstud
             End If
           End If
 End With
Else
 MsgBox "You have to complete the information to save these book!", vbInformation + vbOKOnly, "Saving"
End If


End Sub


Private Sub CMDdelete1_Click()
Dim sw2 As Boolean
With tn1
 .MoveFirst
  Do Until .EOF
   If Text8.Text = !IDNO Then
     sw2 = True
  Exit Do
   End If
 .MoveNext
  Loop

       If sw2 = True Then
        reply3 = MsgBox("  Do you want to delete this record??", vbYesNo + vbQuestion, "SAVE")
          If reply3 = vbYes Then
          
            .Edit
            .delete
            cleanforstud
                
            End If
         Else
            MsgBox "There is no record found !!", vbOKOnly + vbExclamation, "Deletion"
        End If
        
        
End With

End Sub


Public Sub cleanforstud()
FG1.Clear
Do Until b = 0
b = b - 1
FG1.RemoveItem FG1.Row
Loop
ParasaStudent

End Sub
