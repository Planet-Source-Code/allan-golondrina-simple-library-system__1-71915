VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DONATION 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Catalogue"
   ClientHeight    =   7635
   ClientLeft      =   2160
   ClientTop       =   810
   ClientWidth     =   11865
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
   ScaleHeight     =   7635
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CDMrefresh 
      Caption         =   "&Refresh"
      Height          =   855
      Left            =   3240
      Picture         =   "REGEST.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   8760
      TabIndex        =   19
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   8760
      TabIndex        =   18
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   8760
      TabIndex        =   17
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton CMDedit 
      Caption         =   "&Edit"
      Height          =   855
      Left            =   4200
      Picture         =   "REGEST.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton CMDsave 
      Caption         =   "S&ave"
      Height          =   855
      Left            =   5040
      Picture         =   "REGEST.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Delete"
      Height          =   855
      Left            =   5880
      Picture         =   "REGEST.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton CMDclose 
      Caption         =   "&Close"
      Height          =   855
      Left            =   6720
      Picture         =   "REGEST.frx":0B18
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5400
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5400
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox text5 
      ForeColor       =   &H80000001&
      Height          =   360
      ItemData        =   "REGEST.frx":0F5A
      Left            =   5400
      List            =   "REGEST.frx":0F64
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid FG3 
      Height          =   3375
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   10
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
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3600
      Picture         =   "REGEST.frx":0F82
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4365
   End
   Begin VB.Image Image2 
      Height          =   1395
      Left            =   -120
      Picture         =   "REGEST.frx":865C
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   12090
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "REGEST.frx":DCEA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11850
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Donated By:"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   22
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy:"
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date:"
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   20
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      Height          =   255
      Index           =   9
      Left            =   1185
      TabIndex        =   12
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Call No:"
      Height          =   375
      Index           =   8
      Left            =   1185
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title of the book:"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Index           =   11
      Left            =   4785
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher:"
      Height          =   255
      Index           =   10
      Left            =   4545
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Edition Date:"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   7
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Image Image5 
      Height          =   5700
      Left            =   0
      Picture         =   "REGEST.frx":13378
      Stretch         =   -1  'True
      Top             =   840
      Width           =   11835
   End
End
Attribute VB_Name = "DONATION"
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

Private Sub Command1_Click()


End Sub

Private Sub CDMrefresh_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
'text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text1.Refresh
Text2.Refresh
Text3.Refresh
Text4.Refresh
text5.Refresh
Text6.Refresh
Text7.Refresh
Text8.Refresh
End Sub

Private Sub Form_Load()
Set dn = OpenDatabase(App.Path + "\SADconvert.mdb")
Set tn = dn.OpenRecordset("donation", dbOpenDynaset)
Set tn1 = dn.OpenRecordset("BOOK TABLE", dbOpenDynaset)
'Set tn1 = dn.OpenRecordset("studenttable", dbOpenDynaset)

Refreshable
'ParasaStudent
End Sub
'==========================================tools
Public Sub Refreshable()


On Error Resume Next
FG3.Clear

FG3.TextMatrix(0, 1) = "Title of the book"
FG3.TextMatrix(0, 2) = "Aqui. No"
FG3.TextMatrix(0, 3) = "author"
FG3.TextMatrix(0, 4) = "Publisher"
FG3.TextMatrix(0, 5) = "Status"
FG3.TextMatrix(0, 6) = "Edit.Date"
FG3.TextMatrix(0, 7) = "Due Date"
FG3.TextMatrix(0, 8) = "Copy"
FG3.TextMatrix(0, 9) = "Donated By"

FG3.ColWidth(0) = 200
FG3.ColWidth(1) = 2000
FG3.ColWidth(2) = 1000
FG3.ColWidth(3) = 2000
FG3.ColWidth(4) = 2000
FG3.ColWidth(5) = 2000
FG3.ColWidth(6) = 1000
FG3.ColWidth(7) = 1000
FG3.ColWidth(8) = 900
FG3.ColWidth(9) = 1000
FG3.ColAlignment(2) = 4


 tn.MoveFirst
   FG3.TextMatrix(1, 1) = tn!TITLEOFTHEBOOK
   FG3.TextMatrix(1, 2) = tn!AQUISITIONNO
   FG3.TextMatrix(1, 3) = tn!AUTHOR
   FG3.TextMatrix(1, 4) = tn!PUBLISHER
   FG3.TextMatrix(1, 5) = tn!Status
   FG3.TextMatrix(1, 6) = tn!EDITIONDATE
   FG3.TextMatrix(1, 7) = tn!duedate
   FG3.TextMatrix(1, 8) = tn!Copy
   FG3.TextMatrix(1, 9) = tn!donatedby

 tn.MoveNext
 
 Do Until tn.EOF
  a = a + 1
  
   FG3.AddItem vbTab & tn!TITLEOFTHEBOOK & vbTab & tn!AQUISITIONNO & vbTab & _
    tn!AUTHOR & vbTab & tn!PUBLISHER & vbTab & tn!Status & vbTab & _
    tn!EDITIONDATE & vbTab & tn!duedate & vbTab & tn!Copy & vbTab & tn!donatedby
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

'ito
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
               !duedate = Text7.Text
               !Copy = Text8.Text
               !donatedby = Text9.Text
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
                !duedate = Text7.Text
               !Copy = Text8.Text
               !donatedby = Text9.Text
              .Update
              Clean
             End If
           End If
End With
End Sub


'ito
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
               !duedate = Text7.Text
               !Copy = Text8.Text
               !donatedby = Text9.Text
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
                !duedate = Text7.Text
               !Copy = Text8.Text
               !donatedby = Text9.Text
              .Update
              Clean
             End If
           End If
 End With
Else
 MsgBox "You have to complete the information to save these book!", vbInformation + vbOKOnly, "Saving"
End If


''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''
'''''''''''SAVE TO BOOK TABLE



Dim sw2 As Boolean
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text6.Text <> "" Then

With tn1
 .MoveFirst
  Do Until .EOF
   If Text2.Text = !AQUISITIONNO Then
     sw2 = True
  Exit Do
   End If
 .MoveNext
  Loop



       If sw2 = True Then
        reply1 = MsgBox(" There is a current record in the data base! Do you want to edit this entry??", vbYesNo + vbQuestion, "SAVE")
          If reply1 = vbYes Then
          
            .Edit
               !TITLEOFTHEBOOK = Text1.Text
               !AQUISITIONNO = Text2.Text
               !AUTHOR = Text3.Text
               !PUBLISHER = Text4.Text
               !Status = text5.Text
               !EDITIONDATE = Text6.Text
               !duedate = Text7.Text
               !Copy = Text8.Text
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
                !duedate = Text7.Text
               !Copy = Text8.Text
              
              .Update
              Clean
             End If
           End If
 End With
Else
 MsgBox "You have to complete the information to save these book!", vbInformation + vbOKOnly, "Saving"
End If
End Sub


'ito
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
    Text7.Text = !duedate
    Text8.Text = !Copy
    Text9.Text = !donatedby
    

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


















