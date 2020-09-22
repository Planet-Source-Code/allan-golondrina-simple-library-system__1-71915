VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Borrowing Section"
   ClientHeight    =   7050
   ClientLeft      =   3030
   ClientTop       =   1260
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7050
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2175
      Left            =   6600
      TabIndex        =   40
      Top             =   5160
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524288
      _ExtentX        =   4260
      _ExtentY        =   3836
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2002
      Month           =   9
      Day             =   19
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   0
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8880
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   6795
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      SimpleText      =   "klk"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11351
            Text            =   "C.M. Recto High School Library System"
            TextSave        =   "C.M. Recto High School Library System"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "3/28/2009"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "3:12 PM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDClose 
      Caption         =   "&Close"
      Height          =   855
      Left            =   6000
      Picture         =   "borrowing.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton CMDview 
      Caption         =   "&View"
      Height          =   855
      Left            =   5040
      Picture         =   "borrowing.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton CMDrefresh 
      Caption         =   "&Refresh"
      Height          =   855
      Left            =   4080
      Picture         =   "borrowing.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton CMDsearch 
      Caption         =   "S&earch"
      Height          =   855
      Left            =   3120
      Picture         =   "borrowing.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton CMDsave 
      Caption         =   "&Save"
      Height          =   855
      Left            =   2160
      Picture         =   "borrowing.frx":0E10
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Book Information"
      Height          =   4935
      Left            =   4680
      TabIndex        =   27
      Top             =   600
      Width           =   4575
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "borrowing.frx":1252
         Left            =   2160
         List            =   "borrowing.frx":125C
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search"
         Height          =   1335
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   4215
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   7
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Aquisition No:"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   29
            Top             =   600
            Width           =   1215
         End
         Begin VB.Image Image6 
            Height          =   1290
            Left            =   0
            Picture         =   "borrowing.frx":127A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4200
         End
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         TabIndex        =   8
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         TabIndex        =   9
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         TabIndex        =   10
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         TabIndex        =   11
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   4080
         TabIndex        =   39
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         TabIndex        =   13
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Image Image9 
         Height          =   735
         Left            =   240
         Picture         =   "borrowing.frx":152F0
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date: "
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   36
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Title of the book:"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   34
         Top             =   2400
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Aquisition No:"
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   33
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   32
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher"
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   31
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   30
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   4980
         Left            =   0
         Picture         =   "borrowing.frx":1B906
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4635
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Student Information"
      Height          =   4935
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Search"
         Height          =   1335
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   4095
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "I.D. No:"
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   21
            Top             =   600
            Width           =   735
         End
         Begin VB.Image Image7 
            Height          =   1290
            Left            =   0
            Picture         =   "borrowing.frx":62E20
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4080
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         TabIndex        =   2
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         TabIndex        =   3
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         TabIndex        =   4
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         TabIndex        =   5
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Image Image8 
         Height          =   735
         Left            =   240
         Picture         =   "borrowing.frx":76E96
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date and Time:"
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   37
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.D. No:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   25
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Yr / Section:"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   24
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   23
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tell:"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   22
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   38
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Image Image5 
         Height          =   4980
         Left            =   0
         Picture         =   "borrowing.frx":7D4AC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4635
      End
   End
   Begin VB.Image Image10 
      Height          =   825
      Left            =   0
      Picture         =   "borrowing.frx":C49C6
      Top             =   0
      Width           =   3330
   End
   Begin VB.Image Image3 
      Height          =   555
      Left            =   -120
      Picture         =   "borrowing.frx":CD98C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9570
   End
   Begin VB.Image Image2 
      Height          =   1395
      Left            =   0
      Picture         =   "borrowing.frx":D301A
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   9570
   End
   Begin VB.Image Image1 
      Height          =   7455
      Left            =   0
      Picture         =   "borrowing.frx":D86A8
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   9450
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As Database
Dim tn As Recordset
    Dim tn1 As Recordset
    Dim tn2 As Recordset
   
    Dim all, all1
    
    
    Option Compare Text
    
    
    
    
    Private Sub Form_Load()
    Set dn = OpenDatabase(App.Path + "\SADconvert.mdb")
    Set tn = dn.OpenRecordset("book table", dbOpenDynaset)
    Set tn1 = dn.OpenRecordset("student table", dbOpenDynaset)
    Set tn2 = dn.OpenRecordset("borrowed material", dbOpenDynaset)
      
    End Sub

'===================================================Extra
Private Sub Calendar1_Click()

Text12.Text = Calendar1.Value
Calendar1.Visible = False

End Sub


Private Sub Command1_Click()
Calendar1.Visible = True
End Sub




Private Sub Timer1_Timer()
Label3.Caption = Format(Date, "mm/dd/yy")
Label2.Caption = Time
Label2.Caption = Label3.Caption + " , " + Label2.Caption
End Sub


'=====================================================Commnands
Private Sub CMDclose_Click()
Form1.Show
Unload Me
End Sub

Private Sub CMDrefresh_Click()
delete
End Sub


Private Sub CMDview_Click()
Form7.Show
Unload Me
End Sub


Private Sub CMDsearch_Click()
Dim sw As Boolean
Dim sw1 As Boolean
Dim sw2 As Boolean

tn1.MoveFirst
 Do Until tn1.EOF
  If Text2.Text = tn1!IDNO Then
   sw = True
   Exit Do
  End If
tn1.MoveNext
 Loop
  
  If sw = False Then
   MsgBox "No Student Record Found!!!", vbOKOnly + vbInformation, "Search"
   Text6.SetFocus
   SendKeys "{Home}+{End}"
  Else
  all = 1
   With tn1
    Text1.Text = !STUDENTNAME
    Text2.Text = !IDNO
    Text3.Text = !YRSEC
    Text4.Text = !ADDRESS
    Text5.Text = !TELL
   End With
  End If






tn.MoveFirst
 Do Until tn.EOF
  If Text9.Text = tn!AQUISITIONNO Then
   sw1 = True
   
   Exit Do
  End If
tn.MoveNext
 Loop
  
  If sw1 = True Then
  all1 = 1
   With tn
    On Error Resume Next
    Text8.Text = !TITLEOFTHEBOOK
    Text9.Text = !AQUISITIONNO
    Text10.Text = !AUTHOR
    Text11.Text = !PUBLISHER
    Combo1.Text = !Status
    Text12.Text = !duedate
   End With
  Else
   MsgBox "No Book Record Found!!!", vbOKOnly + vbInformation, "Search"
   If all = 1 Then
   Text7.SetFocus
   SendKeys "{Home}+{End}"
   ElseIf all <> 1 And all1 <> 1 Then
   Text6.SetFocus
   SendKeys "{Home}+{End}"
   
   End If

  End If
  
  
  
  
End Sub

Private Sub CMDsave_Click()
Dim sw2, sw1 As Boolean

CMDsearch.Value = True

If Combo1.Text = "Not-Available" Then
 MsgBox "The book is not Available"

 
ElseIf all = 1 And all1 = 1 Then
    tn2.MoveFirst
     Do Until tn2.EOF
      If Text2.Text = tn2!IDNO Then
       sw2 = True
       Exit Do
      End If
    tn2.MoveNext
     Loop
        
      If sw2 = False Then
       If Text12.Text = "" Then
          MsgBox "The due date is not set", vbOKOnly + vbInformation, "Saving"
       Else
            tn2.AddNew
             tn2!STUDENTNAME = Text1.Text
             tn2!IDNO = Text2.Text
             tn2!TITLEOFTHEBOOK = Text8.Text
             tn2!AQUISITIONNO = Text9.Text
             tn2!duedate = Text12.Text
            tn2.Update
            MsgBox "The borrowers information was save!!", vbOKOnly + vbInformation, "Borrowing"
        
        
            tn.MoveFirst
             Do Until tn.EOF
              If Text9.Text = tn!AQUISITIONNO Then
               sw1 = True
               Exit Do
              End If
            tn.MoveNext
             Loop
             
              If sw1 = True Then
               tn.Edit
                tn!Status = "Not-Available"
               tn.Update
              End If
        End If
       
       
      Else
       MsgBox "The borrower has a previos account in borrowing of book. they have to borrow one at a time", vbOKOnly + vbInformation, "Borrowing"
      End If
 
    
       

End If

End Sub


'==============================================others

Private Sub Text6_Change()
Dim sw As Boolean
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
tn1.MoveFirst
 Do Until tn1.EOF
  If tn1!IDNO Like Text6.Text + "*" Then
   sw = True
   Exit Do
  End If
tn1.MoveNext
 Loop
  
  If sw = True Then
   With tn1
    Text1.Text = !STUDENTNAME
    Text2.Text = !IDNO
    Text3.Text = !YRSEC
    Text4.Text = !ADDRESS
    Text5.Text = !TELL
   End With
  End If

End Sub

Private Sub Text7_Change()
Dim sw1 As Boolean
Dim sw2 As Boolean
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""

tn.MoveFirst
 Do Until tn.EOF
  If tn!AQUISITIONNO Like Text7.Text + "*" Then
   sw1 = True
   Exit Do
  End If
tn.MoveNext
 Loop
  
  If sw1 = True Then
   With tn
   On Error Resume Next
    Text8.Text = !TITLEOFTHEBOOK
    Text9.Text = !AQUISITIONNO
    Text10.Text = !AUTHOR
    Text11.Text = !PUBLISHER
    Combo1.Text = !Status
    Text12.Text = !duedate
   End With
  End If
 

End Sub


Public Sub delete()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
End Sub
