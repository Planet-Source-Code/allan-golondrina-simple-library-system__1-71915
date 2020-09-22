VERSION 5.00
Begin VB.Form Borrowingsec 
   Caption         =   "Borrowing Sec"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command17 
      Caption         =   "Searh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      Picture         =   "Borrowingsec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\allan\Desktop\new.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Borrowedandreturning"
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   6720
      TabIndex        =   45
      Text            =   "Text15"
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command5"
      Height          =   495
      Left            =   6600
      TabIndex        =   44
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   600
      TabIndex        =   43
      Text            =   "Text14"
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   480
      TabIndex        =   42
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Index           =   1
      Left            =   6120
      TabIndex        =   20
      Top             =   2400
      Width           =   5655
      Begin VB.TextBox Text12 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   2040
         TabIndex        =   39
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   2040
         TabIndex        =   38
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   2040
         TabIndex        =   37
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   2040
         TabIndex        =   35
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   2040
         TabIndex        =   34
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   2040
         TabIndex        =   33
         Top             =   330
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton Command8 
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Command11"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   480
         ScaleHeight     =   345
         ScaleWidth      =   1305
         TabIndex        =   21
         Top             =   240
         Width           =   1335
         Begin VB.Label Label1 
            Caption         =   "Accession No:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Label Label9 
         Caption         =   "          Due Date:"
         Height          =   375
         Left            =   600
         TabIndex        =   41
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "               Status:"
         Height          =   495
         Left            =   600
         TabIndex        =   40
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         X1              =   0
         X2              =   5640
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         Caption         =   "   Title of the book:"
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "               Author:"
         Height          =   495
         Left            =   600
         TabIndex        =   25
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   " Date Accquired:"
         Height          =   495
         Left            =   600
         TabIndex        =   24
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "              Volume:"
         Height          =   375
         Left            =   600
         TabIndex        =   23
         Top             =   2520
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Picture         =   "Borrowingsec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      Picture         =   "Borrowingsec.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&Veiw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Picture         =   "Borrowingsec.frx":16C6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "Borrowingsec.frx":19D0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000009&
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000009&
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000009&
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000009&
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   5655
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   480
         ScaleHeight     =   345
         ScaleWidth      =   705
         TabIndex        =   14
         Top             =   240
         Width           =   735
         Begin VB.Label Label1 
            Caption         =   "I.D. No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label Label4 
         Caption         =   "         Start Date:"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "                    Tel:"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "             Address:"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "              Name:"
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   0
         X2              =   5640
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   375
      Left            =   6360
      TabIndex        =   47
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   375
      Left            =   4080
      TabIndex        =   46
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5640
      Picture         =   "Borrowingsec.frx":1E12
      Top             =   240
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      Height          =   1335
      Left            =   3960
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Borrowingsec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tf As Recordset
Dim tf1 As Recordset
Dim dn As Database
Option Compare Text
Dim a As String
Dim b As String
Dim sw As Boolean



Private Sub Command12_Click()
'borrower information save
a = MsgBox("Do you want to add this book", vbYesNo + vbQuestion, "save")
If a = vbYes Then
tf.FindFirst "ID_no like'" + Text1.Text + "'"
If tf.NoMatch = False Then
MsgBox "This ID_no is already exist"
Text1.Text = ""
Else
If Text1.Text <> "" And Text6.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text6.Text <> "" And Text2.Text <> "" _
And Text8.Text <> "" And Text9.Text <> "" And Text10.Text <> "" And Text11.Text <> "" And Text12.Text <> "" Then
tf.AddNew
tf.Edit
tf!ID_no = Text1.Text
tf!Name = Text3.Text
tf!Address = Text4.Text
tf!Tel = Text5.Text
tf!StartDate = Text6.Text
tf!Accession_no = Text2.Text
tf!Title = Text8.Text
tf!Author = Text9.Text
tf!Volume = Text10.Text
tf!Date_Accquired = Text11.Text
tf!DueDate = Text12.Text
tf.Update
Else
a = MsgBox("you must complete the information", vbExclamation + vbOK, "save")
End If
End If
End If


End Sub

Private Sub Command13_Click()
inventory.Show
Unload Me
Me.Hide
End Sub

Private Sub Command14_Click()
Borrowingsec.Refresh

End Sub

Private Sub Command15_Click()
Main_menu.Show
Unload Me
Me.Hide
End Sub

Private Sub Image2_Click()
End Sub

Private Sub Command17_Click()
'search
 
tf1.FindFirst "ID_no like'" + Text1.Text + "'"
If tf1!ID_no = Text1.Text Then


Text3.Text = tf1!Name
Text4.Text = tf1!Address
Text5.Text = tf1!Tel

Else
a = MsgBox("No Record found ", vbOKOnly + vbExclamation, "Search")


End If

End Sub

Private Sub Form_Load()
Set dn = OpenDatabase("C:\Documents and Settings\allan\Desktop\new.mdb")
Set tf = dn.OpenRecordset("Borrowedandreturning", dbOpenDynaset)
Set tf1 = dn.OpenRecordset("Addborrower", dbOpenDynaset)
End Sub
