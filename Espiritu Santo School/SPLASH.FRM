VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10b.ocx"
Begin VB.Form Splash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7320
   ClientLeft      =   3555
   ClientTop       =   1380
   ClientWidth     =   8145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SPLASH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "skip"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   6840
      Width           =   855
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _cx             =   4209570
      _cy             =   4207242
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Timer Timer1 
      Interval        =   2300
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim a

Private Sub ac_OnReadyStateChange(newState As Long)

End Sub

Private Sub Command1_Click()
Form1.Show
Form6.Show
Unload Me
End Sub

Private Sub Form_Load()
ShockwaveFlash1.Movie = App.Path + "\plash.swf"
End Sub

Private Sub Timer1_Timer()
a = a + 1
If a = 13 Then
Form1.Enabled = False
Form1.Show
Form6.Show
Unload Me
End If
End Sub
    

