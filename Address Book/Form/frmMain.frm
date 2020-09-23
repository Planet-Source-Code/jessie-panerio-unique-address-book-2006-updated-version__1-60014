VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMain 
   Caption         =   "Address Book by Jessie Panerio"
   ClientHeight    =   2145
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   2955
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1290
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1740
      ExtentX         =   3069
      ExtentY         =   2275
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright April 2005 Jessie Panerio Philippines"
      BeginProperty Font 
         Name            =   "La Bamba LET"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   12720
      TabIndex        =   1
      Top             =   9720
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileTextSidebar 
         Caption         =   $"frmMain.frx":0BC2
      End
      Begin VB.Menu mnuAddrBook 
         Caption         =   "Address &Book"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuAuthor 
         Caption         =   "&Author"
      End
      Begin VB.Menu mnuCredit1 
         Caption         =   "Credit...&1"
      End
      Begin VB.Menu mnuCredit2 
         Caption         =   "Credit...&2"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-{RAISED}Zaphat"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
     
     On Error Resume Next
     Kill App.Path & "\tmpfile.swf"
     LoadDataIntoFile 101, App.Path & "\tmpfile.swf"
     WebBrowser1.Navigate App.Path & "\tmpfile.swf"
     modMenus.HighlightDisabledMenuItems = True
     modMenus.HighlightGradient = True
     modMenus.RaisedIconOnSelect = True
     modMenus.SelectedItemBackColor = vbBlue
     modMenus.SelectedItemTextColor = vbCyan
     modMenus.SeparatorBarColor_Dark = vbBlue
     modMenus.SeparatorBarColor_Light = vbBlue
     SetMenu frmMain.hWnd
                   
End Sub

Private Sub Form_Resize()

     WebBrowser1.Left = Me.ScaleLeft
     WebBrowser1.Top = Me.ScaleTop
     WebBrowser1.Width = Me.ScaleWidth
     WebBrowser1.Height = Me.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)
     
     On Error Resume Next
     Kill App.Path & "\tmpfile.swf"
     cnn.Close
     Set adoview = Nothing
    
End Sub

Private Sub mnuAddrBook_Click()
    
     On Error Resume Next
     Kill App.Path & "\tmpfile.swf"
     LoadDataIntoFile 104, App.Path & "\tmpfile.swf"
     WebBrowser1.Navigate App.Path & "\tmpfile.swf"
     frmMain.Enabled = False
     Load frmPersonalInfo
     frmPersonalInfo.Show
        
End Sub

Private Sub mnuAuthor_Click()
  
     On Error Resume Next
     Kill App.Path & "\tmpfile.swf"
     LoadDataIntoFile 102, App.Path & "\tmpfile.swf"
     WebBrowser1.Navigate App.Path & "\tmpfile.swf"

End Sub

Private Sub mnuCredit1_Click()
  
     On Error Resume Next
     Kill App.Path & "\tmpfile.swf"
     LoadDataIntoFile 103, App.Path & "\tmpfile.swf"
     WebBrowser1.Navigate App.Path & "\tmpfile.swf"

End Sub

Private Sub mnuCredit2_Click()
  
     On Error Resume Next
     Kill App.Path & "\tmpfile.swf"
     LoadDataIntoFile 101, App.Path & "\tmpfile.swf"
     WebBrowser1.Navigate App.Path & "\tmpfile.swf"

End Sub

Private Sub mnuExit_Click()

     On Error Resume Next
     Kill App.Path & "\tmpfile.swf"
     Unload Me
     
End Sub

Private Sub mnuHelp_Click()
  
     Call WinHelp(0, App.HelpFile, HELPC, 0)

End Sub
