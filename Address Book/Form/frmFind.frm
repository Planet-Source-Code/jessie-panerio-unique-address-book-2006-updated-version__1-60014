VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Search"
      TabPicture(0)   =   "frmFind.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   5895
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8655
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmFind.frx":001C
            Height          =   2295
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "LISTING"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "lastname"
               Caption         =   "Lastname"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "firstname"
               Caption         =   "Firstname"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "nickname"
               Caption         =   "Nickname"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "gender"
               Caption         =   "Gender"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   5
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   2204.788
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   2204.788
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2204.788
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1454.74
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame5 
            Height          =   2415
            Left            =   6720
            TabIndex        =   18
            Top             =   3360
            Width           =   1815
            Begin VB.PictureBox Picture1 
               Height          =   2055
               Left            =   120
               ScaleHeight     =   1995
               ScaleWidth      =   1515
               TabIndex        =   19
               Top             =   240
               Width           =   1575
               Begin VB.Image imgpic 
                  Height          =   2055
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   1575
               End
            End
         End
         Begin VB.Frame Frame4 
            Height          =   855
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   5295
            Begin VB.Label lblrecno 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   1440
               TabIndex        =   20
               Top             =   360
               Width           =   2415
            End
         End
         Begin VB.Frame Frame7 
            Height          =   855
            Left            =   5520
            TabIndex        =   16
            Top             =   120
            Width           =   3015
            Begin Project1.lvButtons_H lvbutton7MoveFirst 
               Height          =   495
               Left            =   120
               TabIndex        =   4
               ToolTipText     =   "Move to the First Record"
               Top             =   240
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               Caption         =   "|<"
               CapAlign        =   2
               BackStyle       =   7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Swis721 BlkEx BT"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   65535
               cFHover         =   255
               cBhover         =   16776960
               LockHover       =   3
               cGradient       =   16711680
               Gradient        =   3
               Mode            =   0
               Value           =   0   'False
               cBack           =   0
            End
            Begin Project1.lvButtons_H lvbutton8Previous 
               Height          =   495
               Left            =   840
               TabIndex        =   5
               ToolTipText     =   "Move to the Previous Record"
               Top             =   240
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               Caption         =   "<"
               CapAlign        =   2
               BackStyle       =   7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Swis721 BlkEx BT"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   65535
               cFHover         =   255
               cBhover         =   16776960
               LockHover       =   3
               cGradient       =   16711680
               Gradient        =   3
               Mode            =   0
               Value           =   0   'False
               cBack           =   0
            End
            Begin Project1.lvButtons_H lvbutton9Next 
               Height          =   495
               Left            =   1560
               TabIndex        =   6
               ToolTipText     =   "Move to the Next Record"
               Top             =   240
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               Caption         =   ">"
               CapAlign        =   2
               BackStyle       =   7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Swis721 BlkEx BT"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   65535
               cFHover         =   255
               cBhover         =   16776960
               LockHover       =   3
               cGradient       =   16711680
               Gradient        =   3
               Mode            =   0
               Value           =   0   'False
               cBack           =   0
            End
            Begin Project1.lvButtons_H lvbutton10MoveLast 
               Height          =   495
               Left            =   2280
               TabIndex        =   7
               ToolTipText     =   "Move to the Last Record"
               Top             =   240
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               Caption         =   ">|"
               CapAlign        =   2
               BackStyle       =   7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Swis721 BlkEx BT"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   65535
               cFHover         =   255
               cBhover         =   16776960
               LockHover       =   3
               cGradient       =   16711680
               Gradient        =   3
               Mode            =   0
               Value           =   0   'False
               cBack           =   0
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1335
            Left            =   120
            TabIndex        =   12
            Top             =   4440
            Width           =   6495
            Begin Project1.lvButtons_H cmdClose 
               Height          =   975
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1720
               Caption         =   "&Close"
               CapAlign        =   2
               BackStyle       =   7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   65280
               cFHover         =   65535
               cBhover         =   255
               LockHover       =   3
               cGradient       =   0
               Gradient        =   3
               Mode            =   0
               Value           =   0   'False
               ImgAlign        =   4
               Image           =   "frmFind.frx":0031
               ImgSize         =   32
               cBack           =   16711680
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Jessie Panerio Philippines"
               BeginProperty Font 
                  Name            =   "La Bamba LET"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   2160
               TabIndex        =   22
               Top             =   840
               Width           =   3135
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Copyright April 2005"
               BeginProperty Font 
                  Name            =   "La Bamba LET"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   2520
               TabIndex        =   21
               Top             =   480
               Width           =   2415
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1095
            Left            =   120
            TabIndex        =   11
            Top             =   3360
            Width           =   6495
            Begin VB.ComboBox cmbField 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               ItemData        =   "frmFind.frx":0483
               Left            =   120
               List            =   "frmFind.frx":0490
               TabIndex        =   2
               Text            =   "lastname"
               Top             =   600
               Width           =   2295
            End
            Begin VB.ComboBox cmbSort 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               ItemData        =   "frmFind.frx":04B3
               Left            =   2520
               List            =   "frmFind.frx":04BD
               TabIndex        =   3
               Text            =   "ASC"
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox txtvalue 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4560
               TabIndex        =   1
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label3 
               Caption         =   "Sort Order"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3000
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "Search for?"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   5040
               TabIndex        =   14
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "Fieldnames"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   600
               TabIndex        =   13
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Program by Jessie Panerio of Philippines All Rights Reserved"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   975
            Left            =   3240
            TabIndex        =   25
            Top             =   1920
            Width           =   3255
         End
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Program by Jessie Panerio of Philippines All Rights Reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   2760
         TabIndex        =   24
         Top             =   2760
         Width           =   3255
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program by Jessie Panerio of Philippines All Rights Reserved"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   2880
      TabIndex        =   23
      Top             =   2280
      Width           =   3255
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
     
     findmode = True
     infomode = False
     If adoview.BOF = True And adoview.EOF = True Then
        MsgBox "Empty Database", vbInformation, "Address Book"
        Exit Sub
     Else
        adoview.Close
        Call rs_view
        adoview.MoveFirst
        Call LoadImage
        Set frmFind.DataGrid1.DataSource = adoview
        Call recordcount
     End If
     Set frmFind.DataGrid1.DataSource = adoview

End Sub

Private Sub cmbField_Click()
     
     Call sort
     Call LoadImage
     txtvalue.Text = ""
     
End Sub

Private Sub cmbField_KeyPress(KeyAscii As Integer)
   
     Dim strvalid
     strvalid = ""
     If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
           KeyAscii = 0
        End If
     End If

End Sub

Private Sub cmbSort_Click()
     
     Call sort
     Call LoadImage
     txtvalue.Text = ""

End Sub

Private Sub txtvalue_KeyPress(KeyAscii As Integer)
   
     If KeyAscii = 13 Then
        Call search
     End If
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
          
End Sub

Private Sub cmbSort_KeyPress(KeyAscii As Integer)
  
     Dim strvalid
     strvalid = ""
     If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
           KeyAscii = 0
        End If
     End If

End Sub

Private Sub cmdClose_Click()
     
     infomode = True
     findmode = False
     frmPersonalInfo.imgpic.Picture = frmFind.imgpic.Picture
     Load frmPersonalInfo
     frmPersonalInfo.Show
     Call LoadImage
     Unload Me
  
End Sub

Private Sub txtvalue_Change()
  
     search
     
End Sub

Private Sub search()

     adoview.Find (" " & frmFind.cmbField.Text & " = '" & frmFind.txtvalue.Text & "'")
     Set frmFind.DataGrid1.DataSource = adoview
     Call recordcount
     If adoview.AbsolutePosition = -1 Then
        MsgBox "No file found", vbInformation, "Address Book"
        txtvalue.Text = ""
        txtvalue.SetFocus
        Call LoadImage
     Else
        Call LoadImage
     End If
    
End Sub

Private Sub sort()

     On Error Resume Next
     Dim SQLTEXT As String
     adoview.Close
     SQLTEXT = "SELECT * FROM addressbook ORDER BY" + " " & frmFind.cmbField.Text + " " & frmFind.cmbSort.Text
     adoview.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
     Set frmFind.DataGrid1.DataSource = adoview
     
End Sub
      
Private Sub recordcount()
   
     lblrecno.Caption = " Record   " & adoview.AbsolutePosition & "  of   " & adoview.recordcount
  
End Sub

Private Sub lvbutton7MoveFirst_Click()
  
     If adoview.recordcount <= 1 Then Exit Sub
       txtvalue.Text = ""
       adoview.MoveFirst
       Call recordcount
       Call LoadImage
      
End Sub

Private Sub lvbutton8Previous_Click()
  
     If adoview.AbsolutePosition <= 1 Then Exit Sub
       txtvalue.Text = ""
       adoview.MovePrevious
       Call recordcount
       Call LoadImage
       
End Sub

Private Sub lvbutton9Next_Click()
    
     If adoview.AbsolutePosition >= adoview.recordcount Or adoview.recordcount <= 1 Then Exit Sub
       txtvalue.Text = ""
       adoview.MoveNext
       Call recordcount
       Call LoadImage
       
End Sub

Private Sub lvbutton10MoveLast_Click()
     
     If adoview.recordcount <= 1 Then Exit Sub
       txtvalue.Text = ""
       adoview.MoveLast
       recordcount
       Call LoadImage
       
End Sub

