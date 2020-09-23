VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddNew 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   4440
      TabIndex        =   67
      Top             =   5280
      Width           =   4335
      Begin VB.Label Label16 
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
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1080
         TabIndex        =   69
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label13 
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
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   480
         TabIndex        =   68
         Top             =   480
         Width           =   3495
      End
   End
   Begin Project1.lvButtons_H lvbuttonsEditClose 
      Height          =   975
      Left            =   3360
      TabIndex        =   65
      Top             =   5280
      Visible         =   0   'False
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
      Image           =   "AddNew.frx":0000
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbuttonsEditSave 
      Height          =   975
      Left            =   2280
      TabIndex        =   64
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&Save"
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
      Image           =   "AddNew.frx":0452
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbuttonsEditCancel 
      Height          =   975
      Left            =   1200
      TabIndex        =   63
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "Canc&el"
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
      Image           =   "AddNew.frx":0D2C
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvButtons_H4Cover 
      Height          =   975
      Left            =   3360
      TabIndex        =   60
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
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
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvButtons_H3Cover 
      Height          =   975
      Left            =   2280
      TabIndex        =   59
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
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
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvButtons_H2Cover 
      Height          =   975
      Left            =   1200
      TabIndex        =   58
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
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
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvButtons_H1Cover 
      Height          =   975
      Left            =   120
      TabIndex        =   56
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
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
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      TabHeight       =   582
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal &Information"
      TabPicture(0)   =   "AddNew.frx":1606
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "C&ontact No."
      TabPicture(1)   =   "AddNew.frx":1622
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label24"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "E-&mail Address"
      TabPicture(2)   =   "AddNew.frx":163E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame9 
         Height          =   2535
         Left            =   -74280
         TabIndex        =   31
         Top             =   1080
         Width           =   7335
         Begin VB.Frame Frame8 
            Height          =   2295
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   7095
            Begin VB.Frame Frame3 
               Height          =   2055
               Left            =   120
               TabIndex        =   33
               Top             =   120
               Width           =   6855
               Begin VB.TextBox txtEmail1 
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
                  Left            =   2160
                  MaxLength       =   30
                  TabIndex        =   18
                  Top             =   360
                  Width           =   4575
               End
               Begin VB.TextBox txtEmail2 
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
                  Left            =   2160
                  MaxLength       =   30
                  TabIndex        =   19
                  Top             =   840
                  Width           =   4575
               End
               Begin VB.TextBox txtEmail3 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   2160
                  MaxLength       =   30
                  TabIndex        =   20
                  Top             =   1320
                  Width           =   4575
               End
               Begin VB.Label Label6 
                  Caption         =   "E-mail Address 1"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   36
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.Label Label7 
                  Caption         =   "E-mail Address 3"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   35
                  Top             =   1440
                  Width           =   1815
               End
               Begin VB.Label Label8 
                  Caption         =   "E-mail Address 2"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   34
                  Top             =   960
                  Width           =   1815
               End
            End
         End
         Begin VB.Label Label25 
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
            Left            =   2280
            TabIndex        =   74
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   26
         Top             =   2520
         Width           =   8655
         Begin VB.Frame Frame12 
            Height          =   1935
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   8415
            Begin VB.TextBox txtOfficePhone 
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
               Left            =   2280
               MaxLength       =   20
               TabIndex        =   15
               Top             =   1440
               Width           =   2775
            End
            Begin VB.TextBox txtOfficeName 
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
               Left            =   2280
               MaxLength       =   50
               TabIndex        =   13
               Top             =   240
               Width           =   6015
            End
            Begin VB.TextBox txtOfficeAddr 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   2280
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   14
               Top             =   720
               Width           =   6015
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Office Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   480
               TabIndex        =   30
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Office Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   240
               TabIndex        =   29
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Office Phone #"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   240
               TabIndex        =   28
               Top             =   1440
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   8655
         Begin VB.Frame Frame10 
            Height          =   1935
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   8415
            Begin VB.TextBox txtContactHome 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   2280
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   10
               Top             =   240
               Width           =   6015
            End
            Begin VB.TextBox txtContactLandLine 
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
               Left            =   2280
               MaxLength       =   20
               TabIndex        =   11
               Top             =   960
               Width           =   2775
            End
            Begin VB.TextBox txtContactMobile 
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
               Left            =   2280
               MaxLength       =   20
               TabIndex        =   12
               Top             =   1440
               Width           =   2775
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Home Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label lblContactNoMobile 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile Phone #"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Landline #"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   720
               TabIndex        =   23
               Top             =   1080
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   8655
         Begin VB.Frame Frame2 
            Height          =   4455
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   8415
            Begin VB.Frame Frame7 
               Height          =   2055
               Left            =   4200
               TabIndex        =   54
               Top             =   2280
               Width           =   4095
               Begin Project1.lvButtons_H cmdAddPicture 
                  Height          =   1215
                  Left            =   2640
                  TabIndex        =   9
                  Top             =   480
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   2143
                  Caption         =   "&Add Picture"
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
                  Image           =   "AddNew.frx":165A
                  ImgSize         =   32
                  cBack           =   16711680
               End
               Begin VB.PictureBox picimage 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00808080&
                  ForeColor       =   &H80000008&
                  Height          =   1695
                  Left            =   120
                  ScaleHeight     =   1665
                  ScaleWidth      =   1860
                  TabIndex        =   62
                  Top             =   240
                  Width           =   1890
                  Begin MSComDlg.CommonDialog jessiepanerio 
                     Left            =   8760
                     Top             =   1080
                     _ExtentX        =   847
                     _ExtentY        =   847
                     _Version        =   393216
                  End
                  Begin MSComDlg.CommonDialog jjessiepanerio 
                     Left            =   8760
                     Top             =   600
                     _ExtentX        =   847
                     _ExtentY        =   847
                     _Version        =   393216
                  End
                  Begin VB.Image imgpic 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   1635
                     Left            =   0
                     Stretch         =   -1  'True
                     Top             =   0
                     Width           =   1875
                  End
               End
               Begin VB.TextBox txtPictureName 
                  Height          =   285
                  Left            =   480
                  Locked          =   -1  'True
                  TabIndex        =   61
                  Top             =   840
                  Width           =   1095
               End
               Begin Project1.lvButtons_H cmdChangePicture 
                  Height          =   1215
                  Left            =   2640
                  TabIndex        =   66
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   2143
                  Caption         =   "Change &Picture"
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
                  Image           =   "AddNew.frx":1F34
                  ImgSize         =   32
                  cBack           =   16711680
               End
            End
            Begin VB.Frame Frame6 
               Height          =   2175
               Left            =   4200
               TabIndex        =   50
               Top             =   120
               Width           =   4095
               Begin VB.TextBox txtCitizenship 
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
                  Left            =   1800
                  MaxLength       =   18
                  TabIndex        =   7
                  Top             =   960
                  Width           =   2175
               End
               Begin VB.TextBox txtCivilStatus 
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
                  Left            =   1800
                  MaxLength       =   18
                  TabIndex        =   8
                  Top             =   1560
                  Width           =   2175
               End
               Begin VB.TextBox txtReligion 
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
                  Left            =   1800
                  MaxLength       =   18
                  TabIndex        =   6
                  Top             =   360
                  Width           =   2175
               End
               Begin VB.Label Label21 
                  Caption         =   "Citizenship"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   53
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.Label Label22 
                  Caption         =   "Civil Status"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   52
                  Top             =   1680
                  Width           =   1335
               End
               Begin VB.Label Label20 
                  Caption         =   "Religion"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   51
                  Top             =   480
                  Width           =   975
               End
            End
            Begin VB.Frame Frame11 
               Height          =   4215
               Left            =   120
               TabIndex        =   43
               Top             =   120
               Width           =   3975
               Begin VB.TextBox txtCover 
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
                  Left            =   1680
                  Locked          =   -1  'True
                  TabIndex        =   57
                  Top             =   2760
                  Width           =   1575
               End
               Begin MSMask.MaskEdBox MaskEdBoxBirthDay 
                  Height          =   375
                  Left            =   1680
                  TabIndex        =   4
                  Top             =   2760
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PromptChar      =   "_"
               End
               Begin VB.ComboBox cmbGender 
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
                  ItemData        =   "AddNew.frx":280E
                  Left            =   1680
                  List            =   "AddNew.frx":2818
                  TabIndex        =   5
                  Top             =   3600
                  Width           =   1695
               End
               Begin VB.TextBox txtNickName 
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
                  Left            =   1680
                  MaxLength       =   10
                  TabIndex        =   0
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.TextBox txtLastName 
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
                  Left            =   1680
                  MaxLength       =   18
                  TabIndex        =   1
                  Top             =   960
                  Width           =   2055
               End
               Begin VB.TextBox txtFirstName 
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
                  Left            =   1680
                  MaxLength       =   18
                  TabIndex        =   2
                  Top             =   1560
                  Width           =   2055
               End
               Begin VB.TextBox txtMiddleName 
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
                  Left            =   1680
                  MaxLength       =   18
                  TabIndex        =   3
                  Top             =   2160
                  Width           =   2055
               End
               Begin VB.TextBox txtNickName1 
                  Height          =   375
                  Left            =   2280
                  TabIndex        =   76
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label Label12 
                  Alignment       =   2  'Center
                  Caption         =   "MM - DD - YYYY"
                  ForeColor       =   &H00000080&
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   55
                  Top             =   3240
                  Width           =   1215
               End
               Begin VB.Label Label19 
                  Caption         =   "Gender"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   49
                  Top             =   3720
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Caption         =   "NickName"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   48
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Label Label2 
                  Caption         =   "LastName"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   47
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.Label Label3 
                  Caption         =   "FirstName"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   46
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.Label Label4 
                  Caption         =   "MiddleName"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   45
                  Top             =   2280
                  Width           =   1335
               End
               Begin VB.Label Label5 
                  Caption         =   "BirthDay"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   44
                  Top             =   2880
                  Width           =   1215
               End
            End
            Begin VB.Label Label23 
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
               Left            =   720
               TabIndex        =   72
               Top             =   1800
               Width           =   2415
            End
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
            Left            =   1920
            TabIndex        =   70
            Top             =   1800
            Width           =   2415
         End
      End
      Begin VB.Label Label24 
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
         Left            =   -72480
         TabIndex        =   73
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label18 
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
         Left            =   2640
         TabIndex        =   71
         Top             =   2400
         Width           =   2415
      End
   End
   Begin Project1.lvButtons_H lvbutton4Close 
      Height          =   975
      Left            =   3360
      TabIndex        =   38
      Top             =   5280
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
      Image           =   "AddNew.frx":282A
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbutton3Save 
      Height          =   975
      Left            =   2280
      TabIndex        =   39
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&Save"
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
      Image           =   "AddNew.frx":2C7C
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbutton2Cancel 
      Height          =   975
      Left            =   1200
      TabIndex        =   40
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "Canc&el"
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
      Image           =   "AddNew.frx":3556
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbutton1New 
      Height          =   975
      Left            =   120
      TabIndex        =   41
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&New"
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
      Image           =   "AddNew.frx":3E30
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5160
      Width           =   8895
   End
   Begin VB.Label Label26 
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
      Left            =   3240
      TabIndex        =   75
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "frmAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
     
     infomode = True
     findmode = False
     
End Sub

Private Sub txtNickName_KeyPress(KeyAscii As Integer)
   
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
          
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
  
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
   
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
   
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
   
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
   
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
   
End Sub

Private Sub MaskEdBoxBirthDay_GotFocus()
   
     If lvButtons_H1Cover.Visible = True Then
        MaskEdBoxBirthDay.Mask = "##/##/####"
     End If

End Sub

Private Sub MaskEdBoxBirthDay_KeyPress(KeyAscii As Integer)
   
     Dim strvalid
     strvalid = "0123456789"
     If KeyAscii > 26 Then
       If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
     End If

End Sub

Private Sub cmbGender_KeyPress(KeyAscii As Integer)
   
     Dim strvalid
     strvalid = ""
     If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
           KeyAscii = 0
        End If
     End If

End Sub

Private Sub txtReligion_KeyPress(KeyAscii As Integer)
   
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

Private Sub txtCitizenship_KeyPress(KeyAscii As Integer)
   
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
   
End Sub

Private Sub txtCivilStatus_KeyPress(KeyAscii As Integer)
   
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
   
End Sub

Private Sub txtContactLandLine_KeyPress(KeyAscii As Integer)

     Dim strvalid
     strvalid = "0123456789-+()"
     If KeyAscii > 26 Then
       If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
     End If

End Sub

Private Sub txtContactMobile_KeyPress(KeyAscii As Integer)
  
     Dim strvalid
     strvalid = "0123456789-+()"
     If KeyAscii > 26 Then
       If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
     End If
     
End Sub

Private Sub txtOfficePhone_KeyPress(KeyAscii As Integer)
  
     Dim strvalid
     strvalid = "0123456789-+()"
     If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
           KeyAscii = 0
        End If
     End If

End Sub

Private Sub cmdAddPicture_Click()
    
     With jjessiepanerio
       .InitDir = "C:\My Documents"
       .Filter = "JPEG image|*.jpg|GIF image|*.gif|BITMAP image|*.bmp|Icon image|*.ico|Cursor image|*.cur|Panerio image|*.pan"
       .ShowOpen
          If .FileName <> "" Then
             strImgN = .FileName
             txtPictureName.Text = .FileTitle
             imgpic.Picture = LoadPicture(.FileName)
          End If
     End With
    
End Sub

Private Sub cmdChangePicture_Click()
    
     cmdAddPicture_Click

End Sub

Private Sub lvbutton1New_Click()
        
     Call cmdCover(True, False, False, True, True)
     Call lockcontrols(False)
     txtNickName.SetFocus
     
End Sub

Private Sub lvbutton2Cancel_Click()
 
     Call clearcontrols
     Call cmdCover(False, True, True, False, False)
     Call lockcontrols(True)
          
End Sub

Private Sub lvbutton3Save_Click()
     
     Dim res As VbMsgBoxResult
     If frmAddNew.txtNickName.Text = "" Then
        MsgBox "Nickname field is empty." & vbCrLf & "Please enter a value for the said field.", vbInformation, "Information"
        txtNickName.SetFocus
        Exit Sub
     Else
        Call validate
          If modeval = False Then
             res = MsgBox("Save this to Database?", vbYesNo, "Confirmation")
                If res = vbYes Then
                   adoview.AddNew
                   WriteDataFromControls
                   adoview.Update
                   Call clearcontrols
                   Call cmdCover(False, True, True, False, False)
                   Call lockcontrols(True)
                   Call loaddataforviewing
                   Load frmPersonalInfo
                   frmPersonalInfo.Show
                   Unload Me
                Else
                   Exit Sub
                End If
          Else
             MsgBox "Warning! Duplication of entries is not allowed in this application." & vbCrLf & vbCrLf & "Nickname ''" & txtNickName.Text & "'' already exist.", vbExclamation, "Address Book"
             txtNickName.Text = ""
             txtNickName.SetFocus
          End If
     End If
     Exit Sub

End Sub

Private Sub lvbutton4Close_Click()
 
     Unload Me
     Load frmPersonalInfo
     frmPersonalInfo.Show
         
End Sub

Private Sub lvbuttonsEditCancel_Click()
   
     Call loaddatatoedit
   
End Sub

Private Sub lvbuttonsEditSave_Click()
     
     Dim res As VbMsgBoxResult
     If txtNickName.Text = "" Then
        MsgBox "Nickname field is empty." & vbCrLf & "Please enter a value for the said field.", vbInformation, "Information"
        txtNickName.SetFocus
        Exit Sub
     Else
        Call validate
          If modeval = False Then
save:
             res = MsgBox("Save this to Database?", vbYesNo, "Confirmation")
                If res = vbYes Then
                   WriteDataFromControlsEdit
                   adoview.Update
                   Load frmPersonalInfo
                   frmPersonalInfo.Show
                   Unload Me
                Else
                   Exit Sub
                End If
          Else
             If txtNickName1.Text = txtNickName.Text Then
                GoTo save:
                Exit Sub
             Else
                MsgBox "Warning! Duplication of entries is not allowed in this application." & vbCrLf & vbCrLf & "Nickname ''" & txtNickName.Text & "'' already exist.", vbExclamation, "Address Book"
                txtNickName.Text = ""
                txtNickName.SetFocus
             End If
          End If
     End If
     
End Sub

Private Sub lvbuttonsEditClose_Click()
   
     Call LoadImage
     frmPersonalInfo.imgpic.Picture = frmAddNew.imgpic.Picture
     Load frmPersonalInfo
     frmPersonalInfo.Show
     Unload Me
 
End Sub

