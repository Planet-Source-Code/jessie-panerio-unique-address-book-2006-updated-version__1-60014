VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersonalInfo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   13
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
      TabPicture(0)   =   "PersonalInfo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRecNo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "C&ontact No."
      TabPicture(1)   =   "PersonalInfo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "E-&mail Address"
      TabPicture(2)   =   "PersonalInfo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame17"
      Tab(2).Control(2)=   "Frame13"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame4 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   8655
         Begin VB.Frame Frame10 
            Height          =   1935
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   8415
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
               TabIndex        =   57
               Top             =   360
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
               TabIndex        =   56
               Top             =   1440
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
               TabIndex        =   55
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label lblContactHome 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   900
               Left            =   2160
               TabIndex        =   54
               Top             =   120
               Width           =   6165
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblContactLandLine 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2160
               TabIndex        =   53
               Top             =   1080
               Width           =   3615
            End
            Begin VB.Label lblContactMobile 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2160
               TabIndex        =   52
               Top             =   1440
               Width           =   3615
            End
         End
         Begin VB.Frame Frame11 
            Height          =   615
            Left            =   120
            TabIndex        =   49
            Top             =   120
            Width           =   8415
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Address and Contact No."
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
               Height          =   375
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame Frame12 
            Height          =   1935
            Left            =   120
            TabIndex        =   42
            Top             =   2640
            Width           =   8415
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
               TabIndex        =   48
               Top             =   240
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
               TabIndex        =   47
               Top             =   840
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
               TabIndex        =   46
               Top             =   1560
               Width           =   1695
            End
            Begin VB.Label lblOfficeName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2160
               TabIndex        =   45
               Top             =   240
               Width           =   6135
            End
            Begin VB.Label lblOfficeAddr 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   855
               Left            =   2160
               TabIndex        =   44
               Top             =   600
               Width           =   6135
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblOfficePhone 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2160
               TabIndex        =   43
               Top             =   1560
               Width           =   3615
            End
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
            TabIndex        =   69
            Top             =   1200
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   5
         Top             =   1200
         Width           =   8175
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "3."
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
            Left            =   2160
            TabIndex        =   40
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "2."
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
            Left            =   2160
            TabIndex        =   39
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "1."
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
            Left            =   2160
            TabIndex        =   38
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblEmail2 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   2520
            TabIndex        =   14
            Top             =   1440
            Width           =   4815
         End
         Begin VB.Label lblEmail3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   2520
            TabIndex        =   12
            Top             =   1920
            Width           =   4815
         End
         Begin VB.Label lblEmail1 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   2520
            TabIndex        =   11
            Top             =   960
            Width           =   4815
         End
      End
      Begin VB.Frame Frame17 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   35
         Top             =   480
         Width           =   8415
         Begin VB.Frame Frame14 
            Height          =   615
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   8175
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "E-mail Address"
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
               Height          =   375
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Label Label22 
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
            Left            =   2880
            TabIndex        =   70
            Top             =   2160
            Width           =   2415
         End
      End
      Begin VB.Frame Frame13 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   34
         Top             =   360
         Width           =   8655
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   5760
         TabIndex        =   31
         Top             =   360
         Width           =   3015
         Begin VB.PictureBox picimage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   2775
            Left            =   120
            ScaleHeight     =   2745
            ScaleWidth      =   2745
            TabIndex        =   64
            Top             =   240
            Width           =   2775
            Begin VB.Label lblWanted 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   120
               TabIndex        =   65
               Top             =   2280
               Width           =   2535
            End
            Begin VB.Image imgpic 
               BorderStyle     =   1  'Fixed Single
               Height          =   2775
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2775
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4695
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   5535
         Begin VB.Frame Frame2 
            Height          =   4455
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   5295
            Begin VB.Label lblcivilstatus 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2040
               TabIndex        =   67
               Top             =   3720
               Width           =   3135
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "MM / DD /  YYYY"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   2040
               TabIndex        =   63
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label lblReligion 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2040
               TabIndex        =   62
               Top             =   2760
               Width           =   3015
            End
            Begin VB.Label lblCitizenship 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2040
               TabIndex        =   60
               Top             =   3240
               Width           =   3135
            End
            Begin VB.Label lblGender 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2040
               TabIndex        =   59
               Top             =   2280
               Width           =   3015
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Civil Status"
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
               TabIndex        =   33
               Top             =   3720
               Width           =   1335
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Citizenship"
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
               TabIndex        =   32
               Top             =   3240
               Width           =   1335
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Religion"
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
               TabIndex        =   30
               Top             =   2760
               Width           =   1095
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Gender"
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
               Top             =   2280
               Width           =   975
            End
            Begin VB.Label lblBirthDay 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2040
               TabIndex        =   28
               Top             =   1800
               Width           =   3015
            End
            Begin VB.Label lblMiddleName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2040
               TabIndex        =   27
               Top             =   1320
               Width           =   3255
            End
            Begin VB.Label lblFirstName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2040
               TabIndex        =   26
               Top             =   840
               Width           =   3135
            End
            Begin VB.Label lblLastName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2040
               TabIndex        =   25
               Top             =   360
               Width           =   3255
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "BirthDay"
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
               TabIndex        =   24
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "LastName"
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
               TabIndex        =   23
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "FirstName"
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
               TabIndex        =   22
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "MiddleName"
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
               TabIndex        =   21
               Top             =   1320
               Width           =   1455
            End
         End
      End
      Begin VB.Frame Frame7 
         Height          =   975
         Left            =   5760
         TabIndex        =   18
         Top             =   4080
         Width           =   3015
         Begin Project1.lvButtons_H lvbutton7MoveFirst 
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   615
            _ExtentX        =   873
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
            Height          =   615
            Left            =   840
            TabIndex        =   8
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
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
            Height          =   615
            Left            =   1560
            TabIndex        =   9
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
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
            Height          =   615
            Left            =   2280
            TabIndex        =   10
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
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
      Begin VB.Frame Frame8 
         Height          =   615
         Left            =   5760
         TabIndex        =   15
         Top             =   3480
         Width           =   3015
         Begin VB.Label lblNickName2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label lblNickName1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label lblNickName3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   2655
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
         TabIndex        =   68
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label lblRecNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6360
         TabIndex        =   66
         Top             =   120
         Width           =   1935
      End
   End
   Begin Project1.lvButtons_H lvbutton6Close 
      Height          =   975
      Left            =   5520
      TabIndex        =   6
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
      Image           =   "PersonalInfo.frx":0054
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbutton5Print 
      Height          =   975
      Left            =   4440
      TabIndex        =   4
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&Print"
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
      Image           =   "PersonalInfo.frx":04A6
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbutton4Delete 
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&Delete"
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
      Image           =   "PersonalInfo.frx":0D80
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbutton3Find 
      Height          =   975
      Left            =   2280
      TabIndex        =   2
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&Find"
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
      Image           =   "PersonalInfo.frx":1A5A
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbutton2Edit 
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&Edit"
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
      Image           =   "PersonalInfo.frx":2334
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin Project1.lvButtons_H lvbutton1AddNew 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "Add &New"
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
      Image           =   "PersonalInfo.frx":300E
      ImgSize         =   32
      cBack           =   16711680
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5160
      Width           =   8895
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnuFileTextSidebar 
         Caption         =   $"PersonalInfo.frx":38E8
      End
      Begin VB.Menu mnuPrintCurrentList 
         Caption         =   "&Current List"
      End
      Begin VB.Menu mnuPrintAll 
         Caption         =   "&All"
      End
      Begin VB.Menu mnuPrintGender 
         Caption         =   "By &Gender"
         Begin VB.Menu mnuPrintMale 
            Caption         =   "&Male"
         End
         Begin VB.Menu mnuPrintFemale 
            Caption         =   "&Female"
         End
      End
   End
End
Attribute VB_Name = "frmPersonalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
       
     If mode = True Then
     Call main
     End If
     infomode = True
     findmode = False
     modMenus.HighlightDisabledMenuItems = True
     modMenus.HighlightGradient = True
     modMenus.RaisedIconOnSelect = True
     modMenus.SelectedItemBackColor = vbBlue
     modMenus.SelectedItemTextColor = vbCyan
     modMenus.SeparatorBarColor_Dark = vbBlue
     modMenus.SeparatorBarColor_Light = vbBlue
     SetMenu frmPersonalInfo.hWnd
     Call loaddataforviewing
     mode = False
         
End Sub

Private Sub lvbutton1AddNew_Click()
  
     frmPersonalInfo.Hide
     Load frmAddNew
     frmAddNew.Caption = "Add New"
     frmAddNew.Show
     Call cmdCover(False, True, True, False, False)
     Call lockcontrols(True)
 
End Sub

Private Sub lvbutton2Edit_Click()
   
     If adoview.BOF = True Or adoview.EOF = True Then
        MsgBox "Empty Database!", vbInformation, "Address Book"
        Exit Sub
     Else
        Call cmdCover(True, True, True, True, False)
        Call lvbutton(False)
        Call lockcontrols(False)
        Call loaddatatoedit
        Call editbutton(True)
        frmAddNew.imgpic.Picture = frmPersonalInfo.imgpic.Picture
        Load frmAddNew
        frmAddNew.Show
        frmAddNew.Caption = "Edit"
        Unload Me
     End If
 
End Sub

Private Sub lvbutton3Find_Click()
   
     Load frmFind
     frmFind.Show
     Unload Me

End Sub

Private Sub lvbutton4Delete_Click()
     
     Dim res As VbMsgBoxResult
     With adoview
        If .BOF And .EOF = True Then
           MsgBox "Empty Database!", vbInformation, "Address Book"
           Exit Sub
        Else
           res = MsgBox("Are you sure you want to Delete  " & adoview!firstname & " ''" & adoview!nickname & "'' " & adoview!lastname, vbYesNo, "Confirmation")
              If res = vbYes Then
                 .Delete
                 .Requery
                 Call clearlblcontrols
              Else
                 Exit Sub
              End If
           Call loaddataforviewing
        End If
     End With
    
End Sub

Private Sub lvbutton5Print_Click()

     Call PopupMenu(mnuPrint)

End Sub

Private Sub lvbutton6Close_Click()

     infomode = False
     findmode = False
     mode = True
     With frmMain
       .Enabled = True
       .SetFocus
       On Error Resume Next
       Kill App.Path & "\tmpfile.swf"
       LoadDataIntoFile 101, App.Path & "\tmpfile.swf"
       .WebBrowser1.Navigate App.Path & "\tmpfile.swf"
     End With
     Unload Me
     cnn.Close
     Set adoview = Nothing
     
End Sub

Private Sub lvbutton7MoveFirst_Click()
        
     If adoview.recordcount <= 1 Then Exit Sub
     adoview.MoveFirst
     Call loaddataforviewing
        
End Sub

Private Sub lvbutton8Previous_Click()
        
     If adoview.AbsolutePosition <= 1 Then Exit Sub
     adoview.MovePrevious
     Call loaddataforviewing
       
End Sub

Private Sub lvbutton9Next_Click()
        
     If adoview.AbsolutePosition >= adoview.recordcount Or adoview.recordcount <= 1 Then Exit Sub
     adoview.MoveNext
     Call loaddataforviewing
      
End Sub

Private Sub lvbutton10MoveLast_Click()
        
     If adoview.recordcount <= 1 Then Exit Sub
     adoview.MoveLast
     Call loaddataforviewing
  
End Sub

Private Sub mnuPrintAll_Click()
    
    On Error Resume Next
    If adoview.BOF = True And adoview.EOF = True Then
    MsgBox "Empty Database", vbInformation, "Address Book"
    Exit Sub
    Else
    Set jessiepanerioALL.DataSource = adoview
    jessiepanerioALL.Show
    End If

End Sub

Private Sub mnuPrintCurrentList_Click()
    
    On Error Resume Next
    If adoview.BOF = True And adoview.EOF = True Then
    MsgBox "Empty Database", vbInformation, "Address Book"
    Exit Sub
    Else
    adoview.Close
    Dim SQLTEXT As String
    SQLTEXT = "SELECT * FROM addressbook WHERE Left(lastname," & Len(lblLastName.Caption) & ")='" & lblLastName.Caption & "' And Left(nickname," & Len(lblNickName1.Caption) & ")='" & lblNickName1.Caption & "'"
    Set adoview = New ADODB.Recordset
    adoview.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
    Set jessiepanerioCL.DataSource = adoview
    jessiepanerioCL.Show
    Call rs_view
    End If

End Sub
     
Private Sub mnuPrintMale_Click()
    
    On Error Resume Next
    If adoview.BOF = True And adoview.EOF = True Then
    MsgBox "Empty Database", vbInformation, "Address Book"
    Exit Sub
    Else
    Set adoview = New ADODB.Recordset
    adoview.Open "Select * from addressbook WHERE gender = 'MALE';", cnn, adOpenStatic, adLockPessimistic
    Set jessiepanerioByGender.DataSource = adoview
    jessiepanerioByGender.Show
    End If
        
End Sub

Private Sub mnuPrintFemale_Click()

    On Error Resume Next
    If adoview.BOF = True And adoview.EOF = True Then
    MsgBox "Empty Database", vbInformation, "Address Book"
    Exit Sub
    Else
    Set adoview = New ADODB.Recordset
    adoview.Open "Select * from addressbook WHERE gender = 'FEMALE';", cnn, adOpenStatic, adLockPessimistic
    Set jessiepanerioByGender.DataSource = adoview
    jessiepanerioByGender.Show
    End If

End Sub

