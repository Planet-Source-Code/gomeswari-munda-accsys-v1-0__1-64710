VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About AccSys v1.0"
   ClientHeight    =   4605
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5985
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3178.453
   ScaleMode       =   0  'User
   ScaleWidth      =   5620.225
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   2400
      Top             =   2040
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   10005
      Left            =   0
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   6984.706
      ScaleMode       =   0  'User
      ScaleWidth      =   8048.74
      TabIndex        =   1
      Top             =   0
      Width           =   11520
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   1200
         TabIndex        =   7
         Top             =   -120
         Width           =   5175
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   4080
            Width           =   855
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acknowledgements"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   1815
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   3135
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   $"frmAbout.frx":173338
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Left            =   120
               TabIndex        =   10
               Top             =   480
               Width           =   3060
            End
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Special Thanks To"
            Height          =   735
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Keep Screen Resolution at 800 X 600 pixels"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   150
            Left            =   120
            TabIndex        =   23
            Top             =   4440
            Width           =   3030
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mundSoft Technologies"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1680
            TabIndex        =   22
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Copyright (c) 2004 -- 2005"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   $"frmAbout.frx":1733F8
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   20
            Top             =   3600
            Width           =   3135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "mundSoft Technologies"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Left            =   1680
            TabIndex        =   19
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "All  Rights  Reserved."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1200
            TabIndex        =   18
            Top             =   960
            Width           =   1320
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "AccSys v1.0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1080
            TabIndex        =   17
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Programming : Gomeswari Munda     (@gomesh)"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   3405
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Contact No. : 09830840824 "
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Or E-mail me at"
            Height          =   195
            Left            =   1200
            TabIndex        =   14
            Top             =   2160
            Width           =   1065
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "g _munda@rediffmail.com"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   840
            TabIndex        =   13
            Top             =   2400
            Width           =   1830
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "gomesh_p@yahoo.co.in"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   840
            TabIndex        =   12
            Top             =   2640
            Width           =   1725
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   3585
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   4035
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.399
      Y1              =   2350.192
      Y2              =   2350.192
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1650
      TabIndex        =   3
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1650
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.399
      Y1              =   2277.719
      Y2              =   2277.719
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1650
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   135
      TabIndex        =   4
      Top             =   3585
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------mundSoft Technologies product----------------------------
'Programmer-@gomesh
'Year-2005
'E-mail:(i)  g_munda@rediffmail.com
'       (ii) gomesh_p@yahoo.co.in
'Copyright (c) mundSoft Technologies -- All Rights Reserved
'-----------------------------------------------------------------------
Option Explicit
Private Sub Command1_Click()
Frame2.Visible = True
Label8.Visible = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Frame2.Visible = False
End Sub

Private Sub Timer1_Timer()
Frame2.Visible = False
End Sub

