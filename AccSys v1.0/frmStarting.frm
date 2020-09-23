VERSION 5.00
Begin VB.Form frmStarting 
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "frmStarting.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmStarting.frx":030A
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerTime 
      Interval        =   500
      Left            =   2160
      Top             =   7440
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   8175
      Left            =   0
      TabIndex        =   8
      Top             =   -120
      Width           =   12015
      Begin VB.Frame FrameCreateComp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FFFFC0&
         Height          =   4575
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   7935
         Begin VB.TextBox txtPeriod2 
            Appearance      =   0  'Flat
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
            Left            =   3840
            TabIndex        =   22
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtPeriod1 
            Appearance      =   0  'Flat
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
            Left            =   1920
            TabIndex        =   21
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CommandButton cmdExitCrtComp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Exit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3840
            Width           =   2175
         End
         Begin VB.CommandButton cmdSaveCrtComp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sa&ve"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3840
            Width           =   2295
         End
         Begin VB.TextBox txtFYr 
            Appearance      =   0  'Flat
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
            Left            =   1920
            TabIndex        =   18
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox txtSTno 
            Appearance      =   0  'Flat
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
            Left            =   5280
            TabIndex        =   17
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtITNo 
            Appearance      =   0  'Flat
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
            Left            =   1920
            TabIndex        =   16
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtCAdd 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   1200
            Width           =   5295
         End
         Begin VB.TextBox txtCName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1920
            TabIndex        =   14
            Top             =   720
            Width           =   5295
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Create Company Profile"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   240
            Left            =   2565
            TabIndex        =   46
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label Label6 
            BackColor       =   &H00808080&
            Height          =   495
            Left            =   0
            TabIndex        =   45
            Top             =   120
            Width           =   7935
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   3360
            TabIndex        =   29
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Period                   :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   240
            TabIndex        =   28
            Top             =   3360
            Width           =   1755
         End
         Begin VB.Label Label13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Financial Year  from  :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   240
            TabIndex        =   27
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sales Tax No.    :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3720
            TabIndex        =   26
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "PAN No. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Address                :             :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   240
            TabIndex        =   24
            Top             =   1200
            Width           =   2160
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Company Name :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   5175
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   8175
         Begin VB.Timer TimerWarning 
            Interval        =   500
            Left            =   240
            Top             =   4200
         End
         Begin VB.Frame FrameSelComp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H8000000F&
            Height          =   3135
            Left            =   120
            TabIndex        =   40
            Top             =   1440
            Width           =   7935
            Begin VB.ComboBox cmbSelComp 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   1440
               TabIndex        =   42
               Top             =   1440
               Width           =   4335
            End
            Begin VB.CommandButton cmdSelCompOK 
               BackColor       =   &H00C0C0C0&
               Caption         =   "&OK"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
               Caption         =   "Select a Company"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2490
               TabIndex        =   43
               Top             =   360
               Width           =   2205
            End
            Begin VB.Label Label5 
               BackColor       =   &H00808080&
               Height          =   735
               Left            =   0
               TabIndex        =   44
               Top             =   120
               Width           =   7935
            End
         End
         Begin VB.Image imgWarning 
            Height          =   480
            Left            =   1200
            Picture         =   "frmStarting.frx":103B1C
            Top             =   4200
            Width           =   480
         End
         Begin VB.Label lblWarning 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   " Please Select Company Name and then press OK !"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   1710
            TabIndex        =   48
            Top             =   4320
            Width           =   5595
         End
         Begin VB.Label lblTime 
            BackColor       =   &H00000000&
            Caption         =   "Label6"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   0
            TabIndex        =   39
            Top             =   4680
            Width           =   8175
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   240
            TabIndex        =   38
            Top             =   960
            Width           =   3915
         End
         Begin VB.Label lblPeriod1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "00-00-0000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4830
            TabIndex        =   37
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6120
            TabIndex        =   36
            Top             =   960
            Width           =   225
         End
         Begin VB.Label lblPeriod2 
            AutoSize        =   -1  'True
            Caption         =   "00-00-0000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   35
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Selected Company Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   855
            TabIndex        =   32
            Top             =   240
            Width           =   2385
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Current Period"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   5550
            TabIndex        =   31
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label17 
            BackColor       =   &H00808080&
            Height          =   615
            Left            =   0
            TabIndex        =   33
            Top             =   120
            Width           =   8175
         End
         Begin VB.Label Label21 
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   7935
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   11745
         TabIndex        =   11
         Top             =   720
         Width           =   11775
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
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
            Height          =   1335
            Left            =   3240
            Picture         =   "frmStarting.frx":103F5E
            ScaleHeight     =   1335
            ScaleWidth      =   5175
            TabIndex        =   47
            Top             =   0
            Width           =   5175
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   15
            Left            =   960
            Picture         =   "frmStarting.frx":1465A0
            Top             =   480
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   14
            Left            =   720
            Picture         =   "frmStarting.frx":1466A2
            Top             =   480
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   13
            Left            =   480
            Picture         =   "frmStarting.frx":1467A4
            Top             =   480
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   12
            Left            =   240
            Picture         =   "frmStarting.frx":1468A6
            Top             =   480
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   11
            Left            =   0
            Picture         =   "frmStarting.frx":1469A8
            Top             =   480
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   10
            Left            =   2400
            Picture         =   "frmStarting.frx":146AAA
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   9
            Left            =   2160
            Picture         =   "frmStarting.frx":146BAC
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   8
            Left            =   1920
            Picture         =   "frmStarting.frx":146CAE
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   7
            Left            =   1680
            Picture         =   "frmStarting.frx":146DB0
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   6
            Left            =   1440
            Picture         =   "frmStarting.frx":146EB2
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   5
            Left            =   1200
            Picture         =   "frmStarting.frx":146FB4
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   4
            Left            =   960
            Picture         =   "frmStarting.frx":1470B6
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   3
            Left            =   720
            Picture         =   "frmStarting.frx":1471B8
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   480
            Picture         =   "frmStarting.frx":1472BA
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   240
            Picture         =   "frmStarting.frx":1473BC
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   0
            Picture         =   "frmStarting.frx":1474BE
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   5175
         Left            =   8400
         TabIndex        =   9
         Top             =   2760
         Width           =   3495
         Begin VB.CommandButton cmdIE 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Income and Expenditure A/c"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3840
            Width           =   3135
         End
         Begin VB.CommandButton cmdSelComp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Select Company"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   840
            Width           =   3135
         End
         Begin VB.CommandButton cmdBalSht 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Balance Sheet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   4440
            Width           =   3135
         End
         Begin VB.CommandButton cmdRP 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Receipts and Payments A/c"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3240
            Width           =   3135
         End
         Begin VB.CommandButton cmdTB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Trial Balance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2640
            Width           =   3135
         End
         Begin VB.CommandButton cmdVch 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Voucher"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2040
            Width           =   3135
         End
         Begin VB.CommandButton cmdLdg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Ledger"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1440
            Width           =   3135
         End
         Begin VB.CommandButton cmdCrtComp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Create Company"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Label Label19 
         BackColor       =   &H00808080&
         Caption         =   " Copyright (c)  2004--2005 mundSoft Technologies."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   11775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   11775
      End
   End
End
Attribute VB_Name = "frmStarting"
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
Dim OpBalCash               As Double
Dim OpBalBank               As Double
Dim TotBalCash              As Double
Dim TotBalBank              As Double
Dim Dt                      As Date
Dim per1                    As String
Dim per2                    As String
Dim val1                    As Double
Dim val2                    As Double
Dim RPADr                   As Double
Dim RPACr                   As Double
Dim resComp                 As ADODB.Recordset 'for Company table
Dim res                     As ADODB.Recordset     'for Ledger table
Dim resVch                  As ADODB.Recordset  'for Voucher table
Dim resRP                   As ADODB.Recordset
Dim res1                    As ADODB.Recordset    'for Income & Expenditure
Dim st                      As String
Dim str1                    As String
Dim i                       As Integer
Dim j                       As Integer
Dim Y                       As Integer
Dim z                       As Integer
Dim DiffDrCr                As Double
Dim TotalDrRP               As Double
Dim TotalCrRP               As Double
Dim TotalDrRPSep1           As Double
Dim TotalCrRPSep            As Double
Dim RPADrSep                As Double
Dim RPACrSep                As Double
Dim RPADrSep1               As Double
Dim Values                  As Double
Dim temp                    As Double
Dim recNum                  As Integer


Private Sub cmdBalSht_Click()
    frmBalanceSheet.Show
    frmBalanceSheet.FrameIE.Visible = False
    frmBalanceSheet.FrameBS.Visible = True
    frmBalanceSheet.Frame1.Visible = False
    frmBalanceSheet.Label12.Caption = cmbSelComp.Text + " as on  " + Format(Now, "dd-mm-yyyy")
End Sub

Private Sub cmdBalSht_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub cmdCrtComp_Click()
    FrameCreateComp.Visible = True
Label17.Visible = False
End Sub

Private Sub cmdCrtComp_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub cmdExitCrtComp_Click()
    FrameCreateComp.Visible = False
Label17.Visible = True
End Sub

Private Sub cmdIE_Click()
    frmBalanceSheet.Show
    frmBalanceSheet.FrameBS.Visible = False
    frmBalanceSheet.Frame1.Visible = False
    frmBalanceSheet.Frame3.Visible = False
    Call CalIE
End Sub

Private Sub cmdIE_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub cmdLdg_Click()
    frmIndex.Show
    frmIndex.FrameLedger.Visible = True
    frmIndex.FrameLedger.BorderStyle = 0 '0=None
    frmIndex.FrameVoucher.Visible = False
    frmIndex.Label21.Caption = Label4.Caption
End Sub

Private Sub cmdLdg_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub cmdRP_Click()
frmBalanceSheet.Frame2.Visible = False
    Call FlexRP
    Call Item2
    Call FinalVal
End Sub

Private Sub cmdRP_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub cmdSaveCrtComp_Click()
Dim i As Integer
    resComp.AddNew
    resComp!CompName = txtCName.Text
    resComp!Add = txtCAdd.Text
    resComp!ITaxNo = txtITNo.Text
    resComp!STaxNo = txtSTno.Text
    resComp!FinancialYr = txtFYr.Text
    resComp!Period1 = txtPeriod1.Text
    resComp!Period2 = txtPeriod2.Text
    resComp.Update
    MsgBox "Company Information saved", vbInformation, "Save"
Label17.Visible = True
End Sub

Private Sub cmdSelComp_Click()
    FrameSelComp.Visible = True
    'Picture1.Visible = False
    cmbSelComp.Clear
    Call popuComp
End Sub

Private Sub cmdSelComp_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub cmdSelCompOK_Click()
TimerWarning.Enabled = False
lblWarning.Visible = False
imgWarning.Visible = False
    FrameSelComp.Visible = False
   ' Picture1.Visible = True
    Set resComp = New ADODB.Recordset
    resComp.Open "SELECT Period1,Period2 from Company where CompName='" & cmbSelComp.Text & "'", con, adOpenKeyset, adLockOptimistic
        If Not cmbSelComp.Text = "" Then
            Label4.Caption = cmbSelComp.Text
            lblPeriod1.Caption = Format(resComp!Period1, "dd-mm-yyyy")
            lblPeriod2.Caption = Format(resComp!Period2, "dd-mm-yyyy")
        Else
        Exit Sub
        End If
MDIForm1.mnuL.Enabled = True
MDIForm1.mnuV.Enabled = True
MDIForm1.mnuTB.Enabled = True
MDIForm1.mnuRP.Enabled = True
MDIForm1.mnuIE.Enabled = True
MDIForm1.mnuBS.Enabled = True

cmdLdg.Enabled = True
cmdVch.Enabled = True
cmdTB.Enabled = True
cmdRP.Enabled = True
cmdIE.Enabled = True
cmdBalSht.Enabled = True

End Sub

Private Sub cmdTB_Click()
    frmIndex.Show
    frmIndex.FrameTB.Visible = True
    frmIndex.FrameTB.BorderStyle = 0 '0=NOne
    frmIndex.FrameVoucher.Visible = False
    frmIndex.Label26.Caption = Label4.Caption + "  as on  " + Format(Now, "dd-mm-yyyy")
    frmIndex.FrameJournal.Visible = False
End Sub

Private Sub cmdTB_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub cmdVch_Click()
    frmIndex.Show
    frmIndex.FrameVoucher.Visible = True
    frmIndex.FrameVoucher.BorderStyle = 0 '0=None
    frmIndex.FrameLedger.Visible = False
    frmIndex.Label23.Caption = Label4.Caption
End Sub

Private Sub cmdVch_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub Form_Activate()
    Call popuComp
End Sub


Private Sub Form_KeyPress(keyascii As Integer)
    Call MainEscp(27)
End Sub

Private Sub Form_Load()
    cmdLdg.Enabled = False
    cmdVch.Enabled = False
    cmdTB.Enabled = False
    cmdRP.Enabled = False
    cmdIE.Enabled = False
    cmdBalSht.Enabled = False
    'Label7.Caption = Format(Now, "dd-mm-yyyy")
    FrameCreateComp.Visible = False
    FrameSelComp.Visible = False
        Set resComp = New ADODB.Recordset
            resComp.Open "SELECT * from Company", con, adOpenKeyset, adLockOptimistic
End Sub

Public Sub popuComp()
    On Error Resume Next
        Set resComp = New ADODB.Recordset
        resComp.Open "SELECT * from Company", con, adOpenKeyset, adLockOptimistic
    Do While Not resComp.EOF
        cmbSelComp.AddItem resComp!CompName
        resComp.MoveNext
    Loop
End Sub

Public Sub Calculate()
    Set resVch = New ADODB.Recordset
        resVch.Open "select sum([Amt]) as TotalDrRP from Voucher where LdgName='" & st & "'  and Type='Dr.' ", con, adOpenKeyset, adLockOptimistic
     Do While Not resVch.EOF
        If Not resVch!TotalDrRP = "Null" Then
            RPADr = resVch!TotalDrRP
        Else
            RPADr = 0
        End If
        resVch.MoveNext
      Loop

    Set resVch = Nothing

    Set resVch = New ADODB.Recordset
        resVch.Open "select sum([Amt]) as TotalCrRP from Voucher where LdgName= '" & st & "'  and Type='Cr.' ", con, adOpenKeyset, adLockOptimistic
     Do While Not resVch.EOF
        If Not resVch!TotalCrRP = "Null" Then
            RPACr = resVch!TotalCrRP
        Else
            RPACr = 0
        End If
            resVch.MoveNext
      Loop

    Set resVch = Nothing

        DiffDrCr = RPADr - RPACr
        temp = DiffDrCr
End Sub

'--------Initializing FlexGrid of Receipits & Payments Accounts-------------
Public Sub FlexRP()
        frmBalanceSheet.Show
        frmBalanceSheet.FrameBS.Visible = False
        frmBalanceSheet.FrameIE.Visible = False
        frmBalanceSheet.lblRP = Label4.Caption
'----------------------------------------
    Set res = New ADODB.Recordset
        res.Open "select * from Ledger", con, adOpenKeyset, adLockOptimistic

    If res.RecordCount > 0 Then
        res.MoveLast
        recNum = res.RecordCount
   
    Dim n As Integer
    With frmBalanceSheet.MSFlexGridRP
        .Rows = recNum + 1
        .ColWidth(0) = 3000: .ColWidth(1) = 1500: .ColWidth(2) = 3000: .ColWidth(3) = 1500
        .Row = 0: .Col = 0: .Text = "Receipts"
        .Row = 0: .Col = 1: .Text = "Amount (Rs.)"
        .Row = 0: .Col = 2: .Text = "Payments"
        .Row = 0: .Col = 3: .Text = "Amount (Rs.)"
        
    End With
    End If
    
End Sub

Public Sub Item2()
Dim c As Integer 'c=col
Dim r As Integer 'r=row

    For c = 0 To 3 Step 1
         res.MoveFirst
           r = 1
                frmBalanceSheet.MSFlexGridRP.Col = c
                frmBalanceSheet.MSFlexGridRP.Row = r

        If Not res.EOF And c = 0 Then
        '--------------------------CASH-----------------------------
                res.Close
                res.Open "SELECT AcName from Ledger where AcName='Cash A/c' ", con, adOpenKeyset, adLockOptimistic
           
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                frmBalanceSheet.MSFlexGridRP.Text = "To," + res!AcName
                st = res!AcName
                res.MoveNext
                r = r + 1
            Wend
        '--------------------------BANK-----------------------------
                res.Close
                res.Open "SELECT AcName from Ledger where AcName='Bank A/c' ", con, adOpenKeyset, adLockOptimistic
           
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                frmBalanceSheet.MSFlexGridRP.Text = "To," + res!AcName
                st = res!AcName
                res.MoveNext
                r = r + 1
            Wend
        '------------------------REST(Income)-----------------------
                res.Close
                res.Open "SELECT AcName from Ledger where Nature='Income' ", con, adOpenKeyset, adLockOptimistic
           
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                frmBalanceSheet.MSFlexGridRP.Text = "To," + res!AcName
                st = res!AcName
                res.MoveNext
                r = r + 1
            Wend
        
       End If
      '=======================================================================================================
       If Not res.EOF And c = 1 Then
                res.Close
                res.Open "SELECT AcName,OpBal,OpBalType from Ledger where AcName='Cash A/c' ", con, adOpenKeyset, adLockOptimistic
            
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                st = res!AcName
            
            If Not res!OpBal = "NULL" Then
                frmBalanceSheet.MSFlexGridRP.Text = res!OpBal '<--printing the corresponding values
            Else
                frmBalanceSheet.MSFlexGridRP.Text = 0
            End If
                res.MoveNext
                r = r + 1
            Wend
                
                
                res.Close
                res.Open "SELECT AcName,OpBal,OpBalType from Ledger where AcName='Bank A/c' ", con, adOpenKeyset, adLockOptimistic
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                st = res!AcName
             
            If Not res!OpBal = "NULL" Then
                frmBalanceSheet.MSFlexGridRP.Text = res!OpBal '<--printing the corresponding values
            Else
                frmBalanceSheet.MSFlexGridRP.Text = 0
            End If
                res.MoveNext
                r = r + 1
            Wend
            
                
                
                res.Close
                res.Open "SELECT AcName from Ledger where Nature='Income' ", con, adOpenKeyset, adLockOptimistic
            
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                st = res!AcName
                Call Calculate
                frmBalanceSheet.MSFlexGridRP.Text = DiffDrCr '<--printing the corresponding values
                res.MoveNext
                r = r + 1
            Wend
      
       End If
      '========================================================================================================
       If Not res.EOF And c = 2 Then
                res.Close
                res.Open "SELECT AcName from Ledger where Nature='Expenses' ", con, adOpenKeyset, adLockOptimistic
            
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                frmBalanceSheet.MSFlexGridRP.Text = "By," + res!AcName
                st = res!AcName
                res.MoveNext
                r = r + 1
            Wend
                res.Close
                res.Open "SELECT AcName from Ledger where AcName='Cash A/c' ", con, adOpenKeyset, adLockOptimistic
           
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                frmBalanceSheet.MSFlexGridRP.Text = "By," + res!AcName
                st = res!AcName
                res.MoveNext
                r = r + 1
            Wend
                res.Close
                res.Open "SELECT AcName from Ledger where AcName='Bank A/c' ", con, adOpenKeyset, adLockOptimistic
           
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                frmBalanceSheet.MSFlexGridRP.Text = "By," + res!AcName
                st = res!AcName
                res.MoveNext
                r = r + 1
            Wend
                  
                  
       End If
      '===========================================================================================================
       If Not res.EOF And c = 3 Then
                res.Close
                res.Open "SELECT AcName from Ledger where Nature='Expenses' ", con, adOpenKeyset, adLockOptimistic
            
            While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                st = res!AcName
                Call Calculate
                frmBalanceSheet.MSFlexGridRP.Text = DiffDrCr  '<--printing the corresponding values
                res.MoveNext
                r = r + 1
            Wend
            
                res.Close
                res.Open "SELECT AcName from Ledger where AcName='Cash A/c' ", con, adOpenKeyset, adLockOptimistic
            'While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                st = res!AcName
                Call CalculateSpecial
                frmBalanceSheet.MSFlexGridRP.Text = TotBalCash '<--printing the corresponding values
                r = r + 1
            'Wend
                res.Close
                res.Open "SELECT AcName from Ledger where AcName='Bank A/c' ", con, adOpenKeyset, adLockOptimistic
            'While Not res.EOF
                frmBalanceSheet.MSFlexGridRP.Row = r
                st = res!AcName
                Call CalculateSpecial1
                frmBalanceSheet.MSFlexGridRP.Text = TotBalBank '<--printing the corresponding values
               'r = r + 1
            'Wend
           
       End If
  
Next c

End Sub
Public Sub UpdateRP()
    Set resRP = New ADODB.Recordset
        resRP.Open "select * from RP", con, adOpenKeyset, adLockOptimistic
        resRP.AddNew
        resRP!ActNameI = st
        resRP!ActNameP = st
        resRP!Expenses = DiffDrCr
        resRP!Income = DiffDrCr
End Sub

Public Sub MainEscp(keyascii As Integer)
 Dim resp As Integer
    If keyascii = 27 Then
       resp = MsgBox("Are you sure you want to exit ?", vbYesNo + vbCritical, "Quit")
       If resp = vbYes Then
            Unload MDIForm1
            
       ElseIf resp = vbNo Then
            Exit Sub
       End If
    End If
End Sub

'This Public subroutine is used to calculate the TOTAL of (Cash & Bank A/c's)
'in the Receipts and Payments A/c
Public Sub CalculateSpecial()
    Set resComp = New ADODB.Recordset
        resComp.Open "SELECT * from Company", con, adOpenKeyset, adLockOptimistic
                            
        Dt = Format(Date, "dd-mm-yyyy")
        per1 = resComp!Period1 '3/31/2005
        per2 = resComp!Period2 '3/31/2006
        res.Close

        res.Open "SELECT * from Ledger where AcName='" & st & "'", con, adOpenKeyset, adLockOptimistic
       
      If Not res!OpBal = "NULL" Then
        OpBalCash = res!OpBal
      Else
        OpBalCash = 0
      End If
    
    Set resVch = New ADODB.Recordset
    resVch.Open "SELECT sum([Amt]) as TotalDrRPSep from Voucher where LdgName='" & st & "' and Type='Dr.' and date between #" & per1 & "# and #" & per2 & "# ", con, adOpenKeyset, adLockOptimistic
    Do While Not resVch.EOF
        If Not resVch!TotalDrRPSep = "Null" Then
            RPADrSep = resVch!TotalDrRPSep
        Else
            RPADrSep = 0
        End If
            resVch.MoveNext
    Loop
    TotBalCash = Abs(OpBalCash + RPADrSep)

End Sub


Public Sub CalculateSpecial1()
    Set resComp = New ADODB.Recordset
        resComp.Open "SELECT * from Company", con, adOpenKeyset, adLockOptimistic
                            
        Dt = Format(Date, "dd-mm-yyyy")
        per1 = resComp!Period1 '3/31/2005
        per2 = resComp!Period2 '3/31/2006
    Set res = New ADODB.Recordset
        res.Open "SELECT * from Ledger where AcName='" & st & "'", con, adOpenKeyset, adLockOptimistic
   
    If Not res!OpBal = "NULL" Then
        OpBalBank = res!OpBal
    Else
        OpBalBank = 0
    End If
    
    Set resVch = New ADODB.Recordset
        resVch.Open "SELECT sum([Amt]) as TotalDrRPSep1 from Voucher where LdgName='" & st & "' and Type='Dr.' and date between #" & per1 & "# and #" & per2 & "# ", con, adOpenKeyset, adLockOptimistic

        Do While Not resVch.EOF
           If Not resVch!TotalDrRPSep1 = "Null" Then
              RPADrSep1 = resVch!TotalDrRPSep1
           Else
              RPADrSep1 = 0
           End If
           resVch.MoveNext
        Loop
     TotBalBank = Abs(OpBalBank + RPADrSep1)

End Sub
'This subroutine is used to calculate the total of Receipts and Payments
'made in the Receipts and payments A/c
Public Sub FinalVal()
Dim ColNo, RowNo As Integer
Dim RecpTot, PayTot, DiffTot As Double
    For ColNo = 1 To 3 Step 2
        If ColNo = 1 Then
            For RowNo = 1 To recNum
                frmBalanceSheet.MSFlexGridRP.Row = RowNo
                frmBalanceSheet.MSFlexGridRP.Col = ColNo
                RecpTot = RecpTot + Val(frmBalanceSheet.MSFlexGridRP.Text)
            Next
        End If
                
        If ColNo = 3 Then
            frmBalanceSheet.MSFlexGridRP.Col = ColNo
            For RowNo = 1 To recNum
                frmBalanceSheet.MSFlexGridRP.Row = RowNo
                PayTot = PayTot + Val(frmBalanceSheet.MSFlexGridRP.Text)
            Next
        End If
    Next
                
        frmBalanceSheet.txtRecp.Text = RecpTot
        frmBalanceSheet.txtPay.Text = PayTot
        DiffTot = Abs(RecpTot - PayTot)
        frmBalanceSheet.Text1.Text = DiffTot
        
End Sub
'In Income & Expenditure A/c include natures Income and Expenses only and exclude the rest
Public Sub CalIE()

Set res = New ADODB.Recordset
    res.Open "Select AcName,Nature,OpBal,OpBalType,CompName from Ledger where Nature='Expenses'", con, adOpenKeyset, adLockOptimistic
    Do While Not res.EOF
        frmBalanceSheet.cmbAcNameExp.AddItem res!AcName
        res.MoveNext
    Loop

Set res1 = New ADODB.Recordset
    res1.Open "Select AcName,Nature,OpBal,OpBalType,CompName from Ledger where Nature='Income'", con, adOpenKeyset, adLockOptimistic
    Do While Not res1.EOF
        frmBalanceSheet.cmbAcNameInc.AddItem res1!AcName
        res1.MoveNext
    Loop

End Sub

Private Sub TimerTime_Timer()
lblTime.Caption = "   Time:" & Time & "                      Day:" & Format(Now, "dddd") & "                                   Date:" & Format(Now, "dd-mm-yyyy")
End Sub

Private Sub TimerWarning_Timer()
lblWarning.Visible = Not lblWarning.Visible
imgWarning.Visible = Not imgWarning.Visible
End Sub
