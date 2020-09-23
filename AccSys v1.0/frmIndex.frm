VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIndex 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmIndex.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   7695
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCalculator 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Calculator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Frame FrameJournal 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   7335
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   12015
      Begin VB.TextBox txtTotalbal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   6720
         Width           =   1815
      End
      Begin VB.TextBox txtCrJrnl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   6000
         Width           =   1815
      End
      Begin VB.TextBox txtDrJrnl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   6000
         Width           =   1815
      End
      Begin VB.PictureBox PictureJrnl 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   12105
         TabIndex        =   38
         Top             =   0
         Width           =   12135
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Journal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   285
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   900
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridJrnl 
         Height          =   4095
         Left            =   1320
         TabIndex        =   39
         Top             =   1800
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColor       =   14209995
         ForeColor       =   0
         BackColorBkg    =   -2147483633
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
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Credit Total Rs."
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
         Left            =   6720
         TabIndex        =   78
         Top             =   6120
         Width           =   1485
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Debit Total Rs."
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
         Left            =   2400
         TabIndex        =   77
         Top             =   6120
         Width           =   1410
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmIndex.frx":030A
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
         Height          =   240
         Left            =   1320
         TabIndex        =   76
         Top             =   1440
         Width           =   9285
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL DIFFERENCE (Rs.)"
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
         Left            =   1200
         TabIndex        =   44
         Top             =   6720
         Width           =   2280
      End
      Begin VB.Label lblJrnlHead 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ledger A/c Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4575
         TabIndex        =   41
         Top             =   1200
         Width           =   1785
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ledger A/c details:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1320
         TabIndex        =   40
         Top             =   840
         Width           =   1890
      End
   End
   Begin VB.Frame FrameTB 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTBGo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Click for Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdExitTB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Return to Main"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   840
         Width           =   1695
      End
      Begin VB.PictureBox PictureTB 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   11985
         TabIndex        =   32
         Top             =   0
         Width           =   12015
         Begin VB.Label Label26 
            BackColor       =   &H00808080&
            Caption         =   "Basantapur Education Society"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4680
            TabIndex        =   35
            Top             =   120
            Width           =   5295
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Trial Balance of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2880
            TabIndex        =   34
            Top             =   120
            Width           =   1470
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Trial Balance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCrTB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   6360
         Width           =   2535
      End
      Begin VB.TextBox txtDrTB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   6360
         Width           =   2415
      End
      Begin VB.TextBox txtTBDate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4935
         Left            =   840
         TabIndex        =   23
         Top             =   1320
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   8705
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   14209995
         ForeColor       =   0
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
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
      Begin VB.Label lblTotTb 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL"
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
         Height          =   285
         Left            =   4440
         TabIndex        =   28
         Top             =   6480
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Trial Balance as on"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2040
         TabIndex        =   24
         Top             =   960
         Width           =   1800
      End
   End
   Begin VB.Frame FrameVoucher 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12015
      Begin VB.PictureBox PictureVch 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   12105
         TabIndex        =   29
         Top             =   0
         Width           =   12135
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Company Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3840
            TabIndex        =   31
            Top             =   120
            Width           =   3615
         End
         Begin VB.Label Label22 
            BackColor       =   &H00808080&
            Caption         =   "Voucher Entry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdVchModify 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ListBox lstVoucher 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4350
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtVchId 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cmbLdgName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4800
         TabIndex        =   11
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ComboBox cmbParti 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmIndex.frx":03A4
         Left            =   4800
         List            =   "frmIndex.frx":03A6
         TabIndex        =   10
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtAmt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4800
         TabIndex        =   9
         Top             =   3120
         Width           =   2415
      End
      Begin VB.OptionButton OpDr 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dr."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   3600
         Width           =   615
      End
      Begin VB.OptionButton OpCr 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cr."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtNarr 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   4800
         TabIndex        =   5
         Top             =   4560
         Width           =   4935
      End
      Begin VB.CommandButton cmdSaveVch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdVchEdit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdVchExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Press Modify button above and then select any account to Edit.Always save after each modification."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   795
         Left            =   240
         TabIndex        =   85
         Top             =   6240
         Width           =   3045
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   19
         Top             =   4200
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   18
         Top             =   3600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Voucher ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   17
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ledger Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   16
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Particulars"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   15
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amount     Rs."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   14
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Top             =   4560
         Width           =   795
      End
   End
   Begin VB.Frame FrameLedger 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   12015
      Begin VB.OptionButton OpCrL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cr."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5880
         TabIndex        =   81
         Top             =   4440
         Width           =   615
      End
      Begin VB.OptionButton OpDrL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dr."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   79
         Top             =   4440
         Width           =   615
      End
      Begin VB.ComboBox cmbUnder 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4920
         TabIndex        =   69
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox txtAcID 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4920
         TabIndex        =   68
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton cmdExitLdg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditLdg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveLdg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox txtOpBal 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4920
         TabIndex        =   57
         Top             =   3840
         Width           =   2655
      End
      Begin VB.ComboBox cmbNature 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4920
         TabIndex        =   56
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox txtAcName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4920
         TabIndex        =   55
         Top             =   2400
         Width           =   2655
      End
      Begin VB.ListBox lstLedger 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4155
         Left            =   360
         TabIndex        =   54
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton cmdModyLdg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdDispLdg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Display"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdLdgReset 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   5160
         Width           =   975
      End
      Begin VB.PictureBox PictureLdg 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   12585
         TabIndex        =   48
         Top             =   0
         Width           =   12615
         Begin VB.Label Label20 
            BackColor       =   &H00808080&
            Caption         =   "Ledger Creation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Company Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3960
            TabIndex        =   49
            Top             =   120
            Width           =   2970
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Balance Summary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1935
         Left            =   7800
         TabIndex        =   61
         Top             =   1800
         Width           =   2895
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1680
            TabIndex        =   67
            Top             =   1440
            Width           =   1170
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Difference    Rs."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   1440
            Width           =   1305
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            X1              =   120
            X2              =   2760
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1680
            TabIndex        =   65
            Top             =   840
            Width           =   1185
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1680
            TabIndex        =   64
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cr. Balance   Rs."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dr. Balance   Rs."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Label Label34 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select any account in the above List box and then press enter to view its corresponding Journal entries."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   240
         TabIndex        =   83
         Top             =   6240
         Width           =   3255
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Opening Bal. Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   80
         Top             =   4440
         Width           =   1500
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "A/c ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   74
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Opening Balance Rs."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   73
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nature"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   72
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Under"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   71
         Top             =   3000
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "A/c Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   70
         Top             =   2520
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   120
      Width           =   11655
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Press Escape button to close this form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7920
      TabIndex        =   84
      Top             =   7560
      Width           =   3750
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------mundaSoft Technologies product----------------------------
'Programmer-@gomesh
'Year-2005
'E-mail:(i)  g_munda@rediffmail.com
'       (ii) gomesh_p@yahoo.co.in
'Copyright(c)mundSoft Technologies--All Rights Reserved
'-----------------------------------------------------------------------
Option Explicit
Dim res                 As ADODB.Recordset
Dim resVch              As ADODB.Recordset
Dim DifDrCr(100000)     As Double
Dim TotaltbDr           As Double
Dim TotaltbCr           As Double
Dim stru                As Double
Dim j                   As Integer
Public dd               As String
Dim ee                  As String
Dim countT              As Integer

Private Sub cmbParti_KeyPress(keyascii As Integer)
    res.Close
    ee = UCase(Chr(keyascii)) + "%"            ' code for to find the Ldg names
    If Len(Trim(ee)) > 0 Then                  'on merely pressing  the 1st char and
   
        res.Open "SELECT * from Ledger where AcName LIKE '" & ee & "'"
        'res.Filter = "AcName LIKE '" & ee & "'" ' the phenomenon of droping down of the
        cmbParti.Clear                          'cmbLdgname box
      While Not res.EOF                       '
        cmbParti.AddItem (res!AcName)    '
        res.MoveNext
      Wend
        SendKeys "%+{DOWN}"
    End If

End Sub

Private Sub cmbLdgName_KeyPress(keyascii As Integer)
    res.Close
    dd = UCase(Chr(keyascii)) + "%"            ' code for to find the Ldg names
    If Len(Trim(dd)) > 0 Then                  'on merely pressing  the 1st char and
  
        res.Open "SELECT * from Ledger where AcName LIKE '" & dd & "'"
        ' res.Filter = "AcName LIKE '" & dd & "'" ' the phenomenon of droping down of the
        cmbLdgName.Clear                        ' cmbLdgname box
     While Not res.EOF                       '
        cmbLdgName.AddItem (res!AcName)     '
        res.MoveNext
     Wend
        SendKeys "%+{DOWN}"
End If

End Sub

Private Sub cmdCalculator_Click()
    Calculator.Show
End Sub

Private Sub cmdEditLdg_Click()
    'Set res = New ADODB.Recordset
     '   res.Open "Update Ledger set AcName='" & txtAcName.Text & "' where AcID='" & Me.txtAcID & "'", con, adOpenKeyset, adLockOptimistic
    '  res.Open "Update Ledger set Under='" & cmbUnder.Text & "' where AcID='" & txtAcID.Text & "'", con, adOpenKeyset, adLockOptimistic
    ' res.Open "Update Ledger set Nature='" & cmbNature.Text & "' where AcID='" & txtAcID.Text & "'", con, adOpenKeyset, adLockOptimistic
End Sub

Private Sub cmdExitLdg_Click()
    frmStarting.Show
    Unload Me
End Sub

Private Sub cmdExitTB_Click()
    frmStarting.Show
    Unload Me
End Sub

Private Sub cmdLdgReset_Click()
    txtAcName.Text = ""
    cmbUnder.Text = ""
    cmbNature.Text = ""
    txtAcName.SetFocus
End Sub

Private Sub cmdModyLdg_Click()
    lstLedger.Visible = True
    Frame4.Visible = False
End Sub

Private Sub cmdDispLdg_Click()
    lstLedger.Visible = True
    Frame4.Visible = True
End Sub

Private Sub cmdSaveLdg_Click()
    On Error Resume Next
    res.AddNew
    res!AcID = Trim(txtAcID.Text) 'IIf(Len() = 0, "", Trim(txtAcID.Text))
    res!AcName = Trim(txtAcName.Text) ' IIf(Len(Trim(txtAcName.Text)) = 0, "", Trim(txtAcName.Text))
    res!Under = Trim(cmbUnder.Text) 'IIf(Len(Trim(cmbUnder.Text)) = 0, "", Trim(cmbUnder.Text))
    res!Nature = Trim(cmbNature.Text) ' IIf(Len(Trim(cmbNature.Text)) = 0, "", Trim(cmbNature.Text))
    res!OpBal = Trim(txtOpBal.Text)
    If OpDrL.Value = True Then
        res!OpBalType = "Dr."
    Else
        res!OpBalType = "Cr."
    End If
        res!CompName = Label21.Caption
        res.Update
        txtAcID.Text = ""
        txtAcName.Text = ""
        cmbUnder.Text = ""
        cmbNature.Text = ""
        MsgBox "Success"
        Call populstLedger
        Call AutoIdGenerationLdg
        Call popCmbUnder
        txtAcName.SetFocus
End Sub

Private Sub cmdSaveVch_Click()
    On Error Resume Next
    Set resVch = New ADODB.Recordset
        resVch.Open "select * from Voucher", con, adOpenKeyset, adLockOptimistic
        resVch.AddNew
        resVch!VoucherId = Trim(txtVchId.Text) 'IIf(Len(Trim(txtVchId.Text)) = 0, "", Trim(txtVchId.Text))
        resVch!LdgName = Trim(cmbLdgName.Text) ' IIf(Len(Trim(cmbLdgName.Text)) = 0, "", Trim(cmbLdgName.Text))
        resVch!Particulars = Trim(cmbParti.Text) ' IIf(Len(Trim(cmbParti.Text)) = 0, "", Trim(cmbParti.Text))
        resVch!Amt = Trim(txtAmt.Text) ' IIf(Len(Trim(txtAmt.Text)) = 0, "", Trim(txtAmt.Text))
        resVch!Narration = Trim(txtNarr.Text) ' IIf(Len(Trim(txtNarr.Text)) = 0, "", Trim(txtNarr.Text))
        resVch!Date = Trim(txtDate.Text) ' IIf(Len(Trim(txtDate.Text)) = 0, "", Trim(txtDate.Text))
        resVch!CompanyName = Label23.Caption
        If OpDr.Value = True Then
            resVch!Type = "Dr."
        Else
            resVch!Type = "Cr."
        End If
'---Updating the F.Key(AcID) of Ledger table from Voucher table----------
    Set res = New ADODB.Recordset
        res.Open " select AcID from Ledger where AcName='" & cmbLdgName.Text & "'", con, adOpenKeyset, adLockOptimistic
        MsgBox cmbLdgName.Text
        stru = res!AcID  '<<<<<<----------------------+++++++++++
        resVch!AcID = stru
'--------------------------------------------------------
        resVch.Update
        txtVchId.Text = ""
        cmbLdgName.Text = ""
        cmbParti.Text = ""
        txtAmt.Text = ""
        OpDr.Refresh
        OpCr.Refresh
        txtNarr.Text = ""
        MsgBox "Success I have done it"
        'Call populstLedger
        Call AutoIdGenerationVch
        cmbLdgName.SetFocus
        'res.Close
End Sub

Private Sub cmdTBGo_Click()
        cmdTBGo.MousePointer = vbHourglass
        Call query
Dim k As Integer
Dim l As Integer
        For k = 0 To 2 Step 1
            res.MoveFirst
                For l = 1 To lstLedger.ListCount Step 1
                    MSFlexGrid1.Col = k
                    MSFlexGrid1.Row = l

                    If Not res.EOF And k = 0 Then
                        MSFlexGrid1.Text = res!AcName
                        res.MoveNext
                    ElseIf Not res.EOF And k = 1 And DifDrCr(l - 1) >= 0 Then
                        MSFlexGrid1.Text = Abs(DifDrCr(l - 1)) '<--- Abs()gives only +ve value
                        TotaltbDr = TotaltbDr + DifDrCr(l - 1) '<--- Cal the sum of total Dr in TB
                        res.MoveNext
                    ElseIf Not res.EOF And k = 2 And DifDrCr(l - 1) <= 0 Then
                        MSFlexGrid1.Text = Abs(DifDrCr(l - 1)) '<--- Abs()gives only +ve value
                        TotaltbCr = TotaltbCr + DifDrCr(l - 1) '<--- Cal the sum of total Cr in TB
                        res.MoveNext
                    End If
                Next l
        Next k
    txtDrTB.Visible = True
    txtCrTB.Visible = True
    lblTotTb.Visible = True
    txtDrTB.Text = Abs(TotaltbDr)
    txtCrTB.Text = Abs(TotaltbCr)
End Sub

Private Sub cmdVchExit_Click()
    frmStarting.Show
    Unload Me
End Sub

Private Sub cmdVchModify_Click()
    lstVoucher.Visible = True
    'Call populstVoucher
End Sub

Private Sub Form_Activate()
    Call popCmbUnder
    'Call popCmbVch
    Call populstLedger
    Call populstVoucher
End Sub

Private Sub Form_KeyPress(keyascii As Integer)
    If keyascii = 27 Then
        FrameLedger.Visible = True
        FrameJournal.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Label35.Visible = False
    cmdCalculator.Enabled = True
    FrameLedger.Visible = False
    FrameVoucher.Visible = False
    FrameJournal.Visible = False
    FrameTB.Visible = False
    txtDrTB.Visible = False
    txtCrTB.Visible = False
    lblTotTb.Visible = False
    lstVoucher.Visible = False
    cmbNature.AddItem "Assets"
    cmbNature.AddItem "Liability"
    cmbNature.AddItem "Income"
    cmbNature.AddItem "Expenses"

    txtDate.Text = Format(Now, "dd-mm-yyyy")
    txtTBDate.Text = Format(Now, "dd-mm-yyyy")

    Set res = New ADODB.Recordset
        res.Open "select * from Ledger", con, adOpenKeyset, adLockOptimistic

    Set resVch = New ADODB.Recordset
        resVch.Open "select * from Voucher", con, adOpenKeyset, adLockOptimistic
        Call AutoIdGenerationLdg
        Call AutoIdGenerationVch


Dim recNo As Integer
If res.RecordCount > 0 Then
    res.MoveLast
    recNo = res.RecordCount
    
'*******Initializing the FlexGrid******
Dim i As Integer
With MSFlexGrid1
     .Rows = recNo + 1
     .ColWidth(0) = 4900: .ColWidth(1) = 2400: .ColWidth(2) = 2400
     .Row = 0: .Col = 0: .Text = "         Account Name"
     .Row = 0: .Col = 1: .Text = "          AmountDr.(Rs.)"
     .Row = 0: .Col = 2: .Text = "          AmountCr.(Rs.)"
     
 End With
End If
'*******Initializing the FlexGridJournal***************
If resVch.RecordCount > 0 Then
   resVch.MoveLast
   countT = resVch.RecordCount
Dim p As Integer
   With MSFlexGridJrnl
       .Rows = countT + 1
       .ColWidth(0) = 2500: .ColWidth(1) = 2000: .ColWidth(2) = 2500: .ColWidth(3) = 2000
       .Row = 0: .Col = 0: .Text = "  Particulars"
       .Row = 0: .Col = 1: .Text = "      Dr. Amount (Rs.)"
       .Row = 0: .Col = 2: .Text = "  Particulars"
       .Row = 0: .Col = 3: .Text = "      Cr. Amount (Rs.)"
   End With
End If

End Sub


Public Function populstLedger()
    On Error Resume Next
    lstLedger.Clear
    Set res = New ADODB.Recordset
        res.Open "select * from Ledger", con, adOpenKeyset, adLockOptimistic
    
    Do While Not res.EOF
        lstLedger.AddItem res!AcName
        lstLedger.ItemData(lstLedger.NewIndex) = res.Fields(0).Value
        res.MoveNext
    Loop

End Function



Private Sub lstLedger_Click()
    res.MoveFirst
    res.Find "AcName='" & Trim(lstLedger.Text) & "'"
    Call Display
End Sub

Public Sub Display()
    txtAcID.Text = res!AcID
    txtAcName.Text = res!AcName
    cmbUnder.Text = res!Under
    cmbNature.Text = res!Nature
    txtOpBal.Text = res!OpBal
 '***** Code to find the sum of a particular field in a database**********
 Dim ADr, ACr, DifDrCr As Double
 Dim i As Integer
 For i = lstLedger.ListCount - 1 To 0 Step -1
 
 Set resVch = New ADODB.Recordset
    resVch.Open "select sum([Amt]) as totalDr from Voucher where LdgName='" & Trim(lstLedger.Text) & "'  and Type='Dr.' ", con, adOpenKeyset, adLockOptimistic
     
     Do While Not resVch.EOF
       If Not resVch!totalDr = "Null" Then
         Label8.Caption = resVch!totalDr
         ADr = resVch!totalDr
         resVch.MoveNext
       Else
         Label8.Caption = 0
         ADr = 0
        resVch.MoveNext
       End If
     Loop
     Set resVch = Nothing
    
    Set resVch = New ADODB.Recordset
     resVch.Open "select sum([Amt]) as totalCr from Voucher where LdgName='" & Trim(lstLedger.Text) & "'  and Type='Cr.' ", con, adOpenKeyset, adLockOptimistic
    
    Do While Not resVch.EOF
      If Not resVch!totalCr = "Null" Then
         Label9.Caption = resVch!totalCr
         ACr = resVch!totalCr
         resVch.MoveNext
      Else
         Label9.Caption = 0
         ACr = 0
         resVch.MoveNext
      End If
    Loop
 Set resVch = Nothing
 DifDrCr = ADr - ACr
 Label11.Caption = DifDrCr
Next i
End Sub

Public Sub populstVoucher()
On Error Resume Next
lstVoucher.Clear
Set res = New ADODB.Recordset
    res.Open "select * from Ledger", con, adOpenKeyset, adLockOptimistic
 
    Do While Not res.EOF
    lstVoucher.AddItem res!AcName
    lstVoucher.ItemData(lstLedger.NewIndex) = res.Fields(0).Value
    res.MoveNext
    Loop
End Sub
'--------------------------------------------------------------------------
Public Sub query()
 Dim str As Integer
 Dim ADr, ACr As Double
 Dim i As Integer
 For j = 0 To lstLedger.ListCount Step 1
 For i = lstLedger.ListCount - 1 To 0 Step -1
 lstLedger.Refresh
 Set resVch = New ADODB.Recordset
     resVch.Open "select sum([Amt]) as totalDr from Voucher where LdgName='" & Trim(lstLedger.List(j)) & "'  and Type='Dr.' ", con, adOpenKeyset, adLockOptimistic
     Do While Not resVch.EOF
     If Not resVch!totalDr = "Null" Then
        Label8.Caption = resVch!totalDr
        ADr = resVch!totalDr
        resVch.MoveNext
     Else
        Label8.Caption = 0
        ADr = 0
        resVch.MoveNext
     End If
     Loop
     
    Set resVch = Nothing
    
     Set resVch = New ADODB.Recordset
     resVch.Open "select sum([Amt]) as totalCr from Voucher where LdgName='" & Trim(lstLedger.List(j)) & "'  and Type='Cr.' ", con, adOpenKeyset, adLockOptimistic
    Do While Not resVch.EOF
    
    If Not resVch!totalCr = "Null" Then
       Label9.Caption = resVch!totalCr
       ACr = resVch!totalCr
       resVch.MoveNext
    Else
       Label9.Caption = 0
       ACr = 0
       resVch.MoveNext
    End If
    Loop
  Set resVch = Nothing
   DifDrCr(j) = ADr - ACr
   Label11.Caption = DifDrCr(j)
   str = DifDrCr(0)
Next i
Next j

End Sub

Public Sub popCmbUnder()
    On Error Resume Next
    cmbUnder.Clear
    cmbUnder.AddItem "#PRIMARY"
    Set res = New ADODB.Recordset
        res.Open "select * from Ledger", con, adOpenKeyset, adLockOptimistic
Do While Not res.EOF
   cmbUnder.AddItem res!AcName
   res.MoveNext
Loop
End Sub

Private Sub lstLedger_KeyPress(keyascii As Integer)
On Error Resume Next
Dim a, T1, T2 As Double

If keyascii = 13 Then
    cmdCalculator.Enabled = False
    FrameJournal.Visible = True
    MSFlexGridJrnl.Clear
    lblJrnlHead.Caption = "Ledger name : " + Trim(lstLedger.Text)
    frmIndex.Label35.Visible = True
    FrameJournal.BorderStyle = 0 '0=None
   Dim c, d, g, dd As Integer
   Set resVch = New ADODB.Recordset
       resVch.Open "SELECT * from Voucher", con, adOpenKeyset, adLockOptimistic
   For c = 0 To 3 Step 1
    
            resVch.MoveFirst
           
            d = 1

 If Not resVch.EOF And c = 0 Then
         
          Set resVch = New ADODB.Recordset
              resVch.Open "SELECT LdgName,Particulars,Type from Voucher where LdgName='" & lstLedger.Text & "' and Type='Dr.'", con, adOpenKeyset, adLockOptimistic
        
          While Not resVch.EOF And c = 0
            MSFlexGridJrnl.Col = c
            MSFlexGridJrnl.Row = d
          If Not resVch!Particulars = "NULL" Then
            MSFlexGridJrnl.Text = resVch!Particulars
          Else
            MSFlexGridJrnl.Text = ""
          End If
            d = d + 1
            resVch.MoveNext
          Wend
            MSFlexGridJrnl.Refresh
        
        
        ElseIf Not resVch.EOF And c = 1 Then
         
         Set resVch = New ADODB.Recordset
             resVch.Open "SELECT LdgName,Amt,Particulars from Voucher where LdgName='" & lstLedger.Text & "' and Type='Dr.'", con, adOpenKeyset, adLockOptimistic
        
          While Not resVch.EOF And c = 1
             MSFlexGridJrnl.Col = c
             MSFlexGridJrnl.Row = d
          If Not resVch!Amt = "NULL" Then
             MSFlexGridJrnl.Text = resVch!Amt
          Else
             MSFlexGridJrnl.Text = ""
          End If
             d = d + 1
             resVch.MoveNext
          Wend
             MSFlexGridJrnl.Refresh
        

        
        ElseIf Not resVch.EOF And c = 2 Then
             Set resVch = New ADODB.Recordset
             resVch.Open "SELECT LdgName,Particulars,Type from Voucher where LdgName='" & lstLedger.Text & "' and Type='Cr.'", con, adOpenKeyset, adLockOptimistic
        
        While Not resVch.EOF And c = 2
            MSFlexGridJrnl.Col = c
            MSFlexGridJrnl.Row = d
         If Not resVch!Particulars = "NULL" Then
            MSFlexGridJrnl.Text = resVch!Particulars
         Else
            MSFlexGridJrnl.Text = ""
         End If
            d = d + 1
            resVch.MoveNext
         Wend
            MSFlexGridJrnl.Refresh
'=======================
        ElseIf resVch.EOF And c = 2 Then
             Set resVch = New ADODB.Recordset
             resVch.Open "SELECT LdgName,Particulars,Type from Voucher where LdgName='" & lstLedger.Text & "' and Type='Cr.'", con, adOpenKeyset, adLockOptimistic
        
        While Not resVch.EOF And c = 2
            MSFlexGridJrnl.Col = c
            MSFlexGridJrnl.Row = d
         If Not resVch!Particulars = "NULL" Then
            MSFlexGridJrnl.Text = resVch!Particulars
         Else
            MSFlexGridJrnl.Text = ""
         End If
            d = d + 1
            resVch.MoveNext
         Wend
            MSFlexGridJrnl.Refresh

'=========================
          ElseIf Not resVch.EOF And c = 3 Then
            Set resVch = New ADODB.Recordset
            resVch.Open "SELECT LdgName,Amt,Type from Voucher where LdgName='" & Trim(lstLedger.Text) & "' and Type='Cr.'", con, adOpenKeyset, adLockOptimistic
         While Not resVch.EOF And c = 3
            MSFlexGridJrnl.Col = c
            MSFlexGridJrnl.Row = d
        If Not resVch!Amt = "NULL" Then
            MSFlexGridJrnl.Text = resVch!Amt
        Else
            MSFlexGridJrnl.Text = ""
       
       End If
            d = d + 1
            resVch.MoveNext
       Wend
            MSFlexGridJrnl.Refresh
  End If
 
Next c
Set resVch = New ADODB.Recordset
resVch.Open "select sum([Amt]) as totalDr from Voucher where LdgName='" & lstLedger.Text & "'  and Type='Dr.' ", con, adOpenKeyset, adLockOptimistic
If Not resVch!totalDr = "NULL" Then
    txtDrJrnl.Text = resVch!totalDr
    T1 = resVch!totalDr
Else
    T1 = 0
    txtDrJrnl.Text = 0
End If
Set resVch = Nothing
Set resVch = New ADODB.Recordset
     resVch.Open "select sum([Amt]) as totalCr from Voucher where LdgName='" & lstLedger.Text & "'  and Type='Cr.' ", con, adOpenKeyset, adLockOptimistic
If Not resVch!totalCr = "NULL" Then
     txtCrJrnl.Text = resVch!totalCr
     T2 = resVch!totalCr
Else
     T2 = 0
     txtCrJrnl.Text = 0
End If
    txtTotalbal.Text = Abs(T1 - T2)
    Set resVch = Nothing
End If
'----------make the form invisible on pressing of Esc button------
    If keyascii = 27 Then
        FrameJournal.Visible = False
        cmdCalculator.Enabled = True
    End If
End Sub

Private Sub lstVoucher_Click()
    Set resVch = New ADODB.Recordset
        resVch.Open "select * from Voucher", con, adOpenKeyset, adLockOptimistic
        resVch.MoveFirst
        resVch.Find "LdgName='" & Trim(lstVoucher.Text) & "'"
        Call Display2
End Sub

Public Sub Display2()
    txtVchId.Text = resVch!VoucherId
    cmbLdgName.Text = resVch!LdgName
    cmbParti.Text = resVch!Particulars
    txtAmt.Text = resVch!Amt
    If resVch!Type = "Dr." Then
        OpDr.Value = True
    Else
        OpCr.Value = True
    End If
        txtDate.Text = Format(resVch!Date, "dd-mm-yyyy")
        txtNarr.Text = resVch!Narration
End Sub


Public Sub AutoIdGenerationLdg()
    If res.RecordCount = 0 Then
        txtAcID.Text = "LDG" + Format(1, "000000")                                '
    Else                                        '<------Auto ID Generation
        res.MoveLast                                '
        txtAcID.Text = "LDG" + Format(Val(Right(Trim(res!AcID), 6)) + 1, "000000")
    End If
End Sub

Public Sub AutoIdGenerationVch()
    If resVch.RecordCount = 0 Then
        txtVchId.Text = "VCH" + Format(1, "000000")
    Else                                            '<------Auto ID Generation
        resVch.MoveLast
        txtVchId.Text = "VCH" + Format(Val(Right(Trim(resVch!VoucherId), 6)) + 1, "000000")
    End If
End Sub



Private Sub MSFlexGridJrnl_KeyPress(keyascii As Integer)
    If keyascii = 27 Then
        MsgBox "XXXXXXXXXXX"
        FrameJournal.Visible = False
        'Unload Me
    End If
End Sub



