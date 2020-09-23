VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackupDB 
   Caption         =   "Backup Database"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   Icon            =   "frmBackupDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBackupDB.frx":030A
   ScaleHeight     =   4440
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
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
      Left            =   5160
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtBkpUpFName 
      Height          =   285
      Left            =   3960
      TabIndex        =   8
      Top             =   2040
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   4320
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.Timer TimerBkupDB 
      Interval        =   400
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtBkUpDest 
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   2775
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBkDB 
      Caption         =   "Backup Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Picture         =   "frmBackupDB.frx":78F4C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database Backup Successful"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   3240
      TabIndex        =   12
      Top             =   2640
      Width           =   2400
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter File Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   2040
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path Selected"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Destination of Backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3045
      TabIndex        =   4
      Top             =   0
      Width           =   2805
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Database everyday .Select a Drive,a folder and enter a file name (or use the default name) and then press the button"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   3960
      Width           =   4095
   End
End
Attribute VB_Name = "frmBackupDB"
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
Dim FSO                     As Object
Dim BkpUpFileName           As String

Private Sub cmdBkDB_Click()
BkpUpFileName = "" + Me.txtBkUpDest.Text + "\" + Me.txtBkpUpFName.Text + ".mdb"
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.copyfile App.Path & "\Database\Accounts1.mdb", BkpUpFileName
PrgBar.Visible = True
Me.Drive1.SetFocus
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
Me.txtBkUpDest.Text = "" & Dir1.Path
End Sub

Private Sub Dir1_Click()
cmdBkDB.Enabled = True
End Sub

Private Sub Drive1_Change()
Dim Obj1, Obj2 As Object
Set Obj2 = CreateObject("Scripting.FileSystemObject")
Set Obj1 = Obj2.getdrive(Obj2.getdrivename(Drive1.Drive))

If Obj1.isready Then
    Dir1.Path = Drive1.Drive
    Dir1.SetFocus
Else
    MsgBox "DRIVE  NOT READY!", vbCritical
End If
End Sub
Private Sub Form_Load()
Label6.Visible = False
cmdBkDB.Enabled = False
PrgBar.Visible = False
Me.Drive1.Refresh
Me.Dir1.Refresh
Me.txtBkpUpFName.Text = "AccSysBkUp_" + Format$(Now, "dd-mm-yyyy")
End Sub

Private Sub TimerBkupDB_Timer()
PrgBar.Value = PrgBar.Value + 5
If PrgBar.Value = 100 Then
    Label6.Visible = True
    Label6.Caption = "Database Backup Successful"
    TimerBkupDB.Enabled = False
    PrgBar.Visible = False
End If
'Label6.Visible = False
End Sub
