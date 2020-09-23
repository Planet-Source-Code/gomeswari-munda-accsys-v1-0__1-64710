VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRestoreDB 
   Caption         =   "Restore Database"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   Icon            =   "frmRestoreDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRestoreDB.frx":030A
   ScaleHeight     =   3450
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   2640
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   3360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtRestbkUp 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin MSComctlLib.ProgressBar PBar2 
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdRestoreDb 
      Caption         =   "Restore Database"
      Height          =   855
      Left            =   5520
      Picture         =   "frmRestoreDB.frx":56D4C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Datsbase Restoration Successful"
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
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Destinationto Restore a Database"
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
      Left            =   1965
      TabIndex        =   2
      Top             =   240
      Width           =   4005
   End
End
Attribute VB_Name = "frmRestoreDB"
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
Dim RdestS          As String
Dim RestbkpDBF      As Object

Private Sub cmdBrowse_Click()
CDialog.ShowOpen
RdestS = CDialog.FileName
txtRestbkUp = CDialog.FileName
End Sub

Private Sub cmdRestoreDb_Click()
Set RestbkpDBF = CreateObject("Scripting.FileSystemObject")
RestbkpDBF.copyfile RdestS, "" & App.Path & "\Database\Accounts1.mdb"
PBar2.Visible = True
End Sub

Private Sub Form_Load()
PBar2.Visible = False
Label2.Visible = False
End Sub


Private Sub Timer1_Timer()
PBar2.Value = PBar2.Value + 5
If PBar2.Value = 100 Then
    Label2.Visible = True
    PBar2.Visible = False
End If
End Sub
