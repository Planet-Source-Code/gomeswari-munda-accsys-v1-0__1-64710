VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "AccSys v1.0"
   ClientHeight    =   4095
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8910
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picClose 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      Picture         =   "MDIForm1.frx":030A
      ScaleHeight     =   0
      ScaleWidth      =   8880
      TabIndex        =   1
      Top             =   0
      Width           =   8910
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   16828
            MinWidth        =   16828
            Text            =   "AccSys V1.0  Copyright (c) mundSoft Technologies"
            TextSave        =   "AccSys V1.0  Copyright (c) mundSoft Technologies"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1833
            MinWidth        =   1833
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "     E&xit"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuAcc 
      Caption         =   "Accounts"
      Begin VB.Menu mnuCC 
         Caption         =   "     &Create Company"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSc 
         Caption         =   "     &Select Company"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuL 
         Caption         =   "     &Ledger"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuV 
         Caption         =   "     &Voucher"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTB 
         Caption         =   "     &Trial Balance"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRP 
         Caption         =   "     &Receipts and Payments A/c"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIE 
         Caption         =   "     &Income and Expenditure A/c"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBS 
         Caption         =   "     &Balance Sheet"
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "Reports"
   End
   Begin VB.Menu mnuDb 
      Caption         =   "&DataBase"
      Begin VB.Menu mnuBD 
         Caption         =   "      Backup Database"
      End
      Begin VB.Menu mnusep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRD 
         Caption         =   "      Restore Database"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utility"
      Begin VB.Menu mnuNP 
         Caption         =   "      NotePad"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWE 
         Caption         =   "      Windows Explorer"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOSK 
         Caption         =   "      On Screen Keyboard"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuUM 
         Caption         =   "     User Manual"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAA 
         Caption         =   "     About AccSys"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------mundSoft Technologies product-------------------------------
'Programmer-@gomesh
'Year-2005
'E-mail:(i)  g_munda@rediffmail.com
'       (ii) gomesh_p@yahoo.co.in
'Copyright (c) mundSoft Technologies -- All Rights Reserved
'--------------------------------------------------------------------------
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
    ByVal hBitmapChecked As Long) As Long
Const MF_BITMAP = 4

Private Sub MDIForm_Load()
Dim MenuH1, SMenuH1, MenuId1 As Long
MenuH1 = GetMenu(Me.hWnd)
SMenuH1 = GetSubMenu(MenuH1, 0)
frmStarting.Image1(0).Picture = frmStarting.Image1(0).Picture
MenuId1 = GetMenuItemID(SMenuH1, 0)
Call SetMenuItemBitmaps(MenuH1, MenuId1, MF_BITMAP, frmStarting.Image1(0).Picture, frmStarting.Image1(0).Picture)
'------------------------------------------------------------
Dim MenuH2, SMenuH2, MenuId2 As Long
MenuH2 = GetMenu(Me.hWnd)
SMenuH2 = GetSubMenu(MenuH2, 2)
frmStarting.Image1(3).Picture = frmStarting.Image1(3).Picture
MenuId2 = GetMenuItemID(SMenuH2, 0)
Call SetMenuItemBitmaps(MenuH2, MenuId2, MF_BITMAP, frmStarting.Image1(3).Picture, frmStarting.Image1(3).Picture)

frmStarting.Image1(4).Picture = frmStarting.Image1(4).Picture
MenuId2 = GetMenuItemID(SMenuH2, 2)
Call SetMenuItemBitmaps(MenuH2, MenuId2, MF_BITMAP, frmStarting.Image1(4).Picture, frmStarting.Image1(4).Picture)

frmStarting.Image1(5).Picture = frmStarting.Image1(5).Picture
MenuId2 = GetMenuItemID(SMenuH2, 4)
Call SetMenuItemBitmaps(MenuH2, MenuId2, MF_BITMAP, frmStarting.Image1(5).Picture, frmStarting.Image1(5).Picture)

frmStarting.Image1(6).Picture = frmStarting.Image1(6).Picture
MenuId2 = GetMenuItemID(SMenuH2, 6)
Call SetMenuItemBitmaps(MenuH2, MenuId2, MF_BITMAP, frmStarting.Image1(6).Picture, frmStarting.Image1(6).Picture)

frmStarting.Image1(7).Picture = frmStarting.Image1(7).Picture
MenuId2 = GetMenuItemID(SMenuH2, 8)
Call SetMenuItemBitmaps(MenuH2, MenuId2, MF_BITMAP, frmStarting.Image1(7).Picture, frmStarting.Image1(7).Picture)

frmStarting.Image1(8).Picture = frmStarting.Image1(8).Picture
MenuId2 = GetMenuItemID(SMenuH2, 10)
Call SetMenuItemBitmaps(MenuH2, MenuId2, MF_BITMAP, frmStarting.Image1(8).Picture, frmStarting.Image1(8).Picture)

frmStarting.Image1(9).Picture = frmStarting.Image1(9).Picture
MenuId2 = GetMenuItemID(SMenuH2, 12)
Call SetMenuItemBitmaps(MenuH2, MenuId2, MF_BITMAP, frmStarting.Image1(9).Picture, frmStarting.Image1(9).Picture)

frmStarting.Image1(10).Picture = frmStarting.Image1(10).Picture
MenuId2 = GetMenuItemID(SMenuH2, 14)
Call SetMenuItemBitmaps(MenuH2, MenuId2, MF_BITMAP, frmStarting.Image1(10).Picture, frmStarting.Image1(10).Picture)
'---------------------------------------------------------------------------------------------------------------------------
Dim MenuH3, SMenuH3, MenuId3 As Long
MenuH3 = GetMenu(Me.hWnd)
SMenuH3 = GetSubMenu(MenuH3, 4)
frmStarting.Image1(11).Picture = frmStarting.Image1(11).Picture
MenuId3 = GetMenuItemID(SMenuH3, 0)
Call SetMenuItemBitmaps(MenuH3, MenuId3, MF_BITMAP, frmStarting.Image1(11).Picture, frmStarting.Image1(11).Picture)

frmStarting.Image1(12).Picture = frmStarting.Image1(12).Picture
MenuId3 = GetMenuItemID(SMenuH3, 2)
Call SetMenuItemBitmaps(MenuH3, MenuId3, MF_BITMAP, frmStarting.Image1(12).Picture, frmStarting.Image1(12).Picture)
'--------------------------------------------------------------------------------------------------------------------------
Dim MenuH4, SMenuH4, MenuId4 As Long
MenuH4 = GetMenu(Me.hWnd)
SMenuH4 = GetSubMenu(MenuH4, 5)
frmStarting.Image1(13).Picture = frmStarting.Image1(13).Picture
MenuId4 = GetMenuItemID(SMenuH4, 0)
Call SetMenuItemBitmaps(MenuH4, MenuId4, MF_BITMAP, frmStarting.Image1(13).Picture, frmStarting.Image1(13).Picture)

frmStarting.Image1(14).Picture = frmStarting.Image1(14).Picture
MenuId4 = GetMenuItemID(SMenuH4, 2)
Call SetMenuItemBitmaps(MenuH4, MenuId4, MF_BITMAP, frmStarting.Image1(14).Picture, frmStarting.Image1(14).Picture)

frmStarting.Image1(15).Picture = frmStarting.Image1(15).Picture
MenuId4 = GetMenuItemID(SMenuH4, 4)
Call SetMenuItemBitmaps(MenuH4, MenuId4, MF_BITMAP, frmStarting.Image1(15).Picture, frmStarting.Image1(15).Picture)
'---------------------------------------------------------------------------------------------------------------------
Dim MenuH5, SMenuH5, MenuId5 As Long
MenuH5 = GetMenu(Me.hWnd)
SMenuH5 = GetSubMenu(MenuH5, 6)
frmStarting.Image1(1).Picture = frmStarting.Image1(1).Picture
MenuId5 = GetMenuItemID(SMenuH5, 0)
Call SetMenuItemBitmaps(MenuH5, MenuId5, MF_BITMAP, frmStarting.Image1(1).Picture, frmStarting.Image1(1).Picture)

frmStarting.Image1(2).Picture = frmStarting.Image1(2).Picture
MenuId5 = GetMenuItemID(SMenuH5, 2)
Call SetMenuItemBitmaps(MenuH5, MenuId5, MF_BITMAP, frmStarting.Image1(2).Picture, frmStarting.Image1(2).Picture)

mnuL.Enabled = False
mnuV.Enabled = False
mnuTB.Enabled = False
mnuRP.Enabled = False
mnuIE.Enabled = False
mnuBS.Enabled = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Quit ?", vbYesNo Or vbQuestion) = vbNo Then
        Cancel = True
        frmStarting.Show
    Else
        Call SlideDown
    End If
    Exit Sub
End Sub

Private Sub mnuAA_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuBD_Click()
frmBackupDB.Show vbModal, Me
'frmBackupDB.txtBkUpDest.SetFocus
End Sub

Private Sub mnuBS_Click()
frmBalanceSheet.FrameBS.Visible = True
End Sub

Private Sub mnuCC_Click()
frmStarting.FrameCreateComp.Visible = True
frmStarting.Label17.Visible = False
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub
Public Sub SlideDown()
Dim i As Integer
    Me.WindowState = 0
    For i = 1 To Me.Height
        picClose.Height = picClose.Height / 2 + i
        DoEvents
    Next i
    
    For i = 1 To Me.Width
        Me.Top = Me.Top + i * Screen.TwipsPerPixelX
        DoEvents
    Next i

End Sub



Private Sub mnuIE_Click()
frmBalanceSheet.FrameIE.Visible = True
End Sub

Private Sub mnuL_Click()
frmIndex.FrameLedger.Visible = True
End Sub

Private Sub mnuNP_Click()
On Error Resume Next
Shell ("notepad"), vbNormalFocus
Exit Sub
End Sub

Private Sub mnuOSK_Click()
On Error Resume Next
Shell ("osk"), vbNormalFocus
Exit Sub
End Sub

Private Sub mnuRD_Click()
frmRestoreDB.Show vbModal, Me
End Sub

Private Sub mnuRP_Click()
frmBalanceSheet.Frame1.Visible = True 'Frame1->Receipts & payments frame
frmBalanceSheet.FrameBS.Visible = False
frmBalanceSheet.FrameIE.Visible = False
frmBalanceSheet.Frame2.Visible = False
frmBalanceSheet.Frame3.Visible = False

End Sub

Private Sub mnuSc_Click()
frmStarting.FrameSelComp.Visible = True

End Sub

Private Sub mnuTB_Click()
frmIndex.FrameTB.Visible = True
End Sub

Private Sub mnuV_Click()
frmIndex.FrameVoucher.Visible = True
End Sub

Private Sub mnuWE_Click()
On Error Resume Next
Shell ("explorer"), vbMaximizedFocus
Exit Sub
End Sub
