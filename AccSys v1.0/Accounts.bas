Attribute VB_Name = "Module1"
'---------------------mundSoft Technologies product----------------------------
'Programmer-@gomesh
'Year-2005
'E-mail:(i)  g_munda@rediffmail.com
'       (ii) gomesh_p@yahoo.co.in
'Copyright (c) mundSoft Technologies -- All Rights Reserved
'-----------------------------------------------------------------------
Option Explicit
Public res          As ADODB.Recordset
Public con          As ADODB.Connection
Public mpath        As String

Public Sub Main()
    mpath = App.Path
        Set con = New ADODB.Connection
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mpath + "\Database\Accounts1.mdb"

        frmStarting.Show
End Sub


Public Sub Escape(Index As Integer, keyascii As Integer)
    If keyascii = 27 Then
      '  Unload Me
    End If
End Sub
