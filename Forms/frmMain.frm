VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Software Inventory System of IICT Department"
   ClientHeight    =   9075
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   11955
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnutransaction 
      Caption         =   "&File"
      Begin VB.Menu mnuclient 
         Caption         =   "User Account"
         Shortcut        =   ^N
      End
      Begin VB.Menu uhhhm 
         Caption         =   "-"
      End
      Begin VB.Menu sdsds 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "&Transaction"
      Begin VB.Menu mnubdp 
         Caption         =   "View Software Profiles"
      End
      Begin VB.Menu ghtreee 
         Caption         =   "-"
      End
      Begin VB.Menu dfdfdfd 
         Caption         =   "Search Profiles"
      End
   End
   Begin VB.Menu mnuuser 
      Caption         =   "&View"
      Begin VB.Menu mnuadd 
         Caption         =   "Summary of Profiles"
      End
      Begin VB.Menu dsfteyye 
         Caption         =   "-"
      End
      Begin VB.Menu fdfghrr 
         Caption         =   "Individual Profiles"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "&Help"
      Begin VB.Menu dteeew 
         Caption         =   "About the System"
      End
      Begin VB.Menu olpk 
         Caption         =   "-"
      End
      Begin VB.Menu tyiot 
         Caption         =   "Company Information"
      End
      Begin VB.Menu pqmf 
         Caption         =   "-"
      End
      Begin VB.Menu yiopfd 
         Caption         =   "About the Developers/Author"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dfdfdfd_Click()
Screen.MousePointer = vbHourglass
frmsearch.Show
Screen.MousePointer = vbNormal
End Sub

Private Sub dteeew_Click()
Screen.MousePointer = vbHourglass
frmabout.Show
Screen.MousePointer = vbNormal
End Sub

Private Sub fdfghrr_Click()
Screen.MousePointer = vbHourglass
frmsummary.Show
Screen.MousePointer = vbNormal
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you sure you want to exit?", vbInformation + vbYesNo + vbOKOnly, "Service Monitoring System") = vbYes Then
End
End If
End Sub

Private Sub mnuadd_Click()
Dim zz As New ADODB.Recordset

zz.Open "SELECT * FROM tblsoftware", conn, adOpenStatic, adLockReadOnly

If zz.RecordCount <> 0 Then

    Set rptreport.DataSource = zz
        rptreport.WindowState = vbMaximized
        rptreport.Show
Else

    MsgBox "No record available.", vbExclamation, "Report"
End If
End Sub

Private Sub mnubdp_Click()
Screen.MousePointer = vbHourglass
frmNewClient.Show
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuclient_Click()
Screen.MousePointer = vbHourglass
frmUser.Show
Screen.MousePointer = vbNormal
End Sub

Private Sub sdsds_Click()
If MsgBox("Are you sure you want to exit?", vbInformation + vbYesNo + vbOKOnly, "Software Inventory System of IICT Department") = vbYes Then
End
End If
End Sub

Private Sub tyiot_Click()
Screen.MousePointer = vbHourglass
frmcompany.Show
Screen.MousePointer = vbNormal
End Sub

Private Sub yiopfd_Click()
Screen.MousePointer = vbHourglass
frmdeveloper.Show
Screen.MousePointer = vbNormal
End Sub
