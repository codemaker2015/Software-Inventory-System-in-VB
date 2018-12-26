VERSION 5.00
Begin VB.Form frmsummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Individual Profiles"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4800
   Icon            =   "frmsummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print Preview"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click Here to Preview the Report"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      MaxLength       =   50
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   " Type Software Name Here..."
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmsummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
Dim zz As New ADODB.Recordset

            LocalSQL = "SELECT *" _
                        & " FROM tblsoftware" _
                        & " WHERE SoftwareName LIKE '%" & Trim(txtSearch.Text) & "%'" _
                        & " ORDER BY SoftwareName"

If zz.State = 1 Then zz.Close
   zz.Open LocalSQL, conn

If zz.RecordCount <> 0 Then

    Set rptreport.DataSource = zz
        rptreport.WindowState = vbMaximized
        rptreport.Show
Else

    MsgBox "No record available.", vbExclamation, "Report"

End If
End Sub

Private Sub Form_Load()
Gitna Me
End Sub
Private Sub txtSearch_GotFocus()
txtSearch.SelStart = 0
txtSearch.SelLength = Len(txtSearch.Text)
End Sub
