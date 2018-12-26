VERSION 5.00
Begin VB.Form frmConnection 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.Timer Timer2 
         Interval        =   2000
         Left            =   3360
         Top             =   1680
      End
      Begin VB.Timer Timer1 
         Interval        =   2500
         Left            =   3840
         Top             =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Opening connection to database . . ."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call Connect
End Sub
Private Sub Timer1_Timer()
frmlogin.Show vbModal
End Sub

Private Sub Timer2_Timer()
Label1.Caption = "Connection successful !!!"
End Sub
