VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmcompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About the Company"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7245
   Icon            =   "frmcompany.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7245
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Goals"
      TabPicture(0)   =   "frmcompany.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDisclaimer(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Objectives"
      TabPicture(1)   =   "frmcompany.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDisclaimer(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblDisclaimer(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblDisclaimer(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmcompany.frx":0044
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1065
         Index           =   1
         Left            =   -74520
         TabIndex        =   4
         Top             =   840
         Width           =   5895
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmcompany.frx":00E1
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Index           =   2
         Left            =   -74520
         TabIndex        =   3
         Top             =   1920
         Width           =   5895
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   "3. To develop the student on a particular area of strenght in sub area of IT or in some complementary area of study."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1065
         Index           =   3
         Left            =   -74520
         TabIndex        =   2
         Top             =   3000
         Width           =   5895
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmcompany.frx":01A1
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3705
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmcompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Gitna Me
End Sub
