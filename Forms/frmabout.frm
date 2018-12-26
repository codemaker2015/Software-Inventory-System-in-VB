VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About the System"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5130
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5130
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   9240
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   9480
      Top             =   120
   End
   Begin VB.CommandButton cmdclose 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00C1AB7D&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to Close this Window"
      Top             =   2520
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   9360
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   8640
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmabout.frx":000C
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "All Rights Reserved. Copyright 2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   5205
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9000
      TabIndex        =   6
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   5160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Software Inventory System"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   480
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   4485
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1.0.0"
      Height          =   225
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmabout.frx":0316
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4755
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Gitna Me
End Sub

