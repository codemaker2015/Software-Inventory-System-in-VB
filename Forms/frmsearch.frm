VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Record"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11430
   Icon            =   "frmsearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11430
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   11400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   11430
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
         Left            =   600
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " Type Software Name Here..."
         Top             =   360
         Width           =   5535
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search  Name"
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Click Here to Search a Record"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear Search"
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
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click Here to Reload the Records"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdclose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
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
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Click Here to Close this Form"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   10935
      Begin MSComctlLib.ListView lstCustomerList 
         Height          =   4215
         Left            =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12632256
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Software No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Software Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Software Serial #"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Company"
            Object.Width           =   7056
         EndProperty
      End
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClearSearch_Click()
    txtSearch.Text = " Type Software Name Here..."
    Call LoadCustomer
End Sub

Private Sub cmdSearch_Click()
If Not Me.txtSearch.Text = " Type Software Name Here..." Then

        LocalSQL = ""
        LocalSQL = "SELECT * FROM tblsoftware WHERE SoftwareName LIKE '%" & Replace(Trim(Me.txtSearch.Text), "'", "''") & "%' ORDER BY SoftwareNo"
        
        If rs.State = 1 Then rs.Close
        rs.Open LocalSQL, conn
        
        lstCustomerList.ListItems.Clear
        
        While Not rs.EOF
        
        Set ListItem = lstCustomerList.ListItems.Add(, , IIf(IsNull(rs!SoftwareNo) = True, "", rs!SoftwareNo))
        ListItem.ListSubItems.Add , , IIf(IsNull(rs!SoftwareName) = True, "", rs!SoftwareName)
        ListItem.ListSubItems.Add , , IIf(IsNull(rs!SerialNo) = True, "", rs!SerialNo)
        ListItem.ListSubItems.Add , , IIf(IsNull(rs!Company) = True, "", rs!Company)
    
            rs.MoveNext
            
        Wend

End If
End Sub
Public Function LoadCustomer()

LocalSQL = "SELECT * FROM tblsoftware ORDER BY SoftwareNo"

If rs.State = 1 Then rs.Close
rs.Open LocalSQL, conn

lstCustomerList.ListItems.Clear

While Not rs.EOF

        Set ListItem = lstCustomerList.ListItems.Add(, , IIf(IsNull(rs!SoftwareNo) = True, "", rs!SoftwareNo))
        ListItem.ListSubItems.Add , , IIf(IsNull(rs!SoftwareName) = True, "", rs!SoftwareName)
        ListItem.ListSubItems.Add , , IIf(IsNull(rs!SerialNo) = True, "", rs!SerialNo)
        ListItem.ListSubItems.Add , , IIf(IsNull(rs!Company) = True, "", rs!Company)
    
    rs.MoveNext
    
Wend

End Function
Private Sub Cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Gitna Me
Call LoadCustomer
End Sub

Private Sub txtSearch_GotFocus()
txtSearch.SelStart = 0
txtSearch.SelLength = Len(txtSearch.Text)
End Sub
