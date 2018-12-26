VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Software Information"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
   Icon            =   "frmNewClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmNewClient.frx":000C
   ScaleHeight     =   7905
   ScaleWidth      =   11385
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Customer Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   11175
      Begin VB.TextBox txt3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1440
         Width           =   9135
      End
      Begin VB.TextBox txt2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   2
         Top             =   960
         Width           =   9135
      End
      Begin VB.TextBox txt1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   480
         Width           =   9135
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Software Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Software CD Key"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   11355
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   11385
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   10020
         Picture         =   "frmNewClient.frx":150DE
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Enter New Record"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   10680
         Picture         =   "frmNewClient.frx":163E0
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Save Entry"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtCustomerCode 
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   3300
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Software Code:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   1425
         Left            =   240
         Picture         =   "frmNewClient.frx":176E2
         Top             =   0
         Width           =   1470
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Current Records - Double click on the record to load the details and edit it"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   11175
      Begin VB.TextBox txtSearch 
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
         Left            =   120
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " Type Software Name Here..."
         Top             =   360
         Width           =   5535
      End
      Begin VB.CommandButton cmdSearchLastName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Last Name"
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1815
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin MSComctlLib.ListView lstCustomerList 
         Height          =   3375
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5953
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
Attribute VB_Name = "frmNewClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iselected As Integer

Private Sub cmdClearSearch_Click()

    txtSearch.Text = " Type Software Name Here..."
    Call LoadCustomer

End Sub

Private Sub cmdNew_Click()
    Call ClearDetails
    Call GetClientNo
    Call LoadCustomer
End Sub

Private Sub cmdSave_Click()

If Me.txtCustomerCode.Text = "" Or Me.txt1.Text = "" Or Me.txt2.Text = "" Or Me.txt3.Text = "" _
Then

    MsgBox "Please complete your entries !        ", vbExclamation
    Exit Sub

End If

        LocalSQL = "SELECT * FROM tblsoftware  WHERE SoftwareNo = " & Trim(Me.txtCustomerCode.Text)
        
        If rs.State = 1 Then rs.Close
        rs.Open LocalSQL, conn
        
        
        If rs.EOF Then

            LocalSQL = "INSERT INTO tblsoftware VALUES('" & Replace(Trim(txtCustomerCode.Text), "'", "''") & "','" _
            & Replace(Trim(Me.txt1.Text), "'", "''") & "','" _
            & Replace(Trim(Me.txt2.Text), "'", "''") & "','" _
            & Replace(Trim(Me.txt3.Text), "'", "''") & "')"
            
            conn.Execute LocalSQL
    
            MsgBox "Software Information successfully saved!        ", vbInformation
        
            Call GetClientNo
            Call LoadCustomer
            Call ClearDetails

        Else
        
            LocalSQL = "UPDATE tblsoftware " _
            & "SET SoftwareName = '" & Replace(Trim(Me.txt1.Text), "'", "''") & "', " _
            & "SerialNo = '" & Replace(Trim(Me.txt2.Text), "'", "''") & "', Company = '" & Replace(Trim(Me.txt3.Text), "'", "''") & "' WHERE SoftwareNo = " & Replace(Trim(Me.txtCustomerCode.Text), "'", "''")

            conn.Execute LocalSQL
    
            MsgBox "Software Information successfully updated!        ", vbInformation
        
            Call GetClientNo
            Call LoadCustomer
            Call ClearDetails
        
        End If
End Sub

Private Sub cmdSearchLastName_Click()
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

Private Sub Form_Load()
    Call GetClientNo
    Call LoadCustomer
    Call ClearDetails
    
    Gitna Me
    
End Sub

Private Sub txtSearch_GotFocus()
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
End Sub
Public Sub ClearDetails()
    Me.txt1.Text = ""
    Me.txt2.Text = ""
    Me.txt3.Text = ""

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

Public Function GetClientNo()
    
    LocalSQL = "SELECT MAX(SoftwareNo) FROM tblsoftware"

    If rs.State = 1 Then rs.Close
    rs.Open LocalSQL, conn
    
    If Not rs.EOF Then
    
            If IsNull(rs(0)) = True Then
                txtCustomerCode.Text = "1"
            Else
                txtCustomerCode.Text = rs(0) + 1
            End If
            
    Else
        txtCustomerCode.Text = "1"
    End If

End Function

Private Sub lstCustomerList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    iselected = Item.Index
End Sub

Private Sub lstCustomerList_DblClick()
    
    If iselected > O Then
       

            Me.txtCustomerCode.Text = lstCustomerList.ListItems(iselected).Text
            
            LocalSQL = "SELECT * " _
                     & "FROM tblsoftware where SoftwareNo = " & Replace(Trim(Me.txtCustomerCode.Text), "'", "''")
                     
            If rs.State = 1 Then rs.Close
            rs.Open LocalSQL, conn
        
            If Not rs.EOF Then
        
                txt1.Text = Trim(rs!SoftwareName)
                txt2.Text = Trim(rs!SerialNo)
                txt3.Text = Trim(rs!Company)
   
            End If
            
    End If
    
End Sub




