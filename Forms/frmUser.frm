VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Maintenance"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6255
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   0
      ScaleHeight     =   1770
      ScaleWidth      =   6225
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5520
         Picture         =   "frmUser.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Save Entry"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   4860
         Picture         =   "frmUser.frx":130E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Enter New Record"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   240
         Picture         =   "frmUser.frx":2610
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   6015
      Begin MSComctlLib.ListView lstCustomerList 
         Height          =   2055
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3625
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "User Name"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   6015
      Begin VB.TextBox txtVerifyPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtUserName 
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
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtUserNo 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -120
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -120
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iselected As Integer
Private Sub cmdNew_Click()
    Call GetUserNo
    Call ClearDetails
End Sub
Private Sub cmdSave_Click()
If Me.txtUserName.Text = "" Or Me.txtPassword.Text = "" Then

    MsgBox "Please complete your entries !        ", vbExclamation
    Exit Sub

End If

        LocalSQL = "SELECT * FROM tblUser  WHERE userno = " & Trim(Me.txtUserNo.Text)
        
        If rs.State = 1 Then rs.Close
        rs.Open LocalSQL, conn
        
        If rs.EOF Then

            LocalSQL = "INSERT INTO tblUser VALUES (" & CDbl(Me.txtUserNo.Text) & ",'" & Replace(Trim(txtUserName.Text), "'", "''") & "','" _
            & Replace(Trim(Me.txtPassword.Text), "'", "''") & "')"
            
            conn.Execute LocalSQL
    
            MsgBox "User successfully saved!        ", vbInformation
        
            Call GetUserNo
            Call LoadUser
            Call ClearDetails

        Else
        
            LocalSQL = "UPDATE tblUser " _
            & "SET txtusername = '" & Replace(Trim(Me.txtUserName.Text), "'", "''") & "', txtpassword = '" & Replace(Trim(Me.txtPassword.Text), "'", "''") & "' WHERE userno = " & Replace(Trim(Me.txtUserNo.Text), "'", "''")
        
            conn.Execute LocalSQL
    
            MsgBox "User successfully updated!        ", vbInformation
        
            Call GetUserNo
            Call LoadUser
            Call ClearDetails
        
        End If
End Sub



Private Sub Form_Load()
Gitna Me
    Call LoadUser
    Call GetUserNo
End Sub

Public Function LoadUser()

LocalSQL = "SELECT * FROM tblUser ORDER BY userno"

If rs.State = 1 Then rs.Close
rs.Open LocalSQL, conn

lstCustomerList.ListItems.Clear

While Not rs.EOF

    Set ListItem = lstCustomerList.ListItems.Add(, , IIf(IsNull(rs!userno) = True, "", rs!userno))
    ListItem.ListSubItems.Add , , IIf(IsNull(rs!UserName) = True, "", rs!UserName)
    
    rs.MoveNext
    
Wend

End Function
Public Function GetUserNo()
    
    LocalSQL = "SELECT MAX(userno) FROM tblUser"

    If rs.State = 1 Then rs.Close
    rs.Open LocalSQL, conn
    
    If Not rs.EOF Then
    
            If IsNull(rs(0)) = True Then
                txtUserNo.Text = "1"
            Else
                txtUserNo.Text = rs(0) + 1
            End If
            
    Else
        txtUserNo.Text = "1"
    End If

End Function
Private Sub lstCustomerList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    iselected = Item.Index
End Sub
Private Sub lstCustomerList_DblClick()
    
    If iselected > O Then
       

            Me.txtUserNo.Text = lstCustomerList.ListItems(iselected).Text
            
            LocalSQL = "SELECT * " _
                     & "FROM tblUser where userno = " & Replace(Trim(Me.txtUserNo.Text), "'", "''")
                     
            If rs.State = 1 Then rs.Close
            rs.Open LocalSQL, conn
        
            If Not rs.EOF Then
        
                txtUserName.Text = rs!txtUserName
                txtPassword.Text = rs!txtPassword
                txtVerifyPassword.Text = rs!txtPassword
   
            End If
    End If
End Sub
Public Function ClearDetails()
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtVerifyPassword.Text = ""
End Function


