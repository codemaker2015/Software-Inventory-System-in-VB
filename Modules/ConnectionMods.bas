Attribute VB_Name = "ConnectionMods"
Global conn As New ADODB.Connection
Global rs As New ADODB.Recordset
Global LocalSQL As String

Function Connect()

If conn.State = 1 Then
Else
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & App.Path & "\Software.mdb;"
End If

If conn.State = 1 Then

    frmConnection.Label1.Caption = "Opening connection to database . . ."

End If
End Function

Sub Gitna(frm As Form)
    frm.Left = (frmMain.ScaleWidth - frm.Width) / 2
    frm.Top = (frmMain.ScaleHeight - frm.Height) / 2
End Sub
