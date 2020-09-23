Attribute VB_Name = "Module1"
' _
this is the code which does most of the work _
as you can see, it is very simple _
 _
Copyright (C) 2002 Xerb (Randy Girard) _
 _
you are free to use this code anyway you please, all I ask _
  is that you pls give some credit. Thx _
           - xErB
  
Public MaForms As New Collection
Public MaFormNames As New Collection

Public Sub DeleteForm(frmName As String)
  For i = 1 To MaForms.Count
    If LCase(MaFormNames(i)) = LCase(frmName) Then
      MaForms.Remove (i)
      MaFormNames.Remove (i)
      Exit Sub
    End If
  Next i
  MsgBox "Error: The removal of the form '" & frmName & "' was attempted, but the form was not found"
End Sub

Public Sub AddForm(frm As Form, frmName As String)
  For i = 1 To MaForms.Count
    If LCase(MaFormNames(i)) = LCase(frmName) Then
      MsgBox "Error: Form '" & frmName & "' allready egsists. form will not be created, and this may cause many more errors later in this script."
      Exit Sub
    End If
  Next i
  MaForms.Add frm
  MaFormNames.Add frmName
  frm.Tag = frmName
End Sub

Public Function GetForm(frmName As String) As Form
  For i = 1 To MaForms.Count
    If LCase(MaFormNames(i)) = LCase(frmName) Then
      Set GetForm = MaForms(i)
      Exit Function
    End If
  Next i
  MsgBox "Error: The request for the form '" & frmName & "' was not able to be completed; The form was not found"
End Function
