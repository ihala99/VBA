
Private Sub CommandButton1_Click()
Dim LastRow As Long
LastRow = WorksheetFunction.CountA(Sheets("hala").Range("A:A"))

If Usrfrm.TextBox1.Value = "" Or Usrfrm.TextBox2.Value = "" Or Usrfrm.TextBox3.Value = "" Then
MsgBox "أدخل البيانات كاملة"
Else

Sheets("hala").Cells(LastRow + 1, 1).Value = Usrfrm.TextBox1.Value
Sheets("hala").Cells(LastRow + 1, 2).Value = Usrfrm.TextBox2.Value
Sheets("hala").Cells(LastRow + 1, 3).Value = Usrfrm.TextBox3.Value


Usrfrm.TextBox2.Value = ""
Usrfrm.TextBox3.Value = ""
Usrfrm.TextBox1.Value = ""

End If

End Sub

Private Sub CommandButton2_Click()
psw = InputBox("Please Enter The Password")
If psw = 1234 Then
Application.Visible = True
Sheets("hala").Visible = True
Sheets("hala").Activate
Usrfrm.Hide

Else
MsgBox "WRONG PASSWORD ! ASK DEVELOPER HALA"
End If
End Sub


Private Sub TextBox1_Change()

End Sub

Private Sub TextBox3_Change()

End Sub
Private Sub TextBox2_Change()

End Sub
Private Sub Usrfrm_Terminate()
ActiveWorkbook.Save
Application.Quit
End Sub

Private Sub UserForm_Click()

End Sub
