VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Sheets("BÝLET").Unprotect "1234"
UserForm1.ComboBox1.Text = ActiveCell.Offset(0, 1).Value
UserForm1.TextBox2.Text = ActiveCell.Offset(0, 2).Value
UserForm1.TextBox3.Text = ActiveCell.Offset(0, 3).Value
UserForm1.ComboBox2.Text = ActiveCell.Offset(0, 4).Value
UserForm1.ComboBox14.Text = ActiveCell.Offset(0, 5).Value
UserForm1.ComboBox15.Text = ActiveCell.Offset(0, 6).Value
UserForm1.ComboBox3.Text = ActiveCell.Offset(0, 7).Value
UserForm1.ComboBox4.Text = ActiveCell.Offset(0, 8).Value
UserForm1.ComboBox5.Text = ActiveCell.Offset(0, 9).Value
UserForm1.TextBox6.Text = ActiveCell.Offset(0, 10).Value
UserForm1.TextBox7.Text = ActiveCell.Offset(0, 11).Value
UserForm1.TextBox13.Text = ActiveCell.Offset(0, 12).Value
UserForm1.TextBox12.Text = ActiveCell.Offset(0, 13).Value
UserForm1.ComboBox7.Text = ActiveCell.Offset(0, 14).Value
UserForm1.ComboBox18.Text = ActiveCell.Offset(0, 15).Value
UserForm1.ComboBox17.Text = ActiveCell.Offset(0, 16).Value
UserForm1.ComboBox10.Text = ActiveCell.Offset(0, 17).Value
UserForm1.ComboBox19.Text = ActiveCell.Offset(0, 18).Value
UserForm1.ComboBox16.Text = ActiveCell.Offset(0, 19).Value
UserForm1.TextBox8.Text = ActiveCell.Offset(0, 20).Value
UserForm1.ComboBox12.Text = ActiveCell.Offset(0, 21).Value
UserForm1.TextBox9.Text = ActiveCell.Offset(0, 22).Value
UserForm1.ComboBox13.Text = ActiveCell.Offset(0, 23).Value
UserForm1.TextBox10.Text = ActiveCell.Offset(0, 24).Value
UserForm1.TextBox11.Text = ActiveCell.Offset(0, 25).Value
UserForm1.CommandButton1.Visible = False
Sheets("BÝLET").Protect "1234", AllowFiltering:=True
UserForm1.TextBox7.Enabled = False
UserForm1.TextBox13.Enabled = False
UserForm1.TextBox10.Enabled = False
UserForm1.Show
End Sub
Private Sub CommandButton2_Click()
answer = MsgBox(ActiveCell.Value & " sýra numaralý kaydý silmek istediðinize emin misiniz?", vbYesNo + vbQuestion, "BÝLET SATIÞ")
If answer = vbYes Then
Sheets("BÝLET").Unprotect "1234"
For i = 1 To 25
ActiveCell.Offset(0, i).ClearContents
Next
If ActiveCell.Offset(1, 0).Value = "" Then
ActiveCell.ClearContents
Else
A = ActiveCell.Row
B = ActiveSheet.Range("b655655").End(xlUp).Row
Sheets("BÝLET").Select
Sheets("BÝLET").Range(Cells(B, 2), Cells(A + 1, 26)).Copy
'Selection.Cut
ActiveSheet.Range("b" & A).PasteSpecial Paste:=xlPasteValues
'selectio.PasteSpecial Paste:=xlPasteValues
For i = 0 To 25
Range("a" & B).Offset(0, i).ClearContents
Next
End If
Sheets("BÝLET").Protect "1234", AllowFiltering:=True
Else
End If
End Sub

Private Sub CommandButton3_Click()
UserForm3.Show
End Sub




