VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Sheets("VÝZE").Unprotect "1234"
vize.TextBox7.Text = ActiveCell.Offset(0, 1).Value
vize.TextBox13.Text = ActiveCell.Offset(0, 2).Value
vize.ComboBox1.Text = ActiveCell.Offset(0, 3).Value
vize.ComboBox3.Text = ActiveCell.Offset(0, 4).Value
vize.ComboBox4.Text = ActiveCell.Offset(0, 5).Value
vize.ComboBox5.Text = ActiveCell.Offset(0, 6).Value
vize.ComboBox7.Text = ActiveCell.Offset(0, 7).Value
vize.ComboBox18.Text = ActiveCell.Offset(0, 8).Value
vize.ComboBox17.Text = ActiveCell.Offset(0, 9).Value
vize.ComboBox2.Text = ActiveCell.Offset(0, 10).Value
vize.ComboBox14.Text = ActiveCell.Offset(0, 11).Value
vize.ComboBox15.Text = ActiveCell.Offset(0, 12).Value
vize.TextBox8.Text = ActiveCell.Offset(0, 13).Value
vize.ComboBox12.Text = ActiveCell.Offset(0, 14).Value
vize.TextBox9.Text = ActiveCell.Offset(0, 15).Value
vize.ComboBox13.Text = ActiveCell.Offset(0, 16).Value
vize.ComboBox10.Text = ActiveCell.Offset(0, 17).Value
vize.ComboBox19.Text = ActiveCell.Offset(0, 18).Value
vize.ComboBox16.Text = ActiveCell.Offset(0, 19).Value
vize.TextBox10.Text = ActiveCell.Offset(0, 20).Value
vize.TextBox11.Text = ActiveCell.Offset(0, 21).Value
vize.CommandButton1.Visible = False
Sheets("VÝZE").Protect "1234", AllowFiltering:=True
vize.TextBox7.Enabled = False
vize.TextBox13.Enabled = False
vize.TextBox10.Enabled = False
vize.Show
End Sub

Private Sub CommandButton2_Click()
answer = MsgBox(ActiveCell.Value & " sýra numaralý kaydý silmek istediðinize emin misiniz?", vbYesNo + vbQuestion, "BÝLET SATIÞ")
If answer = vbYes Then
Sheets("VÝZE").Unprotect "1234"
For i = 1 To 25
ActiveCell.Offset(0, i).ClearContents
Next
If ActiveCell.Offset(1, 0).Value = "" Then
ActiveCell.ClearContents
Else
A = ActiveCell.Row
B = ActiveSheet.Range("b655655").End(xlUp).Row
Sheets("VÝZE").Select
Sheets("VÝZE").Range(Cells(B, 2), Cells(A + 1, 26)).Copy
'Selection.Cut
ActiveSheet.Range("b" & A).PasteSpecial Paste:=xlPasteValues
'selectio.PasteSpecial Paste:=xlPasteValues
For i = 0 To 25
Range("a" & B).Offset(0, i).ClearContents
Next
End If
Sheets("VÝZE").Protect "1234", AllowFiltering:=True
Else
End If
End Sub

Private Sub CommandButton3_Click()
UserForm3.Show
End Sub




