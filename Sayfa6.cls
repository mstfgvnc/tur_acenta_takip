VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Sheets("OTEL").Unprotect "1234"
otel.TextBox7.Text = ActiveCell.Offset(0, 1).Value
otel.TextBox13.Text = ActiveCell.Offset(0, 2).Value
otel.ComboBox1.Text = ActiveCell.Offset(0, 3).Value
otel.ComboBox3.Text = ActiveCell.Offset(0, 4).Value
otel.ComboBox4.Text = ActiveCell.Offset(0, 5).Value
otel.ComboBox5.Text = ActiveCell.Offset(0, 6).Value
otel.ComboBox7.Text = ActiveCell.Offset(0, 7).Value
otel.ComboBox18.Text = ActiveCell.Offset(0, 8).Value
otel.ComboBox17.Text = ActiveCell.Offset(0, 9).Value
otel.ComboBox2.Text = ActiveCell.Offset(0, 10).Value
otel.ComboBox14.Text = ActiveCell.Offset(0, 11).Value
otel.ComboBox15.Text = ActiveCell.Offset(0, 12).Value
otel.TextBox8.Text = ActiveCell.Offset(0, 13).Value
otel.ComboBox12.Text = ActiveCell.Offset(0, 14).Value
otel.TextBox9.Text = ActiveCell.Offset(0, 15).Value
otel.ComboBox13.Text = ActiveCell.Offset(0, 16).Value
otel.ComboBox10.Text = ActiveCell.Offset(0, 17).Value
otel.ComboBox19.Text = ActiveCell.Offset(0, 18).Value
otel.ComboBox16.Text = ActiveCell.Offset(0, 19).Value
otel.TextBox10.Text = ActiveCell.Offset(0, 20).Value
otel.TextBox11.Text = ActiveCell.Offset(0, 21).Value
otel.CommandButton1.Visible = False
Sheets("OTEL").Protect "1234", AllowFiltering:=True
otel.TextBox7.Enabled = False
otel.TextBox13.Enabled = False
otel.TextBox10.Enabled = False
otel.Show
End Sub

Private Sub CommandButton2_Click()
answer = MsgBox(ActiveCell.Value & " sýra numaralý kaydý silmek istediðinize emin misiniz?", vbYesNo + vbQuestion, "OTEL SATIÞ")
If answer = vbYes Then
Sheets("OTEL").Unprotect "1234"
For i = 1 To 25
ActiveCell.Offset(0, i).ClearContents
Next
If ActiveCell.Offset(1, 0).Value = "" Then
ActiveCell.ClearContents
Else
A = ActiveCell.Row
B = ActiveSheet.Range("b655655").End(xlUp).Row
Sheets("OTEL").Select
Sheets("OTEL").Range(Cells(B, 2), Cells(A + 1, 26)).Copy
'Selection.Cut
ActiveSheet.Range("b" & A).PasteSpecial Paste:=xlPasteValues
'selectio.PasteSpecial Paste:=xlPasteValues
For i = 0 To 25
Range("a" & B).Offset(0, i).ClearContents
Next
End If
Sheets("OTEL").Protect "1234", AllowFiltering:=True
Else
End If
End Sub

Private Sub CommandButton3_Click()
UserForm3.Show
End Sub




