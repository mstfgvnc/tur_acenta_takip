VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Sheets("TUR").Unprotect "1234"
A = ActiveCell.Offset(0, 2).Value
Range("A2").AutoFilter 2, ActiveCell.Offset(0, 1).Value
Range("A2").AutoFilter 3, ActiveCell.Offset(0, 2).Value
Range("A2").AutoFilter 6, ActiveCell.Offset(0, 5).Value
Range("A2").AutoFilter 7, ActiveCell.Offset(0, 6).Value
Range("A2").AutoFilter 8, ActiveCell.Offset(0, 7).Value
X = WorksheetFunction.CountIf(Range("C4:C10000"), A)
Sheets("TUR").Range("C3:C10000").Find(A).Activate
tur.ComboBox1.Value = ActiveCell.Offset(0, -1)
tur.TextBox1.Value = ActiveCell.Offset(0, 0)
tur.ComboBox3.Value = ActiveCell.Offset(0, 3)
tur.ComboBox4.Value = ActiveCell.Offset(0, 4)
tur.ComboBox5.Value = ActiveCell.Offset(0, 5)
tur.ComboBox394.Value = ActiveCell.Offset(0, 6)
tur.ComboBox395.Value = ActiveCell.Offset(0, 7)
tur.ComboBox393.Value = ActiveCell.Offset(0, 8)
tur.TextBox2.Value = ActiveCell.Offset(0, 21)
tur.TextBox3.Value = ActiveCell.Offset(0, 22)
tur.TextBox4.Value = ActiveCell.Offset(0, 23)
tur.TextBox245.Value = ActiveCell.Offset(0, 20)
For i = 1 To X
tur.Controls("Textbox" & (i * 5)).Value = ActiveCell.Offset(i - 1, 1)
tur.Controls("Textbox" & (i * 5) + 1).Value = ActiveCell.Offset(i - 1, 2)
tur.Controls("Textbox" & (i * 5) + 2).Value = ActiveCell.Offset(i - 1, 19)
tur.Controls("Textbox" & (i * 5) + 3).Value = ActiveCell.Offset(i - 1, 12)
tur.Controls("Textbox" & (i * 5) + 4).Value = ActiveCell.Offset(i - 1, 14)
tur.Controls("ödem" & i).Value = ActiveCell.Offset(i - 1, 9)
tur.Controls("kart" & i).Value = ActiveCell.Offset(i - 1, 10)
tur.Controls("ücretd" & i).Value = ActiveCell.Offset(i - 1, 11)
tur.Controls("gün" & i).Value = ActiveCell.Offset(i - 1, 16)
tur.Controls("ay" & i).Value = ActiveCell.Offset(i - 1, 17)
tur.Controls("yýl" & i).Value = ActiveCell.Offset(i - 1, 18)
Next
ActiveSheet.ShowAllData
Sheets("TUR").Protect "1234", AllowFiltering:=True
tur.CommandButton3.Enabled = False
For i = 5 To 240 Step 5
tur.Controls("TextBox" & i).Enabled = False
Next
For i = 6 To 241 Step 5
tur.Controls("TextBox" & i).Enabled = False
Next
For i = 7 To 242 Step 5
tur.Controls("TextBox" & i).Enabled = False
Next
tur.Show
End Sub

Private Sub CommandButton2_Click()
answer = MsgBox(ActiveCell.Value & " sýra numaralý kaydý silmek istediðinize emin misiniz?", vbYesNo + vbQuestion, "TUR SATIÞ")
If answer = vbYes Then
Sheets("TUR").Unprotect "1234"
For i = 1 To 25
ActiveCell.Offset(0, i).ClearContents
Next
If ActiveCell.Offset(1, 0).Value = "" Then
ActiveCell.ClearContents
Else
A = ActiveCell.Row
B = ActiveSheet.Range("b655655").End(xlUp).Row
Sheets("TUR").Select
Sheets("TUR").Range(Cells(B, 2), Cells(A + 1, 26)).Copy
'Selection.Cut
ActiveSheet.Range("b" & A).PasteSpecial Paste:=xlPasteValues
'selectio.PasteSpecial Paste:=xlPasteValues
For i = 0 To 25
Range("a" & B).Offset(0, i).ClearContents
Next
End If
Sheets("TUR").Protect "1234", AllowFiltering:=True
Else
End If
End Sub

Private Sub CommandButton3_Click()
UserForm3.Show
End Sub




