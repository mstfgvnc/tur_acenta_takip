VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "ANA MEN�"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8265
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
UserForm3.Hide
Sheets("B�LET").Select
UserForm1.CommandButton2.Visible = False
UserForm1.TextBox7.Enabled = False
UserForm1.TextBox13.Enabled = False
UserForm1.TextBox10.Enabled = False
UserForm1.Show
End Sub

Private Sub CommandButton2_Click()
UserForm3.Hide
Sheets("V�ZE").Select
vize.CommandButton2.Visible = False
vize.TextBox7.Enabled = False
vize.TextBox13.Enabled = False
vize.TextBox10.Enabled = False
vize.Show
End Sub

Private Sub CommandButton3_Click()
UserForm3.Hide
Sheets("TUR").Select
For i = 5 To 240 Step 5
tur.Controls("TextBox" & i).Enabled = False
Next
For i = 6 To 241 Step 5
tur.Controls("TextBox" & i).Enabled = False
Next
For i = 7 To 242 Step 5
tur.Controls("TextBox" & i).Enabled = False
Next
tur.CommandButton4.Enabled = False
tur.Show
End Sub

Private Sub CommandButton4_Click()
UserForm3.Hide
Sheets("OTEL").Select
otel.CommandButton2.Visible = False
otel.TextBox7.Enabled = False
otel.TextBox13.Enabled = False
otel.TextBox10.Enabled = False
otel.Show
End Sub

Private Sub CommandButton5_Click()
Application.Quit
End Sub

Private Sub CommandButton6_Click()
UserForm3.Hide
UserForm7.CommandButton10.Enabled = False
UserForm7.CommandButton4.Enabled = False
UserForm7.CommandButton5.Enabled = False
UserForm7.CommandButton6.Enabled = False
UserForm7.Show
End Sub

Private Sub CommandButton7_Click()
Sheets("TUR").Unprotect "1234"
Sheets("B�LET").Unprotect "1234"
Sheets("OTEL").Unprotect "1234"
Sheets("V�ZE").Unprotect "1234"
Sheets("RAPORLA").Unprotect "1234"
Sheets("RAPORLA").Range("a4:z50000").ClearContents
'Sheets("TUR").ShowAllData
'Sheets("B�LET").ShowAllData
'Sheets("OTEL").ShowAllData
'Sheets("V�ZE").ShowAllData
'Sheets("TUR").Activate
If Sheets("TUR").Range("A4").Value = "" Then
Z = 3
Else
Z = Sheets("TUR").Range("A500000").End(xlUp).Row
'MsgBox Z
Sheets("TUR").Range("a4", "a" & Z).Copy
Sheets("RAPORLA").Range("B4").PasteSpecial Paste:=xlPasteValues
Sheets("RAPORLA").Range("a4", "a" & Z).Value = "TUR"
Sheets("RAPORLA").Range("K4", "K" & Z).Value = "TL"
Sheets("RAPORLA").Range("M4", "M" & Z).Value = "TL"
Sheets("TUR").Range("D4", "D" & Z).Copy
Sheets("RAPORLA").Range("C4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("E4", "E" & Z).Copy
Sheets("RAPORLA").Range("D4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("V4", "V" & Z).Copy
Sheets("RAPORLA").Range("E4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("L4", "L" & Z).Copy
Sheets("RAPORLA").Range("G4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("M4", "M" & Z).Copy
Sheets("RAPORLA").Range("H4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("N4", "N" & Z).Copy
Sheets("RAPORLA").Range("I4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("O4", "O" & Z).Copy
Sheets("RAPORLA").Range("J4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("Q4", "Q" & Z).Copy
Sheets("RAPORLA").Range("L4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("S4", "S" & Z).Copy
Sheets("RAPORLA").Range("N4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("T4", "T" & Z).Copy
Sheets("RAPORLA").Range("O4").PasteSpecial Paste:=xlPasteValues
Sheets("TUR").Range("U4", "U" & Z).Copy
Sheets("RAPORLA").Range("P4").PasteSpecial Paste:=xlPasteValues
Sheets("RAPORLA").Range("a4", "P" & Z).Borders.LineStyle = xlContinuous
End If
If Sheets("B�LET").Range("A4").Value = "" Then
Y = 3
Else
Y = Sheets("B�LET").Range("A500000").End(xlUp).Row
'MsgBox Y
Sheets("B�LET").Range("a4", "a" & Y).Copy
Sheets("RAPORLA").Range("B" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("RAPORLA").Range("a" & Z + 1, "a" & Y + Z - 3).Value = "B�LET"
Sheets("RAPORLA").Range("K" & Z + 1, "K" & Y + Z - 3).Value = "TL"
Sheets("RAPORLA").Range("M" & Z + 1, "M" & Y + Z - 3).Value = "TL"
Sheets("B�LET").Range("L4", "L" & Y).Copy
Sheets("RAPORLA").Range("C" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("M4", "M" & Y).Copy
Sheets("RAPORLA").Range("D" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("Y4", "Y" & Y).Copy
Sheets("RAPORLA").Range("E" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("B4", "B" & Y).Copy
Sheets("RAPORLA").Range("F" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("E4", "E" & Y).Copy
Sheets("RAPORLA").Range("G" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("F4", "F" & Y).Copy
Sheets("RAPORLA").Range("H" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("G4", "G" & Y).Copy
Sheets("RAPORLA").Range("I" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("W4", "W" & Y).Copy
Sheets("RAPORLA").Range("J" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("U4", "U" & Y).Copy
Sheets("RAPORLA").Range("L" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("H4", "H" & Y).Copy
Sheets("RAPORLA").Range("N" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("I4", "I" & Y).Copy
Sheets("RAPORLA").Range("O" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("B�LET").Range("J4", "J" & Y).Copy
Sheets("RAPORLA").Range("P" & Z + 1).PasteSpecial Paste:=xlPasteValues
Sheets("RAPORLA").Range("a" & Z + 1, "P" & Y + Z - 3).Borders.LineStyle = xlContinuous
End If
If Sheets("OTEL").Range("A4").Value = "" Then
w = 3
Else
w = Sheets("OTEL").Range("A500000").End(xlUp).Row
'MsgBox W
Sheets("OTEL").Range("a4", "a" & w).Copy
Sheets("RAPORLA").Range("B" & Y + Z + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("RAPORLA").Range("a" & Y + Z + 1 - 3, "a" & w + Y + Z - 6).Value = "OTEL"
Sheets("RAPORLA").Range("K" & Y + Z + 1 - 3, "K" & w + Y + Z - 6).Value = "TL"
Sheets("RAPORLA").Range("M" & Y + Z + 1 - 3, "M" & w + Y + Z - 6).Value = "TL"
Sheets("OTEL").Range("B4", "B" & w).Copy
Sheets("RAPORLA").Range("C" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("C4", "C" & w).Copy
Sheets("RAPORLA").Range("D" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("U4", "U" & w).Copy
Sheets("RAPORLA").Range("E" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("D4", "D" & w).Copy
Sheets("RAPORLA").Range("F" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("K4", "K" & w).Copy
Sheets("RAPORLA").Range("G" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("L4", "L" & w).Copy
Sheets("RAPORLA").Range("H" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("M4", "M" & w).Copy
Sheets("RAPORLA").Range("I" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("N4", "N" & w).Copy
Sheets("RAPORLA").Range("J" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("P4", "P" & w).Copy
Sheets("RAPORLA").Range("L" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("R4", "R" & w).Copy
Sheets("RAPORLA").Range("N" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("S4", "S" & w).Copy
Sheets("RAPORLA").Range("O" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("OTEL").Range("T4", "T" & w).Copy
Sheets("RAPORLA").Range("P" & Z + Y + 1 - 3).PasteSpecial Paste:=xlPasteValues
Sheets("RAPORLA").Range("a" & Z + Y + 1 - 3, "P" & Y + Z + w - 6).Borders.LineStyle = xlContinuous
End If
If Sheets("V�ZE").Range("A4").Value = "" Then
v = 3
Else
v = Sheets("V�ZE").Range("A500000").End(xlUp).Row
'MsgBox V
Sheets("V�ZE").Range("a4", "a" & v).Copy
Sheets("RAPORLA").Range("B" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("RAPORLA").Range("a" & w + Y + Z + 1 - 6, "a" & v + Y + Z + w - 9).Value = "V�ZE"
Sheets("RAPORLA").Range("K" & w + Y + Z + 1 - 6, "K" & v + Y + Z + w - 9).Value = "TL"
Sheets("RAPORLA").Range("M" & w + Y + Z + 1 - 6, "M" & v + Y + Z + w - 9).Value = "TL"
Sheets("V�ZE").Range("B4", "B" & v).Copy
Sheets("RAPORLA").Range("C" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("C4", "C" & v).Copy
Sheets("RAPORLA").Range("D" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("U4", "U" & v).Copy
Sheets("RAPORLA").Range("E" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("K4", "K" & v).Copy
Sheets("RAPORLA").Range("G" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("L4", "L" & v).Copy
Sheets("RAPORLA").Range("H" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("M4", "M" & v).Copy
Sheets("RAPORLA").Range("I" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("N4", "N" & v).Copy
Sheets("RAPORLA").Range("J" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("P4", "P" & v).Copy
Sheets("RAPORLA").Range("L" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("E4", "E" & v).Copy
Sheets("RAPORLA").Range("N" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("F4", "F" & v).Copy
Sheets("RAPORLA").Range("O" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("V�ZE").Range("G4", "G" & v).Copy
Sheets("RAPORLA").Range("P" & w + Y + Z + 1 - 6).PasteSpecial Paste:=xlPasteValues
Sheets("RAPORLA").Range("a" & w + Y + Z + 1 - 6, "P" & v + Y + Z + w - 9).Borders.LineStyle = xlContinuous
End If
Sheets("TUR").Protect "1234", AllowFiltering:=True
Sheets("B�LET").Protect "1234", AllowFiltering:=True
Sheets("OTEL").Protect "1234", AllowFiltering:=True
Sheets("V�ZE").Protect "1234", AllowFiltering:=True
Sheets("RAPORLA").Protect "1234"
Sheets("RAPORLA").Activate
Sheets("RAPORLA").Range("a3").Select
Unload UserForm3
End Sub

Private Sub UserForm_Click()

End Sub
