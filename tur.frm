VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tur 
   Caption         =   "TURLAR"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16215
   OleObjectBlob   =   "tur.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'tur.Hide
UserForm7.ListBox1.MultiSelect = fmMultiSelectExtended
UserForm7.CommandButton1.Enabled = False
UserForm7.CommandButton2.Enabled = False
UserForm7.CommandButton3.Enabled = False
UserForm7.CommandButton4.Enabled = False
UserForm7.CommandButton5.Enabled = False
UserForm7.CommandButton6.Enabled = False
UserForm7.Show
End Sub
Private Sub CommandButton6_Click()
tur.Hide
UserForm3.Show
End Sub
Private Sub CommandButton2_Click()
UserForm_Initialize
End Sub
Private Sub CommandButton3_Click()
If tur.TextBox1.Text = "" Then
MsgBox ("Lütfen Tur Adýný Giriniz...")
Else
If tur.ComboBox1.Text = "" Then
MsgBox ("Lütfen Tur Tipini Giriniz...")
Else
If tur.TextBox5.Text = "" Then
MsgBox ("Lütfen Müþteri Ekle Ýle Müþteri Bilgilerini Giriniz...")
Else
X = 0
For i = 5 To 240 Step 5
If Controls("Textbox" & i).Value = "" Then
Else
X = X + 1
End If
Next
For i = 1 To X
If Controls("ücretd" & i).Value = "" Then
MsgBox ("Lütfen " & i & ". Müþterinin Ücret Durumu Bilgilerini Giriniz...")
GoTo BÝTÝR
Else
If Controls("gün" & i).Value = "" Then
MsgBox ("Lütfen " & i & ". Müþterinin Satýþ Günü Bilgilerini Giriniz...")
GoTo BÝTÝR
Else
If Controls("ay" & i).Value = "" Then
MsgBox ("Lütfen " & i & ". Müþterinin Satýþ Ayý Bilgilerini Giriniz...")
GoTo BÝTÝR
Else
If Controls("yýl" & i).Value = "" Then
MsgBox ("Lütfen " & i & ". Müþterinin Satýþ Yýlý Bilgilerini Giriniz...")
GoTo BÝTÝR
Else
End If
End If
End If
End If
Next
Sheets("TUR").Unprotect "1234"
For i = 1 To X
ActiveSheet.Range("B65536").End(xlUp).Offset(1, 0).Value = tur.ComboBox1.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 1).Value = tur.TextBox1.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 4).Value = tur.ComboBox3.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 5).Value = tur.ComboBox4.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 6).Value = tur.ComboBox5.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 7).Value = tur.ComboBox394.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 8).Value = tur.ComboBox395.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 9).Value = tur.ComboBox393.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 22).Value = tur.TextBox2.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 23).Value = tur.TextBox3.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 24).Value = tur.TextBox4.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 2).Value = tur.Controls("TextBox" & i * 5).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 3).Value = tur.Controls("TextBox" & (i * 5) + 1).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 20).Value = tur.Controls("TextBox" & (i * 5) + 2).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 13).Value = tur.Controls("TextBox" & (i * 5) + 3).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 14).Value = "TL"
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 15).Value = tur.Controls("TextBox" & (i * 5) + 4).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 16).Value = "TL"
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 10).Value = tur.Controls("ödem" & i).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 11).Value = tur.Controls("kart" & i).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 12).Value = tur.Controls("ücretd" & i).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 17).Value = tur.Controls("gün" & i).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 18).Value = tur.Controls("ay" & i).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 19).Value = tur.Controls("yýl" & i).Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 21).Value = tur.TextBox245.Text
If ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Value = "Sýra No" Then
ActiveSheet.Range("a65536").End(xlUp).Offset(2, 0).Value = 1
Else
ActiveSheet.Range("a65536").End(xlUp).Offset(1, 0).Value = ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Value + 1
End If
D = ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Row
Range("a" & D, "z" & D).Borders.LineStyle = xlContinuous
Next
Sheets("TUR").Protect "1234", AllowFiltering:=True
Unload tur
BÝTÝR:
End If
End If
End If
End Sub
Private Sub CommandButton4_Click()
Sheets("TUR").Unprotect "1234"
X = WorksheetFunction.CountIf(Range("C4:C10000"), ActiveCell.Offset(0, 2).Value)
'MsgBox X
For i = 5 To 240 Step 5
If Controls("Textbox" & i).Value = "" Then
Else
Y = Y + 1
End If
Next
'MsgBox Y
For i = 1 To Y
If Controls("ücretd" & i).Value = "" Then
MsgBox ("Lütfen " & i & ". Müþterinin Ücret Durumu Bilgilerini Giriniz...")
GoTo BÝTÝR
Else
If Controls("gün" & i).Value = "" Then
MsgBox ("Lütfen " & i & ". Müþterinin Satýþ Günü Bilgilerini Giriniz...")
GoTo BÝTÝR
Else
If Controls("ay" & i).Value = "" Then
MsgBox ("Lütfen " & i & ". Müþterinin Satýþ Ayý Bilgilerini Giriniz...")
GoTo BÝTÝR
Else
If Controls("yýl" & i).Value = "" Then
MsgBox ("Lütfen " & i & ". Müþterinin Satýþ Yýlý Bilgilerini Giriniz...")
GoTo BÝTÝR
Else
End If
End If
End If
End If
Next
If Y - X = 0 Then
Sheets("TUR").Range("C3:C10000").Find(ActiveCell.Offset(0, 2).Value).Activate
For i = 1 To Y
ActiveCell.Offset(i - 1, -1).Value = tur.ComboBox1.Value
ActiveCell.Offset(i - 1, 0).Value = tur.TextBox1.Value
ActiveCell.Offset(i - 1, 3).Value = tur.ComboBox3.Value
ActiveCell.Offset(i - 1, 4).Value = tur.ComboBox4.Value
ActiveCell.Offset(i - 1, 5).Value = tur.ComboBox5.Value
ActiveCell.Offset(i - 1, 6).Value = tur.ComboBox394.Value
ActiveCell.Offset(i - 1, 7).Value = tur.ComboBox395.Value
ActiveCell.Offset(i - 1, 8).Value = tur.ComboBox393.Value
ActiveCell.Offset(i - 1, 21).Value = tur.TextBox2.Value
ActiveCell.Offset(i - 1, 22).Value = tur.TextBox3.Value
ActiveCell.Offset(i - 1, 23).Value = tur.TextBox4.Value
ActiveCell.Offset(i - 1, 20).Value = tur.TextBox245.Value
ActiveCell.Offset(i - 1, 1).Value = tur.Controls("Textbox" & (i * 5)).Value
ActiveCell.Offset(i - 1, 2).Value = tur.Controls("Textbox" & (i * 5) + 1).Value
ActiveCell.Offset(i - 1, 19).Value = tur.Controls("Textbox" & (i * 5) + 2).Value
ActiveCell.Offset(i - 1, 12).Value = tur.Controls("Textbox" & (i * 5) + 3).Value
ActiveCell.Offset(i - 1, 14).Value = tur.Controls("Textbox" & (i * 5) + 4).Value
ActiveCell.Offset(i - 1, 9).Value = tur.Controls("ödem" & i).Value
ActiveCell.Offset(i - 1, 10).Value = tur.Controls("kart" & i).Value
ActiveCell.Offset(i - 1, 11).Value = tur.Controls("ücretd" & i).Value
ActiveCell.Offset(i - 1, 16).Value = tur.Controls("gün" & i).Value
ActiveCell.Offset(i - 1, 17).Value = tur.Controls("ay" & i).Value
ActiveCell.Offset(i - 1, 18).Value = tur.Controls("yýl" & i).Value
Next
Else
k = ActiveCell.Offset(0, 2).Value
Sheets("TUR").Range("C3:C10000").Find(k).Activate
C = ActiveCell.Offset(X, -2).Row
MsgBox C
D = ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Row
MsgBox D
If C = D + 1 Then
Else
Sheets("TUR").Range("b" & C, "z" & D).Copy
Sheets("TUR").Range("b" & C + Y - X).PasteSpecial (xlPasteValues)
Range("a" & C + Y - X, "z" & D + Y - X).Borders.LineStyle = xlContinuous
Sheets("TUR").Range("C3:C10000").Find(k).Activate
End If
For i = 1 To Y
ActiveCell.Offset(i - 1, -1).Value = tur.ComboBox1.Value
ActiveCell.Offset(i - 1, 0).Value = tur.TextBox1.Value
ActiveCell.Offset(i - 1, 3).Value = tur.ComboBox3.Value
ActiveCell.Offset(i - 1, 4).Value = tur.ComboBox4.Value
ActiveCell.Offset(i - 1, 5).Value = tur.ComboBox5.Value
ActiveCell.Offset(i - 1, 6).Value = tur.ComboBox394.Value
ActiveCell.Offset(i - 1, 7).Value = tur.ComboBox395.Value
ActiveCell.Offset(i - 1, 8).Value = tur.ComboBox393.Value
ActiveCell.Offset(i - 1, 21).Value = tur.TextBox2.Value
ActiveCell.Offset(i - 1, 22).Value = tur.TextBox3.Value
ActiveCell.Offset(i - 1, 23).Value = tur.TextBox4.Value
ActiveCell.Offset(i - 1, 20).Value = tur.TextBox245.Value
ActiveCell.Offset(i - 1, 1).Value = tur.Controls("Textbox" & (i * 5)).Value
ActiveCell.Offset(i - 1, 2).Value = tur.Controls("Textbox" & (i * 5) + 1).Value
ActiveCell.Offset(i - 1, 19).Value = tur.Controls("Textbox" & (i * 5) + 2).Value
ActiveCell.Offset(i - 1, 12).Value = tur.Controls("Textbox" & (i * 5) + 3).Value
ActiveCell.Offset(i - 1, 14).Value = tur.Controls("Textbox" & (i * 5) + 4).Value
ActiveCell.Offset(i - 1, 9).Value = tur.Controls("ödem" & i).Value
ActiveCell.Offset(i - 1, 10).Value = tur.Controls("kart" & i).Value
ActiveCell.Offset(i - 1, 11).Value = tur.Controls("ücretd" & i).Value
ActiveCell.Offset(i - 1, 16).Value = tur.Controls("gün" & i).Value
ActiveCell.Offset(i - 1, 17).Value = tur.Controls("ay" & i).Value
ActiveCell.Offset(i - 1, 18).Value = tur.Controls("yýl" & i).Value
Range("a" & C + Y - X, "z" & D + Y - X).Borders.LineStyle = xlContinuous
Next
For i = 1 To Y - X
ActiveSheet.Range("a65536").End(xlUp).Offset(1, 0).Value = ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Value + 1
Next
End If
Unload tur
BÝTÝR:
Sheets("TUR").Protect "1234", AllowFiltering:=True
End Sub
Private Sub UserForm_Initialize()
'If ComboBox2.Value = "TOPLU" Then
For i = 13 To 243 Step 5
If Controls("Textbox" & i - 3).Value = "" Then
Else
Controls("Textbox" & i).Value = TextBox8.Value
End If
Next
For i = 14 To 244 Step 5
If Controls("Textbox" & i - 4).Value = "" Then
Else
Controls("Textbox" & i).Value = TextBox9.Value
End If
Next
For i = 2 To 48
If Controls("textbox" & i * 5).Value = "" Then
Else
Controls("ödem" & i).Value = ödem1.Value
End If
Next
For i = 2 To 48
If Controls("textbox" & i * 5).Value = "" Then
Else
Controls("kart" & i).Value = kart1.Value
End If
Next
For i = 2 To 48
If Controls("textbox" & i * 5).Value = "" Then
Else
Controls("ücretd" & i).Value = ücretd1.Value
End If
Next
For i = 2 To 48
If Controls("textbox" & i * 5).Value = "" Then
Else
Controls("gün" & i).Value = gün1.Value
End If
Next
For i = 2 To 48
If Controls("textbox" & i * 5).Value = "" Then
Else
Controls("ay" & i).Value = ay1.Value
End If
Next
For i = 2 To 48
If Controls("textbox" & i * 5).Value = "" Then
Else
Controls("yýl" & i).Value = yýl1.Value
End If
Next
'End If
End Sub
