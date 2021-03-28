VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} otel 
   Caption         =   "OTEL ÝÞLEMÝ"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9750
   OleObjectBlob   =   "otel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "otel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
If otel.TextBox7.Text = "" Then
MsgBox ("Lütfen Müþteri Ekle Ýle Müþteri Bilgilerini Giriniz...")
Else
If otel.ComboBox15.Text = "" Then
MsgBox ("Lütfen Ödeme Durumu Bilgilerini Giriniz...")
Else
If otel.ComboBox10.Text = "" Then
MsgBox ("Lütfen Satýþ Gününü Giriniz...")
Else
If otel.ComboBox19.Text = "" Then
MsgBox ("Lütfen Satýþ Ayýný Giriniz...")
Else
If otel.ComboBox16.Text = "" Then
MsgBox ("Lütfen Satýþ Yýlýný Giriniz...")
Else
Sheets("OTEL").Unprotect "1234"
answer = vbYes
On Error GoTo HATA
Sheets("OTEL").Range("c1:c65536").Find(otel.TextBox13, SearchDirection:=xlPrevious).Activate
If ActiveCell.Offset(0, 15).Value = otel.ComboBox10.Text And ActiveCell.Offset(0, 16).Value = otel.ComboBox19.Text And ActiveCell.Offset(0, 17).Value = otel.ComboBox16.Text Then
answer = MsgBox(ActiveCell.Value & " TC NUMARALI VE AYNI SATIÞ TARÝHLÝ BÝR KAYIT BULUNMAKTADIR.BU KAYDI EKLEMEK ÝSTEDÝÐÝNÝZE EMÝN MÝSÝNÝZ ? ", vbYesNo + vbQuestion, "OTEL SATIÞ")
Else
answer = vbYes
End If
HATA:
If answer = vbYes Then
ActiveSheet.Range("B65536").End(xlUp).Offset(1, 0).Value = otel.TextBox7.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 1).Value = otel.TextBox13.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 2).Value = otel.ComboBox1.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 3).Value = otel.ComboBox3.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 4).Value = otel.ComboBox4.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 5).Value = otel.ComboBox5.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 6).Value = otel.ComboBox7.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 7).Value = otel.ComboBox18.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 8).Value = otel.ComboBox17.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 9).Value = otel.ComboBox2.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 10).Value = otel.ComboBox14.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 11).Value = otel.ComboBox15.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 12).Value = otel.TextBox8.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 13).Value = otel.ComboBox12.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 14).Value = otel.TextBox9.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 15).Value = otel.ComboBox13.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 16).Value = otel.ComboBox10.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 17).Value = otel.ComboBox19.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 18).Value = otel.ComboBox16.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 19).Value = otel.TextBox10.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 20).Value = otel.TextBox11.Text
If ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Value = "Sýra No" Then
ActiveSheet.Range("a65536").End(xlUp).Offset(2, 0).Value = 1
Else
ActiveSheet.Range("a65536").End(xlUp).Offset(1, 0).Value = ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Value + 1
End If
X = ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Row
Range("a" & X, "z" & X).Borders.LineStyle = xlContinuous
End If
Sheets("OTEL").Protect "1234", AllowFiltering:=True
Unload otel
End If
End If
End If
End If
End If
End Sub

Private Sub CommandButton2_Click()
If otel.TextBox7.Text = "" Then
MsgBox ("Lütfen Müþteri Ekle Ýle Müþteri Bilgilerini Giriniz...")
Else
If otel.ComboBox15.Text = "" Then
MsgBox ("Lütfen Ödeme Durumu Bilgilerini Giriniz...")
Else
If otel.ComboBox10.Text = "" Then
MsgBox ("Lütfen Satýþ Gününü Giriniz...")
Else
If otel.ComboBox19.Text = "" Then
MsgBox ("Lütfen Satýþ Ayýný Giriniz...")
Else
If otel.ComboBox16.Text = "" Then
MsgBox ("Lütfen Satýþ Yýlýný Giriniz...")
Else
Sheets("OTEL").Unprotect "1234"
ActiveCell.Offset(0, 1).Value = otel.TextBox7.Text
ActiveCell.Offset(0, 2).Value = otel.TextBox13.Text
ActiveCell.Offset(0, 3).Value = otel.ComboBox1.Text
ActiveCell.Offset(0, 4).Value = otel.ComboBox3.Text
ActiveCell.Offset(0, 5).Value = otel.ComboBox4.Text
ActiveCell.Offset(0, 6).Value = otel.ComboBox5.Text
ActiveCell.Offset(0, 7).Value = otel.ComboBox7.Text
ActiveCell.Offset(0, 8).Value = otel.ComboBox18.Text
ActiveCell.Offset(0, 9).Value = otel.ComboBox17.Text
ActiveCell.Offset(0, 10).Value = otel.ComboBox2.Text
ActiveCell.Offset(0, 11).Value = otel.ComboBox14.Text
ActiveCell.Offset(0, 12).Value = otel.ComboBox15.Text
ActiveCell.Offset(0, 13).Value = otel.TextBox8.Text
ActiveCell.Offset(0, 14).Value = otel.ComboBox12.Text
ActiveCell.Offset(0, 15).Value = otel.TextBox9.Text
ActiveCell.Offset(0, 16).Value = otel.ComboBox13.Text
ActiveCell.Offset(0, 17).Value = otel.ComboBox10.Text
ActiveCell.Offset(0, 18).Value = otel.ComboBox19.Text
ActiveCell.Offset(0, 19).Value = otel.ComboBox16.Text
ActiveCell.Offset(0, 20).Value = otel.TextBox10.Text
ActiveCell.Offset(0, 21).Value = otel.TextBox11.Text
Sheets("OTEL").Protect "1234", AllowFiltering:=True
Unload otel
End If
End If
End If
End If
End If
End Sub

Private Sub Label20_Click()

End Sub

Private Sub CommandButton3_Click()
UserForm7.CommandButton10.Enabled = False
UserForm7.CommandButton1.Enabled = False
UserForm7.CommandButton2.Enabled = False
UserForm7.CommandButton3.Enabled = False
UserForm7.CommandButton4.Enabled = False
UserForm7.CommandButton5.Enabled = False
UserForm7.Show
End Sub

Private Sub UserForm_Click()

End Sub
