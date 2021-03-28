VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "BÝLET SATIÞI"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10440
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If UserForm1.ComboBox1.Text = "" Then
MsgBox ("Lütfen Tur Operatörünü Giriniz...")
Else
If UserForm1.ComboBox15.Text = "" Then
MsgBox ("Lütfen Ödeme Durumunu Giriniz...")
Else
If UserForm1.ComboBox3.Text = "" Then
MsgBox ("Lütfen Gün Bilgisini Giriniz...")
Else
If UserForm1.ComboBox4.Text = "" Then
MsgBox ("Lütfen Ay Bilgisini Giriniz...")
Else
If UserForm1.ComboBox5.Text = "" Then
MsgBox ("Lütfen Yýl Bilgisini Giriniz...")
Else
If UserForm1.ComboBox14.Text = "" Then
MsgBox ("Lütfen Kart Tipini Giriniz...")
Else
If UserForm1.TextBox7.Text = "" Then
MsgBox ("Lütfen Müþteri Adýný Giriniz...")
Else
Sheets("BÝLET").Unprotect "1234"
answer = vbYes
On Error GoTo HATA
Sheets("BÝLET").Range("m1:m65536").Find(UserForm1.TextBox13, SearchDirection:=xlPrevious).Activate
If ActiveCell.Offset(0, -5).Value = UserForm1.ComboBox3.Text And ActiveCell.Offset(0, -4).Value = UserForm1.ComboBox4.Text And ActiveCell.Offset(0, -3).Value = UserForm1.ComboBox5.Text Then
answer = MsgBox(ActiveCell.Value & " TC NUMARALI VE AYNI SATIÞ TARÝHLÝ BÝR KAYIT BULUNMAKTADIR.BU KAYDI EKLEMEK ÝSTEDÝÐÝNÝZE EMÝN MÝSÝNÝZ ? ", vbYesNo + vbQuestion, "BÝLET SATIÞ")
Else
answer = vbYes
End If
HATA:
If answer = vbYes Then
ActiveSheet.Range("B65536").End(xlUp).Offset(1, 0).Value = UserForm1.ComboBox1.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 1).Value = UserForm1.TextBox2.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 2).Value = UserForm1.TextBox3.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 3).Value = UserForm1.ComboBox2.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 4).Value = UserForm1.ComboBox14.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 5).Value = UserForm1.ComboBox15.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 6).Value = UserForm1.ComboBox3.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 7).Value = UserForm1.ComboBox4.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 8).Value = UserForm1.ComboBox5.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 9).Value = UserForm1.TextBox6.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 10).Value = UserForm1.TextBox7.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 11).Value = UserForm1.TextBox13.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 12).Value = UserForm1.TextBox12.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 13).Value = UserForm1.ComboBox7.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 14).Value = UserForm1.ComboBox18.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 15).Value = UserForm1.ComboBox17.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 16).Value = UserForm1.ComboBox10.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 17).Value = UserForm1.ComboBox19.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 18).Value = UserForm1.ComboBox16.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 19).Value = UserForm1.TextBox8.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 20).Value = UserForm1.ComboBox12.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 21).Value = UserForm1.TextBox9.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 22).Value = UserForm1.ComboBox13.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 23).Value = UserForm1.TextBox10.Text
ActiveSheet.Range("B65536").End(xlUp).Offset(0, 24).Value = UserForm1.TextBox11.Text
If ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Value = "Sýra No" Then
ActiveSheet.Range("a65536").End(xlUp).Offset(2, 0).Value = 1
Else
ActiveSheet.Range("a65536").End(xlUp).Offset(1, 0).Value = ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Value + 1
End If
X = ActiveSheet.Range("a65536").End(xlUp).Offset(0, 0).Row
Range("a" & X, "z" & X).Borders.LineStyle = xlContinuous
End If
Sheets("BÝLET").Protect "1234", AllowFiltering:=True
Unload UserForm1
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub CommandButton2_Click()
If UserForm1.ComboBox1.Text = "" Then
MsgBox ("Lütfen Tur Operatörünü Giriniz...")
Else
If UserForm1.ComboBox15.Text = "" Then
MsgBox ("Lütfen Ödeme Durumunu Giriniz...")
Else
If UserForm1.ComboBox3.Text = "" Then
MsgBox ("Lütfen Gün Bilgisini Giriniz...")
Else
If UserForm1.ComboBox4.Text = "" Then
MsgBox ("Lütfen Ay Bilgisini Giriniz...")
Else
If UserForm1.ComboBox5.Text = "" Then
MsgBox ("Lütfen Yýl Bilgisini Giriniz...")
Else
If UserForm1.ComboBox14.Text = "" Then
MsgBox ("Lütfen Kart Tipini Giriniz...")
Else
If UserForm1.TextBox7.Text = "" Then
MsgBox ("Lütfen Müþteri Adýný Giriniz...")
Else
Sheets("BÝLET").Unprotect "1234"
ActiveCell.Offset(0, 1).Value = UserForm1.ComboBox1.Text
ActiveCell.Offset(0, 2).Value = UserForm1.TextBox2.Text
ActiveCell.Offset(0, 3).Value = UserForm1.TextBox3.Text
ActiveCell.Offset(0, 4).Value = UserForm1.ComboBox2.Text
ActiveCell.Offset(0, 5).Value = UserForm1.ComboBox14.Text
ActiveCell.Offset(0, 6).Value = UserForm1.ComboBox15.Text
ActiveCell.Offset(0, 7).Value = UserForm1.ComboBox3.Text
ActiveCell.Offset(0, 8).Value = UserForm1.ComboBox4.Text
ActiveCell.Offset(0, 9).Value = UserForm1.ComboBox5.Text
ActiveCell.Offset(0, 10).Value = UserForm1.TextBox6.Text
ActiveCell.Offset(0, 11).Value = UserForm1.TextBox7.Text
ActiveCell.Offset(0, 12).Value = UserForm1.TextBox13.Text
ActiveCell.Offset(0, 13).Value = UserForm1.TextBox12.Text
ActiveCell.Offset(0, 14).Value = UserForm1.ComboBox7.Text
ActiveCell.Offset(0, 15).Value = UserForm1.ComboBox18.Text
ActiveCell.Offset(0, 16).Value = UserForm1.ComboBox17.Text
ActiveCell.Offset(0, 17).Value = UserForm1.ComboBox10.Text
ActiveCell.Offset(0, 18).Value = UserForm1.ComboBox19.Text
ActiveCell.Offset(0, 19).Value = UserForm1.ComboBox16.Text
ActiveCell.Offset(0, 20).Value = UserForm1.TextBox8.Text
ActiveCell.Offset(0, 21).Value = UserForm1.ComboBox12.Text
ActiveCell.Offset(0, 22).Value = UserForm1.TextBox9.Text
ActiveCell.Offset(0, 23).Value = UserForm1.ComboBox13.Text
ActiveCell.Offset(0, 24).Value = UserForm1.TextBox10.Text
ActiveCell.Offset(0, 25).Value = UserForm1.TextBox11.Text
Sheets("BÝLET").Protect "1234", AllowFiltering:=True
Unload UserForm1
End If
End If
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
UserForm7.CommandButton6.Enabled = False
UserForm7.Show
End Sub

Private Sub UserForm_Click()

End Sub
