VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm8 
   Caption         =   "MÜÞTERÝ EKLE"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8940
   OleObjectBlob   =   "UserForm8.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Sheets("MÜÞTERÝ").Unprotect "1234"
On Error Resume Next
If UserForm8.TextBox1.Text = "" Then
MsgBox ("Lütfen Müþteri Adýný Giriniz...")
Else
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(1, 0).Value = UserForm8.TextBox1.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 1).Value = UserForm8.TextBox2.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 2).Value = UserForm8.TextBox3.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 3).Value = UserForm8.TextBox4.Text
If Worksheets("MÜÞTERÝ").Range("a65536").End(xlUp).Offset(0, 0).Value = "NO" Then
Worksheets("MÜÞTERÝ").Range("a65536").End(xlUp).Offset(1, 0).Value = 1
Else
Worksheets("MÜÞTERÝ").Range("a65536").End(xlUp).Offset(1, 0).Value = Worksheets("MÜÞTERÝ").Range("a65536").End(xlUp).Offset(0, 0).Value + 1
End If
If WorksheetFunction.CountIf(Worksheets("MÜÞTERÝ").Range("c2:c50000"), UserForm8.TextBox2.Text) > 1 Then
MsgBox "Hatalý Giriþ Bu Girdiðiniz Kayýt Var", vbCritical
Worksheets("MÜÞTERÝ").Range("a65536").End(xlUp).Offset(0, 0).ClearContents
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 0).ClearContents
Worksheets("MÜÞTERÝ").Range("c65536").End(xlUp).Offset(0, 0).ClearContents
Worksheets("MÜÞTERÝ").Range("d65536").End(xlUp).Offset(0, 0).ClearContents
Worksheets("MÜÞTERÝ").Range("e65536").End(xlUp).Offset(0, 0).ClearContents
End If
If WorksheetFunction.CountIf(Worksheets("MÜÞTERÝ").Range("d2:d50000"), UserForm8.TextBox3.Text) > 1 Then
MsgBox "Hatalý Giriþ Bu Girdiðiniz Kayýt Var", vbCritical
Worksheets("MÜÞTERÝ").Range("a65536").End(xlUp).Offset(0, 0).ClearContents
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 0).ClearContents
Worksheets("MÜÞTERÝ").Range("c65536").End(xlUp).Offset(0, 0).ClearContents
Worksheets("MÜÞTERÝ").Range("d65536").End(xlUp).Offset(0, 0).ClearContents
Worksheets("MÜÞTERÝ").Range("e65536").End(xlUp).Offset(0, 0).ClearContents
End If
Unload UserForm8
UserForm7.CommandButton10.Enabled = False
UserForm7.CommandButton4.Enabled = False
UserForm7.CommandButton5.Enabled = False
UserForm7.CommandButton6.Enabled = False
UserForm7.Show
End If
Sheets("MÜÞTERÝ").Protect "1234"
End Sub

Private Sub CommandButton2_Click()
Sheets("MÜÞTERÝ").Unprotect "1234"
If UserForm8.TextBox1.Text = "" Then
MsgBox ("Lütfen Müþteri Adýný Giriniz...")
Else
A = UserForm7.ListBox1.ListIndex
Sheets("MÜÞTERÝ").Range("B" & A + 1).Value = UserForm8.TextBox1.Text
Sheets("MÜÞTERÝ").Range("C" & A + 1).Value = UserForm8.TextBox2.Text
Sheets("MÜÞTERÝ").Range("D" & A + 1).Value = UserForm8.TextBox3.Text
Sheets("MÜÞTERÝ").Range("E" & A + 1).Value = UserForm8.TextBox4.Text
Unload UserForm8
End If
Sheets("MÜÞTERÝ").Protect "1234"
End Sub

Private Sub Label4_Click()

End Sub

Private Sub UserForm_Click()

End Sub
