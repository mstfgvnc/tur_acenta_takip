VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "MÜÞTERÝ LÝSTESÝ"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13110
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload UserForm7
UserForm8.CommandButton2.Enabled = False
UserForm8.Show
End Sub

Private Sub CommandButton10_Click()
i = 0
A = 0
B = 0
C = 5
X = 0
For i = C To ((UserForm7.ListBox1.ListCount - 1) * 5) + 4 Step 5
'MsgBox UserForm7.ListBox1.ListCount - 1 * 5
'MsgBox UserForm7.ListBox1.ListCount
If tur.Controls("Textbox" & i).Value = "" Then
For B = A To UserForm7.ListBox1.ListCount - 1
If UserForm7.ListBox1.Selected(B) = True Then
tur.Controls("Textbox" & i).Value = UserForm7.ListBox1.List(B, 1)
i = i + 1
tur.Controls("Textbox" & i).Value = UserForm7.ListBox1.List(B, 2)
i = i + 1
tur.Controls("Textbox" & i).Value = UserForm7.ListBox1.List(B, 3)
i = i + 3
If B = UserForm7.ListBox1.ListCount - 1 Then
GoTo GÝT
Else
End If
A = B
Else
A = B
End If
Next
'C = i
Else
'i = i + 4
'C = i
End If
Next
GÝT:
Unload UserForm7
'tur.Show
End Sub

Private Sub CommandButton2_Click()
UserForm8.CommandButton1.Enabled = False
On Error Resume Next
UserForm8.TextBox1.Text = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value
UserForm8.TextBox2.Text = Sheets("MÜÞTERÝ").Range("C" & ListBox1.ListIndex + 1).Value
UserForm8.TextBox3.Text = Sheets("MÜÞTERÝ").Range("D" & ListBox1.ListIndex + 1).Value
UserForm8.TextBox4.Text = Sheets("MÜÞTERÝ").Range("E" & ListBox1.ListIndex + 1).Value
UserForm8.Show
End Sub

Private Sub CommandButton3_Click()
answer = MsgBox(Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value & " müþteriyi silmek istediðinize emin misiniz?", vbYesNo + vbQuestion, "MÜÞTERÝ LÝSTESÝ")
If answer = vbYes Then
Sheets("MÜÞTERÝ").Unprotect "1234"
Dim i, A, B As Integer
A = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 2).Column
B = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 2).Row
X = ListBox1.ListIndex
Sheets("MÜÞTERÝ").Range("c" & X + 1).ClearContents
Sheets("MÜÞTERÝ").Range("d" & X + 1).ClearContents
Sheets("MÜÞTERÝ").Range("e" & X + 1).ClearContents
Sheets("MÜÞTERÝ").Range("B" & X + 1).ClearContents
Sheets("MÜÞTERÝ").Range("A" & Rows.Count).End(xlUp).ClearContents
If Sheets("MÜÞTERÝ").Range("B" & X + 2).Value = "" Then
Else
i = Worksheets("MÜÞTERÝ").Range("b655336").End(xlUp).Row
Sheets("MÜÞTERÝ").Select
Sheets("MÜÞTERÝ").Range(Cells(B, A), Cells(i, A + 7)).Select
Selection.Cut
ActiveSheet.Cells(B - 1, A).Select
ActiveSheet.Paste
End If
Sheets("MÜÞTERÝ").Protect "1234"
Else
End If
Unload UserForm7
UserForm7.CommandButton10.Enabled = False
UserForm7.CommandButton4.Enabled = False
UserForm7.CommandButton5.Enabled = False
UserForm7.CommandButton6.Enabled = False
UserForm7.Show
End Sub

Private Sub CommandButton4_Click()
On Error Resume Next
vize.TextBox7.Text = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value
vize.TextBox13.Text = Sheets("MÜÞTERÝ").Range("C" & ListBox1.ListIndex + 1).Value
vize.TextBox10.Text = Sheets("MÜÞTERÝ").Range("D" & ListBox1.ListIndex + 1).Value
vize.TextBox7.Enabled = False
vize.TextBox13.Enabled = False
vize.TextBox10.Enabled = False
Unload UserForm7
End Sub

Private Sub CommandButton5_Click()
On Error Resume Next
UserForm1.TextBox7.Text = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value
UserForm1.TextBox13.Text = Sheets("MÜÞTERÝ").Range("C" & ListBox1.ListIndex + 1).Value
UserForm1.TextBox10.Text = Sheets("MÜÞTERÝ").Range("D" & ListBox1.ListIndex + 1).Value
UserForm1.TextBox7.Enabled = False
UserForm1.TextBox13.Enabled = False
UserForm1.TextBox10.Enabled = False
Unload UserForm7
End Sub

Private Sub CommandButton6_Click()
otel.TextBox7.Text = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value
otel.TextBox13.Text = Sheets("MÜÞTERÝ").Range("C" & ListBox1.ListIndex + 1).Value
otel.TextBox10.Text = Sheets("MÜÞTERÝ").Range("D" & ListBox1.ListIndex + 1).Value
otel.TextBox7.Enabled = False
otel.TextBox13.Enabled = False
otel.TextBox10.Enabled = False
Unload UserForm7
End Sub

Private Sub CommandButton7_Click()
Set bul = Sheets("MÜÞTERÝ").Range("B1:B65536").Find(TextBox1, lookat:=xlWhole)
If Not bul Is Nothing Then
UserForm7.ListBox1.Selected(bul.Row - 1) = True
Else
MsgBox "aradýðýnýz kayýt bulunamadý."
End If
End Sub

Private Sub CommandButton8_Click()
Set bul = Sheets("MÜÞTERÝ").Range("c1:c65536").Find(TextBox2, lookat:=xlWhole)
If Not bul Is Nothing Then
UserForm7.ListBox1.Selected(bul.Row - 1) = True
Else
MsgBox "aradýðýnýz kayýt bulunamadý."
End If
End Sub

Private Sub CommandButton9_Click()

Set bul = Sheets("MÜÞTERÝ").Range("d1:d65536").Find(TextBox3, lookat:=xlWhole)
If Not bul Is Nothing Then
UserForm7.ListBox1.Selected(bul.Row - 1) = True
Else
MsgBox "aradýðýnýz kayýt bulunamadý."
End If

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'On Error Resume Next
'UserForm8.TextBox1.Text = Sheets("ÜRÜNLER").Range("B" & ListBox1.ListIndex + 1).Value
'UserForm8.TextBox2.Text = Sheets("ÜRÜNLER").Range("C" & ListBox1.ListIndex + 1).Value
'UserForm8.TextBox3.Text = Sheets("ÜRÜNLER").Range("D" & ListBox1.ListIndex + 1).Value
'UserForm8.TextBox4.Text = Sheets("ÜRÜNLER").Range("E" & ListBox1.ListIndex + 1).Value
'UserForm8.TextBox1.Enabled = False
'UserForm8.TextBox2.Enabled = False
'UserForm8.TextBox3.Enabled = False
'UserForm8.TextBox4.Enabled = False
'UserForm8.CommandButton1.Enabled = False
'UserForm8.CommandButton2.Enabled = False
'UserForm8.Show
End Sub


Private Sub UserForm_Initialize()
Dim ts
Set ts = Sheets("MÜÞTERÝ")
ListBox1.Clear
ListBox1.ColumnCount = 5
ListBox1.ColumnWidths = "20;150;80;80;300"
ListBox1.RowSource = "MÜÞTERÝ!A1:e" & ts.Range("B" & Rows.Count).End(xlUp).Row
End Sub
