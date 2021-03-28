VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
'Dim A, B, C, D, E, F As Integer
Sheets("RAPORLA").Unprotect "1234"
If ComboBox1.Value = "" Then
Else
Range("A2").AutoFilter 1, ComboBox1.Value
RAPOR.ListBox1.AddItem "ÝÞLEM TÝPÝ = " & Sheets("RAPORLA").ComboBox1.Value
End If
If TextBox1.Value = "" Then
Else
Range("A2").AutoFilter 2, TextBox1.Value
RAPOR.ListBox1.AddItem "ÝÞLEM SIRA NO  = " & TextBox1.Value
End If
If TextBox2.Value = "" Then
Else
Range("A2").AutoFilter 3, TextBox2.Value
RAPOR.ListBox1.AddItem "MÜÞTERÝ ADI  = " & TextBox2.Value
End If
If TextBox3.Value = "" Then
Else
Range("A2").AutoFilter 4, TextBox3.Value
RAPOR.ListBox1.AddItem "TC NO = " & TextBox3.Value
End If
If TextBox4.Value = "" Then
Else
Range("A2").AutoFilter 5, TextBox4.Value
RAPOR.ListBox1.AddItem "TEL = " & TextBox3.Value
End If
If ComboBox2.Value = "" Then
Else
Range("A2").AutoFilter 6, ComboBox2.Value
RAPOR.ListBox1.AddItem "OPERATÖR = " & ComboBox2.Value
End If
If ComboBox3.Value = "" Then
Else
Range("A2").AutoFilter 7, ComboBox3.Value
RAPOR.ListBox1.AddItem "ÖDEME ÞEKLÝ = " & ComboBox3.Value
End If
If ComboBox4.Value = "" Then
Else
Range("A2").AutoFilter 8, ComboBox4.Value
RAPOR.ListBox1.AddItem "KART = " & ComboBox4.Value
End If
If ComboBox5.Value = "" Then
Else
Range("A2").AutoFilter 9, ComboBox5.Value
RAPOR.ListBox1.AddItem "ÜCRET DURUMU = " & ComboBox5.Value
End If
If ComboBox6.Value = "" Then
Else
Range("A2").AutoFilter 14, ComboBox6.Value
RAPOR.ListBox1.AddItem "SATIÞ / ÝÞLEM GÜNÜ = " & ComboBox6.Value
End If
If ComboBox7.Value = "" Then
Else
Range("A2").AutoFilter 15, ComboBox7.Value
RAPOR.ListBox1.AddItem "SATIÞ / ÝÞLEM AYI = " & ComboBox7.Value
End If
If ComboBox8.Value = "" Then
Else
Range("A2").AutoFilter 16, ComboBox8.Value
RAPOR.ListBox1.AddItem "SATIÞ / ÝÞLEM YILI = " & ComboBox8.Value
End If
If Sheets("RAPORLA").ComboBox5.Value = "" Then
Range("A2").AutoFilter 9, "REZERVASYON"
A = Range("AA1").Value
B = Range("AB1").Value
Range("A2").AutoFilter 9, "SATIN ALINDI / ÖDEME YAPILDI"
C = Range("AA1").Value
D = Range("AB1").Value
Range("A2").AutoFilter 9, "SATIN ALINDI / ÖDEME YAPILMADI"
E = Range("AA1").Value
F = Range("Ab1").Value
Range("A2").AutoFilter 9
Else
If Sheets("RAPORLA").ComboBox5.Value = "REZERVASYON" Then
A = Range("AA1").Value
B = Range("AB1").Value
Else
A = 0
B = 0
End If
If Sheets("RAPORLA").ComboBox5.Value = "SATIN ALINDI / ÖDEME YAPILDI" Then
C = Range("AA1").Value
D = Range("AB1").Value
Else
C = 0
D = 0
End If
If Sheets("RAPORLA").ComboBox5.Value = "SATIN ALINDI / ÖDEME YAPILMADI" Then
E = Range("AA1").Value
F = Range("AB1").Value
Else
E = 0
F = 0
End If
End If
RAPOR.TextBox1.Value = A
RAPOR.TextBox2.Value = B
RAPOR.TextBox3.Value = C
RAPOR.TextBox4.Value = D
RAPOR.TextBox5.Value = E
RAPOR.TextBox6.Value = F
RAPOR.TextBox7.Value = E
RAPOR.TextBox8.Value = D
RAPOR.TextBox9.Value = D - E
RAPOR.Show
Sheets("RAPORLA").Protect "1234"
'MsgBox A & B & C & D & E & F
End Sub

Private Sub CommandButton2_Click()

Sheets("RAPORLA").ComboBox1.Value = ""
Sheets("RAPORLA").ComboBox2.Value = ""
Sheets("RAPORLA").ComboBox3.Value = ""
Sheets("RAPORLA").ComboBox4.Value = ""
Sheets("RAPORLA").ComboBox5.Value = ""
Sheets("RAPORLA").ComboBox6.Value = ""
Sheets("RAPORLA").ComboBox7.Value = ""
Sheets("RAPORLA").ComboBox8.Value = ""
Sheets("RAPORLA").TextBox1.Value = ""
Sheets("RAPORLA").TextBox2.Value = ""
Sheets("RAPORLA").TextBox3.Value = ""
Sheets("RAPORLA").TextBox4.Value = ""
Sheets("RAPORLA").Unprotect "1234"
ActiveSheet.ShowAllData
Sheets("RAPORLA").Protect "1234"
End Sub

Private Sub CommandButton3_Click()
UserForm3.Show
End Sub
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
If Intersect(Target, Range("b:b")) Is Nothing Then Exit Sub
'Sheets("MALZEME LÝSTESÝ").Range("Z1").Value = Target.Cells.Offset(0, 0).Value
A = Target.Cells.Offset(0, -1).Value
B = Target.Cells.Offset(0, 0).Value
Sheets(A).Select
Sheets(A).Range("A4:A500000").Find(B).Activate
End Sub
