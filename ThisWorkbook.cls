VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
Application.DisplayFullScreen = True
Dim wmi, cpu, cpuid
  Set wmi = GetObject("winmgmts:")
  For Each cpu In wmi.InstancesOf("Win32_Processor")
   cpuid = cpuid + cpu.ProcessorID
  Next
  MsgBox cpuid
 If cpuid = "BFEBFBFF000206A7" Then
ThisWorkbook.Activate
Else
If cpuid = "BFEBFBFF000306C3" Then
ThisWorkbook.Activate
Else
If cpuid = "BFEBFBFF000306D4" Then
ThisWorkbook.Activate
Else
MsgBox "Bu dosyayý açmak için yetkiniz bulunmamaktadýr.", vbMsgBoxRtlReading, "MG PROGRAM"
Application.Quit
Application.DisplayAlerts = False
End If
End If
End If
'Dim FS As Object, Surucu As Object
'Set FS = CreateObject("Scripting.FileSystemObject")
'Set Surucu = FS.GetDrive("C:")
'seri = Surucu.SerialNumber
'If seri = 548683664 Then
'ThisWorkbook.Activate
'Else: MsgBox "Bu dosyayý açmak için yetkiniz bulunmamaktadýr.", vbMsgBoxRtlReading, "MG PROGRAM"
'Application.Quit
'End If
'Set Surucu = Nothing
'Set FS = Nothing

UserForm3.Show
End Sub
