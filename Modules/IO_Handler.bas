Attribute VB_Name = "IO_Handler"
Option Explicit

Public Sub KonumYazdir(ws As Worksheet, eksen As String, row As Integer, col As Integer, konum As Double)
    ' Eksenine göre (X, Y, Z) ilgili sütuna konumu yazar
    ' AT, AU, AV sütunları kullanılıyor

    If eksen = "x" Then
        ws.Cells(row, col * 10 + Columns("AT").column).value = konum
    End If
    If eksen = "y" Then
        ws.Cells(row, col * 10 + Columns("AU").column).value = konum
    End If
    If eksen = "z" Then
        ws.Cells(row, col * 10 + Columns("AV").column).value = konum
    End If
End Sub

Public Sub MassYazdir(ws As Worksheet, row As Integer, col As Integer, deger As Double)
   ' Hesaplanan kütleyi AS sütununa yazar
   ws.Cells(row, col * 10 + Columns("AS").column).value = deger
End Sub