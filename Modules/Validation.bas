Attribute VB_Name = "Validation"
Option Explicit

Public Function FunctionsControl(fonksiyonlar As Variant) As Boolean
    'Fonksiyonların formatları Regex ile kontrol edilir
    'Format: Ax+By+Cz=D şeklinde olmalı

    Dim func As Object
    Set func = CreateObject("VBScript.RegExp")
    ' Regex Deseni: (Sayı)x(İşaretSayı)y(İşaretSayı)z=(Sayı)
    func.Pattern = "(-?\d+\.?\d*)x([+-]\d+\.?\d*)y([+-]\d+\.?\d*)z=([+-]?\d+\.?\d*)?"
    
    Dim k As Integer
    k = 1
    
    Dim fonksiyon As Variant
    For Each fonksiyon In fonksiyonlar
        Dim matches As Object
        Set matches = func.Execute(fonksiyon)
        
        If matches.Count = 0 Then
            If fonksiyon = "" Then
                MsgBox ("Hata " & k & " numaralı satır boş bırakılmış. Lütfen kontrol edip tekrar deneyin!")
            Else
                MsgBox ("Hata girdiğiniz " & k & " numaralı fonksiyon uygun değil. Lütfen kontrol edip tekrar deneyin! Girilen değer: " & fonksiyon)
            End If
            FunctionsControl = False
            Exit Function
        End If
        
        k = k + 1
    Next fonksiyon

    FunctionsControl = True
End Function