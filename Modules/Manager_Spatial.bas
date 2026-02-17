Attribute VB_Name = "Manager_Spatial"
Option Explicit

' --- Shape Identification Functions (Eski IdentifyShapes) ---

Public Function IdentifyShape(x As String, y As String, z As String) As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Sayfa adı gerekirse buradan değişir

    Dim minmaxX As Variant
    Dim minmaxY As Variant
    Dim minmaxZ As Variant
    
    minmaxX = EksenAta(ws, x)
    minmaxY = EksenAta(ws, y)
    minmaxZ = EksenAta(ws, z)
    
    Dim cr1 As New Nokta, cr2 As New Nokta, cr3 As New Nokta, cr4 As New Nokta
    Dim cr5 As New Nokta, cr6 As New Nokta, cr7 As New Nokta, cr8 As New Nokta
    
    Dim yuz1 As New Collection, yuz2 As New Collection, yuz3 As New Collection
    Dim yuz4 As New Collection, yuz5 As New Collection, yuz6 As New Collection
    
    cr1.x = minmaxX(0): cr1.y = minmaxY(1): cr1.z = minmaxZ(0)
    cr2.x = minmaxX(1): cr2.y = minmaxY(1): cr2.z = minmaxZ(0)
    cr3.x = minmaxX(1): cr3.y = minmaxY(1): cr3.z = minmaxZ(1)
    cr4.x = minmaxX(0): cr4.y = minmaxY(1): cr4.z = minmaxZ(1)
    cr5.x = minmaxX(0): cr5.y = minmaxY(0): cr5.z = minmaxZ(0)
    cr6.x = minmaxX(1): cr6.y = minmaxY(0): cr6.z = minmaxZ(0)
    cr7.x = minmaxX(1): cr7.y = minmaxY(0): cr7.z = minmaxZ(1)
    cr8.x = minmaxX(0): cr8.y = minmaxY(0): cr8.z = minmaxZ(1)
    
    '-----------------
    yuz1.Add cr1: yuz1.Add cr2: yuz1.Add cr3: yuz1.Add cr4
    yuz2.Add cr2: yuz2.Add cr6: yuz2.Add cr7: yuz2.Add cr3
    yuz3.Add cr5: yuz3.Add cr6: yuz3.Add cr7: yuz3.Add cr8
    yuz4.Add cr1: yuz4.Add cr5: yuz4.Add cr8: yuz4.Add cr4
    yuz5.Add cr4: yuz5.Add cr3: yuz5.Add cr7: yuz5.Add cr8
    yuz6.Add cr1: yuz6.Add cr2: yuz6.Add cr6: yuz6.Add cr5
    
    Dim shape As Collection
    Set shape = New Collection

    shape.Add yuz1: shape.Add yuz2: shape.Add yuz3
    shape.Add yuz4: shape.Add yuz5: shape.Add yuz6
    
    Set IdentifyShape = shape
End Function

Public Function IdentifyPiece(x As String, y As String, z As String, p As Integer, ws As Worksheet) As Collection
'parça tanımı
    Dim pXMaxC As String, pXMinC As String
    Dim pYMaxC As String, pYMinC As String
    Dim pZMaxC As String, pZMinC As String
    
    Dim maxX As Double, minX As Double
    Dim maxY As Double, minY As Double
    Dim minZ As Double, maxZ As Double
   
    pXMaxC = PmaxColumnAta(ws, x): pXMinC = PminColumnAta(ws, x)
    pYMaxC = PmaxColumnAta(ws, y): pYMinC = PminColumnAta(ws, y)
    pZMaxC = PmaxColumnAta(ws, z): pZMinC = PminColumnAta(ws, z)

    maxX = ws.Cells(p, pXMaxC).value: minX = ws.Cells(p, pXMinC).value
    maxY = ws.Cells(p, pYMaxC).value: minY = ws.Cells(p, pYMinC).value
    maxZ = ws.Cells(p, pZMaxC).value: minZ = ws.Cells(p, pZMinC).value
    
    Dim cr1 As New Nokta, cr2 As New Nokta, cr3 As New Nokta, cr4 As New Nokta
    Dim cr5 As New Nokta, cr6 As New Nokta, cr7 As New Nokta, cr8 As New Nokta
    
    Dim yuz1 As New Collection, yuz2 As New Collection, yuz3 As New Collection
    Dim yuz4 As New Collection, yuz5 As New Collection, yuz6 As New Collection
    
    cr1.x = minX: cr1.y = maxY: cr1.z = minZ
    cr2.x = maxX: cr2.y = maxY: cr2.z = minZ
    cr3.x = maxX: cr3.y = maxY: cr3.z = maxZ
    cr4.x = minX: cr4.y = maxY: cr4.z = maxZ
    cr5.x = minX: cr5.y = minY: cr5.z = minZ
    cr6.x = maxX: cr6.y = minY: cr6.z = minZ
    cr7.x = maxX: cr7.y = minY: cr7.z = maxZ
    cr8.x = minX: cr8.y = minY: cr8.z = maxZ
    
    '-----------------
    yuz1.Add cr1: yuz1.Add cr2: yuz1.Add cr3: yuz1.Add cr4
    yuz2.Add cr2: yuz2.Add cr6: yuz2.Add cr7: yuz2.Add cr3
    yuz3.Add cr5: yuz3.Add cr6: yuz3.Add cr7: yuz3.Add cr8
    yuz4.Add cr1: yuz4.Add cr5: yuz4.Add cr8: yuz4.Add cr4
    yuz5.Add cr4: yuz5.Add cr3: yuz5.Add cr7: yuz5.Add cr8
    yuz6.Add cr1: yuz6.Add cr2: yuz6.Add cr6: yuz6.Add cr5
    
    Dim shape As Collection
    Set shape = New Collection
    
    shape.Add yuz1: shape.Add yuz2: shape.Add yuz3
    shape.Add yuz4: shape.Add yuz5: shape.Add yuz6
    
    Set IdentifyPiece = shape
End Function

' --- Axis & Coordinate Helpers (Eski XYZ) ---

Public Function XYZControl(yatay As String, dikey As String) As Boolean
    If yatay = dikey Then
        MsgBox "dikey ve yatay eksen değerleri aynı olamaz!"
        XYZControl = False
        Exit Function
    End If

    Dim eksen(1) As String
    eksen(0) = yatay: eksen(1) = dikey
    Dim XYZ(2) As String
    XYZ(0) = "x": XYZ(1) = "y": XYZ(2) = "z"

    Dim es As Integer
    es = 0
    Dim i As Integer, j As Integer
    
    For i = 0 To 1
        For j = 0 To 2
            If eksen(i) = XYZ(j) Then
                es = es + 1
            End If
        Next j
    Next i

    If es = 2 Then
        XYZControl = True
        Exit Function
    End If

    MsgBox "Lütfen eksen bilgilerini doğru bir şekilde doldurduğunuzdan emin olun. x y ve z değerleri küçük harf ile girilmelidir"
    XYZControl = False
End Function

Public Function EksenAta(ws As Worksheet, eksen As String) As Variant
    Dim mine As Double
    Dim maxe As Double
    Dim sonuc As Variant

    If eksen = "x" Then
        mine = MinVal(ws, "AG")
        maxe = MaxVal(ws, "AP")
    End If

    If eksen = "y" Then
        mine = MinVal(ws, "AH")
        maxe = MaxVal(ws, "AQ")
    End If

    If eksen = "z" Then
        mine = MinVal(ws, "AI")
        maxe = MaxVal(ws, "AR")
    End If

    sonuc = Array(mine, maxe)
    EksenAta = sonuc
End Function

Public Function FindZ(x As String, y As String) As String
    'burada üzerinde çalışılan ve bilinen 2 eksenin dışında 3. eksenin ne olduğu bulunuyor
    If (x = "x" And y = "y") Or (x = "y" And y = "x") Then
        FindZ = "z"
    ElseIf (x = "x" And y = "z") Or (x = "z" And y = "x") Then
        FindZ = "y"
    ElseIf (x = "y" And y = "z") Or (x = "z" And y = "y") Then
        FindZ = "x"
    End If
End Function

Public Function GColumnAta(ws As Worksheet, eksen As String) As String
    If eksen = "x" Then GColumnAta = "F"
    If eksen = "y" Then GColumnAta = "G"
    If eksen = "z" Then GColumnAta = "H"
End Function

Public Function PmaxColumnAta(ws As Worksheet, eksen As String) As String
    If eksen = "x" Then PmaxColumnAta = "AP"
    If eksen = "y" Then PmaxColumnAta = "AQ"
    If eksen = "z" Then PmaxColumnAta = "AR"
End Function

Public Function PminColumnAta(ws As Worksheet, eksen As String) As String
    If eksen = "x" Then PminColumnAta = "AG"
    If eksen = "y" Then PminColumnAta = "AH"
    If eksen = "z" Then PminColumnAta = "AI"
End Function

Public Function MinVal(ws As Worksheet, icolumn As String) As Double
    Dim value As Double
    Dim lRow As Long
    Dim cRange As Range

    lRow = ws.Cells(ws.Rows.Count, icolumn).End(xlUp).row
    Set cRange = ws.Range(icolumn & "1:" & icolumn & lRow)
    value = Application.WorksheetFunction.Min(cRange)
    
    MinVal = value
End Function

Public Function MaxVal(ws As Worksheet, icolumn As String) As Double
    Dim value As Double
    Dim lRow As Long
    Dim cRange As Range

    lRow = ws.Cells(ws.Rows.Count, icolumn).End(xlUp).row
    Set cRange = ws.Range(icolumn & "1:" & icolumn & lRow)
    value = Application.WorksheetFunction.Max(cRange)
    
    MaxVal = value
End Function