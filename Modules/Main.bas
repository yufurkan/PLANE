Public Sub Aktar()


Dim ws As Worksheet

Set ws = ThisWorkbook.Sheets("Sheet1")
    
    
Dim x As String
Dim y As String
x = ws.Range("J8").value
y = ws.Range("J9").value

If XYZControl(x, y) Then
Dim z As String
z = FindZ(x, y)
Dim lRow As Long
    lRow = 0
    lRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    planecount = Sheet1.Range("J7").value
    Dim formulas As Variant
    ReDim formulas(1 To planecount)
    Dim newname As String
    Dim FileName As String

    
    For o = 1 To planecount  'formüller alınıyor
        formulas(o) = ws.Cells(o + 1, "B").value
    Next o

If FunctionsControl(formulas) Then
ThisWorkbook.izin = False


Dim chaptercount As Integer

    Path = ActiveWorkbook.Path
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
 
    ' Get Chapter List
    Sheets.Add.Name = "ChapterList"
    
    
    fileCount = Dir(Path & "\*.txt")
    FileName = Dir(Path & "\*.txt")
    Do While fileCount <> ""
        chaptercount = chaptercount + 1
        fileCount = Dir()
        newname = Left(FileName, Len(FileName) - 4) & "çizim"
    Loop
    
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = Path & "\*txt"
    
    ActiveWorkbook.Names.Add Name:="chaptername", RefersToR1C1:= _
        "=FILES(ChapterList!R1C1)"
    ' ActiveWorkbook.Names("chaptername").Comment = ""
    
    For i = 1 To chaptercount
        Range("b" & (i + 1)).value = i
    Next
    
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(INDEX(chaptername,RC[1]),"""")"
'    Range("A2").Select
'    'On Error Resume Next
'        Selection.AutoFill Destination:=Range("A2:A" & (chaptercount + 1)), Type:=xlFillDefault
    
    
    ' Write Down Each Chapter
    
    For t = 1 To chaptercount
    current_file = Worksheets("ChapterList").Cells(t + 1, 1).value
    Sheets.Add.Name = current_file


    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & Path & "\" & current_file, Destination:=Range("$A$1"))
        .Name = current_file
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1254
        .TextFileStartRow = 7
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "|"
        .TextFileDecimalSeparator = ","
        .TextFileThousandsSeparator = "."
        .TextFileColumnDataTypes = Array(1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    

    Columns("A:A").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$" & i, Selection.End(xlDown)).AutoFilter Field:=1, Criteria1:="=", _
        Operator:=xlOr, Criteria2:="=-*"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
    
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").EntireColumn.AutoFit
    Range("A1").Select
    
    

    Next 'aşşağıdaki kodlar dikkatlice next içine taşınırsa kod dizindeki her txt dosysı için işlem yapabilir. Deneme aşamasında olduğu için aşşağıya yazıldı
    
    Application.DisplayAlerts = False
    Sheets("ChapterList").Delete
    Application.DisplayAlerts = True
    Sheets("Sheet1").Activate
    Columns("D:H").Select
    Selection.ClearContents
    Sheets(current_file).Activate
    
    
    
    Set ws = Sheets(current_file)
    
    ws.Range("E:E").Copy Destination:=ws.Range("AO:AO")
    
    'burada min nokta ile min noktanın uzaklığı toplanarak max noktalar bulunuyor ve x y z için yeni satırlarına yazılıyor
        lRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).row
    ws.Range("AP1:AP" & lRow).formula = "=AG1+AJ1"
        lRow = ws.Cells(ws.Rows.Count, "AH").End(xlUp).row
    ws.Range("AQ1:AQ" & lRow).formula = "=AH1+AK1"
        lRow = ws.Cells(ws.Rows.Count, "AI").End(xlUp).row
    ws.Range("AR1:AR" & lRow).formula = "=AI1+AL1"
    ws.Range("AP1").value = "BBLx[mm]"
    ws.Range("AQ1").value = "BBLy[mm]"
    ws.Range("AR1").value = "BBLz[mm]"
    Dim ar As Integer
    Dim n As Integer

    ar = ws.Range("AR1").column
    

    'çalışma alanı max ve min noktalardan belirlenerek ilk şekil oldu şimdi
    Dim shape As Collection
    Dim shapes As New Collection
    Dim shapesP As New Collection
    
    Set shape = IdentifyShape(x, y, z)
    
    shapeP.Add shape
'----------------
    

    Set shapes = ShapeCutter(shapesP, formulas)
    
    Dim alanAd As String
    Dim alanSayı As Integer
    alanSayı = 0
    
    Dim gXp As String
    Dim gYp As String
    Dim gZp As String
    gXp = GColumnAta(ws, x) 'gravity x point
    gYp = GColumnAta(ws, y)
    gZp = GColumnAta(ws, z)
    
    
    Dim pXMaxC As String
    Dim pXMinC As String
    Dim pYMaxC As String
    Dim pYMinC As String
    Dim pZMaxC As String
    Dim pZMinC As String
   
    
    pXMaxC = PmaxColumnAta(ws, x)
    pXMinC = PminColumnAta(ws, x)
    pYMaxC = PmaxColumnAta(ws, y)
    pYMinC = PminColumnAta(ws, y)
    pZMaxC = PmaxColumnAta(ws, z)
    pZMinC = PminColumnAta(ws, z)
            
    
    
    'Her parça için hesap---------------------------------------------------------------------------->>>>>>><<<<<<------------------------------------------------------------------------
    Dim accept As Integer
    accept = 0 'şuan kullanmadım
    Dim p As Integer
    Dim s As Integer
    lRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    For s = 1 To shapes.Count 'Her alan için
        Dim addcolumns As Boolean
        addcolumns = True
        For p = 2 To lRow 'her parça için
            Dim intercont As Boolean
            intercont = True
            

            Dim piece As New Collection
            Set piece = New Collection
            Set piece = IdentifyPiece(x, y, z, p, ws) 'parça tanımlanıyor
            
            Dim grX As Double
            Dim grY As Double
            Dim grZ As Double

            grX = ws.Cells(p, gXp).value
            grY = ws.Cells(p, gYp).value
            grZ = ws.Cells(p, gZp).value
            
            'burayı ilerde fonksiyona al---
            Dim dd As Integer
            For shk = 1 To shapes.Count
                For dd = 1 To shapes(shk).Count
                    Dim nn As New Nokta
                    nn.x = shapes(shk)(dd).x
                    nn.y = shapes(shk)(dd).y
                    nn.z = shapes(shk)(dd).z
                    If ExceptionCheck(piece, nn) Then 'Hata nokta yerine kenar kontrolü yapılmalı!
                        intercont = False
                    End If
                Next dd
            Next shk
            '----------------------
            
            
        If intercont Then 'plane çakışması parçanın içine gelmiyorsa
        'else ekleyerek satırı kırmızıya boya
        
            'tanımlamalar
            Dim areapiece As Collection
            Set areapiece = New Collection
            Dim mass As Double
            mass = ws.Cells(p, "E").value
            Dim in1 As Boolean
            in1 = False
            Dim in2 As Boolean
            in2 = False
            Dim k As Integer
            k = 1
            
            'bir şekil için
            Dim kesismeler As Collection
            Set kesismeler = New Collection
            Dim yuz As Integer
            For yuz = 1 To 6
                For k = 1 To 4 'parçanın kenarları son nokta aşım olmasın diye

                'Aşağıdaki kodlarda mantık hataları bulunuyor olabilir yeterli test edilmedi 
                '-> Güncelleme: Kodlar Çalışıyor Interia Hesaplamalarını kontrol et

                Dim ptt1 As Nokta, ptt2 As Nokta
                Set ptt1 = piece(k)
                
                If k = piece.Count Then
                    Set ptt2 = piece(1)
                Else
                    Set ptt2 = piece(k + 1)
                End If
                
                in1 = InsideCheck(shapes(s), ptt1)
                in2 = InsideCheck(shapes(s), ptt2)
                
          
                Dim intersectPt As Variant
                For pk = 1 To shapes(s).Count
                    Dim areapt1 As Variant, areapt2 As Variant
                    Set areapt1 = shapes(s)(pk)
                    If pk = shapes(s).Count Then
                        Set areapt2 = shapes(s)(1)
                    Else
                        Set areapt2 = shapes(s)(pk)
                    End If
                    
                    Set intersectPt = intersectPoint(ptt1, ptt2, areapt1.x, areapt1.y, areapt2.x, areapt2.y)
                    If Not intersectPt Is Nothing Then
                        kesismeler.Add intersectPt
                    End If

                Next pk
                
                If Not in1 And Not in2 And kesismeler.Count > 1 Then
                    For b = 1 To kesismeler.Count
                        areapiece.Add kesismeler(b)
                    Next b
                    
                ElseIf Not in2 = in1 Then
                    If in2 Then
                        areapiece.Add kesismeler(1)
                        areapiece.Add ptt2
                    Else
                        areapiece.Add kesismeler(1)
                    End If
                ElseIf in1 And in2 Then
                    areapiece.Add ptt2
                End If '
                                                          
 
                    
            Next k
            Next yuz
            ' Bilgiler buada hesaplanacak ve yazdırılacak
            If areapiece.Count > 2 Then 'öngörülmemiş bir durum oluşmaması için kontrol
                Dim ratio As New Nokta
                Dim gcenter As New Nokta
                Dim newGCenter As New Nokta
                Dim arearatio As Double
                Dim newmass As Double
                gcenter.x = grX
                gcenter.y = grY
                Dim center As Nokta
                Set ratio = FindCenterOfMassRatio(piece, gcenter)
                Set center = FindCentroid(areapiece)
                Set newGCenter = New Nokta
                newGCenter.x = center.x + ratio.x * (gcenter.x - center.x) 'yeni ağırlk merkezi eski şekilin matematiksel merkezinin ağırlık merkezine oranının  yeni şeklin matematiksel merkezine oranı hesaplanıyor
                newGCenter.y = center.y + ratio.y * (gcenter.y - center.y)
                Call KonumYazdir(ws, x, p, s - 1, newGCenter.x)
                Call KonumYazdir(ws, y, p, s - 1, newGCenter.y)
                Call KonumYazdir(ws, z, p, s - 1, grZ)
                Dim piecearea As Double
                Dim newarea As Double
                piecearea = CalculateArea(piece)
                newarea = CalculateArea(areapiece)
                arearatio = newarea / piecearea
                newmass = arearatio * mass
                Call MassYazdir(ws, p, s - 1, newmass)
                
                '--- Son Düzeltme noktası ---
                ' Uçağın Referans Noktasına Göre Momentleri Al
                
                Dim MomentX As Double, MomentY As Double, MomentZ As Double
                
                MomentX = newmass * newGCenter.x  ' Kütle * X Mesafesi
                MomentY = newmass * newGCenter.y  ' Kütle * Y Mesafesi
                MomentZ = newmass * grZ           ' Kütle * Z Mesafesi 
                
                'Momentleri Yazdır 
                ws.Cells(p, "AZ").value = MomentX
                ws.Cells(p, "BA").value = MomentY
                ws.Cells(p, "BB").value = MomentZ
                
                ' INERTIA HESABI 

                Dim Ixx As Double, Iyy As Double, Izz As Double
                Dim Ixy As Double, Ixz As Double, Iyz As Double
                
                ' Ixx: X ekseninde dönmeye karşı direnç (Y ve Z uzaklıklarına bağlı)
                Ixx = newmass * (newGCenter.y ^ 2 + grZ ^ 2)
                
                ' Iyy: Y ekseninde dönmeye karşı direnç (X ve Z uzaklıklarına bağlı)
                Iyy = newmass * (newGCenter.x ^ 2 + grZ ^ 2)
                
                ' Izz: Z ekseninde dönmeye karşı direnç (X ve Y uzaklıklarına bağlı)
                Izz = newmass * (newGCenter.x ^ 2 + newGCenter.y ^ 2)
                
                ' Çarpım Ataletleri (Product of Inertia - Simetri bozukluğu için)
                Ixy = newmass * newGCenter.x * newGCenter.y
                Ixz = newmass * newGCenter.x * grZ
                Iyz = newmass * newGCenter.y * grZ
                
                'Inertia Değerlerini Yazdır
                ws.Cells(p, "BC").value = Ixx
                ws.Cells(p, "BD").value = Iyy
                ws.Cells(p, "BE").value = Izz
                ws.Cells(p, "BF").value = Ixy
                ws.Cells(p, "BG").value = Ixz
                ws.Cells(p, "BH").value = Iyz
                
                ' --- Son Düzeltme noktası Bitti ---
                
            End If
        Else 'intercont
        
        End If 'intercont
        Next p
    Next s

    If alanSayı <= planecount + 1 Then 'hata durumuna göre isimlendirme
        
        alanAd = "Sec."
        Else
        alanAd = "Area."
            
    End If
    
    
    
    
    ws.Cells.HorizontalAlignment = xlLeft
    

    
    
' çizdirme komutları burdan başlıyor
    Set ws = Sheets.Add
    ws.Name = newname
    
    
    


'burası zamanın yetmemesi sebeei ile tamamlanmadı
'yapılacaklar:
'X Y kordinat doğrularını çizin
'parçaları çizdirin
'her şeklin kenarını çizdirin
'foksiyon doğrularını çizdirin
'farklı renkler kulanın
   ' For Each shapeC In shapes
   '    Dim b As Integer
   '     b = shapes.Count
   '     For p = 1 To shapeC.Count
    '        If p >= shapeC.Count - 1 Then
    '            CizgiCiz newname, shapeC(p), shapeC(1)
   '         Else
    '            CizgiCiz newname, shapeC(p), shapeC(p + 1)
   '         End If
   '     Next p
  '  Next shapeC
    
    
    
ThisWorkbook.izin = True

End If
End If
End Sub






