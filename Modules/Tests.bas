Attribute VB_Name = "Tests"
Option Explicit

'buranın altı fonksiyonları test etmek için

Public Sub TestKesisimBul()

    ' Test için noktalar ve yön vektörleri tanımlıyoruz
    Dim rayStart As Nokta
    Dim rayDir As Nokta
    Dim face As Collection
    Dim intersect As Variant
    Dim point1 As Nokta, point2 As Nokta, point3 As Nokta, point4 As Nokta
    
    ' Başlangıç noktası
    Set rayStart = New Nokta
    rayStart.x = 1
    rayStart.y = 1
    rayStart.z = 1
    
    ' Yön vektörü
    Set rayDir = New Nokta
    rayDir.x = 0
    rayDir.y = 0
    rayDir.z = -1
    
    'yüzeyler
    Set face = New Collection
    Set point1 = New Nokta
    point1.x = 0
    point1.y = 0
    point1.z = 0
    face.Add point1
    
    Set point2 = New Nokta
    point2.x = 1
    point2.y = 0
    point2.z = 0
    face.Add point2
    
    Set point3 = New Nokta
    point3.x = 0
    point3.y = 1
    point3.z = 0
    face.Add point3
    
    Set point4 = New Nokta
    point4.x = 0
    point4.y = 1
    point4.z = 0
    face.Add point4

    ' Işın ile düzlemin kesişim noktasını bul
    intersect = KesisimBul(rayStart, rayDir, face)

    ' Sonuçları kontrol et ve yazdır
    If Not intersect Is Nothing Then
        Debug.Print "Kesişim Noktası: X=" & intersect.x & " Y=" & intersect.y & " Z=" & intersect.z
    Else
        Debug.Print "Kesişim yok."
    End If
    
End Sub


Sub TestRayIntersection()
    Dim p1 As New Nokta, p2 As New Nokta
    Dim r0 As New Nokta, direction As New Nokta
    Dim result As Nokta

    ' Doğru parçası 1: (p1, p2)
    p1.x = 5: p1.y = 5: p1.z = 4
    p2.x = 5: p2.y = 1: p2.z = 4
    
    ' Işının başlangıç noktası ve yönü
    r0.x = 4: r0.y = 5: r0.z = 4
    direction.x = 1: direction.y = 0: direction.z = 0
    
    ' Kesişimi kontrol et
    Set result = RayIntersectsEdge(r0, direction, p1, p2)
    
    If Not result Is Nothing Then
        MsgBox "Isin doğru parçası ile kesisiyor."
    Else
        MsgBox "Isin doğru parçası ile kesiSmiyor."
    End If
End Sub


Sub TestPointInPolygon3D()


    Dim face As New Collection
    Dim point1 As New Nokta
    Dim point2 As New Nokta
    Dim point3 As New Nokta
    Dim point4 As New Nokta
    
 
    point1.x = 0: point1.y = 0: point1.z = 0
    point2.x = 2: point2.y = 0: point2.z = 0
    point3.x = 2: point3.y = 2: point3.z = 0
    point4.x = 0: point4.y = 2: point4.z = 0
    
    face.Add point1
    face.Add point2
    face.Add point3
    face.Add point4
    

    Dim testPoint As New Nokta
    testPoint.x = 1: testPoint.y = 0.1: testPoint.z = 0
    
    If PointInPolygon3D(testPoint, face) Then
        MsgBox "Nokta yüzeyin içinde."
    Else
        MsgBox "Nokta yüzeyin dışında."
    End If
    
End Sub



Public Sub Planedene()

    Dim cr1 As New Nokta
    Dim cr2 As New Nokta
    Dim cr3 As New Nokta
    Dim cr4 As New Nokta
    Dim cr5 As New Nokta
    Dim cr6 As New Nokta
    Dim cr7 As New Nokta
    Dim cr8 As New Nokta
    
    Dim yuz1 As New Collection
    Dim yuz2 As New Collection
    Dim yuz3 As New Collection
    Dim yuz4 As New Collection
    Dim yuz5 As New Collection
    Dim yuz6 As New Collection
    
    cr1.x = 0
    cr1.y = 10
    cr1.z = 0
    
    cr2.x = 10
    cr2.y = 10
    cr2.z = 0
    
    cr3.x = 10
    cr3.y = 10
    cr3.z = 10
    
    cr4.x = 0
    cr4.y = 10
    cr4.z = 10
     
    cr5.x = 0
    cr5.y = 0
    cr5.z = 0
    
    cr6.x = 10
    cr6.y = 0
    cr6.z = 0
    
    cr7.x = 10
    cr7.y = 0
    cr7.z = 10
    
    cr8.x = 0
    cr8.y = 0
    cr8.z = 10
    
    '-----------------
    
    yuz1.Add cr1
    yuz1.Add cr2
    yuz1.Add cr3
    yuz1.Add cr4
    
    yuz2.Add cr2
    yuz2.Add cr6
    yuz2.Add cr7
    yuz2.Add cr3
    
    yuz3.Add cr5
    yuz3.Add cr6
    yuz3.Add cr7
    yuz3.Add cr8
    
    yuz4.Add cr1
    yuz4.Add cr5
    yuz4.Add cr8
    yuz4.Add cr4
    
    yuz5.Add cr4
    yuz5.Add cr3
    yuz5.Add cr7
    yuz5.Add cr8
    
    yuz6.Add cr1
    yuz6.Add cr2
    yuz6.Add cr6
    yuz6.Add cr5
    
    Dim shape As Collection
    Set shape = New Collection
   

    shape.Add yuz1
    shape.Add yuz2
    shape.Add yuz3
    shape.Add yuz4
    shape.Add yuz5
    shape.Add yuz6
    Dim rayDir As Nokta
    Dim intersectCount As Integer
    intersectCount = 0
    Dim intersect As Variant
    Dim face As Collection 'DİKKAT: face burada tanımlanmamış, kodda hata verebilir.
    
    Dim point As New Nokta
    point.x = 5
    point.y = 3
    point.z = 1
    
    Set rayDir = New Nokta
    rayDir.x = 1
    rayDir.y = 0
    rayDir.z = 0
    
    
        Set intersect = KesisimBul(point, rayDir, face)
       
        If Not intersect Is Nothing Then ' Eğeer ışın yüzeyle kesişiyorsa
            
            ' DİKKAT: PointOnFace fonksiyonu diğer modüllerde yok. Hata verebilir.
            ' If PointOnFace(face, intersect) Then ' Eğer nokta yüzey üzerinde ise içeride sayılmamalı
            '    MsgBox "İçerde"
            '    Exit Sub
            ' End If
            intersectCount = intersectCount + 1
        End If
    
   
    MsgBox "dışarda"
End Sub

Sub TestIntersection()
    Dim p1 As New Nokta, p2 As New Nokta
    Dim q1 As New Nokta, q2 As New Nokta
    Dim result As Boolean

    ' Doğru parçası 1: (p1, p2)
    p1.x = 0: p1.y = 5: p1.z = 4345345
    p2.x = 10: p2.y = 5: p2.z = 4445654
    
    ' Doğru parçası 2: (q1, q2)
    q1.x = 6: q1.y = 0: q1.z = 54535
    q2.x = 5: q2.y = 10: q2.z = 5453
    
    ' Kesişimi kontrol et
    ' DİKKAT: IsLineSegmentsIntersect fonksiyonu Core_Geometry'de yok. Hata verebilir.
    ' result = IsLineSegmentsIntersect(p1, p2, q1, q2)
    
    If result Then
        MsgBox "Dogru parcalari kesisiyor."
    Else
        MsgBox "Dogru parcalari kesismiyor."
    End If
End Sub
Public Sub Fdene()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim nokta1 As New Nokta
    Dim nokta2 As New Nokta
    nokta1.x = 5
    nokta2.x = 6
    nokta1.y = 5
    nokta2.y = 6
    
    Dim point As New Nokta
    ' CalculateIntersection olarak güncellendi (Eski: YFuncControl)
    Set point = CalculateIntersection(ws.Range("B2").value, nokta1, nokta2)
    
    If point Is Nothing Then ' "Noting" hatası düzeltildi
        MsgBox "kesişmeme"
    Else
        MsgBox point.x & ", " & point.y
    End If
 
End Sub
Public Sub shapedene()
    Dim cr1 As New Nokta
    Dim cr2 As New Nokta
    Dim cr3 As New Nokta
    Dim cr4 As New Nokta
    cr1.x = 5
    cr1.y = 5
    cr2.x = 5
    cr2.y = 6
    cr3.x = 7
    cr3.y = 9
    cr4.x = 6
    cr4.y = 9
    Dim shape As Collection
    Set shape = New Collection
    Dim shapesP As New Collection

    shape.Add cr1
    shape.Add cr2
    shape.Add cr3
    shape.Add cr4
    shapesP.Add shape
    Dim shapes As New Collection
    shapes.Add shape
    
    MsgBox shapes(1)(3).x
End Sub

Public Sub InsideDene()

'burası Insidecheck fonksiyonunu test etmek için
Dim point As New Nokta
Dim result As Boolean
 point.y = 25
 point.x = 10
 
 
         Dim cr1 As New Nokta
    Dim cr2 As New Nokta
    Dim cr3 As New Nokta
    Dim cr4 As New Nokta
    cr1.x = 10
    cr1.y = 10
    cr2.x = 50
    cr2.y = 10
    cr3.x = 50
    cr3.y = 50
    cr4.x = 10
    cr4.y = 50
    Dim shape As Collection
    Set shape = New Collection
     
     
    shape.Add cr1
    shape.Add cr2
    shape.Add cr3
    shape.Add cr4

    result = InsideCheck(shape, point)
    If result = False Then
    MsgBox "içerde değil"
    Else
    MsgBox "içerde"
    End If
End Sub