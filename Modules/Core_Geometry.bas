Attribute VB_Name = "Core_Geometry"
Option Explicit

' --- Shape Cutting & Intersection Logic (Eski ShapeFuncs ve Tests) ---

Public Function ShapeCutter(shapes As Collection, planes As Variant) As Collection
    Dim allIntersects As New Collection
    Dim newShapes As New Collection
    Dim shape1 As Collection
    Dim shape2 As Collection
    Dim samePlanePoints As New Collection
    Dim plane As String
    
    Dim shape As Collection
    Dim k As Integer
    Dim currentPoint As Nokta, nextPoint As Nokta
    Dim side1 As Double, side2 As Double
    Dim intersectPoint As Nokta
    Dim intersectPoints As New Collection
    
    For Each plane In planes 'her fonksiyon
        For Each shape In shapes 'her şekil için
            Set shape1 = New Collection
            Set shape2 = New Collection
            Set intersectPoints = New Collection
        

            For k = 1 To shape.Count 'her yüzey
                Set currentPoint = shape(k)
                If k = shape.Count Then
                    Set nextPoint = shape(1) ' son nokta için
                Else
                    Set nextPoint = shape(k + 1)
                End If
            
                'noktaların hangi tarafta olduğunu bul
                side1 = PointPlaneSide(currentPoint, plane)
                side2 = PointPlaneSide(nextPoint, plane)
            
            ' noktalar farkı tarafta ise kesişim bul ikisi de farlı tarafta ise sonuç negatif olur
                If side1 * side2 < 0 Then
                    Set intersectPoint = CalculateIntersection(currentPoint, nextPoint, plane)
                    intersectPoints.Add intersectPoint
                    shape1.Add intersectPoint
                    shape2.Add intersectPoint
                End If
            
                If side1 > 0 Then
                    shape1.Add currentPoint
                ElseIf side1 < 0 Then
                    shape2.Add currentPoint
                Else
                    shape1.Add currentPoint
                    shape2.Add currentPoint
                    samePlanePoints.Add currentPoint
                End If
            Next k
        
            If intersectPoints.Count > 1 Then
                For Each intersectPoint In intersectPoints
                    shape1.Add intersectPoint
                    shape2.Add intersectPoint
                Next intersectPoint
            End If
        
            If shape1.Count > 2 Then
                newShapes.Add shape1
            End If
            If shape2.Count > 2 Then
                newShapes.Add shape2
            End If
        Next shape
    Next plane
    
    Set ShapeCutter = newShapes
End Function

Public Function CalculateIntersection(ByVal fonksiyon As String, ByVal nokta1 As Nokta, ByVal nokta2 As Nokta) As Nokta 'eski ismi ile YFuncControl

'plane düzlemi ile alanların yüzey kenarlarının kesişimini hesaplar

    Dim a As Double, b As Double, c As Double, d As Double
    Dim x As Double, y As Double, z As Double
    Dim t As Double
    b
    Dim func As Object
    Set func = CreateObject("VBScript.RegExp")
    func.Pattern = "(-?\d+\.?\d*)x([+-]\d+\.?\d*)y([+-]\d+\.?\d*)z=([+-]?\d+\.?\d*)"
    
    Dim matches As Object
    Set matches = func.Execute(fonksiyon)
    
    If matches.Count = 0 Then
        Set YFuncControl = Nothing
        MsgBox "Şekiller hesaplanırken bir hata oluştu"
        Exit Function
    End If
    
    a = CDbl(matches(0).submatches(0))
    b = CDbl(matches(0).submatches(1))
    c = CDbl(matches(0).submatches(2))
    d = CDbl(matches(0).submatches(3))
    
    ' Çizgi denklemi (parametrik formda)
    Dim dx As Double, dy As Double, dz As Double
    dx = nokta2.x - nokta1.x
    dy = nokta2.y - nokta1.y
    dz = nokta2.z - nokta1.z
    
    ' sabit parametresini (d) al
    Dim numerator As Double, denominator As Double
    numerator = d - (a * nokta1.x + b * nokta1.y + c * nokta1.z)
    denominator = a * dx + b * dy + c * dz
    
    ' Eğer  0 ise doğru ve düzlem paraleldir, kesişim yoktur
    If denominator = 0 Then
        Set YFuncControl = Nothing
        Exit Function
    End If
    
    t = numerator / denominator
    
    ' Kesişim noktasını bul
    x = nokta1.x + t * dx
    y = nokta1.y + t * dy
    z = nokta1.z + t * dz
    
    ' Kesişim noktası iki nokta arasında mı kontrol et
    If (x >= WorksheetFunction.Min(nokta1.x, nokta2.x) And x <= WorksheetFunction.Max(nokta1.x, nokta2.x) And _
        y >= WorksheetFunction.Min(nokta1.y, nokta2.y) And y <= WorksheetFunction.Max(nokta1.y, nokta2.y) And _
        z >= WorksheetFunction.Min(nokta1.z, nokta2.z) And z <= WorksheetFunction.Max(nokta1.z, nokta2.z)) Then
        
        Dim sonucson As New Nokta
        sonucson.x = x
        sonucson.y = y
        sonucson.z = z
        Set YFuncControl = sonucson
    Else
        Set YFuncControl = Nothing
    End If
End Function

Public Function PointPlaneSide(a As Double, b As Double, c As Double, d As Double, p As Nokta) As Integer
'NOKTANIN PLANE DÜZLEMİNE GÖRE KALDIĞI BÖLGEYİ BULUR
    Dim result As Double
    result = a * p.x + b * p.y + c * p.z + d

    If result > 0 Then
        PointPlaneSide = 1 ' Nokta düzlemin ön tarafında
    ElseIf result < 0 Then
        PointPlaneSide = -1 ' Nokta düzlemin arka tarafında
    Else
        PointPlaneSide = 0 ' Nokta düzlemin üzerinde
    End If
End Function

Public Function intersectPoint(p1 As Nokta, p2 As Nokta, x3 As Double, y3 As Double, x4 As Double, y4 As Double) As Nokta
    ' Bu fonksiyon 2 Boyutlu uzayda (X-Y düzleminde) iki doğru parçasının kesişimini bulur.
    ' Hatayı çözen sihirli değnek burasıdır.

    Dim d As Double
    Dim ua As Double
    Dim ub As Double
    
    ' Payda hesabı (Determinant)
    d = (y4 - y3) * (p2.x - p1.x) - (x4 - x3) * (p2.y - p1.y)
    
    ' Eğer d = 0 ise çizgiler paraleldir, kesişim yoktur.
    If d = 0 Then
        Set intersectPoint = Nothing
        Exit Function
    End If
    
    ' ua ve ub parametrelerini hesapla
    ua = ((x4 - x3) * (p1.y - y3) - (y4 - y3) * (p1.x - x3)) / d
    ub = ((p2.x - p1.x) * (p1.y - y3) - (p2.y - p1.y) * (p1.x - x3)) / d
    
    ' Eğer 0 <= ua <= 1 VE 0 <= ub <= 1 ise, kesişim çizgilerin üzerindedir.
    If ua >= 0 And ua <= 1 And ub >= 0 And ub <= 1 Then
        Dim newP As New Nokta
        
        ' Kesişim koordinatlarını hesapla
        newP.x = p1.x + ua * (p2.x - p1.x)
        newP.y = p1.y + ua * (p2.y - p1.y)
        
        ' Z ekseni için de enterpolasyon yapalım (Eksik kalmasın)
        newP.z = p1.z + ua * (p2.z - p1.z)
        
        Set intersectPoint = newP
    Else
        ' Kesişim var ama çizgilerin uzantısında (yani parça üzerinde değil)
        Set intersectPoint = Nothing
    End If

End Function

' --- Point Inside & Ray Casting Logic (Eski PieceDivide) ---

Public Function InsideCheck(shape As Collection, point As Nokta) As Boolean  'PointInShape3D
    Dim face As Collection
    Dim intersectCount As Integer
    Dim rayDir As Nokta
    intersectCount = 0
    
    Set rayDir = New Nokta
    rayDir.x = 1
    rayDir.y = 0
    rayDir.z = 0
    
    ' Şeklin her yüzeyi için
    For Each face In shape

        Dim intersect As Nokta
        intersect = KesisimBul(point, rayDir, face)
        
        If Not intersect Is Nothing Then
            
            If KesisimBul(face, intersect) Then 'DİKKAT BURA YENİDEN TASARLANMALI Eğer nokta yüzey üzerinde ise içeride sayılmamalı
                 
                intersectCount = intersectCount + 1
            End If
            
        End If
    Next face
    
    ' Kesişim sayısına göre içeride olup olmadığına karar ver
    If intersectCount Mod 2 = 1 Then
        PointInShape3D = True ' Tek kesişim varsa içeride
    Else
        PointInShape3D = False ' Çift kesişim varsa dışarıda
    End If
End Function

Public Function KesisimBul(rayStart As Nokta, rayDir As Nokta, face As Collection) As Variant  'İsmi RayIntersectsPlane idi
'plane ile kenar kesişimini hesaplar eğer bulursa sonucu PointInpoligon3D ye göndermek için döndürür

    Dim normal As Nokta
    Dim d As Double
    Dim t As Double
    Dim intersect As Nokta
    Set intersect = New Nokta
    
    
    ' Yüzeyin normal vektörünü hesapla
    Set normal = FaceNormal(face)
    ' Burada şeklin içinde olduğu düzlem ile doğrunun çakışmasını bulduktan sonra noktanın şeklin içinde olup olmadığını kontrol etmeye karar verdim
    
    ' Düzlem denklemi: Ax + By + Cz + D = 0
    d = -normal.x * face(1).x - normal.y * face(1).y - normal.z * face(1).z
    
    ' Ray: P = rayStart + t * rayDir
    Dim payda As Double
    payda = normal.x * rayDir.x + normal.y * rayDir.y + normal.z * rayDir.z
    
    If payda = 0 Then
        KesisimBul = Nothing ' Işın düzleme paralel, kesişim yok
        Exit Function
    End If
    
    ' t = -(Ax0 + By0 + Cz0 + D) / (A*dx + B*dy + C*dz)
    t = -(normal.x * rayStart.x + normal.y * rayStart.y + normal.z * rayStart.z + d) / payda
    
    ' Eğer t negatifse, ışın ters yönde gidiyor demektir, kesişim yok
    If t < 0 Then
        Set KesisimBul = Nothing
        Exit Function
    End If
    
    ' Kesişim noktasını hesapla: P = rayStart + t * rayDir
    Dim intersect As Nokta
    Set intersect = New Nokta
    intersect.x = rayStart.x + t * rayDir.x
    intersect.y = rayStart.y + t * rayDir.y
    intersect.z = rayStart.z + t * rayDir.z
    
    Set KesisimBul = intersect
    
End Function

' NIKTANIN DÜZLEMSEL ŞEKİL İÇİNDE OLUP OLMADIĞINI
Public Function PointInPolygon3D(point As Nokta, face As Collection) As Boolean
    
    Dim normal As Nokta
    Dim i As Integer
    Dim nextIndex As Integer
    Dim intersectionCount As Integer
    Dim edgeStart As Nokta
    Dim edgeEnd As Nokta
    Dim rayDir As New Nokta 'ray direction
    Dim intersection As Boolean
    Dim iPoint As Nokta
    Dim noktalar As New Collection
    
    Set normal = FaceNormal(face)
    
    ' ışın yönü
    Set rayDir = GetPerpendicularVector(normal)
    
    intersectionCount = 0
    
    ' Her kenar için kesişim
    For i = 1 To face.Count
        If i = face.Count Then
            nextIndex = 1
        Else
            nextIndex = i + 1
        End If
        
        Set edgeStart = face(i)
        Set edgeEnd = face(nextIndex)
        Set iPoint = RayIntersectsEdge(point, rayDir, edgeStart, edgeEnd)
        
        If IsPointOnEdge(point, edgeStart, edgeEnd) Then 'nokta kenarın üzerinde mi
            PointInPolygon3D = False
            Exit Function
        End If
        
        ' Işın bu kenarla kesişiyor mu
        If Not iPoint Is Nothing Then
            intersectionCount = intersectionCount + 1
            noktalar.Add iPoint
        End If
    Next i
    
    ' Aynı olan noktaların çıkarılması işlemi
    Dim cq As Integer
    cq = intersectionCount
    
Dim a As Integer, b As Integer
For a = 1 To cq
    ' Döngüde ters yönde gitmek ve silme işlemi sonrası indeks sorunlarını önlemek için geriye doğru sayıyoruz
    For b = cq To a + 1 Step -1
        ' Aynı noktayı tespit etme
        If ArePointsEqual(noktalar(a), noktalar(b)) Then
            noktalar.Remove b
            intersectionCount = intersectionCount - 1
            cq = cq - 1 ' Collection sayısını güncelle
        End If
    Next b
Next a
    
    ' Eğer kesişim sayısı tekse, nokta yüzeyin içindedir
    PointInPolygon3D = (intersectionCount Mod 2 = 1)
    
End Function

'    BURASI KESİŞİM NOKTASI VARSA NOKTAYI YOKSA NOTHİNG DÖNDÜRECEK ŞEKİLDE GÜNCELLENLELİ-UNUTMA
Public Function RayIntersectsEdge(point As Nokta, rayDir As Nokta, p1 As Nokta, p2 As Nokta) As Nokta

    Dim t As Double, u As Double, denom As Double
    Dim edge As Nokta, w As Nokta, crossEdgeRayDir As Nokta
    Dim crossStartEdge As Nokta
    Dim intersectionPoint As Nokta
    
    ' p1 ve p2'nin yön vektörü
    Set edge = New Nokta
    edge.x = p2.x - p1.x
    edge.y = p2.y - p1.y
    edge.z = p2.z - p1.z
    
    ' Nokta ile p1 arasındaki vektörü hesapla
    Set w = New Nokta
    w.x = point.x - p1.x
    w.y = point.y - p1.y
    w.z = point.z - p1.z
    
    ' Başlangıç noktasının kenar üzerinde olup olmadığını çapraz çarpımla kontrol et
    Set crossStartEdge = New Nokta
    crossStartEdge.x = w.y * edge.z - w.z * edge.y
    crossStartEdge.y = w.z * edge.x - w.x * edge.z
    crossStartEdge.z = w.x * edge.y - w.y * edge.x
    
    ' IŞININ KENAR ÜZERİNDE OLUP OLMADIĞININ KONTROLÜ_BURAYA GEREK YOK KENAR KONTROLÜ POINTINPOLIGON3D DE YAPILMALI-HALLEDİLDİ AMABAŞKA BİRYERDE KULLANILMA İHTİMALİ İÇİN KALDIRMADIM
    ' Eğer çapraz çarpım sıfırsa, ışın kenarın doğrultusunda başlıyor demektir
    If crossStartEdge.x = 0 And crossStartEdge.y = 0 And crossStartEdge.z = 0 Then
        ' Bu durumda, ışının p1 ve p2 arasında olup olmadığını kontrol etmeliyiz
        If point.x >= Min(p1.x, p2.x) And point.x <= Max(p1.x, p2.x) And _
            point.y >= Min(p1.y, p2.y) And point.y <= Max(p1.y, p2.y) And _
            point.z >= Min(p1.z, p2.z) And point.z <= Max(p1.z, p2.z) Then
            ' Eğer nokta p1 ve p2 arasında ise, ışın kenarın üzerinde başlıyor ve kesişim yoktur
            Set RayIntersectsEdge = Nothing
            Exit Function
        End If
    End If
    
    ' Kenar ile ışının yön vektörlerinin dış çarpımını hesapla
    Set crossEdgeRayDir = New Nokta
    crossEdgeRayDir.x = edge.y * rayDir.z - edge.z * rayDir.y
    crossEdgeRayDir.y = edge.z * rayDir.x - edge.x * rayDir.z
    crossEdgeRayDir.z = edge.x * rayDir.y - edge.y * rayDir.x

    ' Payda (denom)
    denom = crossEdgeRayDir.x ^ 2 + crossEdgeRayDir.y ^ 2 + crossEdgeRayDir.z ^ 2
    
    ' Eğer denom sıfırsa, ışın kenara paraleldir ve kesişim yoktur
    If denom = 0 Then
        Set RayIntersectsEdge = Nothing
        Exit Function
    End If
    
    ' En yakın mesafeyi kontrol eden fonksiyonu çağır ve eğer mesafe uygun değilse kesişim olmadığını döndür
    If Not ClosestDistanceBetweenLines(point, rayDir, p1, edge) Then
        Set RayIntersectsEdge = Nothing
        Exit Function
    End If
    
    ' t parametresini hesapla (kesişim ışının neresinde olacak)
    t = (w.x * crossEdgeRayDir.x + w.y * crossEdgeRayDir.y + w.z * crossEdgeRayDir.z) / denom
    
    ' Eğer t < 0 ise, ışın ters yönde gidiyor ve kesişim yok
    If t < 0 Then
        Set RayIntersectsEdge = Nothing
        Exit Function
    End If
    
    ' Şimdi u'yu hesaplayacağız, ancak edge.x sıfırsa bu durumda X ekseni üzerinden yapamayız,
    ' bu yüzden diğer eksenleri kullanarak u'yu hesaplayacağız.
    
    If edge.x <> 0 Then
        u = (w.x + t * rayDir.x) / edge.x
    ElseIf edge.y <> 0 Then
        u = (w.y + t * rayDir.y) / edge.y
    ElseIf edge.z <> 0 Then
        u = (w.z + t * rayDir.z) / edge.z
    Else
        ' Eğer tüm kenar eksenleri sıfırsa (kenar noktaları aynı), bu bir hata durumudur
        Set RayIntersectsEdge = Nothing
        Exit Function
    End If

    ' Eğer u geçerli bir aralıkta ise (0 <= u <= 1), kesişim vardır ve kesişim noktasını döndür
    If u >= 0 And u <= 1 Then
        ' Kesişim noktasını hesapla
        Set intersectionPoint = New Nokta
        intersectionPoint.x = point.x + t * rayDir.x
        intersectionPoint.y = point.y + t * rayDir.y
        intersectionPoint.z = point.z + t * rayDir.z
        Set RayIntersectsEdge = intersectionPoint
    Else
        Set RayIntersectsEdge = Nothing
    End If

End Function

Public Function ClosestDistanceBetweenLines(rayStart As Nokta, rayDir As Nokta, edgeStart As Nokta, edgeDir As Nokta) As Boolean 'kesişim olup olmadığına karar vermek için kullanılır
    Dim w0 As Nokta
    Dim a As Double, b As Double, c As Double, d As Double, e As Double, sc As Double, tc As Double
    Dim distSquared As Double
    Dim denom As Double
    
    
    ' İki doğru arasındaki vektör
    Set w0 = New Nokta
    w0.x = rayStart.x - edgeStart.x
    w0.y = rayStart.y - edgeStart.y
    w0.z = rayStart.z - edgeStart.z
    
    a = rayDir.x * rayDir.x + rayDir.y * rayDir.y + rayDir.z * rayDir.z ' rayDir ile kendisinin skaler çarpımı
    b = rayDir.x * edgeDir.x + rayDir.y * edgeDir.y + rayDir.z * edgeDir.z ' rayDir ile edgeDir'in skaler çarpımı
    c = edgeDir.x * edgeDir.x + edgeDir.y * edgeDir.y + edgeDir.z * edgeDir.z ' edgeDir ile kendisinin skaler çarpımı
    d = rayDir.x * w0.x + rayDir.y * w0.y + rayDir.z * w0.z ' rayDir ile w0'ın skaler çarpımı
    e = edgeDir.x * w0.x + edgeDir.y * w0.y + edgeDir.z * w0.z ' edgeDir ile w0'ın skaler çarpımı

    denom = a * c - b * b ' Doğruların paralel olup olmadığını kontrol etmek için payda

    ' Eğer doğrular paralelse
    If Abs(denom) < tolerance Then
        ' Doğrular paralelse, en yakın mesafeyi kontrol et
        distSquared = (w0.x ^ 2 + w0.y ^ 2 + w0.z ^ 2) - ((d - b * e / c) ^ 2) / a
    Else
        ' Doğrular paralel değilse, en kısa mesafe için katsayıları hesaplayalım
        sc = (b * e - c * d) / denom
        tc = (a * e - b * d) / denom

        ' En kısa mesafe karesi
        distSquared = (w0.x + sc * rayDir.x - tc * edgeDir.x) ^ 2 + _
                      (w0.y + sc * rayDir.y - tc * edgeDir.y) ^ 2 + _
                      (w0.z + sc * rayDir.z - tc * edgeDir.z) ^ 2
    End If

    ' Eğer mesafe belirli bir toleransın altındaysa, doğrular kesişiyor demektir
    If Sqr(distSquared) = 0 Then
        ClosestDistanceBetweenLines = True
    Else
        ClosestDistanceBetweenLines = False
    End If
End Function

Public Function IsPointOnEdge(point As Nokta, p1 As Nokta, p2 As Nokta) As Boolean

    Dim edge As New Nokta
    Dim w As New Nokta
    Dim crossProduct As New Nokta
    
    ' Kenarın yön vektörü
    edge.x = p2.x - p1.x
    edge.y = p2.y - p1.y
    edge.z = p2.z - p1.z
    
    ' Başlangıç noktası ile kenarın başlangıç noktası arasındaki vektör
    w.x = point.x - p1.x
    w.y = point.y - p1.y
    w.z = point.z - p1.z
    
    ' w ve edge vektörlerinin çapraz çarpımı
    crossProduct.x = w.y * edge.z - w.z * edge.y
    crossProduct.y = w.z * edge.x - w.x * edge.z
    crossProduct.z = w.x * edge.y - w.y * edge.x

    ' Çapraz çarpımın tüm bileşenleri sıfırsa, vektörler doğrultudadır (çapraz çarpım sıfır vektördür)
    If crossProduct.x = 0 And crossProduct.y = 0 And crossProduct.z = 0 Then
        ' Şimdi, noktanın p1 ve p2 arasında olup olmadığını kontrol etmeliyiz
        If point.x >= Min(p1.x, p2.x) And point.x <= Max(p1.x, p2.x) And _
            point.y >= Min(p1.y, p2.y) And point.y <= Max(p1.y, p2.y) And _
            point.z >= Min(p1.z, p2.z) And point.z <= Max(p1.z, p2.z) Then
            IsPointOnEdge = True
            Exit Function
        End If
    End If
    
    ' Eğer nokta kenar üzerinde değilse
    IsPointOnEdge = False
End Function

Public Function ArePointsEqual(p1 As Nokta, p2 As Nokta) As Boolean
    ' X, Y ve Z koordinatlarını kontrol ederek iki noktanın eşit olup olmadığını döndürür.
    
    ' Koordinatların eşit olup olmadığını kontrol ediyoruz
    If p1.x = p2.x And p1.y = p2.y And p1.z = p2.z Then
        ArePointsEqual = True
    Else
        ArePointsEqual = False
    End If
End Function

' Düzleme paralel bir radyal vektör döndüren fonksiyon
Function GetPerpendicularVector(normal As Nokta) As Nokta
    Dim perpendicular As New Nokta
    
    ' Normal vektöre dik olacak bir yön seçelim.
    If normal.x <> 0 Then
        ' X eksenine göre dik bir vektör üret
        perpendicular.x = 0
        perpendicular.y = 1
        perpendicular.z = (-normal.y) / normal.x
    ElseIf normal.y <> 0 Then
        ' Y eksenine göre dik bir vektör üret
        perpendicular.x = 1
        perpendicular.y = 0
        perpendicular.z = (-normal.x) / normal.y
    Else
        ' Z eksenine paralel bir normal vektör için (X=0, Y=0, Z?0)
        perpendicular.x = 1
        perpendicular.y = 0
        perpendicular.z = 0
    End If
    
    Set GetPerpendicularVector = perpendicular
End Function

Public Function FaceNormal(face As Collection) As Nokta
'Normal vektörü
    Dim u As Nokta, v As Nokta
    Dim normal As Nokta
    
    ' Yüzeyi oluşturan ilk üç noktayı kullanarak iki vektör oluştur
    Set u = New Nokta
    Set v = New Nokta
    Set normal = New Nokta
    
    u.x = face(2).x - face(1).x
    u.y = face(2).y - face(1).y
    u.z = face(2).z - face(1).z
    
    v.x = face(3).x - face(1).x
    v.y = face(3).y - face(1).y
    v.z = face(3).z - face(1).z
    
    ' Çarpraz çarpım ile normal vektörü hesapla
    normal.x = u.y * v.z - u.z * v.y
    normal.y = u.z * v.x - u.x * v.z
    normal.z = u.x * v.y - u.y * v.x

    Set FaceNormal = normal
End Function

' --- Selection Exception Functions ---

Public Function ExceptionCheck(piece As Collection, obstacleShape As Collection) As Boolean
    ' piece: Kontrol edilen parça (Collection of Points)
    ' obstacleShape: Yasaklı bölge/Diğer şekil (Collection of Points/Faces)
    
    Dim k As Integer
    
    ' 1. NOKTA KONTROLÜ: Parçanın herhangi bir noktası yasaklı şeklin içinde mi?
    For k = 1 To piece.Count
        ' InsideCheck fonksiyonu zaten şeklin tamamını (obstacleShape) ve noktayı alıp bakıyor
        ' Type Kontrolü: obstacleShape (Collection) - piece(k) (Nokta) -> UYUMLU ✅
        If InsideCheck(obstacleShape, piece(k)) Then
            ExceptionCheck = True ' Hata! Nokta yasaklı bölgede.
            Exit Function
        End If
    Next k

    ' Not: Kenar kesişim kontrolü (Edge Intersection) ilerde buraya eklenebilir.
    ' Şu anki haliyle "Point-in-Polygon" mantığıyla çalışır.
    
    ExceptionCheck = False
End Function

'-------- Burayı iptal ettim ----------

' Public Function ExceptionCheck(piece As Collection, shape As Collection) As Boolean

'     Dim pt1 As Nokta
'     Dim pt2 As Nokta
'     Dim i As Integer
'     Dim j As Integer
'     Dim h As Integer

'     For i = 1 To shape.Count
'         For j = 1 To shape(i).Count

'             Set pt1 = shape(i)(j)
 
'             If j = shape(i).Count Then
'                 Set pt2 = shape(i)(1)
'             Else
'                 Set pt2 = shape(i)(j + 1)
'             End If

'             For h = 1 To piece.Count - 1
'                 Dim kesisim As Nokta
'                 Set kesisim = New Nokta
                
                
'                 Set kesisim = KesisimBul(pt1, pt2, piece(h))
                
'                 ' Eğer kesişim varsa ve kesişim noktası yüzeyde ise
'                 If Not kesisim Is Nothing Then
'                     If PointInPolygon3D(kesisim, shape(i)) Then
'                         ExceptionCheck = True ' Kesişim bulundu
'                         Exit Function
'                     End If
'                 End If
'             Next h
'         Next j
'     Next i
    
'     ' Eğer hiçbir kesişim bulunmazsa False döndür
'     ExceptionCheck = False
' End Function

    '-------- İptal Bitiş ----------

' --- Helper Math Functions (from XYZ) ---
' Bu fonksiyonlar intersectPoint ve diğerlerinde kullanıldığı için burada olmalı

' Min fonksiyonu: İki sayıdan küçük olanı döndürür
Public Function Min(a As Double, b As Double) As Double
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

' Max fonksiyonu: İki sayıdan büyük olanı döndürür
Public Function Max(a As Double, b As Double) As Double
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function


    ' --- Collision Manager Function (Main içindeki çağrı yaptığım yer) ---
Public Function CheckCollisionForPiece(piece As Collection, shapes As Collection, currentShapeIndex As Integer) As Boolean
    Dim shk As Integer
    
    ' Tüm şekilleri gez
    For shk = 1 To shapes.Count
        ' Kendi kendine çarpmasını engelle (currentShapeIndex = s)
        If shk <> currentShapeIndex Then
            
            ' shapes(shk) -> Diğer şeklin tamamı (Collection) gönderiliyor
            ' ExceptionCheck artık (Collection, Collection) kabul ediyor
            If ExceptionCheck(piece, shapes(shk)) Then
                CheckCollisionForPiece = True ' Çakışma VAR
                Exit Function
            End If
            
        End If
    Next shk
    
    ' Hiçbir çakışma yoksa
    CheckCollisionForPiece = False
End Function