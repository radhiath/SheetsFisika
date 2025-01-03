Attribute VB_Name = "Module1"
Public Function K(RataRata As Double, Delta As Double, Optional Mode As String = "Default") As String
    Dim Desimal As Double: Desimal = Delta / RataRata
    Dim Persen As String
    Dim criteriaValues() As Double
    Dim criteriaLabels() As String
    
    ' Tentukan set nilai dan label berdasarkan mode
    Select Case Mode
        Case "CustomSet"
            ' Set nilai dan label sesuai kebutuhan
            criteriaValues = Array(0.1, 1, 10, 100)
            criteriaLabels = Array(" (4 AP)", " (3 AP)", " (2 AP)", " (1 AP / ERROR)")
        Case Else
            ' Gunakan set nilai dan label default
            criteriaValues = Array(0.001, 0.01, 0.1, 1)
            criteriaLabels = Array(" (4 AP)", " (3 AP)", " (2 AP)", " (1 AP / ERROR)")
    End Select
    
    ' Hitung persentase
    Persen = Format(Desimal, "0.00 %")
    Persen = Replace(Persen, ".", ",") ' Ganti titik jd koma
    
    ' Evaluasi kriteria
    Dim i As Integer
    For i = LBound(criteriaValues) To UBound(criteriaValues)
        If Desimal <= criteriaValues(i) Then
            K = Persen & criteriaLabels(i)
            Exit Function
        End If
    Next i
    
    ' Jika Desimal melebihi semua set nilai
    K = "LAH GEDE AMAT, KOCAK NI ORANG, CEK LAGI"
End Function

'========== PENGOLAHAN DATA ==========
' 1. Menghitung Ketidakpastian dengan banyak data n
Public Function Delta_n(SigmaX As Double, SigmaKuadratX As Double, Optional n As Integer = 5) As Double

    Delta_n = (1 / n) * Sqr((n * SigmaKuadratX - SigmaX ^ 2) / (n - 1))

End Function

'4. Menghitung Ketidakpastian dengan menggunakan range data
Public Function DeltaAdv(RangeData As Range) As Double
    Dim n As Integer
    n = RangeData.Count 'dapet n
 
    Dim SumSigma As Double
    Dim DataKuadrat As Double
    Dim SumSigmaKuadrat As Double
    SumSigma = 0
    SumSigmaKuadrat = 0
    
    For i = 1 To n
        SumSigma = SumSigma + RangeData(i) 'dapet sigma n
        DataKuadrat = RangeData(i) ^ 2
        SumSigmaKuadrat = SumSigmaKuadrat + DataKuadrat 'dapet sigma n^2
    Next i
    
    DeltaAdv = (1 / n) * Sqr((n * SumSigmaKuadrat - SumSigma ^ 2) / (n - 1))
    
End Function

' 4. Menghitung Ketidakpastian Relatif (KSR) dan Angka Penting
Public Function KSR(RataRata As Double, Delta As Double, Optional AngkaBlkgKoma As Integer = 2) As String
    Dim Desimal As Double
    Dim Persen As String
    
    Desimal = Delta / RataRata
    Persen = Format(Desimal, "0.00 %")
    Persen = FormatPercent(Desimal, AngkaBlkgKoma)
    Persen = Replace(Persen, ".", ",") ' Replace period with comma
    
    If Desimal <= 0.001 Then
        KSR = Persen & " (4 AP)"
    ElseIf Desimal > 0.001 And Desimal <= 0.01 Then
        KSR = Persen & " (3 AP)"
    ElseIf Desimal > 0.01 And Desimal <= 0.1 Then
        KSR = Persen & " (2 AP)"
    ElseIf Desimal > 0.1 And Desimal <= 1 Then
        KSR = Persen & " (1 AP / ERROR)"
    Else
        KSR = "LAH GEDE AMAT, KOCAK NI ORANG, CEK LAGI"
    End If
End Function
' 4. Menghitung Ketidakpastian Relatif (KSR) dan Angka Penting
Public Function KSR_2(RataRata As Double, Delta As Double) As String
    Dim Desimal As Double
    Dim Persen As String
    
    Desimal = Delta / RataRata
    Persen = Format(Desimal, "0.0000 %")
    Persen = Replace(Persen, ".", ",") ' Replace period with comma
    
    If Desimal < 0.001 Then
        KSR_2 = Persen & " (4 AP)"
    ElseIf Desimal > 0.001 And Desimal < 0.01 Then
        KSR_2 = Persen & " (3 AP)"
    ElseIf Desimal > 0.01 And Desimal < 0.1 Then
        KSR_2 = Persen & " (2 AP)"
    ElseIf Desimal > 0.1 And Desimal < 1 Then
        KSR_2 = Persen & " (1 AP / ERROR)"
    Else
        KSR_2 = "LAH GEDE AMAT, KOCAK NI ORANG, CEK LAGI"
    End If
End Function

'5. Menghitung HASIL
Public Function Hasil(RataRata As Double, Delta As Double, AP As String, Optional Penulisan As String = "std") As String
    Dim FormatRata As String
    Dim FormatDelta As String
    Dim Desimal As Integer
    
    AP = Trim(Split(AP, "(")(1))
    AP = Replace(AP, "AP", "")
    Desimal = Val(AP) - 1
    
    If Desimal = 0 Then
    FormatRata = Format(RataRata, "0" & "E+0")
    FormatDelta = Format(Delta, "0" & "E+0")
    
    Else
    FormatRata = Format(RataRata, "0." & String(Desimal, "0") & "E+0")
    FormatDelta = Format(Delta, "0." & String(Desimal, "0") & "E+0")
    
    End If
    
    FormatRata = Replace(FormatRata, ".", ",")
    FormatRata = Replace(FormatRata, "+", "")
    
    FormatDelta = Replace(FormatDelta, ".", ",")
    FormatDelta = Replace(FormatDelta, "+", "")
    
    'Penulisan = LCase(Penulisan)
    
    If LCase(Penulisan) = "std" Then
        FormatRata = Replace(FormatRata, "E", " × 10^")
        FormatDelta = Replace(FormatDelta, "E", " × 10^")
        Hasil = "(" & FormatRata & " ± " & FormatDelta & ")"
    ElseIf LCase(Penulisan) = "dev" Then
        FormatRata = Replace(FormatRata, "E", "\bullet10^")
        FormatDelta = Replace(FormatDelta, "E", "\bullet10^")
        Hasil = "(" & FormatRata & "±" & FormatDelta & ")"
    End If
    
    If InStr(Hasil, " × 10^0") Or InStr(Hasil, "\bullet10^0") Then
        Hasil = Replace(Hasil, " × 10^0", "")
        Hasil = Replace(Hasil, "\bullet10^0", "")
    End If

End Function




'========== GRAFIK ==========
'6. Menghitung a (bisa juga b) (GRAFIK)
Public Function aGrafik(n As Integer, x As Double, y As Double, X_KUADRAT As Double, XY As Double) As Double
    Dim Atas As Double
    Dim Bawah As Double
    
    Atas = (y * X_KUADRAT) - (x * XY)
    Bawah = (n * X_KUADRAT) - (x ^ 2)
    aGrafik = Atas / Bawah

End Function

'7. Menghitung b (bisa juga m atau gradien) (GRAFIK)
Public Function bGrafik(n As Integer, x As Double, y As Double, X_KUADRAT As Double, XY As Double) As Double
    Dim Atas As Double
    Dim Bawah As Double
    
    Atas = (n * XY) - (x * y)
    Bawah = (n * X_KUADRAT) - (x ^ 2)
    bGrafik = Atas / Bawah

End Function

'8. Menghitung y (GRAFIK)
Public Function yGrafik(a As Double, b As Double, x As Double) As Double

    yGrafik = a + (b * x)

End Function



'========== REGRESI LINEAR ========== (to be added soon~)
'9. Menghitung SSxx (REGRESI LINEAR)
Public Function SSxx(ByRef RangeDataX As Range, ByVal RataX As Double) As Variant()
    Dim n As Long, i As Long
    n = RangeDataX.Cells.Count

    Dim values() As Variant
    values = RangeDataX.Value

    Dim SSxxArr() As Double
    ReDim SSxxArr(1 To n)

    For i = 1 To n
        SSxxArr(i) = (values(i, 1) - RataX) ^ 2
    Next i

    SSxx = Application.Transpose(SSxxArr)
End Function

'10. Menghitung SSyy (REGRESI LINEAR)
Public Function SSyy(ByRef RangeDataY As Range, ByVal RataY As Double) As Variant()
    Dim n As Long, i As Long
    n = RangeDataY.Cells.Count

    Dim values() As Variant
    values = RangeDataY.Value

    Dim SSyyArr() As Double
    ReDim SSyyArr(1 To n)

    For i = 1 To n
        SSyyArr(i) = (values(i, 1) - RataY) ^ 2
    Next i

    SSyy = Application.Transpose(SSyyArr)
End Function

'11. Menghitung SSxy (REGRESI LINEAR)
Public Function SSxy(ByRef RangeDataX As Range, ByVal RataX As Double, ByRef RangeDataY As Range, ByVal RataY As Double) As Variant()
    Dim n As Long, i As Long
    n = RangeDataX.Cells.Count

    Dim valuesX() As Variant
    Dim valuesY() As Variant
    
    valuesX = RangeDataX.Value
    valuesY = RangeDataY.Value

    Dim SSxyArr() As Double
    ReDim SSxyArr(1 To n)

    For i = 1 To n
        SSxyArr(i) = (valuesX(i, 1) - RataX) * (valuesY(i, 1) - RataY)
    Next i

    SSxy = Application.Transpose(SSxyArr)
End Function

'12. Menghitung SSe (REGRESI LINEAR)
Public Function SSe(ByRef RangeDataY As Range, ByRef yFitArr As Variant) As Variant()
    Dim n As Long, i As Long
    n = RangeDataY.Cells.Count

    Dim valuesY() As Variant
    valuesY = RangeDataY.Value

    Dim SSeArr() As Double
    ReDim SSeArr(1 To n)

    For i = 1 To n
        SSeArr(i) = (valuesY(i, 1) - yFitArr(i, 1)) ^ 2
    Next i

    SSe = Application.Transpose(SSeArr)
End Function


'13. Menghitung mRegLin (REGRESI LINEAR)
Public Function MRegLin(SSxy As Double, SSxx As Double) As Double

    MRegLin = SSxy / SSxx

End Function

'14. Menghitung bRegLin (REGRESI LINEAR)
Public Function BRegLin(RataX As Double, RataY As Double, MRegLin As Double) As Double

    BRegLin = RataY - MRegLin * RataX

End Function

'15. Menghitung yRegLin (REGRESI LINEAR)
Public Function yRegLin(MRegLin As String, BRegLin As String, Optional Penulisan As String = "std") As String
    Dim FormatMRegLin As String
    Dim FormatBRegLin As String
    
    FormatMRegLin = Format(MRegLin, "0." & String(3, "0") & "E+0")
    FormatMRegLin = Replace(FormatMRegLin, ".", ",")
    FormatMRegLin = Replace(FormatMRegLin, "+", "")
    
    FormatBRegLin = Format(BRegLin, "0." & String(3, "0") & "E+0")
    FormatBRegLin = Replace(FormatBRegLin, ".", ",")
    If InStr(FormatBRegLin, "+") Then
        FormatBRegLin = Replace(FormatBRegLin, "+", "")

    End If
    
    If LCase(Penulisan) = "std" Then
        FormatMRegLin = Replace(FormatMRegLin, "E", " × 10^")
        FormatBRegLin = Replace(FormatBRegLin, "E", " × 10^")
        
       
    ElseIf LCase(Penulisan) = "dev" Then
        FormatMRegLin = Replace(FormatMRegLin, "E", "\bullet10^")
        FormatBRegLin = Replace(FormatBRegLin, "E", "\bullet10^")

    End If
    
    If CDbl(BRegLin) > 0 Then
        yRegLin = "(" & FormatMRegLin & ")" & "x" & " + " & "(" & FormatBRegLin & ")"
    Else
        yRegLin = "(" & FormatMRegLin & ")" & "x" & " - " & "(" & LStripByChar(FormatBRegLin, "-") & ")"
    End If

    If InStr(yRegLin, " × 10^0") Or InStr(yRegLin, "\bullet10^0") Then
        yRegLin = Replace(yRegLin, " × 10^0", "")
        yRegLin = Replace(yRegLin, "\bullet10^0", "")
    End If

End Function
Function LStripByChar(inputString As String, charToStrip As String) As String
    Dim i As Integer
    i = 1
    While Mid(inputString, i, 1) = charToStrip And i <= Len(inputString)
        i = i + 1
    Wend
    LStripByChar = Mid(inputString, i)
End Function

'16. Menghitung y-Fit (REGRESI LINEAR)
Public Function YFit(ByRef RangeDataX As Range, MRegLin As Double, BRegLin As Double) As Variant()
    Dim n As Long, i As Long
    n = RangeDataX.Cells.Count

    Dim values() As Variant
    values = RangeDataX.Value

    Dim yFitArr() As Double
    ReDim yFitArr(1 To n)

    For i = 1 To n
        yFitArr(i) = (MRegLin * values(i, 1)) + BRegLin
    Next i

    YFit = Application.Transpose(yFitArr)
End Function

'17. Menghitung MSE (REGRESI LINEAR)
Public Function MSE(SSe As Double, n As Double) As Double

    MSE = SSe / (n - 2)

End Function

'19. S^2m
Public Function S2m(MSE As Double, SSxx As Double) As Double

    S2m = MSE / SSxx

End Function

'18. Menghitung Delta_mRegLin (REGRESI LINEAR) to be added soon
Public Function Delta_mRegLin(MSE As Double, SSxx As Double) As Double

    Delta_mRegLin = Sqr(MSE / SSxx)

End Function

'19. Menghitung Delta_bRegLin (REGRESI LINEAR) to be added soon
Public Function Delta_bRegLin(MSE As Double, SigmaXKuadrat As Double, n As Double, SSxx As Double) As Double

    Delta_bRegLin = Sqr(MSE * (SigmaXKuadrat / (n * SSxx)))
    
End Function

'20. Menghitung hasil untuk Regresi Linear
Public Function HasilRegLin(x As Double, DeltaX As Double, Optional Penulisan As String = "std") As String
    Dim Formatx As String
    Dim FormatDeltaX As String
    
    Formatx = Format(x, "0." & String(3, "0") & "E+0")
    Formatx = Replace(Formatx, ".", ",")
    
    Formatx = Replace(Formatx, "+", "")
    
    FormatDeltaX = Format(DeltaX, "0." & String(3, "0") & "E+0")
    FormatDeltaX = Replace(FormatDeltaX, ".", ",")
    FormatDeltaX = Replace(FormatDeltaX, "+", "")
    
    If LCase(Penulisan) = "std" Then
        Formatx = Replace(Formatx, "E", " × 10^")
        FormatDeltaX = Replace(FormatDeltaX, "E", " × 10^")
        HasilRegLin = "(" & Formatx & " ± " & FormatDeltaX & ")"
       
    ElseIf LCase(Penulisan) = "dev" Then
        Formatx = Replace(Formatx, "E", "\bullet10^")
        FormatDeltaX = Replace(FormatDeltaX, "E", "\bullet10^")
        HasilRegLin = "(" & Formatx & "±" & FormatDeltaX & ")"
    End If
    
    If InStr(HasilRegLin, " × 10^0") Or InStr(HasilRegLin, "\bullet10^0") Then
        HasilRegLin = Replace(HasilRegLin, " × 10^0", "")
        HasilRegLin = Replace(HasilRegLin, "\bullet10^0", "")
    End If
End Function

Function DATA_MAJEMUK(RangeData As Range, Optional Penulisan As String = "std") As Variant
    Dim n As Integer
    n = RangeData.Count
    
    Dim HasilSum As Double
    Dim HasilAve As Double
    Dim HasilDelta As Double
    Dim HasilKSRAwal As String
    Dim HasilKSR As String
    Dim HasilAkhir As String
 
    '1. Sum
    HasilSum = Application.Sum(RangeData)
    
    '2. Average
    HasilAve = Application.WorksheetFunction.Average(RangeData)
    
    '3. Delta X
    Dim DataKuadrat As Double
    Dim SumSigmaKuadrat As Double
    
    For i = 1 To n
        DataKuadrat = RangeData(i) ^ 2
        SumSigmaKuadrat = SumSigmaKuadrat + DataKuadrat
    Next i
    
    HasilDelta = Delta_n(HasilSum, SumSigmaKuadrat, n)
    
    '4. KSR dan Angka Penting
    HasilKSRAwal = KSR(HasilAve, HasilDelta)
    HasilKSR = HasilKSRAwal
    
    '5. Penulisan hasil akhir
    If LCase(Penulisan) = "std" Then
        HasilAkhir = Hasil(HasilAve, HasilDelta, HasilKSR, "std")
        
    ElseIf LCase(Penulisan) = "dev" Then
        HasilAkhir = Hasil(HasilAve, HasilDelta, HasilKSR, "dev")
    
    End If
    
    Dim Hasilnya(1 To 5) As Variant
    
    Hasilnya(1) = HasilSum
    Hasilnya(2) = HasilAve
    Hasilnya(3) = HasilDelta
    Hasilnya(4) = HasilKSRAwal
    Hasilnya(5) = HasilAkhir
    
    DATA_MAJEMUK = Application.Transpose(Hasilnya)

End Function

Public Function REGRESI_LINEAR(ByRef RangeDataX As Range, ByRef RangeDataY As Range) As Variant()
    Dim n As Long
    Dim RataX As Double
    Dim RataY As Double
    Dim XKuadrat() As Double
    Dim YKuadrat() As Double
    Dim SSxxArr() As Variant
    Dim SSyyArr() As Variant
    Dim SSxyArr() As Variant
    Dim HasilmRegLin As Double
    Dim HasilbRegLin As Double
    Dim yFitArr() As Variant
    Dim SSeArr() As Variant
    Dim Hasil() As Variant
    Dim i As Long

    n = RangeDataX.Cells.Count

    RataX = Application.WorksheetFunction.Average(RangeDataX)
    RataY = Application.WorksheetFunction.Average(RangeDataY)
    
    ReDim XKuadrat(1 To n)
    ReDim YKuadrat(1 To n)
    For i = 1 To n
        XKuadrat(i) = RangeDataX.Cells(i, 1).Value ^ 2
        YKuadrat(i) = RangeDataY.Cells(i, 1).Value ^ 2
    Next i

    SSxxArr = SSxx(RangeDataX, RataX)
    SSyyArr = SSyy(RangeDataY, RataY)
    SSxyArr = SSxy(RangeDataX, RataX, RangeDataY, RataY)

    HasilmRegLin = MRegLin(Application.Sum(SSxyArr), Application.Sum(SSxxArr))
    HasilbRegLin = BRegLin(RataX, RataY, HasilmRegLin)

    yFitArr = YFit(RangeDataX, HasilmRegLin, HasilbRegLin)
    SSeArr = SSe(RangeDataY, yFitArr)

    Dim Strm As String
    Strm = CStr(HasilmRegLin)
    Dim Strb As String
    Strb = CStr(HasilbRegLin)
    
    Dim jumlahsel As Double
    jumlahsel = n
    
    ReDim Hasil(1 To n + 1, 1 To 10)

    For i = 1 To n
        Hasil(i, 1) = XKuadrat(i)
        Hasil(i, 2) = YKuadrat(i)
        Hasil(i, 3) = SSxxArr(i, 1)
        Hasil(i, 4) = SSyyArr(i, 1)
        Hasil(i, 5) = SSxyArr(i, 1)
        Hasil(i, 6) = yFitArr(i, 1)
        Hasil(i, 7) = SSeArr(i, 1)
        
    Next i
    
    Hasil(n + 1, 1) = Application.Sum(XKuadrat)
    Hasil(n + 1, 2) = Application.Sum(YKuadrat)
    Hasil(n + 1, 3) = Application.Sum(SSxxArr)
    Hasil(n + 1, 4) = Application.Sum(SSyyArr)
    Hasil(n + 1, 5) = Application.Sum(SSxyArr)
    Hasil(n + 1, 6) = Application.Sum(yFitArr)
    Hasil(n + 1, 7) = Application.Sum(SSeArr)

    REGRESI_LINEAR = Hasil
End Function

Public Function REGRESI_LINEAR_2(ByRef RangeDataX As Range, ByRef RangeDataY As Range) As Variant()
    Dim n As Long
    Dim RataX As Double
    Dim RataY As Double
    Dim XKuadrat() As Double
    Dim YKuadrat() As Double
    Dim SSxxArr() As Variant
    Dim SSyyArr() As Variant
    Dim SSxyArr() As Variant
    Dim HasilmRegLin As Double
    Dim HasilbRegLin As Double
    Dim yFitArr() As Variant
    Dim SSeArr() As Variant
    Dim Hasil() As Variant
    Dim i As Long

    n = RangeDataX.Cells.Count

    RataX = Application.WorksheetFunction.Average(RangeDataX)
    RataY = Application.WorksheetFunction.Average(RangeDataY)
    
    ReDim XKuadrat(1 To n)
    ReDim YKuadrat(1 To n)
    For i = 1 To n
        XKuadrat(i) = RangeDataX.Cells(i, 1).Value ^ 2
        YKuadrat(i) = RangeDataY.Cells(i, 1).Value ^ 2
    Next i

    SSxxArr = SSxx(RangeDataX, RataX)
    SSyyArr = SSyy(RangeDataY, RataY)
    SSxyArr = SSxy(RangeDataX, RataX, RangeDataY, RataY)

    HasilmRegLin = MRegLin(Application.Sum(SSxyArr), Application.Sum(SSxxArr))
    HasilbRegLin = BRegLin(RataX, RataY, HasilmRegLin)

    yFitArr = YFit(RangeDataX, HasilmRegLin, HasilbRegLin)
    SSeArr = SSe(RangeDataY, yFitArr)

    Dim Strm As String
    Strm = CStr(HasilmRegLin)
    Dim Strb As String
    Strb = CStr(HasilbRegLin)
    
    Dim jumlahsel As Double
    jumlahsel = n
    
    ReDim Hasil(1 To 6, 1 To 8)

        Dim HasilMSE As Double
        HasilMSE = MSE(Application.Sum(SSeArr), jumlahsel)
        Dim HasilSm2 As Double
        HasilSm2 = S2m(HasilMSE, Application.Sum(SSxxArr))
        Dim DeltamRegLin As Double
        DeltamRegLin = Delta_mRegLin(HasilMSE, Application.Sum(SSxxArr))
        Dim DeltabRegLin As Double
        DeltabRegLin = Delta_bRegLin(HasilMSE, Application.Sum(XKuadrat), jumlahsel, Application.Sum(SSxxArr))
        
        Hasil(1, 1) = "Rata x"
        Hasil(2, 1) = "Rata y"
        Hasil(4, 1) = "m"
        Hasil(5, 1) = "c"
        Hasil(6, 1) = "y"
        
        Hasil(1, 2) = Application.WorksheetFunction.Average(RangeDataX)
        Hasil(2, 2) = Application.WorksheetFunction.Average(RangeDataY)
        Hasil(4, 2) = HasilmRegLin
        Hasil(5, 2) = HasilbRegLin
        Hasil(6, 2) = yRegLin(Strm, Strb)
        
        Hasil(1, 4) = "MSE / Syx^2"
        Hasil(2, 4) = "Sm^2"
        Hasil(4, 4) = ChrW(&H394) & "m"
        Hasil(5, 4) = ChrW(&H394) & "c"
        
        Hasil(1, 5) = HasilMSE
        Hasil(2, 5) = HasilSm2
        Hasil(4, 5) = DeltamRegLin
        Hasil(5, 5) = DeltabRegLin
        
        Hasil(4, 7) = "(m ± " & ChrW(&H394) & "m)"
        Hasil(5, 7) = "(c ± " & ChrW(&H394) & "c)"
        Hasil(4, 8) = HasilRegLin(HasilmRegLin, DeltamRegLin)
        Hasil(5, 8) = HasilRegLin(HasilbRegLin, DeltabRegLin)
        For i = LBound(Hasil, 1) To UBound(Hasil, 1)
            For j = LBound(Hasil, 2) To UBound(Hasil, 2)
                If IsEmpty(Hasil(i, j)) Then
                    Hasil(i, j) = ""
                End If
            Next j
        Next i
        REGRESI_LINEAR_2 = Hasil
        End Function
Function FormatKali(Saintifik As Double) As String
    Dim splitNumber() As String
    Dim baseNumber As String
    Dim exponent As String
    
    ' Memisahkan angka dasar dan eksponen
    splitNumber = Split(Saintifik, "E")
    
    ' Mengambil angka dasar dan eksponen
    baseNumber = splitNumber(0)
    exponent = splitNumber(1)
    
    ' Mengubah format menjadi bentuk perkalian dengan 10^n
    FormatKali = baseNumber & " x 10^" & exponent
End Function







