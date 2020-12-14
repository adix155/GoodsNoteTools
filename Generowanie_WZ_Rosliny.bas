Attribute VB_Name = "Module1"
Sub generowanie_wz()

Worksheets("èrÛd≥o").Activate

Dim jp As Integer 'kolumna z sumπ palet
Dim jd As Integer 'kolumna z datπ
Dim jn As Integer 'kolumna z nr WZ
Dim tresc_komorki As String

jp = 1
jd = 1
jn = 1

Do While Cells(1, jp) <> ""
tresc_komorki = Cells(1, jp)
    If InStr(1, tresc_komorki, "SUMA") > 0 Then
    Exit Do
    Else
    jp = jp + 1
    End If
Loop
MsgBox (jp)

Do While Cells(1, jd) <> ""
tresc_komorki = Cells(1, jd)
    If InStr(1, tresc_komorki, "Data") > 0 Then
    Exit Do
    Else
    jd = jd + 1
    End If
Loop
MsgBox (jd)

Do While Cells(1, jn) <> ""
tresc_komorki = Cells(1, jn)
    If InStr(1, tresc_komorki, "WZ") > 0 Then
    Exit Do
    Else
    jn = jn + 1
    End If
Loop
MsgBox (jn)

Dim rok As String
rok = Right(Cells(2, jd), 4)
MsgBox (rok)

Worksheets("WZ").Activate

Dim i As Integer
Dim ii As Integer
Dim iii As Integer
Dim j As Integer

i = 2
ii = 3

Worksheets("Pomoc").Cells(12, 5) = "Auchan Magazyn Wola Bykowska"
Worksheets("Pomoc").Cells(16, 2) = 2812
Worksheets("Pomoc").Cells(16, 3) = "samochÛd"

Worksheets("WZ").Activate

Do While Worksheets("èrÛd≥o").Cells(i, 1) <> ""

    Worksheets("Pomoc").Range("A11:P28").Copy Cells(ii, 1)
    Cells(ii + 5, 1) = Worksheets("èrÛd≥o").Cells(i, 1)
    Cells(ii + 4, 5) = Worksheets("èrÛd≥o").Cells(i, 2)
    Cells(ii + 5, 10) = Worksheets("èrÛd≥o").Cells(i, jd)
    Cells(ii + 3, 15) = Worksheets("èrÛd≥o").Cells(i, jd)
    Cells(ii + 1, 12) = CStr(Worksheets("èrÛd≥o").Cells(i, jn)) & "/" & rok
    Cells(ii + 3, 12) = Cells(ii + 1, 12)
    Cells(ii + 5, 12) = Worksheets("èrÛd≥o").Cells(i, jp)
    
    j = 3
    iii = ii
    
    Do While Worksheets("èrÛd≥o").Cells(1, j) = "Opis"
        If Worksheets("èrÛd≥o").Cells(i, j) = "" Then
        j = j + 2
        Else
        Cells(iii + 8, 2) = Worksheets("èrÛd≥o").Cells(i, j)
            If Cells(iii + 8, 2) = "BOR”WKA C1,5" Then
            Cells(iii + 8, 2) = "BOR”WKA WYSOKA C 1,5L MA£A"
            Cells(iii + 8, 1) = "2000001854396"
            ElseIf Cells(iii + 8, 2) = "JEØYNA" Then
            Cells(iii + 8, 2) = "JEØYNA BEZKOLCOWA"
            Cells(iii + 8, 1) = "2000001854402"
            ElseIf Cells(iii + 8, 2) = "BOR”WKA C5" Then
            Cells(iii + 8, 2) = "BOR”WKA WYSOKA C 5L DUØA"
            Cells(iii + 8, 1) = "5908237600060"
            ElseIf Cells(iii + 8, 2) = "WINOROåL" Then
            Cells(iii + 8, 1) = "2000001854440"
            ElseIf Cells(iii + 8, 2) = "JAGODA KAMCZACKA C1,5L" Then
            Cells(iii + 8, 2) = "JAGODA KAMCZACKA C 1,5L MA£A"
            Cells(iii + 8, 1) = "2000001854426"
            ElseIf Cells(iii + 8, 2) = "JAGODA KAMCZACKA C7L" Then
            Cells(iii + 8, 2) = "JAGODA KAMCZACKA C 7L DUØA"
            Cells(iii + 8, 1) = "5908237600077"
            End If
        Cells(iii + 8, 6) = Worksheets("èrÛd≥o").Cells(i, j + 1)
        Cells(iii + 8, 8) = "szt."
        Cells(iii + 8, 10) = Worksheets("èrÛd≥o").Cells(i, j + 1)
        j = j + 2
        iii = iii + 1
        End If
        
    
    Loop
    
    Range(Cells(ii, 1), Cells(ii + 18, 16)).Copy Cells(ii + 22, 1)
    
    i = i + 1
    ii = ii + 44
Loop


End Sub

Sub Czyszczenie_WZ()

Worksheets("WZ").Activate

Range(Cells(1, 1), Cells(3000, 20)).Clear

End Sub

Sub Dopisz_EAN_Kamczacka_Ma≥a()

Dim i As Integer

Dim a As Integer
a = InputBox("Podaj ostatni wiersz")

For i = 1 To a Step 1
    If Cells(i, 2) = "JAGODA KAMCZACKA" Then
    Cells(i, 1) = "2000001854426"
    End If
Next i

End Sub
Sub Dopisz_EAN_Kamczacka_Duza_Zeruj_Ilosc()

Dim i As Integer

Dim a As Integer
a = InputBox("Podaj ostatni wiersz")

For i = 1 To a Step 1
    If Cells(i, 2) = "JAGODA KAMCZACKA" Then
    Cells(i, 1) = "5908237600077"
    Cells(i, 10) = "0"
    End If
Next i

End Sub

