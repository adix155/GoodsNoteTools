Attribute VB_Name = "Module1"
Sub generujWZ()

Worksheets("WZ").Activate
Range("E2:I3").ClearContents
Range("L2:N2").ClearContents
Range("L4:P4").ClearContents
Range("A6:D6").ClearContents
Range("J6:P6").ClearContents
Range("A9:P13").ClearContents
Range("P14:P14").ClearContents
Range("A22:P38").clear

Worksheets("Pomoc").Activate
Range("B4:B6").ClearContents

Dim WS As Worksheet
Set WS = Sheets.Add(After:=Worksheets("Pomoc"))
WS.Name = "XML"

Worksheets("XML").Activate
ActiveWorkbook.XmlImport Url:=Excel.Application.GetOpenFilename("Pliki XML (*.xml), *.xml", , "Wska¿ plik .xml z danymi zamówienia"), ImportMap:=Nothing, _
                         Overwrite:=True, Destination:=Range("$A$1")

Dim ia As Integer
ia = 1

Do While Cells(ia, 1) <> ""
ia = ia + 1
Loop
ia = ia - 1

Dim cena_cwiartka As Double
Dim cena_polkilowka As Double
Dim cena_kg As Double

If ia = 1 Then
    Worksheets("WZ").Activate

    Dim sciezka As String
    Dim data_dostawy As String
    Dim rok As String
    Dim NR_WZ As String
    Dim numer As String
    Dim nr_zamowienia As String
    Dim nr_partii As String
    Dim jednostka_miary As String
    Dim ilosc_palet_EURO As String

    sciezka = ThisWorkbook.Path
    data_dostawy = Worksheets("XML").Cells(1, 3)
    rok = Right(data_dostawy, 4)
    NR_WZ = InputBox("Podaj bie¿¹cy numer wystawianego dokumentu WZ." & vbCrLf & "Numer poprzedniego dokumentu WZ wystawionego w niniejszym programie to " & (Worksheets("Pomoc").Cells(1, 2) - 1) & ".", , Worksheets("Pomoc").Cells(1, 2))
    numer = NR_WZ & " / " & rok
    nr_zamowienia = Worksheets("XML").Cells(1, 1)
    nr_partii = InputBox("Podaj numer partii." & vbCrLf & "Poprzedni numer partii to " & (Worksheets("Pomoc").Cells(2, 2) - 1) & ".", , Worksheets("Pomoc").Cells(2, 2))
    ilosc_palet_EURO = InputBox("Podaj iloœæ palet EURO")
    
    Cells(6, 1) = nr_zamowienia
    Cells(6, 2) = "2812"
    Cells(6, 3) = "samochód"
    Cells(2, 5) = "Auchan Magazyn Grójec"
    Cells(6, 10) = data_dostawy
    Cells(4, 15) = data_dostawy
    Cells(2, 12) = numer
    Cells(4, 12) = numer
    Cells(6, 15) = nr_partii
    Cells(6, 12) = ilosc_palet_EURO
    
    
    
    Cells(9, 1) = Worksheets("XML").Cells(1, 22) 'nr EAN
    nazwa_produktu = Worksheets("XML").Cells(1, 24)
    Cells(9, 2) = nazwa_produktu 'nazwa produktu
    Cells(9, 6) = Worksheets("XML").Cells(1, 26) 'ilosc zamówiona
    jednostka_miary = Worksheets("XML").Cells(1, 29)
        If jednostka_miary = "PCE" Then
        Cells(9, 8) = "szt."
        End If
    Cells(9, 10) = InputBox("Podaj wydan¹ iloœæ dla " & Cells(9, 2), , Cells(9, 6))
        If InStr(1, nazwa_produktu, "250", vbTextCompare) > 0 Then
        cwiartki = Cells(9, 10)
        Cells(9, 15) = "106"
        Cells(9, 16) = Cells(9, 10) / (Worksheets("XML").Cells(1, 27))
        cena_cwiartka = Worksheets("XML").Cells(1, 30)
        ElseIf InStr(1, nazwa_produktu, "500", vbTextCompare) > 0 Then
        polkilowki = Cells(9, 10)
        Cells(9, 15) = "106"
        Cells(9, 16) = Cells(9, 10) / (Worksheets("XML").Cells(1, 27))
        cena_polkilowka = Worksheets("XML").Cells(1, 30)
        ElseIf InStr(1, nazwa_produktu, "KG", vbTextCompare) > 0 Then
        kg = Cells(9, 10)
        Cells(9, 15) = "154"
        Cells(9, 8) = "kg."
        Cells(9, 16) = Cells(9, 10) / 2.5
        cena_kg = Worksheets("XML").Cells(1, 30)
        End If
    Cells(14, 16) = Cells(9, 16)
    
    
    
Else

    Worksheets("WZ").Activate

    

    sciezka = ThisWorkbook.Path
    data_dostawy = Worksheets("XML").Cells(2, 3)
    rok = Right(data_dostawy, 4)
    NR_WZ = InputBox("Podaj bie¿¹cy numer wystawianego dokumentu WZ." & vbCrLf & "Numer poprzedniego dokumentu WZ wystawionego w niniejszym programie to " & (Worksheets("Pomoc").Cells(1, 2) - 1) & ".", , Worksheets("Pomoc").Cells(1, 2))
    numer = NR_WZ & " / " & rok
    nr_zamowienia = Worksheets("XML").Cells(2, 1)
    nr_partii = InputBox("Podaj numer partii." & vbCrLf & "Poprzedni numer partii to " & (Worksheets("Pomoc").Cells(2, 2) - 1) & ".", , Worksheets("Pomoc").Cells(2, 2))
    ilosc_palet_EURO = InputBox("Podaj iloœæ palet EURO")

    Cells(6, 1) = nr_zamowienia
    Cells(6, 2) = "2812"
    Cells(6, 3) = "samochód"
    Cells(2, 5) = "Auchan Magazyn Grójec"
    Cells(6, 10) = data_dostawy
    Cells(4, 15) = data_dostawy
    Cells(2, 12) = numer
    Cells(4, 12) = numer
    Cells(6, 15) = nr_partii
    Cells(6, 12) = ilosc_palet_EURO

    Dim i As Integer
    Dim IX As Integer
    
    
    
    IX = 2
    i = 9

    Do While Worksheets("XML").Cells(IX, 1) <> ""
        Cells(i, 1) = Worksheets("XML").Cells(IX, 17) 'nr EAN
        nazwa_produktu = Worksheets("XML").Cells(IX, 24)
        Cells(i, 2) = Worksheets("XML").Cells(IX, 19) 'nazwa produktu
        Cells(i, 6) = Worksheets("XML").Cells(IX, 21) 'ilosc zamówiona
        jednostka_miary = Worksheets("XML").Cells(IX, 24)
            If jednostka_miary = "PCE" Then
            Cells(i, 8) = "szt."
            End If
        Cells(i, 10) = InputBox("Podaj wydan¹ iloœæ dla " & Cells(i, 2), , Cells(i, 6))
            If InStr(1, nazwa_produktu, "250", vbTextCompare) > 0 Then
            cwiartki = Cells(i, 10)
            Cells(i, 15) = "106"
            Cells(i, 16) = Cells(i, 10) / (Worksheets("XML").Cells(IX, 22))
            cena_cwiartka = Worksheets("XML").Cells(IX, 25)
            ElseIf InStr(1, nazwa_produktu, "500", vbTextCompare) > 0 Then
            polkilowki = Cells(i, 10)
            Cells(i, 15) = "106"
            Cells(i, 16) = Cells(i, 10) / (Worksheets("XML").Cells(IX, 22))
            cena_polkilowka = Worksheets("XML").Cells(IX, 25)
            ElseIf InStr(1, nazwa_produktu, "KG", vbTextCompare) > 0 Then
            kg = Cells(i, 10)
            Cells(i, 15) = "154"
            Cells(i, 16) = Cells(i, 10) / 2.5
            Cells(i, 8) = "kg."
            cena_kg = Worksheets("XML").Cells(IX, 25)
            End If
        IX = IX + 1
        i = i + 1
    Loop

    Cells(14, 16) = Cells(9, 16) + Cells(10, 16) + Cells(11, 16) + Cells(12, 16) + Cells(13, 16)


   
End If

Worksheets("Pomoc").Cells(1, 2) = NR_WZ + 1
Worksheets("Pomoc").Cells(2, 2) = nr_partii + 1
Worksheets("Pomoc").Cells(4, 2) = cena_cwiartka
Worksheets("Pomoc").Cells(5, 2) = cena_polkilowka
Worksheets("Pomoc").Cells(6, 2) = cena_kg

Range("A1:P17").Copy Range("A22")

Application.DisplayAlerts = False
Worksheets("XML").Delete
Application.DisplayAlerts = True

ActiveWorkbook.Save

MsgBox ("Zakoñczono wype³nianie dokumentu WZ. SprawdŸ poprawnoœæ wprowadzonych danych.")

Dim strPrompt As String
Dim strStyle As String
Dim StrResponse As String
 
strPrompt = "Czy wprowadziæ dane z dokumentu WZ do zestawienia?"
strStyle = vbQuestion + vbYesNo
 
StrResponse = MsgBox(strPrompt, strStyle)
 
If StrResponse = vbYes Then
    Application.ScreenUpdating = False

    Workbooks.Open (sciezka & "\" & "Zestawienie.xlsm")

    Dim iy As Integer
    iy = 1
    Do While Cells(iy, 1) <> ""
        iy = iy + 1
    Loop

    Cells(iy + 1, 4).clear
    Cells(iy + 1, 5).clear
    Cells(iy + 1, 6).clear
    Cells(iy + 1, 7).clear
    Cells(iy + 1, 8).clear
    Cells(iy + 1, 9).clear
    Cells(iy + 1, 10).clear
    Cells(iy + 1, 11).clear
    
    Range("C2:C200").NumberFormat = "@"

    Cells(iy, 1) = nr_partii
    Cells(iy, 2) = nr_zamowienia
    Cells(iy, 3) = numer
    Cells(iy, 4) = data_dostawy
    Cells(iy, 5) = cwiartki
        If cwiartki = "" Then
        Cells(iy, 6) = ""
        Else
        Cells(iy, 6) = cena_cwiartka
        End If
    Cells(iy, 7) = polkilowki
        If polkilowki = "" Then
        Cells(iy, 8) = ""
        Else
        Cells(iy, 8) = cena_polkilowka
        End If
    Cells(iy, 9) = kg
        If kg = "" Then
        Cells(iy, 10) = ""
        Else
        Cells(iy, 10) = cena_kg
        End If
    Cells(iy, 11) = (Cells(iy, 5) / 4) + (Cells(iy, 7) / 2) + Cells(iy, 9)
    Cells(iy, 12) = Cells(iy, 5) * Cells(iy, 6)
    Cells(iy, 13) = Cells(iy, 7) * Cells(iy, 8)
    Cells(iy, 14) = Cells(iy, 9) * Cells(iy, 10)
    Cells(iy, 15) = Cells(iy, 12) + Cells(iy, 13) + Cells(iy, 14)
    Cells(iy, 16) = Cells(iy, 15) * 1.05
    
    Cells(iy + 1, 4) = "Suma"
    Cells(iy + 1, 5) = Application.Sum(Range(Cells(2, 5), Cells(iy, 5)))
    Cells(iy + 1, 7) = Application.Sum(Range(Cells(2, 7), Cells(iy, 7)))
    Cells(iy + 1, 9) = Application.Sum(Range(Cells(2, 9), Cells(iy, 9)))
    Cells(iy + 1, 11) = Application.Sum(Range(Cells(2, 11), Cells(iy, 11)))
    Cells(iy + 1, 12) = Application.Sum(Range(Cells(2, 12), Cells(iy, 12)))
    Cells(iy + 1, 13) = Application.Sum(Range(Cells(2, 13), Cells(iy, 13)))
    Cells(iy + 1, 14) = Application.Sum(Range(Cells(2, 14), Cells(iy, 14)))
    Cells(iy + 1, 15) = Application.Sum(Range(Cells(2, 15), Cells(iy, 15)))
    Cells(iy + 1, 16) = Application.Sum(Range(Cells(2, 16), Cells(iy, 16)))
    

    Range("A1:Z200").HorizontalAlignment = xlCenter

    ActiveWorkbook.Save
    ActiveWorkbook.Close

    Workbooks("Generowanie_WZ_1.0.xlsm").Activate
    Worksheets("WZ").Activate
    Application.ScreenUpdating = True
    
    MsgBox ("Dane do zestawienia zosta³y wprowadzone i pomyœlnie zapisane")
    ActiveWorkbook.Save
End If


Worksheets("WZ").Range("a1:p38").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
sciezka & "\WZ\" & rok & "\" & NR_WZ & " " & rok & ".pdf", Quality:= _
xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
OpenAfterPublish:=False


End Sub
Sub go_to_list()

Dim sciezka As String
sciezka = ThisWorkbook.Path

Workbooks.Open (sciezka & "\" & "Zestawienie.xlsm")

End Sub


