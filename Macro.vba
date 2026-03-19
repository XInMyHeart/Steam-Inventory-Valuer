Sub PobierzEkwipunekCS()
    Dim http As Object
    Dim url As String, response As String
    Dim json As Object, item As Object
    Dim i As Integer, steamID As String
    Dim cena As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    steamID = InputBox("Podaj swój SteamID64 (długi numer):", "Pobieranie ekwipunku", "76561198303875908")
    If steamID = "" Then Exit Sub
    
    url = "https://steamcommunity.com/inventory/" & steamID & "/730/2?l=polish&count=1000"
    
    With http
        .Open "GET", url, False
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Referer", "https://steamcommunity.com/"
        .SetRequestHeader "Cache-Control", "no-cache"
        .Send
        response = .ResponseText
    End With

    If Len(response) < 100 Then
        MsgBox "Błąd: Steam nałożył Rate Limit (za krótka odpowiedź). Zmień IP (hotspot) lub odczekaj 15 min."
        Exit Sub
    End If

    On Error Resume Next
    Set json = JsonConverter.ParseJson(response)
    If Err.Number <> 0 Then
        MsgBox "Błąd parsowania JSON."
        Exit Sub
    End If
    On Error GoTo 0

    Sheets("Arkusz1").Range("A:D").ClearContents
    Range("A1:C1").Value = Array("Nazwa przedmiotu", "Typ", "Cena Rynkowa")
    Range("A1:C1").Font.Bold = True
    
    i = 2
    
    i = 2
    Dim asset As Object
    Dim ilosc As Integer
    Dim classID As String

    For Each item In json("descriptions")
        classID = item("classid")
        ilosc = 0
        
        For Each asset In json("assets")
            If asset("classid") = classID Then
                ilosc = ilosc + 1
            End If
        Next asset
        
        Cells(i, 1).Value = item("market_hash_name")
        Cells(i, 2).Value = item("type")
        Cells(i, 4).Value = ilosc
        
        If item("marketable") = 1 Then
            cenaTekst = PobierzCeneZMarketu(CStr(item("market_hash_name")))
            Cells(i, 3).Value = WyczyscCene(CStr(cenaTekst))
            DoEvents
            Czekaj 10
        Else
            Cells(i, 3).Value = "Niesprzedawalny"
        End If
        
        i = i + 1
    Next item
    Columns("A:C").AutoFit
    MsgBox "Sukces! Wypisano dane."
End Sub

Function PobierzCeneZMarketu(nazwa As String) As String
    Dim h As Object, r As String, j As Object
    Dim u As String
    Dim proba As Integer
    Dim sukces As Boolean
    
    u = "https://steamcommunity.com/market/priceoverview/?appid=730&currency=6&market_hash_name=" & WorksheetFunction.EncodeURL(nazwa)
    sukces = False
    
    For proba = 1 To 2
        Set h = CreateObject("MSXML2.XMLHTTP")
        
        On Error Resume Next
        h.Open "GET", u, False
        h.SetRequestHeader "User-Agent", "Mozilla/5.0"
        h.Send
        r = h.ResponseText
        On Error GoTo 0
        
        If InStr(r, "lowest_price") > 0 Then
            Set j = JsonConverter.ParseJson(r)
            PobierzCeneZMarketu = j("lowest_price")
            sukces = True
            Exit For
        Else
            DoEvents
            Czekaj 2
        End If
    Next proba
    
    If Not sukces Then
        PobierzCeneZMarketu = "Limit/Błąd"
    End If
End Function

Function WyczyscCene(txt As String) As Double
    Dim wynik As String
    Dim i As Integer
    Dim znak As String
    
    If InStr(txt, "Limit") > 0 Or txt = "" Then
        WyczyscCene = 0
        Exit Function
    End If
    
    wynik = ""
    For i = 1 To Len(txt)
        znak = Mid(txt, i, 1)
        If IsNumeric(znak) Or znak = "." Or znak = "," Then
            wynik = wynik & znak
        End If
    Next i
    
    wynik = Replace(wynik, ".", ",")
    
    If wynik <> "" Then
        WyczyscCene = CDbl(wynik)
    Else
        WyczyscCene = 0
    End If
End Function

Sub Czekaj(sekundy As Double)
    Dim koniec As Double
    koniec = Timer + (sekundy * 0.1)
    Do While Timer < koniec
        DoEvents
        If Timer < 0.1 Then koniec = koniec - 86400
    Loop
End Sub

Sub OdswiezBrakujaceCeny()
    Dim ostatniWiersz As Long
    Dim i As Long
    Dim nazwaPrzedmiotu As String
    Dim cenaTekst As String
    Dim licznik As Integer
        Dim Seconds As Integer
    
    Seconds = InputBox("Podaj czas na pojedyncze sprawdzenie: (10 to 1 sec)")
    
    ostatniWiersz = Sheets("Arkusz1").Cells(Rows.Count, "A").End(xlUp).Row
    
    If ostatniWiersz < 2 Then
        MsgBox "Brak przedmiotów do sprawdzenia!", vbExclamation
        Exit Sub
    End If
    
    licznik = 0
    Application.StatusBar = "Rozpoczynam sprawdzanie brakujących cen..."
    
    For i = 2 To ostatniWiersz
        If (Cells(i, 3).Value = 0 Or Cells(i, 3).Value = "") And Cells(i, 3).Value <> "Niesprzedawalny" Then
            
            nazwaPrzedmiotu = Cells(i, 1).Value
            Application.StatusBar = "Odświeżanie: " & nazwaPrzedmiotu & "..."
            
            cenaTekst = PobierzCeneZMarketu(CStr(nazwaPrzedmiotu))
            
            If InStr(cenaTekst, "Limit") = 0 And cenaTekst <> "" Then
                Cells(i, 3).Value = WyczyscCene(CStr(cenaTekst))
                licznik = licznik + 1
            End If
            
            Czekaj (Seconds)
            DoEvents
        End If
    Next i
    
    Application.StatusBar = False
    If licznik > 0 Then
        MsgBox "Pomyślnie uzupełniono " & licznik & " cen!", vbInformation
    Else
        MsgBox "Nie znaleziono brakujących cen lub Steam nadal blokuje zapytania.", vbInformation
    End If
End Sub
