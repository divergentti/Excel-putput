Option Compare Text
Option Explicit

' Sub joka pyörittää sanalistaa ja vertailee ilman capseja
Sub TestaaMua()
' Pyörittelee Sanat-sivulta sanapareja. 30.3.2022 Jari Hiltunen
Dim SanaAlue As Range
Dim SanaEng As Variant
Dim SanaSuo As Variant
Dim Kieli As String
Dim Vastaus As Variant
Dim KysyRivi() As Integer
Dim Kysytty As Integer
Dim Oikein As Integer
Dim OikeinRivi() As Integer
Dim Vaarin As Integer
Dim Rivit As Integer
Dim SatunnainenRivi As Integer
Dim i As Integer
Dim x As Integer
Dim z As Integer
Dim Duplikaatti As Variant

' Sanojen viimeinen rivi alkaen Sanat-sivun A3:sta
Rivit = Worksheets("Sanat").Range("A3").End(xlDown).Row
Set SanaAlue = Worksheets("Sanat").Range("A3:B" & Rivit)
ReDim KysyRivi(0 To 1)

' Randomisoidaan rivit
Do Until z = Int(SanaAlue.Rows.Count)
    ' Arvotaan rivi
    SatunnainenRivi = Int(SanaAlue.Rows.Count * Rnd + 1)
    Duplikaatti = False
    ' Testataan onko rivinumero ennestään arvottuna
    For x = LBound(KysyRivi) To UBound(KysyRivi)
        If SatunnainenRivi = KysyRivi(x) Then
            'Rivi on duplikaatti, älä lisää taulukkoa
            Duplikaatti = True
        End If
    Next x
    ' Rivi ei ole duplikaatti, lisätään taulukkoon rivi
    If Duplikaatti = False Then
        ReDim Preserve KysyRivi(z)
        KysyRivi(z) = SatunnainenRivi
        z = z + 1
    End If
Loop
    

Do Until Vastaus = "stop" Or Kysytty = (UBound(KysyRivi) - LBound(KysyRivi) + 1)

    ' Katsotaan mitä käyttäjä haluaa, SVE-FIN, FIN-SVE vai randomi
    If Worksheets("Aloitus").Range("C5").Value = "x" Then
        Kieli = "svefin"
    ElseIf Worksheets("Aloitus").Range("C6").Value = "x" Then
        Kieli = "finsve"
    ElseIf Worksheets("Aloitus").Range("F5").Value = "x" Then
        Kieli = "randomi"
    ' Valitaan randomi jos mitään ei ole merkattu valittavaksi!
    Else
        Kieli = "randomi"
    End If
        
    If Kieli = "svefin" Then
      SanaEng = SanaAlue.Range("B" & KysyRivi(Kysytty)).Value
      SanaSuo = SanaAlue.Range("A" & KysyRivi(Kysytty)).Value
    ElseIf Kieli = "finsve" Then
      SanaEng = SanaAlue.Range("A" & KysyRivi(Kysytty)).Value
      SanaSuo = SanaAlue.Range("B" & KysyRivi(Kysytty)).Value
    ElseIf Kieli = "randomi" Then
        'Randomisoidaan kysytäänkö sve-suo vai suo-sve
        x = Int(2 * Rnd + 1) - 1
        If x = 0 Then
            SanaEng = SanaAlue.Range("B" & KysyRivi(Kysytty)).Value
            SanaSuo = SanaAlue.Range("A" & KysyRivi(Kysytty)).Value
        Else
            SanaEng = SanaAlue.Range("A" & KysyRivi(Kysytty)).Value
            SanaSuo = SanaAlue.Range("B" & KysyRivi(Kysytty)).Value
        End If
    End If
        
    Vastaus = InputBox("Sana: " & SanaEng & vbNewLine & vbNewLine & "Oikein: " & Oikein & ", " & "väärin: " & Vaarin & vbNewLine & "Rivi # " & KysyRivi(Kysytty) & ", " & "jäljellä: " & ((UBound(KysyRivi) - LBound(KysyRivi) + 1) - Kysytty), "Suomi - Englanti - Suomi käännös")
            If StrPtr(Vastaus) = 0 Then
                Vastaus = "stop"
            ElseIf Vastaus = vbNullString Then
                MsgBox ("Kirjoita jotain!")
            ElseIf Vastaus = SanaSuo Then
                MsgBox ("Oikein! Todella hienoa!")
            ReDim Preserve OikeinRivi(Oikein)
            OikeinRivi(Oikein) = KysyRivi(Kysytty)
            Oikein = Oikein + 1
            Else
                MsgBox ("Väärin. Sana oli: " & SanaSuo)
                Vaarin = Vaarin + 1
            End If
    Kysytty = Kysytty + 1
Loop

End Sub
