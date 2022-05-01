Sub Aakkosta()
'
' Aakkosta Makro
' Aakkostaa jäsenrekisterin sukunimen mukaiseen laskevaan järjestykseen.
' Jari Hiltunen 2015
' Pikanäppäin: Ctrl+a
'
Dim sheetti As Worksheet
Dim viimeinenrivi As Long
Set sheetti = ThisWorkbook.Worksheets("Tietokanta")
'Etsitään viimeinen täytetty rivi Ctrl + Shift + End
  viimeinenrivi = sheetti.Cells(sheetti.Rows.Count, "A").End(xlUp).Row
    Rows("3:4").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Tietokanta").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Tietokanta").Sort.SortFields.Add Key:=Range("A3") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tietokanta").Sort
    ' Valitaan alueeseen oletuksena isoin riviarvo
        .SetRange Range("A3:L" & viimeinenrivi)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    MsgBox "Tiedot aakkostettu."
End Sub
Sub Tallenna()
'
' Tallenna Makro
' Tallentaa työkirjan ja luo siitä varmuuskopion
'
' Pikanäppäin: Ctrl+s
'
    ActiveWorkbook.Save
    
    ' Tekee varmuuskopion
    Application.EnableEvents = False
    thisPath = ThisWorkbook.Path
    myName = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".") - 1))
    ext = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, "."))
    backupdirectory = myName & " varmuuskopiot"

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' Jos varmuuskopiohakemistoa ei ole, niin luodaan se
    If Not FSO.FolderExists(ThisWorkbook.Path & "/" & backupdirectory) Then
        FSO.CreateFolder (ThisWorkbook.Path & "/" & backupdirectory)
    End If

    ' Nimetään tiedosto
    T = Format(Now, "mmm dd yyyy hh mm ss")
    ThisWorkbook.SaveCopyAs thisPath & "\" & backupdirectory & "\" & myName & " " & T & "." & ext

    Application.EnableEvents = True
    ActiveSheet.Protect Password:="SalaSana" ' Suojataan työkirja
      
    MsgBox "Tiedot tallennettu ja varmuuskopioitu onnistuneesti!"
End Sub

Sub LisaaHenkilo()
'
' LisaaHenkilo Makro
' Lisää henkilön pohjadatan ja generoi uuden jäsennumeron.
'
 Dim Alue As Range
 Dim maximi As Double
 Dim sheetti As Worksheet
 Dim viimeinenrivi As Long
   'Etsitään viimeinen täytetty rivi eli vastaa painallusta Ctrl + Shift + End
   Set sheetti = ThisWorkbook.Worksheets("Tietokanta")
   viimeinenrivi = sheetti.Cells(sheetti.Rows.Count, "A").End(xlUp).Row
 
 Sheets("Tietokanta").Select
  ActiveSheet.Unprotect Password:="SalaSana" ' Poistetaan työkirjan suojaus
     Set Alue = Range("C3:C" & viimeinenrivi) ' Alue josta suurinta arvoa etsitään
     maximi = WorksheetFunction.Max(Range("C3:C" & viimeinenrivi)) 'Palautetaan suurin jäsennumero
  Worksheets("Parametrit").Visible = True
  Sheets("Parametrit").Select
    Range("C4").Value = maximi + 1 ' Lisätään jäsennumeroa yhdellä
                    
 Sheets("Parametrit").Select
    Sheets("Parametrit").Range("A4:L4").Copy ' Valitaan pohjatiedot ilman jäsennumeroa
                                             
 Sheets("Tietokanta").Select
    ActiveSheet.Range("A" & Rows.Count).End(xlUp).Offset(1).Select ' Etsitään ensimmäinen tyhjä rivi
         ActiveSheet.Paste ' Lisätään tieto valittuun riviin
   Worksheets("Parametrit").Visible = False ' Piilota parametrit
 ActiveSheet.Protect Password:="SalaSana" ' Suojataan työkirja
                                            
End Sub

Sub PoistaHenkilo()
'
' PoistaHenkilo Makro
' Poistaa taulukon suojauksen, kysyy poistettavan hennkilön sukunimeä ja maalaa poistettavan rivin valmiiksi.
'
'
 ActiveSheet.Unprotect Password:="SalaSana" ' Poistetaan työkirjan suojaus

 Dim KysySuku As Variant
 Dim Alue As Range
 Dim sheetti As Worksheet
 Dim viimeinenrivi As Long
 Dim Vastaus As Integer
   'Etsitään viimeinen täytetty rivi eli vastaa painallusta Ctrl + Shift + End
   Set sheetti = ThisWorkbook.Worksheets("Tietokanta")
   viimeinenrivi = sheetti.Cells(sheetti.Rows.Count, "A").End(xlUp).Row
   KysySuku = InputBox("Kirjoita poistettavan henkilön sukunimi") ' Kysytään etsittävä sukunimi
   
   
   ' Loopataan niin pitkään kunnes ei löydy sukunimeä poistettavaksi
    Set Alue = Worksheets("Tietokanta").Range("A3:A" & viimeinenrivi).Find( _
              What:=KysySuku, LookIn:=xlFormulas) ' Mistä etsitään
     If Not Alue Is Nothing Then
        Alue.Parent.Activate ' Toiminnot mitä tehdään jos löyty
        Alue.EntireRow.Select
          
       ' Kysyy onko oikea ja näyttää vaihtoehdot Kyllä ja Ei
        Vastaus = MsgBox(prompt:="Onko oikea poistettava henkilö? Valitse 'Kyllä' or 'Ei'.", Buttons:=vbYesNo)
      
         ' Mikäli Kyllä oli valittu
         If Vastaus = vbYes Then
         ' Valittiin Kyllä, eli poistetaan rivi
         Rows(ActiveCell.Row).EntireRow.Delete
         Else
         ' Valittiin Ei.
         ' Ei etsitä seuraavaa samaa sukunimeä olevaa
         MsgBox ("Voit poistaa henkilön valitsemalla koko rivin")
         End If
       
        
    Else
        MsgBox "Sukunimeä ei löytynyt."
    End If
 
    
End Sub

