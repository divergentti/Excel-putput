Option Compare Text
Option Explicit

Dim objPresentaion As Presentation
Dim objSlide As Slide
Dim objTextBox As Shape
Private Sub Os1Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    TietoLomake.Os1Omin.Visible = True
End If

End Sub
Private Sub Os1Omin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
 ' Tarkistetaan onko ensimmäiset kirjaimet samat
If KeyCode = vbKeyReturn Then
    If Left(TietoLomake.Os1Nimi, 1) = Left(TietoLomake.Os1Omin, 1) Then
        TietoLomake.Os2Nimi.Visible = True
        TietoLomake.Os2l.Visible = True
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If

End Sub

Private Sub Os2Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    If (TietoLomake.Os2Nimi = "") Then
        Osallistujia = 1
        TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
        TietoLomake.Osallistuu.Visible = True
        TietoLomake.Os1l2 = TietoLomake.Os1Nimi
        TietoLomake.Piirre1_1.Visible = True
        TietoLomake.Piirre1_1.SetFocus
        TietoLomake.Os1l2.Visible = True
    Else
        TietoLomake.Os2Omin.Visible = True
        TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
        TietoLomake.Osallistuu.Visible = True
        TietoLomake.Os1l2 = TietoLomake.Os1Nimi
    End If
End If

End Sub
Private Sub Os2Omin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    If Left(TietoLomake.Os2Nimi, 1) = Left(TietoLomake.Os2Omin, 1) Then
        TietoLomake.Os3Nimi.Visible = True
        TietoLomake.Os3l.Visible = True
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If

End Sub
Private Sub Os3Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
 If KeyCode = vbKeyReturn Then
   If (TietoLomake.Os3Nimi = "") Then
    Osallistujia = 2
    TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
    TietoLomake.Piirre1_1.Visible = True
    TietoLomake.Piirre1_1.SetFocus
    TietoLomake.Os2l2 = TietoLomake.Os2Nimi
    TietoLomake.Os1l2.Visible = True
   Else
    TietoLomake.Os3Omin.Visible = True
   End If
End If
End Sub
Private Sub Os3Omin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    If Left(TietoLomake.Os3Nimi, 1) = Left(TietoLomake.Os3Omin, 1) Then
        TietoLomake.Os4Nimi.Visible = True
        TietoLomake.Os4l.Visible = True
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If
End Sub

Private Sub Os4Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
   If (TietoLomake.Os4Nimi = "") Then
    Osallistujia = 3
    TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
    TietoLomake.Piirre1_1.Visible = True
    TietoLomake.Piirre1_1.SetFocus
    TietoLomake.Os3l2 = TietoLomake.Os3Nimi
    TietoLomake.Os1l2.Visible = True
   Else
    TietoLomake.Os4Omin.Visible = True
   End If
End If
End Sub
Private Sub Os4Omin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    If Left(TietoLomake.Os4Nimi, 1) = Left(TietoLomake.Os4Omin, 1) Then
        TietoLomake.Os5Nimi.Visible = True
        TietoLomake.Os5l.Visible = True
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If
End Sub

Private Sub Os5Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
 If KeyCode = vbKeyReturn Then
   If (TietoLomake.Os5Nimi = "") Then
    Osallistujia = 4
    TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
    TietoLomake.Piirre1_1.Visible = True
    TietoLomake.Piirre1_1.SetFocus
    TietoLomake.Os1l2.Visible = True
    TietoLomake.Os4l2 = TietoLomake.Os4Nimi
   Else
    TietoLomake.Os5Omin.Visible = True
   End If
End If
End Sub
Private Sub Os5Omin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    If Left(TietoLomake.Os5Nimi, 1) = Left(TietoLomake.Os5Omin, 1) Then
        TietoLomake.Os6Nimi.Visible = True
        TietoLomake.Os6l.Visible = True
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If
End Sub
Private Sub Os6Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
 If KeyCode = vbKeyReturn Then
   If (TietoLomake.Os6Nimi = "") Then
    Osallistujia = 5
    TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
    TietoLomake.Piirre1_1.Visible = True
    TietoLomake.Piirre1_1.SetFocus
    TietoLomake.Os1l2.Visible = True
    TietoLomake.Os5l2 = TietoLomake.Os5Nimi
   Else
    TietoLomake.Os6Omin.Visible = True
   End If
End If
End Sub

Private Sub Os6Omin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    If Left(TietoLomake.Os6Nimi, 1) = Left(TietoLomake.Os6Omin, 1) Then
        TietoLomake.Os7Nimi.Visible = True
        TietoLomake.Os7l.Visible = True
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If
End Sub

Private Sub Os7Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
 If KeyCode = vbKeyReturn Then
   If (TietoLomake.Os7Nimi = "") Then
    Osallistujia = 6
    TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
    TietoLomake.Piirre1_1.Visible = True
    TietoLomake.Piirre1_1.SetFocus
    TietoLomake.Os1l2.Visible = True
    TietoLomake.Os6l2 = TietoLomake.Os6Nimi
   Else
    TietoLomake.Os7Omin.Visible = True
   End If
End If
End Sub

Private Sub Os7Omin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    If Left(TietoLomake.Os7Nimi, 1) = Left(TietoLomake.Os7Omin, 1) Then
        TietoLomake.Os8Nimi.Visible = True
        TietoLomake.Os8l.Visible = True
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If
End Sub
Private Sub Os8Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
   If (TietoLomake.Os8Nimi = "") Then
    Osallistujia = 7
    TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
    TietoLomake.Piirre1_1.Visible = True
    TietoLomake.Piirre1_1.SetFocus
    TietoLomake.Os1l2.Visible = True
    TietoLomake.Os7l2 = TietoLomake.Os7Nimi
   Else
    TietoLomake.Os8Omin.Visible = True
    Osallistujia = 8
    TietoLomake.Os8l2 = TietoLomake.Os8Nimi
    TietoLomake.Osallistuu = "Nimiä: " & Osallistujia
    ' Maksimi saatuvettu
   End If
End If
End Sub

Private Sub Os8Omin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    If Left(TietoLomake.Os8Nimi, 1) = Left(TietoLomake.Os8Omin, 1) Then
        TietoLomake.Os8Nimi.Visible = True
        TietoLomake.Piirre1_1.Visible = True
        TietoLomake.Piirre1_1.SetFocus
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If
End Sub

Private Sub Piirre1_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre2_1.Visible = True
End If

End Sub

Private Sub Piirre2_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre3_1.Visible = True
End If

End Sub

Private Sub Piirre3_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre4_1.Visible = True
End If
End Sub
Private Sub Piirre4_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre5_1.Visible = True
End If
End Sub

Private Sub Piirre5_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Os1Tulos = TietoLomake.Os1Nimi & " " & TietoLomake.Os1Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_1 & " " & TietoLomake.Piirre2_1 & " " & TietoLomake.Piirre3_1 & " " & TietoLomake.Piirre4_1 & " " & TietoLomake.Piirre5_1 & ". Ok. Hyvä!"
 TietoLomake.Os1Tulos.Visible = True
 ' Powerpoint sliden rakentaminen valmiiksi
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(4) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = TietoLomake.Os1Omin & " " & TietoLomake.Os1Nimi & " on selvästi puolustaja!" ' Nimi ja ominaisuus"
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Puolustaja on yleensä " & TietoLomake.Piirre1_1 & ", mikä on epätyypillinen luonteenpiirre introverteille persoonille. " _
    & "Hän on " & TietoLomake.Piirre2_1 & ", mutta hän ei käytä sitä tiedon ja nippelitiedon tallentamiseen, vaan painaa mieleensä asioita ihmisistä ja heidän elämästään." & vbNewLine & _
     "Mitä tulee lahjojen antamiseen, puolustaja on vertaansa vailla. Hän on " & TietoLomake.Piirre3_1 & " ja " & TietoLomake.Piirre4_1 & " lisäämään herkkyyttä ilmaisemaan anteliaisuuttaan tavalla, joka koskettaa vastapuolta." & vbNewLine & _
     "Toisinaan " & TietoLomake.Piirre5_1 & " tulee esiin vain työntekijöiden keskuudessa, joita puolustaja todella pitää henkilökohtaisena ystävänään, perheen keskuudessa puolustajan tunneilmaisu pääse todella valloilleen. "
 ActivePresentation.Slides(4).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
    If Osallistujia >= 2 Then
        TietoLomake.Piirre1_2.Visible = True
        TietoLomake.Piirre1_2.SetFocus
        ElseIf Osallistujia = 1 Then
        Valmista
    End If
End If


End Sub

Private Sub Piirre1_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre2_2.Visible = True
End If

End Sub

Private Sub Piirre2_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre3_2.Visible = True
End If

End Sub

Private Sub Piirre3_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre4_2.Visible = True
End If
End Sub
Private Sub Piirre4_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre5_2.Visible = True
End If
End Sub

Private Sub Piirre5_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Os2Tulos = TietoLomake.Os2Nimi & " " & TietoLomake.Os2Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_2 & " " & TietoLomake.Piirre2_2 & " " & TietoLomake.Piirre3_2 & " " & TietoLomake.Piirre4_2 & " " & TietoLomake.Piirre5_2 & ". Ok. Hyvä!"
 TietoLomake.Os2Tulos.Visible = True
 ' Powerpoint sliden rakentaminen valmiiksi
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(5) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = TietoLomake.Os2Omin & " " & TietoLomake.Os2Nimi ' Nimi ja ominaisuus
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Hänen luonteepiirteistään vahvin on " & TietoLomake.Piirre1_2 & vbNewLine & _
     "Miesten seurassa hän on " & TietoLomake.Piirre2_2 & vbNewLine & _
     "Naisten seurassa hän on " & TietoLomake.Piirre3_2 & vbNewLine & _
     "Seurustellessaan hän on " & TietoLomake.Piirre4_2 & vbNewLine & _
     "Kotona hän on " & TietoLomake.Piirre5_2
 ActivePresentation.Slides(5).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
 If Osallistujia >= 3 Then
     TietoLomake.Piirre1_3.Visible = True
     TietoLomake.Piirre1_3.SetFocus
 ElseIf Osallistujia = 2 Then
     Valmista
 End If
End If

End Sub

Private Sub Piirre1_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre2_3.Visible = True
End If

End Sub

Private Sub Piirre2_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre3_3.Visible = True
End If

End Sub

Private Sub Piirre3_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre4_3.Visible = True
End If
End Sub
Private Sub Piirre4_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre5_3.Visible = True
End If
End Sub

Private Sub Piirre5_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Os3Tulos = TietoLomake.Os3Nimi & " " & TietoLomake.Os3Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_3 & " " & TietoLomake.Piirre2_3 & " " & TietoLomake.Piirre3_3 & " " & TietoLomake.Piirre4_3 & " " & TietoLomake.Piirre5_3 & ". Ok. Hyvä!"
 TietoLomake.Os3Tulos.Visible = True
 ' Powerpoint sliden rakentaminen valmiiksi
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(6) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = TietoLomake.Os3Omin & " " & TietoLomake.Os3Nimi ' Nimi ja ominaisuus
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Hänen luonteepiirteistään vahvin on " & TietoLomake.Piirre1_3 & vbNewLine & _
     "Miesten seurassa hän on " & TietoLomake.Piirre2_3 & vbNewLine & _
     "Naisten seurassa hän on " & TietoLomake.Piirre3_3 & vbNewLine & _
     "Seurustellessaan hän on " & TietoLomake.Piirre4_3 & vbNewLine & _
     "Kotona hän on " & TietoLomake.Piirre5_3
 ActivePresentation.Slides(6).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
 If Osallistujia >= 4 Then
     TietoLomake.Piirre1_4.Visible = True
     TietoLomake.Piirre1_4.SetFocus
  ElseIf Osallistujia = 3 Then
     Valmista
  End If
End If


End Sub
Private Sub Piirre1_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre2_4.Visible = True
End If

End Sub

Private Sub Piirre2_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre3_4.Visible = True
End If

End Sub

Private Sub Piirre3_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre4_4.Visible = True
End If
End Sub
Private Sub Piirre4_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre5_4.Visible = True
End If
End Sub

Private Sub Piirre5_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Os4Tulos = TietoLomake.Os4Nimi & " " & TietoLomake.Os4Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_4 & " " & TietoLomake.Piirre2_4 & " " & TietoLomake.Piirre3_4 & " " & TietoLomake.Piirre4_4 & " " & TietoLomake.Piirre5_4 & ". Ok. Hyvä!"
 TietoLomake.Os4Tulos.Visible = True
 ' Powerpoint sliden rakentaminen valmiiksi
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(7) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = TietoLomake.Os4Omin & " " & TietoLomake.Os4Nimi ' Nimi ja ominaisuus
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Hänen luonteepiirteistään vahvin on " & TietoLomake.Piirre1_4 & vbNewLine & _
     "Miesten seurassa hän on " & TietoLomake.Piirre2_4 & vbNewLine & _
     "Naisten seurassa hän on " & TietoLomake.Piirre3_4 & vbNewLine & _
     "Seurustellessaan hän on " & TietoLomake.Piirre4_4 & vbNewLine & _
     "Kotona hän on " & TietoLomake.Piirre5_4
 ActivePresentation.Slides(7).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
   If Osallistujia >= 5 Then
     TietoLomake.Piirre1_5.Visible = True
     TietoLomake.Piirre1_5.SetFocus
  ElseIf Osallistujia = 4 Then
     Valmista
  End If
End If
End Sub

Private Sub Piirre1_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre2_5.Visible = True
End If

End Sub

Private Sub Piirre2_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre3_5.Visible = True
End If

End Sub

Private Sub Piirre3_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre4_5.Visible = True
End If
End Sub
Private Sub Piirre4_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre5_5.Visible = True
End If
End Sub

Private Sub Piirre5_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Os5Tulos = TietoLomake.Os5Nimi & " " & TietoLomake.Os5Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_5 & " " & TietoLomake.Piirre2_5 & " " & TietoLomake.Piirre3_5 & " " & TietoLomake.Piirre4_5 & " " & TietoLomake.Piirre5_5 & ". Ok. Hyvä!"
 TietoLomake.Os5Tulos.Visible = True
 ' Powerpoint sliden rakentaminen valmiiksi
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(8) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = TietoLomake.Os5Omin & " " & TietoLomake.Os5Nimi ' Nimi ja ominaisuus
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Hänen luonteepiirteistään vahvin on " & TietoLomake.Piirre1_5 & vbNewLine & _
     "Miesten seurassa hän on " & TietoLomake.Piirre2_5 & vbNewLine & _
     "Naisten seurassa hän on " & TietoLomake.Piirre3_5 & vbNewLine & _
     "Seurustellessaan hän on " & TietoLomake.Piirre4_5 & vbNewLine & _
     "Kotona hän on " & TietoLomake.Piirre5_5
 ActivePresentation.Slides(8).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
  If Osallistujia >= 6 Then
     TietoLomake.Piirre1_6.Visible = True
     TietoLomake.Piirre1_6.SetFocus
  ElseIf Osallistujia = 5 Then
     Valmista
  End If
End If

End Sub

Private Sub Piirre1_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre2_6.Visible = True
End If

End Sub

Private Sub Piirre2_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre3_6.Visible = True
End If

End Sub

Private Sub Piirre3_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre4_6.Visible = True
End If
End Sub
Private Sub Piirre4_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre5_6.Visible = True
End If
End Sub

Private Sub Piirre5_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Os6Tulos = TietoLomake.Os6Nimi & " " & TietoLomake.Os6Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_6 & " " & TietoLomake.Piirre2_6 & " " & TietoLomake.Piirre3_6 & " " & TietoLomake.Piirre4_6 & " " & TietoLomake.Piirre5_6 & ". Ok. Hyvä!"
 TietoLomake.Os6Tulos.Visible = True
 ' Powerpoint sliden rakentaminen valmiiksi
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(9) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = TietoLomake.Os6Omin & " " & TietoLomake.Os6Nimi ' Nimi ja ominaisuus
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Hänen luonteepiirteistään vahvin on " & TietoLomake.Piirre1_6 & vbNewLine & _
     "Miesten seurassa hän on " & TietoLomake.Piirre2_6 & vbNewLine & _
     "Naisten seurassa hän on " & TietoLomake.Piirre3_6 & vbNewLine & _
     "Seurustellessaan hän on " & TietoLomake.Piirre4_6 & vbNewLine & _
     "Kotona hän on " & TietoLomake.Piirre5_6
 ActivePresentation.Slides(9).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
 If Osallistujia >= 7 Then
     TietoLomake.Piirre1_7.Visible = True
     TietoLomake.Piirre1_7.SetFocus
  ElseIf Osallistujia = 6 Then
     Valmista
  End If
End If

End Sub

Private Sub Piirre1_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre2_7.Visible = True
End If

End Sub

Private Sub Piirre2_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre3_7.Visible = True
End If

End Sub

Private Sub Piirre3_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre4_7.Visible = True
End If
End Sub
Private Sub Piirre4_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre5_7.Visible = True
End If
End Sub

Private Sub Piirre5_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Os7Tulos = TietoLomake.Os7Nimi & " " & TietoLomake.Os7Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_7 & " " & TietoLomake.Piirre2_7 & " " & TietoLomake.Piirre3_7 & " " & TietoLomake.Piirre4_7 & " " & TietoLomake.Piirre5_7 & ". Ok. Hyvä!"
 TietoLomake.Os7Tulos.Visible = True
 ' Powerpoint sliden rakentaminen valmiiksi
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(10) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = TietoLomake.Os7Omin & " " & TietoLomake.Os7Nimi ' Nimi ja ominaisuus
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Hänen luonteepiirteistään vahvin on " & TietoLomake.Piirre1_7 & vbNewLine & _
     "Miesten seurassa hän on " & TietoLomake.Piirre2_7 & vbNewLine & _
     "Naisten seurassa hän on " & TietoLomake.Piirre3_7 & vbNewLine & _
     "Seurustellessaan hän on " & TietoLomake.Piirre4_7 & vbNewLine & _
     "Kotona hän on " & TietoLomake.Piirre5_7
 ActivePresentation.Slides(10).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
 If Osallistujia = 8 Then
     TietoLomake.Piirre1_8.Visible = True
     TietoLomake.Piirre1_8.SetFocus
  ElseIf Osallistujia = 7 Then
     Valmista
  End If
End If

End Sub
Private Sub Piirre1_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre2_8.Visible = True
End If

End Sub

Private Sub Piirre2_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre3_8.Visible = True
End If

End Sub

Private Sub Piirre3_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre4_8.Visible = True
End If
End Sub
Private Sub Piirre4_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Piirre5_8.Visible = True
End If
End Sub

Private Sub Piirre5_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 TietoLomake.Os8Tulos = TietoLomake.Os8Nimi & " " & TietoLomake.Os8Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_8 & " " & TietoLomake.Piirre2_8 & " " & TietoLomake.Piirre3_8 & " " & TietoLomake.Piirre4_8 & " " & TietoLomake.Piirre5_8 & ". Ok. Hyvä!"
 TietoLomake.Os8Tulos.Visible = True
  ' Powerpoint sliden rakentaminen valmiiksi
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(11) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = TietoLomake.Os8Omin & " " & TietoLomake.Os8Nimi ' Nimi ja ominaisuus
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Hänen luonteepiirteistään vahvin on " & TietoLomake.Piirre1_8 & vbNewLine & _
     "Miesten seurassa hän on " & TietoLomake.Piirre2_8 & vbNewLine & _
     "Naisten seurassa hän on " & TietoLomake.Piirre3_8 & vbNewLine & _
     "Seurustellessaan hän on " & TietoLomake.Piirre4_8 & vbNewLine & _
     "Kotona hän on " & TietoLomake.Piirre5_8
 ActivePresentation.Slides(11).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
 Valmista
End If

End Sub

Private Sub Valmista()
    TietoLomake.ValmisNappi.Visible = True

End Sub

Private Sub ValmisNappi_Click()

Unload Me

End Sub

