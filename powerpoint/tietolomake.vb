Option Compare Text
Option Explicit
Dim objPresentaion As Presentation
Dim objSlide As Slide
Dim objTextBox As Shape

Private Sub Os1Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
    NimetJaOminaisuudet(1, 1) = TietoLomake.Os1Nimi.Value
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
        NimetJaOminaisuudet(1, 2) = TietoLomake.Os1Omin.Value
    Else
        MsgBox ("Nimen ensimmäinen kirjain ja ominaisuuden " & vbCrLf & "ensimmäinen kirjain ei ole sama!")
    End If
End If
End Sub

Private Sub Os2Nimi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = vbKeyReturn Then
  NimetJaOminaisuudet(2, 1) = TietoLomake.Os2Nimi.Value
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
        NimetJaOminaisuudet(2, 2) = TietoLomake.Os2Omin.Value
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
  NimetJaOminaisuudet(3, 1) = TietoLomake.Os3Nimi.Value
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
        NimetJaOminaisuudet(3, 2) = TietoLomake.Os3Omin.Value
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
 NimetJaOminaisuudet(4, 1) = TietoLomake.Os4Nimi.Value
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
        NimetJaOminaisuudet(4, 2) = TietoLomake.Os4Omin.Value
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
  NimetJaOminaisuudet(5, 1) = TietoLomake.Os5Nimi.Value
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
        NimetJaOminaisuudet(5, 2) = TietoLomake.Os5Omin.Value
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
  NimetJaOminaisuudet(6, 1) = TietoLomake.Os6Nimi.Value
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
        NimetJaOminaisuudet(6, 2) = TietoLomake.Os6Omin.Value
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
  NimetJaOminaisuudet(7, 1) = TietoLomake.Os7Nimi.Value
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
        NimetJaOminaisuudet(7, 2) = TietoLomake.Os7Omin.Value
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
 NimetJaOminaisuudet(8, 1) = TietoLomake.Os8Nimi.Value
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
        NimetJaOminaisuudet(8, 2) = TietoLomake.Os8Omin.Value
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
  If Len(TietoLomake.Piirre1_1) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
    TietoLomake.Piirre2_1.Visible = True
    NimetJaOminaisuudet(1, 3) = TietoLomake.Piirre1_1.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
  End If
End If

End Sub

Private Sub Piirre2_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
  If Len(TietoLomake.Piirre2_1) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
    TietoLomake.Piirre3_1.Visible = True
    NimetJaOminaisuudet(1, 4) = TietoLomake.Piirre2_1.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
  End If
End If

End Sub

Private Sub Piirre3_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
  If Len(TietoLomake.Piirre3_1) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre4_1.Visible = True
   NimetJaOminaisuudet(1, 5) = TietoLomake.Piirre3_1.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
  End If
End If
End Sub
Private Sub Piirre4_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
  If Len(TietoLomake.Piirre4_1) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
    TietoLomake.Piirre5_1.Visible = True
    NimetJaOminaisuudet(1, 6) = TietoLomake.Piirre4_1.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
  End If
End If
End Sub

Private Sub Piirre5_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre5_1) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
 Else
  TietoLomake.Os1Tulos = TietoLomake.Os1Nimi & " " & TietoLomake.Os1Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_1 & " " & TietoLomake.Piirre2_1 & " " & TietoLomake.Piirre3_1 & " " & TietoLomake.Piirre4_1 & " " & TietoLomake.Piirre5_1 & ". Ok. Hyvä!"
  TietoLomake.Os1Tulos.Visible = True
  NimetJaOminaisuudet(1, 7) = TietoLomake.Piirre5_1.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
    If Osallistujia >= 2 Then
        TietoLomake.Piirre1_2.Visible = True
        TietoLomake.Piirre1_2.SetFocus
        ElseIf Osallistujia = 1 Then
        Valmista
    End If
 End If
End If

End Sub

Private Sub Piirre1_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre1_2) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre2_2.Visible = True
   NimetJaOminaisuudet(2, 3) = TietoLomake.Piirre1_2.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre2_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre2_2) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre3_2.Visible = True
   NimetJaOminaisuudet(2, 4) = TietoLomake.Piirre2_2.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre3_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre3_2) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre4_2.Visible = True
   NimetJaOminaisuudet(2, 5) = TietoLomake.Piirre3_2.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub
Private Sub Piirre4_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre4_2) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre5_2.Visible = True
   NimetJaOminaisuudet(2, 6) = TietoLomake.Piirre4_2.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre5_2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre5_2) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
 Else
  TietoLomake.Os2Tulos = TietoLomake.Os2Nimi & " " & TietoLomake.Os2Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_2 & " " & TietoLomake.Piirre2_2 & " " & TietoLomake.Piirre3_2 & " " & TietoLomake.Piirre4_2 & " " & TietoLomake.Piirre5_2 & ". Ok. Hyvä!"
  TietoLomake.Os2Tulos.Visible = True
  NimetJaOminaisuudet(2, 7) = TietoLomake.Piirre5_2.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
  If Osallistujia >= 3 Then
      TietoLomake.Piirre1_3.Visible = True
      TietoLomake.Piirre1_3.SetFocus
  ElseIf Osallistujia = 2 Then
      Valmista
  End If
 End If
End If

End Sub

Private Sub Piirre1_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre1_3) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre2_3.Visible = True
   NimetJaOminaisuudet(3, 3) = TietoLomake.Piirre1_3.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre2_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre2_3) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre3_3.Visible = True
   NimetJaOminaisuudet(3, 4) = TietoLomake.Piirre2_3.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre3_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre3_3) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
 Else
   TietoLomake.Piirre4_3.Visible = True
   NimetJaOminaisuudet(3, 5) = TietoLomake.Piirre3_3.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub
Private Sub Piirre4_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre4_3) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre5_3.Visible = True
   NimetJaOminaisuudet(3, 6) = TietoLomake.Piirre4_3.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre5_3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre5_3) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
 Else
  TietoLomake.Os3Tulos = TietoLomake.Os3Nimi & " " & TietoLomake.Os3Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_3 & " " & TietoLomake.Piirre2_3 & " " & TietoLomake.Piirre3_3 & " " & TietoLomake.Piirre4_3 & " " & TietoLomake.Piirre5_3 & ". Ok. Hyvä!"
  TietoLomake.Os3Tulos.Visible = True
  NimetJaOminaisuudet(3, 7) = TietoLomake.Piirre5_3.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
  If Osallistujia >= 4 Then
      TietoLomake.Piirre1_4.Visible = True
      TietoLomake.Piirre1_4.SetFocus
   ElseIf Osallistujia = 3 Then
      Valmista
   End If
 End If
End If

End Sub
Private Sub Piirre1_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre1_4) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre2_4.Visible = True
   NimetJaOminaisuudet(4, 3) = TietoLomake.Piirre1_4.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre2_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre2_4) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre3_4.Visible = True
   NimetJaOminaisuudet(4, 4) = TietoLomake.Piirre2_4.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre3_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre3_4) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre4_4.Visible = True
   NimetJaOminaisuudet(4, 5) = TietoLomake.Piirre3_4.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub
Private Sub Piirre4_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre4_4) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre5_4.Visible = True
   NimetJaOminaisuudet(4, 6) = TietoLomake.Piirre4_4.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre5_4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre5_4) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
 Else
  TietoLomake.Os4Tulos = TietoLomake.Os4Nimi & " " & TietoLomake.Os4Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_4 & " " & TietoLomake.Piirre2_4 & " " & TietoLomake.Piirre3_4 & " " & TietoLomake.Piirre4_4 & " " & TietoLomake.Piirre5_4 & ". Ok. Hyvä!"
  TietoLomake.Os4Tulos.Visible = True
  NimetJaOminaisuudet(4, 7) = TietoLomake.Piirre5_4.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
    If Osallistujia >= 5 Then
      TietoLomake.Piirre1_5.Visible = True
      TietoLomake.Piirre1_5.SetFocus
   ElseIf Osallistujia = 4 Then
      Valmista
   End If
  End If
 End If
End Sub

Private Sub Piirre1_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre1_5) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre2_5.Visible = True
   NimetJaOminaisuudet(5, 3) = TietoLomake.Piirre1_5.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre2_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre2_5) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
 Else
  TietoLomake.Piirre3_5.Visible = True
  NimetJaOminaisuudet(5, 4) = TietoLomake.Piirre2_5.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre3_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre3_5) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre4_5.Visible = True
   NimetJaOminaisuudet(5, 5) = TietoLomake.Piirre3_5.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub
Private Sub Piirre4_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre4_5) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre5_5.Visible = True
   NimetJaOminaisuudet(5, 6) = TietoLomake.Piirre4_5.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre5_5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
  If Len(TietoLomake.Piirre5_5) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Os5Tulos = TietoLomake.Os5Nimi & " " & TietoLomake.Os5Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_5 & " " & TietoLomake.Piirre2_5 & " " & TietoLomake.Piirre3_5 & " " & TietoLomake.Piirre4_5 & " " & TietoLomake.Piirre5_5 & ". Ok. Hyvä!"
   TietoLomake.Os5Tulos.Visible = True
   NimetJaOminaisuudet(5, 7) = TietoLomake.Piirre5_5.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
    If Osallistujia >= 6 Then
       TietoLomake.Piirre1_6.Visible = True
       TietoLomake.Piirre1_6.SetFocus
    ElseIf Osallistujia = 5 Then
       Valmista
    End If
 End If
End If

End Sub

Private Sub Piirre1_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre1_6) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre2_6.Visible = True
   NimetJaOminaisuudet(6, 3) = TietoLomake.Piirre1_6.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre2_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre2_6) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre3_6.Visible = True
   NimetJaOminaisuudet(6, 4) = TietoLomake.Piirre2_6.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre3_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre3_6) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre4_6.Visible = True
   NimetJaOminaisuudet(6, 5) = TietoLomake.Piirre3_6.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub
Private Sub Piirre4_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre4_6) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre5_6.Visible = True
   NimetJaOminaisuudet(6, 6) = TietoLomake.Piirre4_6.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre5_6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre5_6) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Os6Tulos = TietoLomake.Os6Nimi & " " & TietoLomake.Os6Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_6 & " " & TietoLomake.Piirre2_6 & " " & TietoLomake.Piirre3_6 & " " & TietoLomake.Piirre4_6 & " " & TietoLomake.Piirre5_6 & ". Ok. Hyvä!"
   TietoLomake.Os6Tulos.Visible = True
   NimetJaOminaisuudet(6, 7) = TietoLomake.Piirre5_6.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
   If Osallistujia >= 7 Then
       TietoLomake.Piirre1_7.Visible = True
       TietoLomake.Piirre1_7.SetFocus
    ElseIf Osallistujia = 6 Then
       Valmista
    End If
 End If
End If

End Sub

Private Sub Piirre1_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre1_7) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre2_7.Visible = True
   NimetJaOminaisuudet(7, 3) = TietoLomake.Piirre1_7.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre2_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre2_7) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre3_7.Visible = True
   NimetJaOminaisuudet(7, 4) = TietoLomake.Piirre2_7.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre3_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre3_7) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre4_7.Visible = True
   NimetJaOminaisuudet(7, 5) = TietoLomake.Piirre3_7.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub
Private Sub Piirre4_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre4_7) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre5_7.Visible = True
   NimetJaOminaisuudet(7, 6) = TietoLomake.Piirre4_7.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub

Private Sub Piirre5_7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre5_7) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Os7Tulos = TietoLomake.Os7Nimi & " " & TietoLomake.Os7Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_7 & " " & TietoLomake.Piirre2_7 & " " & TietoLomake.Piirre3_7 & " " & TietoLomake.Piirre4_7 & " " & TietoLomake.Piirre5_7 & ". Ok. Hyvä!"
   TietoLomake.Os7Tulos.Visible = True
   NimetJaOminaisuudet(7, 7) = TietoLomake.Piirre5_7.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
   If Osallistujia = 8 Then
       TietoLomake.Piirre1_8.Visible = True
       TietoLomake.Piirre1_8.SetFocus
    ElseIf Osallistujia = 7 Then
       Valmista
    End If
 End If
End If

End Sub
Private Sub Piirre1_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
  If Len(TietoLomake.Piirre1_8) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre2_8.Visible = True
   NimetJaOminaisuudet(8, 3) = TietoLomake.Piirre1_8.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
  End If
End If

End Sub

Private Sub Piirre2_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre2_8) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre3_8.Visible = True
   NimetJaOminaisuudet(8, 4) = TietoLomake.Piirre2_8.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If

End Sub

Private Sub Piirre3_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre3_8) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre4_8.Visible = True
   NimetJaOminaisuudet(8, 5) = TietoLomake.Piirre3_8.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
 End If
End If
End Sub
Private Sub Piirre4_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
  If Len(TietoLomake.Piirre4_8) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Piirre5_8.Visible = True
   NimetJaOminaisuudet(8, 6) = TietoLomake.Piirre4_8.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
  End If
End If
End Sub

Private Sub Piirre5_8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
' Kysytään adjektiiveja yksikön perusmuodossa perusmuodossa
' Piirre numero _ käyttäjä numero
If KeyCode = vbKeyReturn Then
 If Len(TietoLomake.Piirre5_8) < MinimiMerkit Then
   MsgBox ("Merkkejä tulee olla " & MinimiMerkit & " tai enemmän!")
  Else
   TietoLomake.Os8Tulos = TietoLomake.Os8Nimi & " " & TietoLomake.Os8Omin & " kertoo luonteenpiirteikseen: " & TietoLomake.Piirre1_8 & " " & TietoLomake.Piirre2_8 & " " & TietoLomake.Piirre3_8 & " " & TietoLomake.Piirre4_8 & " " & TietoLomake.Piirre5_8 & ". Ok. Hyvä!"
   TietoLomake.Os8Tulos.Visible = True
   NimetJaOminaisuudet(8, 7) = TietoLomake.Piirre5_8.Value 'Nimi=1, Ominaisuus=3, 3-> piirteet
   Valmista
  End If
End If

End Sub

Private Sub LuoTarinat(x As Integer)
' Tarina 1
Tarinat(1, 1) = " on selvästi puolustaja!"
Tarinat(1, 2) = "Puolustajana " & NimetJaOminaisuudet(x, 1) & " on yleensä |" & NimetJaOminaisuudet(x, 3) & "|, mikä on epätyypillinen luonteenpiirre introverteille persoonille. " _
      & NimetJaOminaisuudet(x, 1) & " on |" & NimetJaOminaisuudet(x, 4) & "|, mutta " & NimetJaOminaisuudet(x, 1) & "  ei käytä sitä tiedon ja nippelitiedon tallentamiseen, vaan painaa mieleensä asioita ihmisistä ja heidän elämästään." & vbNewLine & _
      "Mitä tulee lahjojen antamiseen, " & NimetJaOminaisuudet(x, 1) & " on vertaansa vailla. Hän on |" & NimetJaOminaisuudet(x, 5) & "| ja |" & NimetJaOminaisuudet(x, 6) & "| lisäämään herkkyyttä ilmaisemaan anteliaisuuttaan tavalla, joka koskettaa vastapuolta." & vbNewLine & _
      "Toisinaan |" & NimetJaOminaisuudet(x, 7) & "| tulee esiin vain työntekijöiden keskuudessa, joita puolustaja todella pitää henkilökohtaisena ystävänään, perheen keskuudessa puolustajan tunneilmaisu pääse todella valloilleen. "
' Tarina 2
Tarinat(2, 1) = " on loogikko!"
Tarinat(2, 2) = "Loogikkopersoona rakastaa kaavoja ja analysoi mielellään tarkkaan sanottua ja sanomatta jäänyttä, minkä vuoksi loogikolle ei kannata lähteä valehtemaan! " & _
      "Tämän vuoksi on ironista, että loogikon |" & NimetJaOminaisuudet(x, 3) & "| puoli pitäisi aina ottaa pienellä varauksella." & vbNewLine & _
      "Loogikko " & NimetJaOminaisuudet(x, 1) & " saattaa ulkopuolisesta vaikuttaa olevan jatkuvasti muissa maailmoissa, mutta loogikon |" & NimetJaOminaisuudet(x, 4) & "| luonteenpiirre on lakkaamatonta ja hänen mielensä syytää ideoita heti varhaisesta aamusta." & vbNewLine & _
      "Kun loogikko " & NimetJaOminaisuudet(x, 1) & " on erityisen |" & NimetJaOminaisuudet(x, 5) & "|, keskustelu voi olla epäjohdonmukaista hänen yrittäessään selittää uusimman keksintönsä loogisten päätelmien ketjua ." & vbNewLine & _
      "Loogikko " & NimetJaOminaisuudet(x, 1) & " ei ymmärrä miksi ihmiset valittavat, vaikka hän on |" & NimetJaOminaisuudet(x, 6) & "| ja |" & NimetJaOminaisuudet(x, 7) & "|."
' Tarina 3
Tarinat(3, 1) = " on päällikkö!"
Tarinat(3, 2) = "Päällikkö " & NimetJaOminaisuudet(x, 1) & " on todellinen tehopakkaus ja viljelee itsestään usein elämää suurempaa kuvaa – mitä usein onkin. " & _
     "Päällikön on kuitenkin muistettava, että hänen |" & NimetJaOminaisuudet(x, 3) & "| luonteenpiirre ei ole yksinomaan hänen omien tekojensa ansiota, vaan myös häntä tukevan tiimin aikaansaannosta." & vbNewLine & _
     "Erityisesti työympäristössä päällikkö " & NimetJaOminaisuudet(x, 1) & " haluaa yksinkertaisesti murskata sellaisten henkilöiden herkkyyden, joiden persoonallisuuspiirteisiin kuuluu |" & NimetJaOminaisuudet(x, 4) & "|, |" & NimetJaOminaisuudet(x, 5) & "| tai |" & NimetJaOminaisuudet(x, 6) & "| piirre. " & vbNewLine & _
     "Päällikölle persoonallisuuspiirre |" & NimetJaOminaisuudet(x, 7) & "| liittyy heikkouteen, mikä on asenne, jolla päällikön on helppo saada vihamiehiä."
' Tarina 4
Tarinat(4, 1) = " on väittelijä!"
Tarinat(4, 2) = "Väittelijä " & NimetJaOminaisuudet(x, 1) & " on varsinainen paholaisen asianajaja, joka on |" & NimetJaOminaisuudet(x, 3) & "| ja | " & NimetJaOminaisuudet(x, 4) & "| ja antaa liuskojen lepattaa tuulessa kaikkien nähtävillä. " & vbNewLine & _
      "Väittelijän kyky väittelyyn voi olla ärsyttävä. Tätä asiaa mutkistaa vielä |" & NimetJaOminaisuudet(x, 5) & "| luonteenpiirre, sillä " & NimetJaOminaisuudet(x, 1) & " ei säästele sanoja ja vähät välittää siitä, että häntä pidetään hienotuntoisena tai hyväsydämisenä." & vbNewLine & _
      "Väittelijän älyllinen riippumattomuus ja |" & NimetJaOminaisuudet(x, 6) & "| luonteenpiirre ovat uskomattoman arvokkaita hänen pitäessään ohjaita käsissään, tai jos hänen kykyjään arvostetaan korkealta taholta." & vbNewLine & _
      "Saavuttaessaan tällaisen aseman, väittelijän on hyvä muistaa, että jotta hänen |" & NimetJaOminaisuudet(x, 7) & "| luonteenpiirteensä kantaisi hedelmää, he tarvitsevat avuksi henkilöitä, jotka laittavat palaset kohdilleen."
' Tarina 5
Tarinat(5, 1) = " on sovittelija!"
Tarinat(5, 2) = "Sovittelija " & NimetJaOminaisuudet(x, 1) & " on todellinen idealisti ja pyrkii aina löytämään edes jotain hyvää pahimmistakin ihmisistä tai ikävimmistäkin tapahtumista ja tekemään asioita aikaisempaa paremmin. " & vbNewLine & _
      "Sovittelijan persoonallisuuspiirteet |" & NimetJaOminaisuudet(x, 3) & "| ja |" & NimetJaOminaisuudet(x, 4) & "|  ilmentävät sisäistä liekkiä ja intohimoa." & vbNewLine & _
      "Sovittelijaa ohjaa periaatteet pikemminkin kuin logiikka, innostus tai käytännöllisyys. Sovittelijan tekemistä määrittää persoonallisuuspiirteistä |" & NimetJaOminaisuudet(x, 5) & "|, |" & NimetJaOminaisuudet(x, 6) & "| ja |" & NimetJaOminaisuudet(x, 7) & "|" & vbNewLine
' Tarina 6
Tarinat(6, 1) = " on aktivisti!"
Tarinat(6, 2) = "Aktivistipersoona " & NimetJaOminaisuudet(x, 1) & " on todellinen vapaa sielu. " & NimetJaOminaisuudet(x, 1) & " on usein juhlien sielu, mutta siitä huolimatta hän ei ole niinkään kiinnostunut pelkästä huvista ja hetken mielihyvästä, vaan nauttii sosiaalisista kontakteista ja saavuttamastaan tunneyhteydestä." & vbNewLine & _
       "Aktivistin persoonallisuuspiirre |" & NimetJaOminaisuudet(x, 3) & "| auttaa lukemaan ihmisiä rivien välistä uteliaina ja energisinä." & vbNewLine & _
       "Aktivisti " & NimetJaOminaisuudet(x, 1) & " on pelkäämättömän itsenäinen ja persoonallisuuspiirteet |" & NimetJaOminaisuudet(x, 4) & "|, |" & NimetJaOminaisuudet(x, 5) & "| ja |" & NimetJaOminaisuudet(x, 6) & "| luovat turvallisuutta." & vbNewLine & _
       "Aktivisti " & NimetJaOminaisuudet(x, 1) & " uskoo, että jokaisen pitäisi ottaa aikaa tunteidensa tutkimiseen ja ilmaisemiseen, ja persoonallisuuspiirteen |" & NimetJaOminaisuudet(x, 7) & "| vuoksi se on hänelle luonnollinen keskustelunaihe."
' Tarina 7
Tarinat(7, 1) = " on asianajaja!"
Tarinat(7, 2) = "Asianajaja " & NimetJaOminaisuudet(x, 1) & " omaa sisäsyntyisen idealismin ja moraalin tajun, mutta hän on muita idealisteja päättäväisempi ja määrätietoisempi. " & NimetJaOminaisuudet(x, 1) & " ei ole toimeton uneksija, vaan ihminen, joka kykenee ottamaan konkreettisia askeleita tavoitteen saavuttamiseksi ja pysyvän positiivisen vaikutuksen tekemiseksi." & vbNewLine & _
      "Asianajajalla on hyvin ainutlaatuisia persoonallisuuspiirteitä, kuten |" & NimetJaOminaisuudet(x, 3) & "| ja |" & NimetJaOminaisuudet(x, 4) & "|." & vbNewLine & _
      "Tämän persoonallisuustyypin yksilö on päättäväinen ja voimakastahtoinen, mutta käyttää energiaansa harvoin oman henkilökohtaisen etunsa tavoitteluun." & vbNewLine & _
      "Sen sijaan " & NimetJaOminaisuudet(x, 1) & " toimii luovasti, käyttää mielikuvitusta, on |" & NimetJaOminaisuudet(x, 5) & "| ja |" & NimetJaOminaisuudet(x, 6) & "| luodakseen tasapainoa." & vbNewLine & _
      "Karma ja |" & NimetJaOminaisuudet(x, 7) & "| piirre ovat asianajajalle erittäin houkuttelevia käsitteitä, ja he uskovat siihen, että maailmalla toimivien tyrannien sydämet voitaisiin parhaiten pehmittää rakkaudella ja myötätunnolla."
' Tarina 8
Tarinat(8, 1) = " on protagonisti!"
Tarinat(8, 2) = "Protagonisti " & NimetJaOminaisuudet(x, 1) & "  on luontainen johtaja, täynnä intohimoa ja viehätysvoimaa. Luontaisen itsevarma ja vaikutusvaltainen protagonisti on ylpeä voidessaan opastaa muita yhteistyöhön, joka antaa lähtökohdat oman itsen ja yhteisön kehittämiseen." & vbNewLine & _
      "Ihmiset tuntevat vetoa vahvoihin persoonallisuuksiin, ja " & NimetJaOminaisuudet(x, 1) & "  on |" & NimetJaOminaisuudet(x, 3) & "|, |" & NimetJaOminaisuudet(x, 4) & "| ja " & NimetJaOminaisuudet(x, 5) & "| ja sanoo sanottavansa, jos siihen on aihetta." & vbNewLine & _
      "Protagonistin on luonnollista ja helppoa keskustella muiden kanssa, etenkin kahden kesken, ja hänen |" & NimetJaOminaisuudet(x, 6) & "| luonteenpiirre auttaa häntä tavoittamaan keskustelukumppaninsa joko faktoilla ja logiikalla tai tunteisiin vetoamalla." & vbNewLine & _
      "Protagonistin kiinnostus toisiin on aitoa – lähes epätervettä. Kun " & NimetJaOminaisuudet(x, 1) & " uskoo johonkuhun, hän voi ottaa kyseisen henkilön ongelmat liikaa omakseen ja luottaa häneen liikaa. Onneksi tämä luottamus on itseään toteuttava ennustus, sillä protagonistin luonteenpiirre |" & NimetJaOminaisuudet(x, 7) & "| inspiroi muita kehittymään ihmisenä."

End Sub

Private Sub LuoEsitys()
Dim x As Integer, y As Integer

For x = 1 To Osallistujia
 LuoTarinat (x)
 y = Int((8 - 1 + 1) * Rnd + 1) ' satunnaisluku 1-8 valittavalle tarinalle
  
' Powerpoint sliden rakentaminen valmiiksi
  Set objPresentaion = ActivePresentation
  Set objSlide = objPresentaion.Slides.Item(3 + x) ' Monesko slide
  Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
  ' Arvot luetaan 2-uloitteisesta taulukosta, jossa arvo 1 = nimi, arvo 2 = ominaisuus, arvo 3 = piirre1, arvo 4 = piirre2 jne
  objTextBox.TextFrame.TextRange.Text = NimetJaOminaisuudet(x, 2) & " " & NimetJaOminaisuudet(x, 1) & Tarinat(y, 1) ' Tarina array 1 on otsikon arvo
  Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
  ' Seuraavassa on tarina, johon jutut lisätään
  objTextBox.TextFrame.TextRange.Text = Tarinat(y, 2) ' Tarina array 2 on tarina
  ActivePresentation.Slides(3 + x).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
Next x
 
 
End Sub

Private Sub Valmista()
    TietoLomake.ValmisNappi.Visible = True
    LuoEsitys
End Sub

Private Sub ValmisNappi_Click()

Unload Me

End Sub

