' (C) 2022 Jari Hiltunen - MIT Licence
'
' Programmed for Laurea University of Applied Science to be used for course
' V1413-3033 Luovuus ja toiminnallisuus asiakastyössä.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
Option Compare Text
Option Explicit
Public Osallistujia As Integer

Sub LomakeTayta()
' Nostetaan Tietolomake ylimmäiseksi ja kysellään tiedot
' Ensimmäisen täytettävän sliden järjestysnumero on oltava 4!
' Älä manipuloi lomakkeen tai kenttien nimiä. Ne on osana koodia!
'
On Error GoTo Whoa

Application.WindowState = ppWindowMinimized ' ppWindowMaximized

If SlideShowWindows.Count = 0 Then
  MsgBox ("Käynnistä Slideshow F5-näppäimellä ensin!")
  Exit Sub
End If

NollaaLomake
Osallistujia = 0
TietoLomake.Show

With TietoLomake.Os1Nimi
.SetFocus
.SelStart = 0
End With

If Osallistujia > 0 Then
    With SlideShowWindows(1).View
        .GotoSlide 4
    End With
Else
   MsgBox ("Osallistujia ei ollut. Moikka.")
End If

LetsContinue:
    Exit Sub

Whoa:
    MsgBox ("Pahus, tuli virhe: " & Err.Description)
    Resume LetsContinue


End Sub

Private Sub NollaaLomake()
' Tämä proseduuri asettaa lomakkeen ja slidet 4-12 alkutilanteeseen.

Dim objPresentaion As Presentation
Dim objSlide As Slide
Dim objTextBox As Shape
Dim x As Integer

For x = 4 To 12
 ' Nollataan tiedot
 Set objPresentaion = ActivePresentation
 Set objSlide = objPresentaion.Slides.Item(x) ' Monesko slide
 Set objTextBox = objSlide.Shapes.Item(1) ' Tämä on ylin otsikkorivi
 objTextBox.TextFrame.TextRange.Text = "Tyhjä "
 Set objTextBox = objSlide.Shapes.Item(2) ' Tämä on tarinan alue
 ' Seuraavassa on tarina, johon jutut lisätään
 objTextBox.TextFrame.TextRange.Text = _
    "Tyhjä " & vbNewLine & _
    "Tyhjä " & vbNewLine & _
    "Tyhjä " & vbNewLine & _
    "Tyhjä " & vbNewLine & _
    "Tyhjä "
 ActivePresentation.Slides(x).SlideShowTransition.Hidden = msoFalse ' Näytetään slide
 ' Slide valmis
Next x

' Piilotetaan slidet 4-12
For x = 4 To 12
    ActivePresentation.Slides(x).SlideShowTransition.Hidden = msoTrue 'Muuta msoFalse jos haluat näyttää
Next x


' Nollataan lomake alkuasetuksille
TietoLomake.Os2Nimi.Visible = False
TietoLomake.Os3Nimi.Visible = False
TietoLomake.Os4Nimi.Visible = False
TietoLomake.Os5Nimi.Visible = False
TietoLomake.Os6Nimi.Visible = False
TietoLomake.Os7Nimi.Visible = False
TietoLomake.Os8Nimi.Visible = False
' Ominaisuudet
TietoLomake.Os1Omin.Visible = False
TietoLomake.Os2Omin.Visible = False
TietoLomake.Os3Omin.Visible = False
TietoLomake.Os4Omin.Visible = False
TietoLomake.Os5Omin.Visible = False
TietoLomake.Os6Omin.Visible = False
TietoLomake.Os7Omin.Visible = False
TietoLomake.Os8Omin.Visible = False
' Labelit nollille
TietoLomake.Os2l.Visible = False
TietoLomake.Os3l.Visible = False
TietoLomake.Os4l.Visible = False
TietoLomake.Os5l.Visible = False
TietoLomake.Os6l.Visible = False
TietoLomake.Os7l.Visible = False
TietoLomake.Os8l.Visible = False
TietoLomake.Os1l2.Visible = False
TietoLomake.Os2l2.Visible = False
TietoLomake.Os3l2.Visible = False
TietoLomake.Os4l2.Visible = False
TietoLomake.Os5l2.Visible = False
TietoLomake.Os6l2.Visible = False
TietoLomake.Os7l2.Visible = False
TietoLomake.Os8l2.Visible = False
TietoLomake.Osallistuu.Visible = False

' Piilotetaan tuloskentät kunnes tiedot on täytetty
TietoLomake.Os1Tulos.Visible = False
TietoLomake.Os2Tulos.Visible = False
TietoLomake.Os3Tulos.Visible = False
TietoLomake.Os4Tulos.Visible = False
TietoLomake.Os5Tulos.Visible = False
TietoLomake.Os6Tulos.Visible = False
TietoLomake.Os7Tulos.Visible = False
TietoLomake.Os8Tulos.Visible = False
' Piilotetaan piirrekentät kunnes nimi on kirjoitettu
' Henkilö #1
TietoLomake.Piirre1_1.Visible = False
TietoLomake.Piirre2_1.Visible = False
TietoLomake.Piirre3_1.Visible = False
TietoLomake.Piirre4_1.Visible = False
TietoLomake.Piirre5_1.Visible = False
' Henkilö #2
TietoLomake.Piirre1_2.Visible = False
TietoLomake.Piirre2_2.Visible = False
TietoLomake.Piirre3_2.Visible = False
TietoLomake.Piirre4_2.Visible = False
TietoLomake.Piirre5_2.Visible = False
' Henkilö #3
TietoLomake.Piirre1_3.Visible = False
TietoLomake.Piirre2_3.Visible = False
TietoLomake.Piirre3_3.Visible = False
TietoLomake.Piirre4_3.Visible = False
TietoLomake.Piirre5_3.Visible = False
' Henkilö #4
TietoLomake.Piirre1_4.Visible = False
TietoLomake.Piirre2_4.Visible = False
TietoLomake.Piirre3_4.Visible = False
TietoLomake.Piirre4_4.Visible = False
TietoLomake.Piirre5_4.Visible = False
' Henkilö #5
TietoLomake.Piirre1_5.Visible = False
TietoLomake.Piirre2_5.Visible = False
TietoLomake.Piirre3_5.Visible = False
TietoLomake.Piirre4_5.Visible = False
TietoLomake.Piirre5_5.Visible = False
' Henkilö #6
TietoLomake.Piirre1_6.Visible = False
TietoLomake.Piirre2_6.Visible = False
TietoLomake.Piirre3_6.Visible = False
TietoLomake.Piirre4_6.Visible = False
TietoLomake.Piirre5_6.Visible = False
' Henkilö #7
TietoLomake.Piirre1_7.Visible = False
TietoLomake.Piirre2_7.Visible = False
TietoLomake.Piirre3_7.Visible = False
TietoLomake.Piirre4_7.Visible = False
TietoLomake.Piirre5_7.Visible = False
' Henkilö #8
TietoLomake.Piirre1_8.Visible = False
TietoLomake.Piirre2_8.Visible = False
TietoLomake.Piirre3_8.Visible = False
TietoLomake.Piirre4_8.Visible = False
TietoLomake.Piirre5_8.Visible = False

End Sub
