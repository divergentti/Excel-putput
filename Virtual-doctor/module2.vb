' This is the learning module
Option Explicit

Sub readYesNo()
' This sub is triggered by value change in range AnswerYesNo at User sheet.
' Trigger is activated at User sheet macro.
On Error GoTo Whoa 'Disable this if you would like to debug
Application.EnableEvents = False ' Enable this in case you would like to debug
Dim answerSelection As String ' for the answer
Dim correctDisease As String 'User inputted disease name for listing
Dim selectedDisease As String 'User selected disease from drop list
Dim diseaseItems() As String ' Found disease names
Dim i As Integer 'Iterator

answerSelection = ThisWorkbook.Sheets("User").Range("AnswerYesNo").Value ' Either Y or N for the question false positive

If answerSelection = "Y" And ThisWorkbook.Sheets("User").Range("correctDiseaseName").Value <> "" Then
    ' Predicted disease is wrong selected and partial or full name written.
    correctDisease = ThisWorkbook.Sheets("User").Range("correctDiseaseName").Value
    ' Build a list of diseases matching for the search criteria
    diseaseItems = findDiseaseName(correctDisease) ' Here we receive an array of diseases
    'Display selection box at the user sheet
    ' THIS CODE BELOW BRAKES EXCEL FILE TOTALLY! IT MUST BE EXCEL BUG! REPORTED TO MICROSFOT!
    ' ----- SNIP ----
    'For i = 0 To (UBound(diseaseItems, 1))
    '  myList = myList & diseaseItems(i) & ","
    'Next i
    '  myList = Mid(myList, 1, Len(myList) - 1)
    'With ThisWorkbook.Sheets("User").Range("E24").Validation
    '  .Delete
    '  .Add _
    '   Type:=xlValidateList, _
    '   AlertStyle:=xlValidAlertStop, _
    '   Operator:=xlBetween, _
    '   Formula1:=myList
    '   .IgnoreBlank = True
    '   .InCellDropdown = True
    '   .InputTitle = ""
    '   .ErrorTitle = ""
    '   .InputMessage = ""
    '   .ErrorMessage = ""
    '   .ShowInput = True
    '   .ShowError = True
    'End With
    ' ----- SNIP ----
    
    'Due to Excel bug, let's try to make another type of validation
    Worksheets("Learning").Activate
    
    'First we clear old list items
    ThisWorkbook.Sheets("Learning").Range("N10:N10000").Clear ' Shall be plenty
    deleteRange ("correctDiseaseList") 'Delete old range
    
    ' Now list items into sheet
    For i = 0 To (UBound(diseaseItems, 1))
      ThisWorkbook.Sheets("Learning").Cells(i + 10, 14).Value = diseaseItems(i) 'Start from N10 range
    Next i
    
    ' Make selection and name it
     ThisWorkbook.Sheets("Learning").Cells(10, 14).Select 'Make selection and expand it
     Selection.Resize(i).Select
     Selection.Name = "correctDiseaseList"
    
    Worksheets("User").Activate
    
    With ThisWorkbook.Sheets("User").Range("E24").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Operator:=xlBetween, Formula1:="=correctDiseaseList"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
    End With
      
End If

' Move selection to list box
ThisWorkbook.Sheets("User").Range("selectedDiseaseName").Select

' Completed, next monitored event is if person makes selection from the list box

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue
    

End Sub

Public Function findDiseaseName(ByVal correctDisease As String) As String()
' On Error GoTo Whoa
' This function will search for PARTIAL text given and return address of the hit
' Input is the disease name as a string
' Output is array of of the diseases matching to searches disease
On Error GoTo Whoa 'Disable this if you would like to debug
Application.EnableEvents = True ' Enable this in case you would like to debug

Dim foundDisease As Range 'Found word address
Dim firstDisease As String 'Where we see first search hit
Dim ws As Worksheet
Dim lRow As Long 'Range definition row number
Dim aCell As Range, range2 As Range 'Range definition
Dim disName() As String 'We will redim this every time
Dim i As Integer
Set ws = ThisWorkbook.Sheets("db") ' Select db sheet
' Check if db sheet has disease column and read range of the data

With ws
        ' Find the cell which has the name - now just A1 (change this if you modify database or make this dynamic)
        Set aCell = .Range("A1").Find("diseasename")
        'If the cell is found
        If Not aCell Is Nothing Then
            'Get the last row in that column and check if the last row is > 1
            lRow = .Range(Split(.Cells(, aCell.Column).Address, "$")(1) & .rows.Count).End(xlUp).row
            If lRow > 1 Then
                ' Set Range
                Set range2 = .Range(aCell.Offset(1), .Cells(lRow, aCell.Column))
            End If
        End If
End With
 
i = 0 ' Reset counter value
 
With range2 'This is range of the diseases ~column
' Here we define what we are trying to find

Set foundDisease = range2.Find(what:=correctDisease, LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)

  If Not foundDisease Is Nothing Then
      ' First address is referral for subsequent searches
      firstDisease = foundDisease.Value
   Do
          ReDim Preserve disName(i) ' Resize the array
          disName(i) = foundDisease.Value 'Array of disease names
            i = i + 1 'Increase counter
       Set foundDisease = .FindNext(foundDisease)
       If foundDisease Is Nothing Then
         MsgBox "!!Not found" ' Check why
       End If
     Loop While Not foundDisease Is Nothing And foundDisease.Value <> firstDisease
   End If

End With

' Inform user if nothing found
If i > 0 Then
   findDiseaseName = disName ' Return an array of found names
   Else 'Return error
      MsgBox "Not Found. Try another wording and press Ctrl+r"
       ' Reason I do this this way is fact that otherwise function will return null, which is not properly handled in calling subs
     End 'Stopping all macros
End If



LetsContinue:
    Application.EnableEvents = True
    Exit Function

Whoa:
    MsgBox Err.Description
    Resume LetsContinue
   

End Function

Public Function deleteRange(ByVal rangename As String)
' This function will delete named range

Dim raName As String

On Error Resume Next
raName = rangename
Range(raName).Select
  DoEvents
    If Err.Number = "1004" Then
    ' Does not exist, we do not care
    Else
        Range(raName).Delete
    ' Deleted
    End If
End Function

Sub listDetails()
' This sub will be activated when user makes selection from the consider result false positive list box
On Error GoTo Whoa 'Disable this if you would like to debug

Application.EnableEvents = True ' Enable this in case you would like to debug
 
 'Check what user has selected item
Dim disAddress As String 'Address of the disease
Dim selectedDisease As String 'User selected disease from drop list
Dim col, row 'For columsn and rows

If ThisWorkbook.Sheets("User").Range("selectedDiseaseName") <> "" Then
          selectedDisease = ThisWorkbook.Sheets("User").Range("selectedDiseaseName").Value
    'Bring in information of the disease starting from B25
    ' Colour of the range
    Range("B25:E33").Interior.ColorIndex = 7
     ' Let we see which address it came from
      disAddress = findDiseaseAddress(selectedDisease)
      ' Split
      col = Split(disAddress, "$")(1)
      row = Split(disAddress, "$")(2)
    '  Disease reference
      ThisWorkbook.Sheets("User").Range("B25").Value = ThisWorkbook.Sheets("db").Range("B" & row).Value
    'Summary
      ThisWorkbook.Sheets("User").Range("B26").Value = ThisWorkbook.Sheets("db").Range("c" & row).Value
    'Causes
      ThisWorkbook.Sheets("User").Range("B27").Value = ThisWorkbook.Sheets("db").Range("d" & row).Value
    'Full list of symptoms
      ThisWorkbook.Sheets("User").Range("B28").Value = ThisWorkbook.Sheets("db").Range("e" & row).Value
    'Exams and tests
      ThisWorkbook.Sheets("User").Range("B29").Value = ThisWorkbook.Sheets("db").Range("f" & row).Value
    'Treatment
      ThisWorkbook.Sheets("User").Range("B30").Value = ThisWorkbook.Sheets("db").Range("g" & row).Value
    'Prognosis
      ThisWorkbook.Sheets("User").Range("B31").Value = ThisWorkbook.Sheets("db").Range("h" & row).Value
    'Possible complications
      ThisWorkbook.Sheets("User").Range("B32").Value = ThisWorkbook.Sheets("db").Range("i" & row).Value
    'Alternative Names
      ThisWorkbook.Sheets("User").Range("B33").Value = ThisWorkbook.Sheets("db").Range("j" & row).Value

End If


LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue


End Sub

Sub readsureYesNo()
' This is a final decision to add disease and vectors to the Learning sheet
' We should have disease selected in the list item E24.

Dim lastRow As Long ' For lastrow handling
Dim lastCol As Long ' For lastcolumn
Dim disNumber As Long ' Item #
Dim selectedDisease As String 'User selected disease from drop list

' Double checking we have all set
If ThisWorkbook.Sheets("User").Range("sureYesNo").Value = "Y" And ThisWorkbook.Sheets("User").Range("selectedDiseaseName").Value <> "" Then


disNumber = ThisWorkbook.Sheets("Learning").Range("I8").Value 'Here we have value which number is biggest

 If disNumber > 0 Then 'There seems to be items
' If there is already item listed, then do this
    Worksheets("Learning").Activate
    ' add one to the disease number and to row number
    disNumber = disNumber + 1
    ' Fill in items
    
    ' Highest item #
    ThisWorkbook.Sheets("Learning").Range("J" & disNumber).Value = disNumber - 9 ' Reduce 9 rows from disease numbers (10th = 0)
    ' Disease name
    ThisWorkbook.Sheets("Learning").Range("A" & disNumber).Value = ThisWorkbook.Sheets("User").Range("B13").Value
    ' Vector 1
    ThisWorkbook.Sheets("Learning").Range("B" & disNumber).Value = ThisWorkbook.Sheets("User").Range("symptom1").Value
    ' Vector 2
    ThisWorkbook.Sheets("Learning").Range("C" & disNumber).Value = ThisWorkbook.Sheets("User").Range("symptom2").Value
    ' Vector 3
    ThisWorkbook.Sheets("Learning").Range("D" & disNumber).Value = ThisWorkbook.Sheets("User").Range("symptom3").Value
    ' Vector 4
    ThisWorkbook.Sheets("Learning").Range("E" & disNumber).Value = ThisWorkbook.Sheets("User").Range("symptom4").Value
    ' Vector 5
    ThisWorkbook.Sheets("Learning").Range("F" & disNumber).Value = ThisWorkbook.Sheets("User").Range("symptom5").Value
    ' Vector 6
    ThisWorkbook.Sheets("Learning").Range("G" & disNumber).Value = ThisWorkbook.Sheets("User").Range("symptom6").Value
    ' Probalitity %
    ThisWorkbook.Sheets("Learning").Range("H" & disNumber).Value = ThisWorkbook.Sheets("User").Range("B23").Value
    ' % of maximums
    ThisWorkbook.Sheets("Learning").Range("I" & disNumber).Value = ThisWorkbook.Sheets("User").Range("D23").Value
    ' User selected correct disease
    ThisWorkbook.Sheets("Learning").Range("K" & disNumber).Value = ThisWorkbook.Sheets("User").Range("selectedDiseaseName")
   
    Worksheets("User").Activate
    
    ThisWorkbook.Sheets("User").Range("AnswerYesNo").Value = "N" 'Reset setting back to No
    ThisWorkbook.Sheets("User").Range("sureYesNo").Value = "N" 'Reset setting back to No
    ThisWorkbook.Sheets("User").Range("correctDiseaseName").Value = "" 'Reset setting back to No
    ThisWorkbook.Sheets("User").Range("selectedDiseaseName").Value = "" 'Reset setting back to No
    ThisWorkbook.Sheets("User").Range("B25:E33").Interior.ColorIndex = 0
    ThisWorkbook.Sheets("User").Range("B25:E33").Clear


  Else ' Let's add first

    Worksheets("Learning").Activate
    ' Number 1 to count number
    ThisWorkbook.Sheets("Learning").Range("J10").Value = 1
    ' Disease name
    ThisWorkbook.Sheets("Learning").Range("A10").Value = ThisWorkbook.Sheets("User").Range("B13").Value
    ' Vector 1
    ThisWorkbook.Sheets("Learning").Range("B10").Value = ThisWorkbook.Sheets("User").Range("symptom1").Value
    ' Vector 2
    ThisWorkbook.Sheets("Learning").Range("C10").Value = ThisWorkbook.Sheets("User").Range("symptom2").Value
    ' Vector 3
    ThisWorkbook.Sheets("Learning").Range("D10").Value = ThisWorkbook.Sheets("User").Range("symptom3").Value
    ' Vector 4
    ThisWorkbook.Sheets("Learning").Range("E10").Value = ThisWorkbook.Sheets("User").Range("symptom4").Value
    ' Vector 5
    ThisWorkbook.Sheets("Learning").Range("F10").Value = ThisWorkbook.Sheets("User").Range("symptom5").Value
    ' Vector 6
    ThisWorkbook.Sheets("Learning").Range("G10").Value = ThisWorkbook.Sheets("User").Range("symptom6").Value
    ' Probalitity %
    ThisWorkbook.Sheets("Learning").Range("H10").Value = ThisWorkbook.Sheets("User").Range("B23").Value
    ' % of maximums
    ThisWorkbook.Sheets("Learning").Range("I10").Value = ThisWorkbook.Sheets("User").Range("D23").Value
    ' User selected correct disease
    ThisWorkbook.Sheets("Learning").Range("K10").Value = ThisWorkbook.Sheets("User").Range("selectedDiseaseName")
        
   ' ThisWorkbook.Sheets("Learning").Range("A10:K10").Select 'Make selection and expand it
   ' Selection.Name = "supervisedList"
    
    Worksheets("User").Activate
    
    ThisWorkbook.Sheets("User").Range("AnswerYesNo").Value = "N" 'Reset setting back to No
    ThisWorkbook.Sheets("User").Range("sureYesNo").Value = "N" 'Reset setting back to No
    ThisWorkbook.Sheets("User").Range("correctDiseaseName").Value = "" 'Reset setting back to No
    ThisWorkbook.Sheets("User").Range("selectedDiseaseName").Value = "" 'Reset setting back to No
    ThisWorkbook.Sheets("User").Range("B25:E33").Interior.ColorIndex = 0
    ThisWorkbook.Sheets("User").Range("B25:E33").Clear

   
  End If

  
End If ' First IF

LetsContinue:
    Application.EnableEvents = True
    Exit Sub

Whoa:
    MsgBox Err.Description
    Resume LetsContinue



End Sub
