Private Sub Worksheet_Change(ByVal Target As Range)
' This sub monitor changes in symptom names and fire subroutines if changed
' Some subs may use END statement, which stops this as well. Then ctrl+r will restart this sub as well.
     If Target.Address = "$A$5" Or Target.Address = "$A$6" Or Target.Address = "$A$7" _
     Or Target.Address = "$A$8" Or Target.Address = "$A$9" Or Target.Address = "$A$10" Then
     ' Execute readdisease sub
        Module1.readisease
 End If
 ' Monitor Consider this result as a false positive change and name change
     If Target.Address = "$B$24" Or Target.Address = "$D$24" Then ' Execute readYesNo sub
        Module2.readYesNo
     End If
     
     If Target.Address = "$E$24" Then ' Execute when disease is selected
        Module2.listDetails
     End If
     
     If Target.Address = "$B$34" Then ' Execute when disease is selected AND person is sure that wants to add item
        Module2.readsureYesNo
     End If
End Sub
