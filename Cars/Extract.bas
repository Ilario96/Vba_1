Attribute VB_Name = "Modulo1"
Sub estrai()

Dim cellaincolla As Excel.Range
Dim foglio_orig As Excel.Worksheet
Dim Casa As String
Dim iter_r As Excel.Range
Set iter_r = Range("a2")
Set foglio_orig = Foglio2


    Do Until iter_r.Value = ""
        
        Casa = iter_r.Offset(0, 1).Value
        
    
        Range(iter_r, iter_r.End(xlToRight)).Copy
        Worksheets(Casa).Activate
        Set cellaincolla = Worksheets(Casa).Range("A1000").End(xlUp).Offset(1, 0)
        cellaincolla.Select
        ActiveCell.PasteSpecial
        foglio_orig.Activate
        
        
        
        Set iter_r = iter_r.Offset(1, 0)
          
    Loop
End Sub














