Attribute VB_Name = "Modulo2"
Sub puliscifogliConIlDoLoop()

Foglio3.Activate
    Do
    Range("a2:e6689").Clear
    If ActiveSheet.Next Is Nothing Then Exit Do
    
    
    
    ActiveSheet.Next.Select

    Loop


Foglio2.Activate



End Sub


