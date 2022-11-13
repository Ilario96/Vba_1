Attribute VB_Name = "Modulo2"
Sub Pulisci()
Attribute Pulisci.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Cella_colorata As Range

' Pulisci Macro
'


    Range("A2", Range("c2").End(xlDown)).Select
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        
        
        
    End With
    
    With Selection.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
    End With
    
    Set Cella_colorata = Range("c2").End(xlDown).Offset(1, 0)
    While Cella_colorata.Interior.Color <> vbWhite
        
        Range("A" & Cella_colorata.Row, "C" & Cella_colorata.Row).Interior.Color = vbWhite
        Set Cella_colorata = Cella_colorata.Offset(1, 0)
        
        
    Wend
    
    
    
    
    
    
    
    
    
    
    
    
End Sub
