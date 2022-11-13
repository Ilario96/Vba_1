Attribute VB_Name = "Modulo1"
Option Explicit

Sub Formattazione_con_ciclo_for_next()

Dim i As Long


    For i = 1 To Range("A2", Range("A2").End(xlDown)).Count
    
     If Range("C" & i + 1) < 100 Then
        Range("A" & i + 1, "C" & i + 1).Interior.Color = Range("i2").Interior.Color
        Else
        If Range("C" & i + 1) < 200 Then
            Range("A" & i + 1, "C" & i + 1).Interior.Color = Range("i3").Interior.Color
            Else
            If Range("C" & i + 1) < 300 Then
                Range("A" & i + 1, "C" & i + 1).Interior.Color = Range("i4").Interior.Color
                Else
                If Range("C" & i + 1) < 400 Then
                    Range("A" & i + 1, "C" & i + 1).Interior.Color = Range("i5").Interior.Color
                Else
                Range("A" & i + 1, "C" & i + 1).Interior.Color = Range("i6").Interior.Color
    
                End If
            End If
        End If
    End If
        
       
    
    Next i




End Sub

