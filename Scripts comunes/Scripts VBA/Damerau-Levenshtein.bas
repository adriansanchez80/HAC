Attribute VB_Name = "Módulo1"
Option Explicit

'algoritmo de Damerau-Levenshtein (distancia de)
'Fuente: https://www.enmimaquinafunciona.com/pregunta/44588/la-formula-para-encontrar-casi-perfecto-coincide-con
'http://stackoverflow.com/questions/4243036/levenshtein-distance-in-excel
'https://x443.wordpress.com/2012/06/25/levenshtein-distance-in-vba/

Private Function Levenshtein(S1 As String, S2 As String)

Dim i As Integer, j As Integer
Dim l1 As Integer, l2 As Integer
Dim d() As Integer
Dim min1 As Integer, min2 As Integer

l1 = Len(S1)
l2 = Len(S2)
ReDim d(l1, l2)
For i = 0 To l1
    d(i, 0) = i
Next
For j = 0 To l2
    d(0, j) = j
Next
For i = 1 To l1
    For j = 1 To l2
        If Mid(S1, i, 1) = Mid(S2, j, 1) Then
            d(i, j) = d(i - 1, j - 1)
        Else
            min1 = d(i - 1, j) + 1
            min2 = d(i, j - 1) + 1
            If min2 < min1 Then
                min1 = min2
            End If
            min2 = d(i - 1, j - 1) + 1
            If min2 < min1 Then
                min1 = min2
            End If
            d(i, j) = min1
        End If
    Next
Next
Levenshtein = d(l1, l2)
End Function

End Function
