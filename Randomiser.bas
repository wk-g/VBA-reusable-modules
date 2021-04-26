Attribute VB_Name = "Randomiser"
Option Explicit
Function weighted_roulette(aPopFitArr As Variant, Optional mateSize As Integer = 2) As Variant
'PopFitArr should be [[pos,fitness] ...]
'Returns [pos1,pos2]
'arrays cannot be passed byval
Dim PopFitArr As Variant
PopFitArr = aPopFitArr

'pick mates
ReDim selMate(0 To mateSize - 1) As Variant '
Dim tmpMate As Integer
Dim x As Integer

For x = 1 To mateSize
    Dim i As Integer
    Dim fitTotal As Double
    'weighted roulette shouldn't be used here, but this basic GA test was to try out weighted roulette
    fitTotal = 0
    'fitTotal needs to be calculated each time because fitness of chosen ones are set to 0
    For i = LBound(PopFitArr) To UBound(PopFitArr)
        fitTotal = fitTotal + PopFitArr(i)(1)
    Next i
    
    Dim rnd_val As Double, eff_rnd As Double
    rnd_val = Rnd()
    eff_rnd = fitTotal * rnd_val
    
    Dim tmpFitSum As Double, exit_loop As Boolean
    tmpFitSum = 0
    exit_loop = False
    For i = LBound(PopFitArr) To UBound(PopFitArr)
        If exit_loop = False Then
            tmpFitSum = tmpFitSum + PopFitArr(i)(1)
            If tmpFitSum >= eff_rnd Then
                'x-1 because x starts at 1, and selMate is index 0
                selMate(x - 1) = PopFitArr(i)(0)
                'don't let the same chromosome mate with itself
                'set fitness of chosen to 0 - weighted roulette only works with non-zero fitness functions
                PopFitArr(i)(1) = 0
                exit_loop = True
            End If
        End If
    Next i
Next x

weighted_roulette = selMate
End Function

