Attribute VB_Name = "ArrayHelp"
Option Explicit
'Module includes algorithms for basic tasks (e.g. sorting an array)
Public Sub arrQuicksort(vArray As Variant, arrLBound As Long, arrUBound As Long, Optional nestArraypos As Variant)
'Apparently ArrayList is a thing, but it requires an external library
'nestArraypos declared as Variant to allow use of IsMissing()
'Quicksort prob slower than simpler sorts (e.g. bubble) after first sort, but only QS implemented for now
'Edited from: https://wellsr.com/vba/2018/excel/vba-quicksort-macro-to-sort-arrays-fast/
'sorts in-place, unstable
Dim pivotVal As Variant
Dim vSwap    As Variant
Dim tmpLow   As Long
Dim tmpHi    As Long
 
tmpLow = arrLBound
tmpHi = arrUBound
'pivotVal is using middle pos of Array
If IsMissing(nestArraypos) Then
    pivotVal = vArray((arrLBound + arrUBound) \ 2)
Else
    pivotVal = vArray((arrLBound + arrUBound) \ 2)(nestArraypos)
End If

If IsMissing(nestArraypos) Then
'non-nested array
    While (tmpLow <= tmpHi) 'divide
        While (vArray(tmpLow) < pivotVal And tmpLow < arrUBound)
        tmpLow = tmpLow + 1
        Wend
  
        While (pivotVal < vArray(tmpHi) And tmpHi > arrLBound)
        tmpHi = tmpHi - 1
        Wend
 
        If (tmpLow <= tmpHi) Then
        vSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = vSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
        End If
    Wend
 
    If (arrLBound < tmpHi) Then arrQuicksort vArray, arrLBound, tmpHi 'conquer
    If (tmpLow < arrUBound) Then arrQuicksort vArray, tmpLow, arrUBound 'conquer
Else
'nested array
    While (tmpLow <= tmpHi) 'divide
        While (vArray(tmpLow)(nestArraypos) < pivotVal And tmpLow < arrUBound)
        tmpLow = tmpLow + 1
        Wend
  
        While (pivotVal < vArray(tmpHi)(nestArraypos) And tmpHi > arrLBound)
        tmpHi = tmpHi - 1
        Wend
 
        If (tmpLow <= tmpHi) Then
        vSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = vSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
        End If
    Wend
 
    If (arrLBound < tmpHi) Then arrQuicksort vArray, arrLBound, tmpHi, nestArraypos 'conquer
    If (tmpLow < arrUBound) Then arrQuicksort vArray, tmpLow, arrUBound, nestArraypos 'conquer
End If
End Sub
Public Function CountOccurance(vArray As Variant, occurance As Variant) As Variant
'occurance can be an array. Will return as arr(occ1,occ2,occ3...) As Integer array

If IsArray(occurance) = True Then
    Dim occLBound As Integer
    Dim occUBound As Integer
    occLBound = LBound(occurance)
    occUBound = UBound(occurance)
    ReDim arrFinalCount(occLBound To occUBound) As Integer
    
    Dim i As Integer, nowOcc As Variant, tmpCount As Integer
    For i = occLBound To occUBound
        nowOcc = occurance(i)
        tmpCount = CountOccurance_Helper(vArray, nowOcc)
        arrFinalCount(i) = tmpCount
    Next i
    CountOccurance = arrFinalCount
ElseIf IsArray(occurance) = False Then
    Dim FinalCount As Integer
    FinalCount = CountOccurance_Helper(vArray, occurance)
    CountOccurance = FinalCount
End If


End Function
Private Function CountOccurance_Helper(vArray As Variant, occurance As Variant) As Integer
Dim arrLBound As Integer
Dim arrUBound As Integer
arrLBound = LBound(vArray)
arrUBound = UBound(vArray)

Dim i As Integer
Dim nowCount As Integer
nowCount = 0
For i = arrLBound To arrUBound
    If vArray(i) = occurance Then
        nowCount = nowCount + 1
    End If
Next i

CountOccurance_Helper = nowCount
End Function
Function check_unique(checkArray As Variant, newInput As Variant)
'check if newInput already exists in checkArray
Dim i As Variant
Dim x As Boolean

x = True
For Each i In checkArray
    If newInput = i Then
        x = False
    End If
Next i

check_unique = x
End Function
Function check_equal(Array_1 As Variant, Array_2 As Variant) As Boolean
'check if 2 arrays are equal
'if lbound/ubound/len are different, no point continuing
If LBound(Array_1) <> LBound(Array_2) Or UBound(Array_1) <> UBound(Array_2) Then
    check_equal = False
    Exit Function
End If

Dim i As Integer

'check for differences and exit function if needed - faster for false results
For i = LBound(Array_1) To UBound(Array_1)
    'if one if the items is an object/array and the other isn't, equal = false
    If IsObject(Array_1(i)) <> IsObject(Array_2(i)) Or _
    IsArray(Array_1(i)) <> IsArray(Array_2(i)) Then
        check_equal = False
        Exit Function
    End If
    
    'If objects, check if they REFERENCE the same object
    'does not check if objects are equal in content
    If IsObject(Array_1(i)) = True And IsObject(Array_2(i)) = True Then
        If (Array_1(i) Is Array_2(i)) = False Then
            check_equal = False
            Exit Function
        End If
    'handle arrays, recurse on this function
    ElseIf IsArray(Array_1(i)) = True And IsArray(Array_2(i)) = True Then
        If check_equal(Array_1(i), Array_2(i)) = False Then
            check_equal = False
            Exit Function
        End If
    'handle values
    Else
        If Array_1(i) <> Array_2(i) Then
            check_equal = False
            Exit Function
        End If
    End If
    
    
Next i

'if function not exited at the end, then arrays are equal
check_equal = True
End Function
Function printable_arr(arr As Variant)
'returns a printable array [a,b,c...]. For 1D arrays only
Dim i As Integer
Dim tmpString As String
tmpString = "[" & arr(LBound(arr))
For i = LBound(arr) + 1 To UBound(arr)
    tmpString = tmpString & " , " & arr(i)
Next i

tmpString = tmpString & "]"
printable_arr = tmpString
End Function
