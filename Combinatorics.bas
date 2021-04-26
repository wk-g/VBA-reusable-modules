Attribute VB_Name = "Combinatorics"
Option Explicit
Dim rColl As Collection
Function gen_permutations(slots As Integer, pool As Integer, Optional nowSlot As Integer = 0, _
Optional tmpResult As Variant, Optional outputFinal As Boolean = True, _
Optional startIndex As Integer = 0, Optional poolShift As Integer = 0, _
Optional repAllowed As Boolean = False) As Variant
'Returns an array of all arrays of permutation
Dim slotUBound As Integer
slotUBound = slots + startIndex - 1

'If tmpResult dimensions don't match output, redim tmpResult
If IsMissing(tmpResult) = False Then
    If LBound(tmpResult) <> startIndex Or UBound(tmpResult) <> slotUBound Then
        ReDim tmpResult(startIndex To slotUBound) As Integer
    End If
Else
    ReDim tmpResult(startIndex To slotUBound) As Integer
End If

'mem_tmpResult holds the original value of tmpResult from args
Dim mem_tmpResult() As Integer
mem_tmpResult = tmpResult

'If outputFinal, then function should output collection
'Clear the module-level Collection so that data can be stored inside
If outputFinal = True Then
    Set rColl = New Collection
End If

Dim i As Integer
Dim skip_i As Boolean
'Pool will be 1,2,3... Poolshifts (e.g. 0,1,2,3) have to be implemented later as _
arrays are intialized with value = 0
For i = 1 To pool
    tmpResult = mem_tmpResult
    If repAllowed = False Then
        If ArrayHelp.check_unique(tmpResult, i) = True Then
        'no repeats
            skip_i = False
            tmpResult(nowSlot) = i
        ElseIf ArrayHelp.check_unique(tmpResult, i) = False Then
            skip_i = True
        End If
    ElseIf repAllowed = True Then
        skip_i = False
        tmpResult(nowSlot) = i
    End If

    If skip_i = False Then
        If nowSlot = slotUBound Then
            rColl.Add tmpResult
        ElseIf nowSlot < slotUBound Then
            tmpResult = gen_permutations(slots, pool, nowSlot + 1, tmpResult, False, startIndex, poolShift, repAllowed)
        End If
    End If
Next i

Dim x As Variant, y As Variant
'rColl is always a collection of arrays: Coll(arr(),arr()...)
If outputFinal = True Then
    If poolShift <> 0 Then
        'Need a separate collection to store the shifted arrays _
        you cannot change arrays added to collection, a copy of the array will be passed
        Dim scoll As Collection
        Set scoll = New Collection
        ReDim nArr(startIndex To slotUBound) As Integer
        For Each x In rColl
            For y = LBound(x) To UBound(x)
                nArr(y) = x(y) + poolShift
            Next y
            scoll.Add nArr
        Next
        Set gen_permutations = scoll
    Else
        Set gen_permutations = rColl
    End If
ElseIf outputFinal = False Then
    'if outputFinal is false, means it is a recursive call. return tmpResult for later use.
    gen_permutations = tmpResult
End If
End Function
Function gen_partition_odds(goal As Integer, groups As Integer) As Collection
'Returns Coll(arr(partarr[1,0,...],odds))
'Trying to hit the goal is the same as generating all permutations(with repetiton) _
then calculating the number of times each number appears.
'e.g. Goal of 3 with 2 groups can be [3,0], [2,1], [1,2], or [0,3]
'ways are: [1,1,1],[1,1,2],[1,2,1],[1,2,2],[2,1,1],[2,1,2],[2,2,1], [2,2,2] = perm(3slots,2pool) _
thus [3,0]: 1 way ; [2,1]: 3 ways ; [1,2]: 3 ways ; [0,3]: 1 way
If goal = 0 Then
    'edge case of 0 should return Coll(arr(partarr(0,0,0...), 1))
    'Need to be Variant array for Join to work
    ReDim zeroArr(0 To groups - 1) As Variant
    Dim i As Integer
    For i = 0 To UBound(zeroArr)
        zeroArr(i) = 0
    Next i
    Dim finalColl As Collection
    Set finalColl = New Collection
    finalColl.Add Array(zeroArr, 1)
Else
    Dim permColl As Collection
    Set permColl = gen_permutations(goal, groups, , , True, , -1, True)

    'Need to be Variant array for Join to work
    ReDim joinArr(0 To groups - 1) As Variant
    Dim permItem As Variant, tmpCount As Integer
    Dim dkeyStr As String
    Dim countDict As Scripting.Dictionary
    Set countDict = New Scripting.Dictionary
    
    For Each permItem In permColl
        For i = 0 To groups - 1
            tmpCount = ArrayHelp.CountOccurance(permItem, i)
            joinArr(i) = tmpCount
        Next i
            dkeyStr = Join(joinArr)
            countDict(dkeyStr) = countDict(dkeyStr) + 1
    Next
    
    Dim dkey As Variant
    Dim tPerm As Integer
    tPerm = permColl.Count

    Dim inArr(0 To 1) As Variant
    Dim nowOdds As Double
    Set finalColl = New Collection
    
    ReDim tmpArr(0 To groups - 1) As String
    For Each dkey In countDict
        tmpArr = Split(dkey)
        nowOdds = countDict(dkey) / tPerm
        inArr(0) = tmpArr
        inArr(1) = nowOdds
        finalColl.Add inArr
    Next
End If
Set gen_partition_odds = finalColl
End Function
Function part_odds_dict(GoalArray As Variant, groups As Integer) As Scripting.Dictionary
'returns Goal -> Coll(partition odds) for all goals in array
Dim finalDict As Scripting.Dictionary
Set finalDict = New Scripting.Dictionary

Dim i As Integer, nowGoal As Integer, nowColl As Collection
For i = LBound(GoalArray) To UBound(GoalArray)
    nowGoal = GoalArray(i)
    Set nowColl = gen_partition_odds(nowGoal, groups)
    finalDict.Add nowGoal, nowColl
Next i
Set part_odds_dict = finalDict
End Function
Private Function Arr_Combis_Helper(Arr1 As Variant, Arr2 As Variant) As Variant
'returns an array of arrays: arr(arr(1,2),arr(1,3)...)
Dim Arr1Len As Integer, Arr2Len As Integer, fArrLen As Integer
Arr1Len = UBound(Arr1) - LBound(Arr1) + 1
Arr2Len = UBound(Arr2) - LBound(Arr2) + 1
fArrLen = Arr1Len * Arr2Len
ReDim fArr(0 To fArrLen - 1)

Dim fArrInd As Integer
fArrInd = 0
Dim i As Integer, j As Integer
For i = LBound(Arr1) To UBound(Arr1)
    ReDim tmpArr(0 To 1) As Variant
    tmpArr(0) = Arr1(i)
    For j = LBound(Arr2) To UBound(Arr2)
        tmpArr(1) = Arr2(j)
        fArr(fArrInd) = tmpArr
        fArrInd = fArrInd + 1
    Next j
Next i
Arr_Combis_Helper = fArr
End Function
Function Arr_Combis(ArrOfArr As Variant) As Variant
'Returns an array
Dim fArr() As Variant

Dim i As Integer
'collections are 1-based
For i = LBound(ArrOfArr) To UBound(ArrOfArr) - 1
    If i = 0 Then
        'initial step
         fArr = Arr_Combis_Helper(ArrOfArr(i), ArrOfArr(i + 1))
    Else
        fArr = Arr_Combis_Helper(fArr, ArrOfArr(i + 1))
        'at this point fArr will be Arr(Combi(1,2,..,i),i+1)
        'unpack combi to put i + 1 into the combi array
        Dim j As Integer
        For j = LBound(fArr) To UBound(fArr)
            Dim tmpOldCombi() As Variant, tmpNewEntry As Variant, ocU As Integer, ocL As Integer
            tmpOldCombi = fArr(j)(0)
            tmpNewEntry = fArr(j)(1)
            ocU = UBound(tmpOldCombi)
            ocL = LBound(tmpOldCombi)
            'expand tmpOldCombi by 1, put new entry into the last slot
            ReDim Preserve tmpOldCombi(ocL To ocU + 1)
            tmpOldCombi(UBound(tmpOldCombi)) = tmpNewEntry
            fArr(j) = tmpOldCombi
        Next j
    End If
Next i

Arr_Combis = fArr
End Function

