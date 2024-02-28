Attribute VB_Name = "modArray"
Option Explicit

Public Sub QuickSort(ByRef varArray As Variant, Optional lngFirst As Long = -1, _
                     Optional lngLast As Long = -1)
'--------------------------------------------------------------------
'Date: 10/23/2000
'Author: Dave Richkun
'Description: QuickSort algorithm used to sort the items in the varArray array.
'Parameters: varArray - The array to be sorted
'            lngFirst - Optional value identifying the first element to begin sorting with
'            lngLast - Optional value identifying the last element to begin sorting with
'Returns: The sorted array by reference parameter
'--------------------------------------------------------------------
    
    Dim lngLow      As Long
    Dim lngHigh     As Long
    Dim lngMiddle   As Long
    Dim varTempVal  As Variant
    Dim varTestVal  As Variant
    
    If lngFirst = -1 Then lngFirst = LBound(varArray)
    If lngLast = -1 Then lngLast = UBound(varArray)
        
    If lngFirst < lngLast Then
        lngMiddle = (lngFirst + lngLast) / 2
        varTestVal = varArray(lngMiddle)
        lngLow = lngFirst
        lngHigh = lngLast
        Do
            Do While varArray(lngLow) < varTestVal
                lngLow = lngLow + 1
            Loop
            Do While varArray(lngHigh) > varTestVal
                lngHigh = lngHigh - 1
            Loop
            If (lngLow <= lngHigh) Then
                varTempVal = varArray(lngLow)
                varArray(lngLow) = varArray(lngHigh)
                varArray(lngHigh) = varTempVal
                lngLow = lngLow + 1
                lngHigh = lngHigh - 1
            End If
        Loop While (lngLow <= lngHigh)
        
        If lngFirst < lngHigh Then QuickSort varArray, lngFirst, lngHigh
        If lngLow < lngLast Then QuickSort varArray, lngLow, lngLast
    End If
    
End Sub
Public Sub QuickSort2D(ByRef varArray As Variant, ByVal lng2DWidth As Long, ByVal lngSortKey As Long, Optional lngFirst As Long = -1, _
                     Optional lngLast As Long = -1)
'--------------------------------------------------------------------
'Date: 10/23/2000
'Author: Dave Richkun
'Description: Modified "QuickSort" (function defined above) to sort the items in a 2D varArray array.
'Parameters: varArray - The 2D array to be sorted
'            lngFirst - Optional value identifying the first element to begin sorting with
'            lngLast - Optional value identifying the last element to begin sorting with
'Returns: The sorted array by reference parameter
'--------------------------------------------------------------------
    
    Dim lngLow      As Long
    Dim lngHigh     As Long
    Dim lngMiddle   As Long
    Dim varTempVal()  As Variant
    Dim varTestVal  As Variant
    Dim lngCtr
    
    If lngFirst = -1 Then lngFirst = LBound(varArray)
    If lngLast = -1 Then lngLast = UBound(varArray)
    ReDim varTempVal(lng2DWidth)
    
    If lngFirst < lngLast Then
        lngMiddle = (lngFirst + lngLast) / 2
        varTestVal = varArray(lngMiddle, lngSortKey)
        lngLow = lngFirst
        lngHigh = lngLast
        Do
            Do While varArray(lngLow, lngSortKey) < varTestVal
                lngLow = lngLow + 1
            Loop
            Do While varArray(lngHigh, lngSortKey) > varTestVal
                lngHigh = lngHigh - 1
            Loop
            If (lngLow <= lngHigh) Then
                'copy to temp array
                For lngCtr = 0 To lng2DWidth
                    varTempVal(lngCtr) = varArray(lngLow, lngCtr)
                Next lngCtr
                
                'swap positions
                For lngCtr = 0 To lng2DWidth
                    varArray(lngLow, lngCtr) = varArray(lngHigh, lngCtr)
                Next lngCtr
                
                'restore from temp array
                For lngCtr = 0 To lng2DWidth
                    varArray(lngHigh, lngCtr) = varTempVal(lngCtr)
                Next lngCtr
                
                lngLow = lngLow + 1
                lngHigh = lngHigh - 1
            End If
        Loop While (lngLow <= lngHigh)
        
        If lngFirst < lngHigh Then QuickSort2D varArray, lng2DWidth, lngSortKey, lngFirst, lngHigh
        If lngLow < lngLast Then QuickSort2D varArray, lng2DWidth, lngSortKey, lngLow, lngLast
    End If
    
End Sub




