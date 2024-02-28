Option Strict Off
Option Explicit On
Module modArray
	
	Public Sub QuickSort(ByRef varArray As Object, Optional ByRef lngFirst As Integer = -1, Optional ByRef lngLast As Integer = -1)
		'--------------------------------------------------------------------
		'Date: 10/23/2000
		'Author: Dave Richkun
		'Description: QuickSort algorithm used to sort the items in the varArray array.
		'Parameters: varArray - The array to be sorted
		'            lngFirst - Optional value identifying the first element to begin sorting with
		'            lngLast - Optional value identifying the last element to begin sorting with
		'Returns: The sorted array by reference parameter
		'--------------------------------------------------------------------
		
		Dim lngLow As Integer
		Dim lngHigh As Integer
		Dim lngMiddle As Integer
		Dim varTempVal As Object
		Dim varTestVal As Object
		
		If lngFirst = -1 Then lngFirst = LBound(varArray)
		If lngLast = -1 Then lngLast = UBound(varArray)
		
		If lngFirst < lngLast Then
			lngMiddle = (lngFirst + lngLast) / 2
			'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varTestVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varTestVal = varArray(lngMiddle)
			lngLow = lngFirst
			lngHigh = lngLast
			Do 
				'UPGRADE_WARNING: Couldn't resolve default property of object varTestVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object varArray(lngLow). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Do While varArray(lngLow) < varTestVal
					lngLow = lngLow + 1
				Loop 
				'UPGRADE_WARNING: Couldn't resolve default property of object varTestVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object varArray(lngHigh). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Do While varArray(lngHigh) > varTestVal
					lngHigh = lngHigh - 1
				Loop 
				If (lngLow <= lngHigh) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varTempVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varTempVal = varArray(lngLow)
					'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varArray(lngLow) = varArray(lngHigh)
					'UPGRADE_WARNING: Couldn't resolve default property of object varTempVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varArray(lngHigh) = varTempVal
					lngLow = lngLow + 1
					lngHigh = lngHigh - 1
				End If
			Loop While (lngLow <= lngHigh)
			
			If lngFirst < lngHigh Then QuickSort(varArray, lngFirst, lngHigh)
			If lngLow < lngLast Then QuickSort(varArray, lngLow, lngLast)
		End If
		
	End Sub
	Public Sub QuickSort2D(ByRef varArray As Object, ByVal lng2DWidth As Integer, ByVal lngSortKey As Integer, Optional ByRef lngFirst As Integer = -1, Optional ByRef lngLast As Integer = -1)
		'--------------------------------------------------------------------
		'Date: 10/23/2000
		'Author: Dave Richkun
		'Description: Modified "QuickSort" (function defined above) to sort the items in a 2D varArray array.
		'Parameters: varArray - The 2D array to be sorted
		'            lngFirst - Optional value identifying the first element to begin sorting with
		'            lngLast - Optional value identifying the last element to begin sorting with
		'Returns: The sorted array by reference parameter
		'--------------------------------------------------------------------
		
		Dim lngLow As Integer
		Dim lngHigh As Integer
		Dim lngMiddle As Integer
		Dim varTempVal() As Object
		Dim varTestVal As Object
		Dim lngCtr As Object
		
		If lngFirst = -1 Then lngFirst = LBound(varArray)
		If lngLast = -1 Then lngLast = UBound(varArray)
		ReDim varTempVal(lng2DWidth)
		
		If lngFirst < lngLast Then
			lngMiddle = (lngFirst + lngLast) / 2
			'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varTestVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varTestVal = varArray(lngMiddle, lngSortKey)
			lngLow = lngFirst
			lngHigh = lngLast
			Do 
				'UPGRADE_WARNING: Couldn't resolve default property of object varTestVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object varArray(lngLow, lngSortKey). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Do While varArray(lngLow, lngSortKey) < varTestVal
					lngLow = lngLow + 1
				Loop 
				'UPGRADE_WARNING: Couldn't resolve default property of object varTestVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object varArray(lngHigh, lngSortKey). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Do While varArray(lngHigh, lngSortKey) > varTestVal
					lngHigh = lngHigh - 1
				Loop 
				If (lngLow <= lngHigh) Then
					'copy to temp array
					For lngCtr = 0 To lng2DWidth
						'UPGRADE_WARNING: Couldn't resolve default property of object lngCtr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object varTempVal(lngCtr). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						varTempVal(lngCtr) = varArray(lngLow, lngCtr)
					Next lngCtr
					
					'swap positions
					For lngCtr = 0 To lng2DWidth
						'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						varArray(lngLow, lngCtr) = varArray(lngHigh, lngCtr)
					Next lngCtr
					
					'restore from temp array
					For lngCtr = 0 To lng2DWidth
						'UPGRADE_WARNING: Couldn't resolve default property of object lngCtr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object varTempVal(lngCtr). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						varArray(lngHigh, lngCtr) = varTempVal(lngCtr)
					Next lngCtr
					
					lngLow = lngLow + 1
					lngHigh = lngHigh - 1
				End If
			Loop While (lngLow <= lngHigh)
			
			If lngFirst < lngHigh Then QuickSort2D(varArray, lng2DWidth, lngSortKey, lngFirst, lngHigh)
			If lngLow < lngLast Then QuickSort2D(varArray, lng2DWidth, lngSortKey, lngLow, lngLast)
		End If
		
	End Sub
End Module