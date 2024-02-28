Option Strict Off
Option Explicit On
Module modPosPay
	Private Const CONST_ACCOUNT_NUMBER As String = "06300058271"
	Public Function ExportPositivePayFile(ByRef arystrCheckNum() As String, ByRef arycurAmount() As Decimal, ByRef aryblnVoid() As Boolean, ByRef arystrNotes() As String) As Boolean
		'------------------------------------------------------------------------------------
		'Date: 01/18/01
		'Author: Eric Pena
		'Description:   Outputs the given check information into a Positive Pay formatted text file
		'Parameters:    arystrCheckNum() - Array of check numbers
		'                       arycurAmount() - Array of check amounts
		'                       aryblnVoid() - Array of booleans (true if the check is voided)
		'                       arystrNotes() - Array of misc notes to output w/each check
		'Returns:       true if the process was sucessful, false otherwise
		'------------------------------------------------------------------------------------
		Dim intFileNum As Short
		Dim strBuffer As String
		Dim strTotalRec As New VB6.FixedLengthString(80)
		Dim strTrailerRec As New VB6.FixedLengthString(80)
		Dim lngNumIssues As Integer
		Dim curTemp As Decimal
		Dim strTemp As String
		Dim lngCtr As Integer
		
		On Error GoTo ErrHand
		
		
		
		ExportPositivePayFile = False
		
		'Validate Data ************************************************************************************************************************************************
		lngCtr = UBound(arystrCheckNum)
		If lngCtr <> UBound(arycurAmount) Or lngCtr <> UBound(aryblnVoid) Or lngCtr <> UBound(arystrNotes) Then
			MsgBox("Check numbers, amounts, notes, and void flags are not equal.  Cannot proceed.", MsgBoxStyle.Critical)
			Exit Function
		End If
		
		If MsgBox("Would you like to delete all existing positive pay files?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.Yes Then Kill(My.Application.Info.DirectoryPath & "\util\*.ppy")
		
		'Init Vars ************************************************************************************************************************************************
		intFileNum = FreeFile
		FileOpen(intFileNum, My.Application.Info.DirectoryPath & "\util\" & VB6.Format(Now, "mmddyy_hhmmss") & ".ppy", OpenMode.Output)
		
		strTotalRec.Value = Replace(strTotalRec.Value, Chr(0), " ")
		strTrailerRec.Value = Replace(strTrailerRec.Value, Chr(0), " ")
		
		curTemp = 0
		
		'print p/w
		PrintLine(intFileNum, "$$ADD ID=HXRED4LP BATCHID='QP05PYQ'")
		
		'Body ***************************************************************************************************************************************************
		For lngCtr = 0 To UBound(arystrCheckNum)
			PrintLine(intFileNum, ParseRecord(arystrCheckNum(lngCtr), arycurAmount(lngCtr), aryblnVoid(lngCtr), arystrNotes(lngCtr)))
			curTemp = curTemp + arycurAmount(lngCtr)
		Next lngCtr
		
		'Chase is changing their mind  - no need to submit whole blocks
		'For lngCtr = 1 To (80 - (UBound(arystrCheckNum) Mod 80)) 'num recs to fill as blanks
		'    Print #intFileNum, strTrailerRec
		'Next lngCtr
		
		'Total Rec ***************************************************************************************************************************************************
		'format currency
		strTemp = VB6.Format(System.Math.Abs(curTemp), "Currency")
		strTemp = Replace(strTemp, "$", "")
		strTemp = Replace(strTemp, ",", "")
		strTemp = Replace(strTemp, ".", "")
		If curTemp < 0 Then strTemp = "-" & strTemp
		'parse string
		Mid(strTotalRec.Value, 1) = "T"
		Mid(strTotalRec.Value, 26) = "000000000000"
		Mid(strTotalRec.Value, 26 + (12 - Len(strTemp))) = strTemp
		
		PrintLine(intFileNum, strTotalRec.Value)
		
		'Trailer ***************************************************************************************************************************************************
		Mid(strTrailerRec.Value, 1) = "***EOF***"
		Mid(strTrailerRec.Value, 11) = "0000000000"
		Mid(strTrailerRec.Value, 20 - Len(CStr(UBound(arystrCheckNum) + 2)) + 1) = CStr(UBound(arystrCheckNum) + 2)
		PrintLine(intFileNum, strTrailerRec.Value)
		
		'Cleanup ***************************************************************************************************************************************************
		FileClose(intFileNum)
		ExportPositivePayFile = True
		
		Exit Function
ErrHand: 
		If Err.Number = 53 Then
			'no files existed.  ok to continue
			Resume Next
		Else
			RaiseError(Err)
		End If
	End Function
	Private Function getNumIssues(ByRef aryblnVoid() As Boolean) As Object
		'------------------------------------------------------------------------------------
		'Date: 01/18/01
		'Author: Eric Pena
		'Description:   Determines the number of check 'issues' to submit.
		'Parameters:    aryblnVoid() - Array of booleans (true if the check is voided)
		'Returns:       # of false values in the array
		'------------------------------------------------------------------------------------
		Dim lngCtr As Integer
		Dim lngTotal As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object getNumIssues. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getNumIssues = -1
		lngTotal = 0
		
		For lngCtr = 0 To UBound(aryblnVoid)
			If Not aryblnVoid(lngCtr) Then lngTotal = lngTotal + 1
		Next lngCtr
		
		'UPGRADE_WARNING: Couldn't resolve default property of object getNumIssues. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getNumIssues = lngTotal
		
	End Function
	Private Function ParseRecord(ByVal strCheckNum As String, ByVal curAmount As Decimal, ByVal blnVoid As Boolean, ByRef strNotes As String) As String
		'------------------------------------------------------------------------------------
		'Date: 01/18/01
		'Author: Eric Pena
		'Description:   Parses out a string formatted to comply with Chase's Positive Pay format for the given check information.
		'Parameters:    strCheckNum - Check number to use
		'                       curAmount - Check amount to use
		'                       blnVoid - True if we are voiding the check
		'Returns:       a string formatted to comply with Chase's Positive Pay format
		'------------------------------------------------------------------------------------
		Dim strBuffer As New VB6.FixedLengthString(80)
		Dim strTemp As String
		
		'init vars
		strBuffer.Value = Replace(strBuffer.Value, Chr(0), " ")
		strCheckNum = Left(Trim(strCheckNum), 10)
		If blnVoid Then curAmount = 0
		
		'format currency
		strTemp = VB6.Format(System.Math.Abs(curAmount), "Currency")
		strTemp = Replace(strTemp, "$", "")
		strTemp = Replace(strTemp, ",", "")
		strTemp = Replace(strTemp, ".", "")
		If curAmount < 0 Then strTemp = "-" & strTemp
		
		'parse string
		Mid(strBuffer.Value, 2) = CONST_ACCOUNT_NUMBER
		Mid(strBuffer.Value, 16) = "0000000000000000000000"
		Mid(strBuffer.Value, 25 - Len(strCheckNum) + 1) = strCheckNum
		Mid(strBuffer.Value, 37 - Len(strTemp) + 1) = strTemp
		Mid(strBuffer.Value, 38) = VB6.Format(Today, "mmddyy")
		Mid(strBuffer.Value, 46) = Left(Trim(strNotes), 15)
		If blnVoid Then Mid(strBuffer.Value, 44) = "26"
		
		ParseRecord = strBuffer.Value
	End Function
End Module