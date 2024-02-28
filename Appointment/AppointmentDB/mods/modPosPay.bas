Attribute VB_Name = "modPosPay"
Option Explicit
Private Const CONST_ACCOUNT_NUMBER As String = "06300058271"
Public Function ExportPositivePayFile(ByRef arystrCheckNum() As String, ByRef arycurAmount() As Currency, ByRef aryblnVoid() As Boolean, ByRef arystrNotes() As String) As Boolean
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
    Dim intFileNum As Integer
    Dim strBuffer As String
    Dim strTotalRec As String * 80
    Dim strTrailerRec As String * 80
    Dim lngNumIssues As Long
    Dim curTemp As Currency
    Dim strTemp As String
    Dim lngCtr As Long
    
    On Error GoTo ErrHand
    

    
    ExportPositivePayFile = False
    
    'Validate Data ************************************************************************************************************************************************
    lngCtr = UBound(arystrCheckNum)
    If lngCtr <> UBound(arycurAmount) Or lngCtr <> UBound(aryblnVoid) Or lngCtr <> UBound(arystrNotes) Then
        MsgBox "Check numbers, amounts, notes, and void flags are not equal.  Cannot proceed.", vbCritical
        Exit Function
    End If
    
    If MsgBox("Would you like to delete all existing positive pay files?", vbYesNo + vbQuestion) = vbYes Then Kill App.Path & "\util\*.ppy"
    
    'Init Vars ************************************************************************************************************************************************
    intFileNum = FreeFile()
    Open App.Path & "\util\" & Format(Now, "mmddyy_hhmmss") & ".ppy" For Output As #intFileNum
    
    strTotalRec = Replace(strTotalRec, Chr(0), " ")
    strTrailerRec = Replace(strTrailerRec, Chr(0), " ")
    
    curTemp = 0
    
    'print p/w
    Print #intFileNum, "$$ADD ID=HXRED4LP BATCHID='QP05PYQ'"
    
    'Body ***************************************************************************************************************************************************
    For lngCtr = 0 To UBound(arystrCheckNum)
        Print #intFileNum, ParseRecord(arystrCheckNum(lngCtr), arycurAmount(lngCtr), aryblnVoid(lngCtr), arystrNotes(lngCtr))
        curTemp = curTemp + arycurAmount(lngCtr)
    Next lngCtr
    
    'Chase is changing their mind  - no need to submit whole blocks
    'For lngCtr = 1 To (80 - (UBound(arystrCheckNum) Mod 80)) 'num recs to fill as blanks
    '    Print #intFileNum, strTrailerRec
    'Next lngCtr
    
    'Total Rec ***************************************************************************************************************************************************
    'format currency
    strTemp = Format(Abs(curTemp), "Currency")
    strTemp = Replace(strTemp, "$", "")
    strTemp = Replace(strTemp, ",", "")
    strTemp = Replace(strTemp, ".", "")
    If curTemp < 0 Then strTemp = "-" & strTemp
    'parse string
    Mid(strTotalRec, 1) = "T"
    Mid(strTotalRec, 26) = "000000000000"
    Mid(strTotalRec, 26 + (12 - Len(strTemp))) = strTemp
    
    Print #intFileNum, strTotalRec
    
    'Trailer ***************************************************************************************************************************************************
    Mid(strTrailerRec, 1) = "***EOF***"
    Mid(strTrailerRec, 11) = "0000000000"
    Mid(strTrailerRec, 20 - Len(CStr(UBound(arystrCheckNum) + 2)) + 1) = UBound(arystrCheckNum) + 2
    Print #intFileNum, strTrailerRec
    
    'Cleanup ***************************************************************************************************************************************************
    Close #intFileNum
    ExportPositivePayFile = True
    
    Exit Function
ErrHand:
    If Err.Number = 53 Then
        'no files existed.  ok to continue
        Resume Next
    Else
        RaiseError Err
    End If
End Function
Private Function getNumIssues(ByRef aryblnVoid() As Boolean)
'------------------------------------------------------------------------------------
'Date: 01/18/01
'Author: Eric Pena
'Description:   Determines the number of check 'issues' to submit.
'Parameters:    aryblnVoid() - Array of booleans (true if the check is voided)
'Returns:       # of false values in the array
'------------------------------------------------------------------------------------
    Dim lngCtr As Long
    Dim lngTotal As Long
    
    getNumIssues = -1
    lngTotal = 0
    
    For lngCtr = 0 To UBound(aryblnVoid)
        If Not aryblnVoid(lngCtr) Then lngTotal = lngTotal + 1
    Next lngCtr
    
    getNumIssues = lngTotal
    
End Function
Private Function ParseRecord(ByVal strCheckNum As String, ByVal curAmount As Currency, ByVal blnVoid As Boolean, ByRef strNotes As String) As String
'------------------------------------------------------------------------------------
'Date: 01/18/01
'Author: Eric Pena
'Description:   Parses out a string formatted to comply with Chase's Positive Pay format for the given check information.
'Parameters:    strCheckNum - Check number to use
'                       curAmount - Check amount to use
'                       blnVoid - True if we are voiding the check
'Returns:       a string formatted to comply with Chase's Positive Pay format
'------------------------------------------------------------------------------------
    Dim strBuffer As String * 80
    Dim strTemp As String
    
    'init vars
    strBuffer = Replace(strBuffer, Chr(0), " ")
    strCheckNum = Left(Trim(strCheckNum), 10)
    If blnVoid Then curAmount = 0
    
    'format currency
    strTemp = Format(Abs(curAmount), "Currency")
    strTemp = Replace(strTemp, "$", "")
    strTemp = Replace(strTemp, ",", "")
    strTemp = Replace(strTemp, ".", "")
    If curAmount < 0 Then strTemp = "-" & strTemp
    
    'parse string
    Mid(strBuffer, 2) = CONST_ACCOUNT_NUMBER
    Mid(strBuffer, 16) = "0000000000000000000000"
    Mid(strBuffer, 25 - Len(strCheckNum) + 1) = strCheckNum
    Mid(strBuffer, 37 - Len(strTemp) + 1) = strTemp
    Mid(strBuffer, 38) = Format(Date, "mmddyy")
    Mid(strBuffer, 46) = Left(Trim(strNotes), 15)
    If blnVoid Then Mid(strBuffer, 44) = "26"
    
    ParseRecord = strBuffer
End Function
