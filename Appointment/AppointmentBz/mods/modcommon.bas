'-------------------------------------------------------------------------------
'Module Name: modCommon
'Author: Dave Richkun
'Date: 11/04/1999
'Description: This module is intended to encapsulate generic routines
'             that can be used in a wide range of applications.
'-------------------------------------------------------------------------------
'Revision History:
'
'-------------------------------------------------------------------------------

Option Explicit
'DEMO
'->Public Const CONST_PSYQUEL_DSN As String = "PsyquelSQL"
'->Public Const CONST_PSYQUEL_DATABASE As String = "PsyquelDemo"
'-old>Public Const CONST_PSYQUEL_UA As String = "psyquel_login"
'-old>Public Const CONST_PSYQUEL_AC As String = "DBSecure"
'-old>Public Const CONST_PSYQUEL_CNN As String = "PsyquelSQL"
'->Public Const CONST_PSYQUEL_UA As String = "sa"
'->Public Const CONST_PSYQUEL_AC As String = "psy1234!"
'->Public Const CONST_PSYQUEL_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=sa;Initial Catalog=PsyquelDemo;Data Source=192.168.4.25"

'->Public Const CONST_PSYREPL_DSN As String = "PsyquelSQL"
'->Public Const CONST_PSYREPL_DATABASE As String = "PsyquelDemo"
'-old>Public Const CONST_PSYREPL_CNN As String = "PsyquelSQL"
'->Public Const CONST_PSYREPL_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=sa;Initial Catalog=PsyquelDemo;Data Source=192.168.4.25"

'Test
Public Const CONST_TEST_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=sa;Initial Catalog=PsyquelTemp;Data Source=192.168.4.25"
'PsyquelDirect
Public Const CONST_DIRECT_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=sa;Initial Catalog=PsyquelDirect;Data Source=192.168.4.25"

'Production
Public Const CONST_PSYQUEL_DSN As String = "PsyquelSQL"
Public Const CONST_PSYQUEL_DATABASE As String = "PsyquelProd"
Public Const CONST_PSYQUEL_UA As String = "sa"
Public Const CONST_PSYQUEL_AC As String = "psy1234!"
'Local ----> This one works for both Local and Mgmt Server in Production <----
''Public Const CONST_PSYQUEL_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=psylocaladmin;Initial Catalog=PsyquelProd;Data Source=192.168.0.28"
'Public Const CONST_PSYQUEL_CNN As String = "Provider=SQLOLEDB;Data Source=192.168.4.25;Initial Catalog=PsyquelProd;User Id=sa;Password=psy1234!;"
Public Const CONST_PSYQUEL_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=sa;Initial Catalog=PsyquelProd;Data Source=192.168.4.25"

'Replication Database
Public Const CONST_PSYREPL_DSN As String = "PsyquelRpl"
Public Const CONST_PSYREPL_DATABASE As String = "PsyquelProd"
''Public Const CONST_PSYREPL_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=psylocaladmin;Initial Catalog=PsyquelProd;Data Source=192.168.0.29"
'Public Const CONST_PSYREPL_CNN As String = "Provider=SQLOLEDB;Data Source=192.168.4.25;Initial Catalog=PsyquelProd;User Id=sa;Password=psy1234!;"
Public Const CONST_PSYREPL_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=sa;Initial Catalog=PsyquelProd;Data Source=192.168.4.26"

'web --> This works  -->>> WEB SERVERS HAVE TO HAVE THE ODBC CONNECTION SET AS "Windows Authentication"
'Public Const CONST_PSYQUEL_CNN As String = "PsyquelSQL"
'Public Const CONST_PSYREPL_CNN As String = "PsyquelRpl"

'Windows authentication
'-->nope Public Const CONST_PSYQUEL_CNN As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PsyquelProd;Data Source=192.168.4.25"
'-->nope Public Const CONST_PSYREPL_CNN As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PsyquelProd;Data Source=192.168.4.26"

'-->nope Public Const CONST_PSYQUEL_CNN As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PsyquelProd;Data Source=192.168.0.23"
'-->nope Public Const CONST_PSYREPL_CNN As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PsyquelProd;Data Source=192.168.0.22"
'Local Authentication
'-->nope Public Const CONST_PSYQUEL_CNN As String = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=psylocaladmin;Initial Catalog=PsyquelProd;Data Source=192.168.0.23"
'-->nope Public Const CONST_PSYQUEL_CNN As String = "Provider=SQLOLEDB.1;Password=sqladmin;Persist Security Info=True;User ID=sa;Initial Catalog=PsyquelProd;Data Source=192.168.0.23"

Type RandomWord
   Word As String * 4
End Type


Public Sub ShowError(ByRef objErr As Object)
'-------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 11/04/1999
'Description: This procedure simply displays error information contained in
'             the objErr object within a message box.
'Parameters: objErr - An object reference to a VB Error object.
'Returns: Null
'-------------------------------------------------------------------------------
'Revision History:
'
'-------------------------------------------------------------------------------

    Call MsgBox("Error: " & objErr.Description & vbLf & vbLf & "Error Number: " _
                & objErr.Number, vbOKOnly + vbCritical, "Error")

End Sub


Public Sub RaiseError(ByVal objErr As Object, Optional ByVal lngSQLErrorNum As Long, _
                      Optional ByVal strSource, Optional ByVal strDescription As String)
'-------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 12/27/1999
'Description: This procedure raises errors to the calling environment.
'             This routine captures intrinsic VB errors and SQL Server errors.
'             raised within SQL stored procedures.
'Parameters:  objErr - The generic VB Error object containing information about
'                    'normal' errors
'            lngSQLErrorNum - The error number raised by the data provider.
'            strSource - The place where the error occured.
'            strDescription - The description of the error message.
'Returns: Null
'-------------------------------------------------------------------------------
'Revision History:
'
'-------------------------------------------------------------------------------

    Dim lngErrorNum As Long
    Dim strMessage As String

    If objErr.Number <> 0 Then
        Err.Raise objErr.Number, objErr.Source, objErr.Description
    Else
        lngErrorNum = vbObjectError + lngSQLErrorNum
        strMessage = "Database Error " & lngSQLErrorNum & ": " & strDescription
        strMessage = strMessage & vbLf & "Occuring in module: " & strSource
        objErr.Raise lngErrorNum, strSource, strMessage
    End If

End Sub




Public Sub UnloadAllForms()
'-----------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 12/09/1997
'Description: This procedure iterates through the applications Forms Collection
'             object and unloads each form. This routine is typically called from
'             an applications 'Exit' menu, or from the MDI forms QueryUnload event.
'Parameters: None
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------
    
    Dim intCTR As Integer
    
    On Error GoTo Err_Trap
    
    For intCTR = Forms.Count - 1 To 1 Step -1
      Unload Forms(intCTR)
    Next intCTR
    
    Exit Sub
    
Err_Trap:
    Resume Next
    
End Sub


Public Sub ClearAllFormControls(ByRef frm As Form)
'-----------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 12/18/1997
'Description: This procedure clears all control values from a form by iterating
'             through the form's Controls Collection and setting the control to the
'             most neutral of values.
'Parameters: frm - A reference to the form whose controls are to be cleared.
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------

    Dim intCTR As Integer
    
    On Error GoTo ErrTrap:
    
    If frm Is Nothing Then
        Exit Sub
    End If
    
    For intCTR = 0 To frm.Controls.Count - 1
        If TypeOf frm.Controls(intCTR) Is TextBox Then
            frm.Controls(intCTR).Text = ""
        End If
        If TypeOf frm.Controls(intCTR) Is ComboBox Then
            frm.Controls(intCTR).Clear
            If frm.Controls(intCTR).Style <> vbComboDropdownList Then
                frm.Controls(intCTR).Text = ""
            End If
        End If
        If TypeOf frm.Controls(intCTR) Is CheckBox Then
            frm.Controls(intCTR).Value = vbUnchecked
        End If
    Next intCTR
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrTrap:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Public Function FormatPhoneNumber(ByVal varPhoneNum As Variant) As Variant
'-----------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 12/18/1997
'Description: This procedure formats a string representing a phone number
'             based on the number of characters passed to it.
'Parameters: strPhoneNum - The string value representing the phone number to be
'             formatted.
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------

    Dim strString As String
    Dim intLength As Integer
    Dim strPhone As String
    Dim intCTR As Integer
    Dim strTest As String
    Dim strPhoneNum As String
      
      
    If IsNull(varPhoneNum) Then
        FormatPhoneNumber = ""
        Exit Function
    Else
        strPhoneNum = CStr(varPhoneNum)
    End If
    
    intLength = Len(strPhoneNum)
    
    For intCTR = 1 To intLength
        If IsNumeric(Mid(strPhoneNum, intCTR)) Then
            strString = strString & Mid(strPhoneNum, intCTR, 1)
        End If
    Next intCTR
    
    intLength = Len(strString)
    
    'Prevent phone numbers consisting of all zeroes from making it through.
    For intCTR = 1 To intLength
        strTest = strTest & "0"
    Next intCTR
    If strString = strTest Then
        FormatPhoneNumber = ""
        Exit Function
    End If
    
    Select Case intLength
      Case 10
        strPhone = "(" & Mid(strString, 1, 3) & ") " & Mid(strString, 4, 3) & "-" & Mid(strString, 7)
      Case 7
        strPhone = Mid(strString, 1, 3) & "-" & Mid(strString, 4)
      Case Is > 10
        strPhone = Mid(strString, 1, intLength - 10) & " " & Mid(strString, intLength - 9, 3) & " " & Mid(strString, intLength - 6, 3) & "-" & Mid(strString, intLength - 3)
      Case Else
        strPhone = strString
    End Select
    
    FormatPhoneNumber = strPhone

End Function

Public Function FormatZipCode(ByVal varZipCode As Variant) As Variant
'-----------------------------------------------------------------------------------
'Author: Eric Pena
'Date: 4/28/2000
'Description: This procedure formats a string representing a zip code
'             based on the number of characters passed to it.
'Parameters: strZipCode - The string value representing the zip code to be
'             formatted.
'Returns: Formatted zip code
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------

    Dim strString As String
    Dim intLength As Integer
    Dim strZip As String
    Dim intCTR As Integer
    Dim strTest As String
    Dim strZipCode As String
    
    If IsNull(varZipCode) Then
        FormatZipCode = ""
        Exit Function
    Else
        strZipCode = Trim(CStr(varZipCode))
    End If
    
    intLength = Len(strZipCode)
    strString = strZipCode

    'Get rid of non-numeric chars
    For intCTR = 1 To intLength
        If Not IsNumeric(Mid(strZipCode, intCTR)) Then
            strString = Replace(strZipCode, Mid(strZipCode, intCTR), "")
        End If
    Next intCTR
    
    intLength = Len(strString)
    
    'Prevent zip codes consisting of all zeroes or nines from making it through.
    For intCTR = 1 To intLength
        strTest = strTest & "0"
    Next intCTR

    If strString = strTest Then
        FormatZipCode = ""
        Exit Function
    End If
    
    For intCTR = 1 To intLength
        strTest = strTest & "9"
    Next intCTR

    If strString = strTest Then
        FormatZipCode = ""
        Exit Function
    End If
    
    Select Case intLength
      Case Is > 5
        strZip = Left(strString, 5) & "-" & Mid(strString, 6, Len(strString) - 5)
      Case Else
        strZip = strString
    End Select
    
    FormatZipCode = strZip

End Function


Public Sub ForceUpperCase(ByRef intKeyPress As Integer)
'-----------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 12/18/1997
'Description: This procedure forces the passed KeyCode characer to uppercase.
'Parameters: intKeyPress - Key code representing the character to be forced to
'               uppercase.
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------
    
    intKeyPress = Asc(UCase(Chr(intKeyPress)))
  
End Sub


Public Function FormatSSN(ByVal varSSN As String) As Variant
'-----------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 02/01/2000
'Description: This procedure formats a string representing a social security number
'             (SSN) based on the number of characters passed to it.
'Parameters: strSSN - The string value representing the SSN to be
'             formatted.
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------

  Dim strString As String
  Dim intLength As Integer
  Dim strSSNumber As String
  Dim strSSN As String
  
  
  If IsNull(varSSN) Then
    FormatSSN = ""
    Exit Function
  Else
    strSSN = Trim(CStr(varSSN))
  End If
    
  strString = NumbersOnly(strSSN)
  intLength = Len(strString)
  
  Select Case intLength
    Case 11
      'Remove preceding 2 digits - for military status only - no longer used.
      strSSNumber = Mid(strString, 3, 3) & "-" & Mid(strString, 6, 2) & "-" & Mid(strString, 8)
    Case 9
      strSSNumber = Mid(strString, 1, 3) & "-" & Mid(strString, 4, 2) & "-" & Mid(strString, 6)
    Case Else
      strSSNumber = strString
  End Select
  
  FormatSSN = strSSNumber

End Function

Public Function FormatCC(ByVal varCC As String) As Variant
'-----------------------------------------------------------------------------------
'Author: Duane C Orth
'Date: 06/01/2018
'Description: This procedure formats a string representing a Credit Card number
'             (CC) based on the number of characters passed to it.
'Parameters: strCC - The string value representing the CC to be
'             formatted.
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------

  Dim strString As String
  Dim intLength As Integer
  Dim strCCNumber As String
  Dim strCC As String
  
  
  If IsNull(varCC) Then
    FormatCC = ""
    Exit Function
  Else
    strCC = Trim(CStr(varCC))
  End If
    
  strString = NumbersOnly(strCC)
  intLength = Len(strString)
  
  Select Case intLength
    Case 14
      'Remove preceding 2 digits - for military status only - no longer used.
      strCCNumber = "****" & " " & "****" & " " & "****" & " " & " " & Mid(strString, 13, 2)
    Case 15
      'Remove preceding 2 digits - for military status only - no longer used.
      strCCNumber = "****" & " " & "****" & " " & "****" & " " & " " & Mid(strString, 13, 3)
    Case 16
'      strCCNumber = Mid(strString, 1, 4) & " " & Mid(strString, 5, 4) & " " & Mid(strString, 9, 4) & " " & Mid(strString, 13, 4)
      strCCNumber = "****" & " " & "****" & " " & "****" & " " & Mid(strString, 13, 4)
    Case Else
      strCCNumber = strString
  End Select
  
  FormatCC = strCCNumber

End Function

Public Sub FriendlyNumberBox(Optional intStart As Integer = 0)
'-----------------------------------------------------------------------------------
'Author: Rick "Boom Boom" Segura
'Date: 02/10/2000
'Description: This procedure places the cursor of an "empty" control at the
'             "beginning" of a control. For example, if a masked text box is
'             clicked on and contains only the mask charatcters, the cursor
'             will be placed in the position given by intStart
'Parameters: intStart - default starting position
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------
    Dim ctl As Control
    Dim var As Variant
    
    Set ctl = Screen.ActiveControl
    
    var = NumbersOnly(ctl.Text)
    var = IIf(IsNull(var), "", var)
    
    If Len(Trim((CStr(var)))) = 0 Then
        ctl.SelStart = intStart
    End If
    Set ctl = Nothing
    
End Sub


Public Function CleanNumber(ByVal varString) As String
'-----------------------------------------------------------------------------------
'Author: Rick "Boom Boom" Segura
'Date: 02/14/2000
'Description: This procedure returns a given string of numbers less any number
'             format and symbols
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------
    Dim str As String
    
    If Not IsNull(varString) Then
        ' Remove mask characters for string comparison
        str = varString
        str = Replace(str, "(", "")
        str = Replace(str, ")", "")
        str = Replace(str, "-", "")
        str = Replace(str, "#", "")
        str = Replace(str, "%", "")
        str = Replace(str, "$", "")
        str = Replace(str, " ", "")
        
        CleanNumber = str
    Else
        CleanNumber = ""
    End If
    
End Function

Public Sub InputDigitsOnly(ByRef KeyCode As Integer)
'-----------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 02/14/2000
'Description: This routine allows only digits to be entered.  The call
'             is typically placed in the KeyPress event of TextBox controls.
'Returns: Null
'-----------------------------------------------------------------------------------
'Revision History:
'
'-----------------------------------------------------------------------------------

    Select Case KeyCode
        Case vbKeyBack
            'allow the keystroke
        Case Is < 48
            KeyCode = 0
        Case Is > 57
            KeyCode = 0
        Case Else
            'allow the keystroke
    End Select
    
End Sub

Public Function ParseTrim(ByRef Source As String, _
                       ByVal Delim As Integer) As String
'------------------------------------------------------------
'Author: Rick "Boom boom" Segura, supplied by Ken White     '
'Date: 03/08/2000                                           '
'Description: Returns the substring from the the begining of'
'             to the first insyance of the delimeter and    '
'             removes the token plus the 1st delimeter from '
'             the original string                           '
'Returns: Substring                                         '
'------------------------------------------------------------
'Revision History:                                          '
'                                                           '
'------------------------------------------------------------

    Dim d       As String
    Dim Dlnxt   As Integer
    Dim Slen    As Integer

    d = Chr(Delim)

    Slen = Len(Source)
    Dlnxt = InStr(Source, d)
    
    Select Case Dlnxt
    Case 0
        ParseTrim = Source
        Source = ""
    Case 1
        ParseTrim = ""
        If Slen >= 1 Then
            ParseTrim = ""
            Source = Trim(Right(Source, Slen - 1))
        End If
    Case Slen
        If Slen > 1 Then
            ParseTrim = Trim(Left(Source, Slen - 1))
        Else
            ParseTrim = ""
        End If
        Source = ""
    Case Else
        ParseTrim = Trim(Left(Source, Dlnxt - 1))
        Source = Right(Source, Slen - Dlnxt)
    End Select

End Function

Public Function GetConnectionString() As String
'------------------------------------------------------------
'Author: Dave Richkun
'Date: 03/30/2000
'Description: Returns the OLEDB Connection string from the Registry.
'             The Connection string must have been previously set using
'             the SetConnection utility.
'Returns: Substring
'------------------------------------------------------------
'Revision History:                                          '
'                                                           '
'------------------------------------------------------------

    Dim strConnect As String
    
''    strConnect = "Provider=SQLOLEDB;"
''    strConnect = strConnect & "Data Source=" & GetSetting("Psyquel Application", "DBConnection", "DataSource") & ";"
''    strConnect = strConnect & "Initial Catalog=" & GetSetting("Psyquel Application", "DBConnection", "Catalog") & ";"
''    strConnect = strConnect & "User ID=" & GetSetting("Psyquel Application", "DBConnection", "UserID") & ";"
''    strConnect = strConnect & "Password=" & GetSetting("Psyquel Application", "DBConnection", "Password") & ";"
''
''    GetConnectionString = strConnect
 '''''   GetConnectionString = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=psylocaladmin;Initial Catalog=PsyquelProd;Data Source=192.168.0.28"
    GetConnectionString = "Provider=SQLOLEDB.1;Password=psy1234!;Persist Security Info=True;User ID=sa;Initial Catalog=PsyquelProd;Data Source=192.168.4.25"
'--old--    GetConnectionString = "Provider=SQLOLEDB;Data Source=PSYQUEL-BDC;Initial Catalog=PsyquelTest;User ID=psyquel_login;Password=dbsecure;"

End Function

Public Function GenerateRandomPassword(ByVal strWordFile As String) As String
'-------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 04/17/2000
'Description: This procedure generates a random password by concatenating 2
'             random four character words listed in a text file.  The list
'             of words used for this function was downloaded from the University
'             of Oakland.  Be warned, this function may generate some
'             'questionable' passwords.
'Parameters: strWordFile - The path that points to the text file containing
'             the list of 4 character words.
'Returns: The random password produced by this function.
'-------------------------------------------------------------------------------
'Revision History:
'
'-------------------------------------------------------------------------------

    Dim intRndm As Integer
    Dim strPart1 As String
    Dim strPart2 As String
    Dim intFileNum As Integer
    Dim typWord As RandomWord

    intFileNum = FreeFile()
    Open strWordFile For Random As #intFileNum Len = 6

    Randomize
    intRndm = Int((3678 * Rnd) + 1)
    Get intFileNum, intRndm, typWord
    strPart1 = typWord.Word
    
    intRndm = Int((3678 * Rnd) + 1)
    Get intFileNum, intRndm, typWord
    strPart2 = typWord.Word
    Close #intFileNum
    
    GenerateRandomPassword = strPart1 & strPart2

End Function

Public Function NumbersOnly(varValue As Variant) As Variant
'-------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 04/19/2000
'Description: This procedure parses a variant parameter and returns only
'             the numeric characters to the calling procedure.
'Parameters: varValue - The Variant parameter to be parsed.
'Returns: The parsed value containing only the numeric characters
'-------------------------------------------------------------------------------
'Revision History:
'
'-------------------------------------------------------------------------------

    Dim intCTR As Integer
    Dim strValue As String
    Dim intLength As Integer
    Dim strNumber As String
    
    If IsNull(varValue) Then
        NumbersOnly = ""
        Exit Function
    End If
    
    strValue = CStr(varValue)
    intLength = Len(strValue)
    
    For intCTR = 1 To intLength
        If IsNumeric(Mid(strValue, intCTR, 1)) Then
            strNumber = strNumber & Mid(strValue, intCTR, 1)
        End If
    Next intCTR
    
    NumbersOnly = strNumber

End Function

Public Function IsAnArray(ByVal varAry As Variant) As Boolean
'--------------------------------------------------------------------------------
'Author: Rick "Boom Boom" Segura                                                '
'Date: 05/18/2000                                                               '
'Description: This procedure determines wether or not the passed argument is an '
'             array.  Created  because VB did a piss poor job on IsArray().     '
'Parameters: varAry - variant to be checked                                     '
'Returns: True if passed argument is an array, false otherwise                  '
'--------------------------------------------------------------------------------
'Revision History:                                                              '
'                                                                               '
'--------------------------------------------------------------------------------
    
    ' Assume parameter is not an array
    IsAnArray = False
    
    If IsEmpty(varAry) Then GoTo Exit_Function
    If TypeName(varAry) = "Nothing" Then GoTo Exit_Function
    If IsNull(varAry) Then GoTo Exit_Function
    
    If IsArray(varAry) Then IsAnArray = True
    
Exit_Function:
    
End Function

Public Function Max(ByVal lngVal1 As Long, ByVal lngVal2 As Long) As Long
'--------------------------------------------------------------------------------
'Author: Rick "Boom Boom" Segura                                                '
'Date: 01/25/2001                                                               '
'Description: Calculates the maximum of 2 numbers                               '
'Parameters: lngVal1 - 1st value of comparison pair                             '
'            lngVal2 - 2nd value of comparison pair                             '
'Returns: The maximum of the pair                                               '
'--------------------------------------------------------------------------------
'Revision History:                                                              '
'                                                                               '
'--------------------------------------------------------------------------------
    Max = IIf((lngVal2 > lngVal1), lngVal2, lngVal1)
End Function

Public Function Min(ByVal lngVal1 As Long, ByVal lngVal2 As Long) As Long
'--------------------------------------------------------------------------------
'Author: Rick "Boom Boom" Segura                                                '
'Date: 01/25/2001                                                               '
'Description: Calculates the minimum of 2 numbers                               '
'Parameters: lngVal1 - 1st value of comparison pair                             '
'            lngVal2 - 2nd value of comparison pair                             '
'Returns: The minimum of the pair                                               '
'--------------------------------------------------------------------------------
'Revision History:                                                              '
'                                                                               '
'--------------------------------------------------------------------------------
    Min = IIf((lngVal2 > lngVal1), lngVal1, lngVal2)
End Function

Public Sub StrCat(ByRef strMain As String, ByVal strSub As String)
'--------------------------------------------------------------------------------
'Author: Rick "Boom Boom" Segura                                                '
'Date: 01/25/2001                                                               '
'Description: Concetenates two strings                                          '
'Parameters: strMain - string to perform concatenation on                       '
'            strSub -  string to concatenate                                    '
'--------------------------------------------------------------------------------
'Revision History:                                                              '
'                                                                               '
'--------------------------------------------------------------------------------
    If strMain > "" Then
        strMain = strMain & strSub
    Else
        strMain = strSub
    End If
End Sub

Public Sub StrCatDel(ByRef strMain As String, ByVal strSub As String, _
                     Optional ByVal strDel As String = ",")
'--------------------------------------------------------------------------------
'Author: Rick "Boom Boom" Segura                                                '
'Date: 01/25/2001                                                               '
'Description: Concetenates two strings with a delimeter                         '
'Parameters: strMain - string to perform concatenation on                       '
'            strSub -  string to concatenate                                    '
'            strDel - delimeter to use                                          '
'--------------------------------------------------------------------------------
'Revision History:                                                              '
'                                                                               '
'--------------------------------------------------------------------------------
    If strMain > "" Then
        strMain = strMain & strDel & strSub
    Else
        strMain = strSub
    End If
End Sub

Public Function NNs(ByVal varString As Variant) As String
'--------------------------------------------------------------------------------
'Author: Rick "Boom Boom" Segura                                                '
'Date: 01/25/2001                                                               '
'Description: Ensures that a value is Not Null                                  '
'Parameters: A non-Null string value                                            '
'--------------------------------------------------------------------------------
'Revision History:                                                              '
'                                                                               '
'--------------------------------------------------------------------------------
    If IsNull(varString) Then
        NNs = ""
    Else
        NNs = CStr(varString)
    End If
End Function


Public Function RPad(ByVal strString As String, ByVal intPad As Integer) As String
'--------------------------------------------------------------------------------
'Author: Dave Richkun
'Date: 09/17/2001
'Description: Pads space characters to the end of a string
'Parameters: strString - The original string
'            intPad - The number of space characters to append to the string
'Returns: The padded string
'--------------------------------------------------------------------------------
'Revision History:                                                              '
'                                                                               '
'--------------------------------------------------------------------------------
    Dim strCopy As String
    Dim intCTR As Integer
    
    strCopy = strString
    
    For intCTR = 1 To intPad
        strCopy = strCopy & Chr(32)
    Next intCTR

    'Return the padded string
    RPad = strCopy

End Function
Public Function ParamType(ByVal strReportName As String, ByVal lngParamNum As Long) As String
'--------------------------------------------------------------------------------
'Author: Eric Pena
'Date: 10/04/2001
'Description: Returns a datatype of a specific paramater in a crystal report
'Parameters: strReportName - name of crystal report
'            lngParamNum - Parameter number to check
'Returns: datatype name (string)
'--------------------------------------------------------------------------------
'Revision History:                                                              '
'                                                                               '
'--------------------------------------------------------------------------------
    Select Case strReportName
        Case "rptProviderDenial", "rptOutstandingPatAcct", "rptPatientInvoice", "rptProgressNote", _
             "rptMisdirectedPmtRejected", "rptNewProviderStats", "rptAgedClaims", "rptHCFA", "rptUB04"
            ParamType = "Long"
        Case "rptBookClosing", "rptProviderAR", "rptCommission", "rptCollectOrigins", "rptCollectionsHistogram", "rptProjectedRevenue", "rptClaimCount", "rptEmployeeStats", "rptEmployeeStatsDetail", "rptMPPostings"
            ParamType = "Date"
        Case "rptSuperbill", "rptPayerSummary", "rptARAgingProviders"
            If lngParamNum = 1 Then ParamType = "Date" Else ParamType = "Long"
        Case "rptCPCSummary", "rptPostingHistory", "rptProviderIncome", "rptBillingAccountDetail", "rptPatientPaymentLog", "rptPayerSessions", "rptSvcGrpPostings", "rptDenialLog", "rptWriteoffLog"
            If lngParamNum < 3 Then ParamType = "Date" Else ParamType = "Long"
        Case "rptHCFAReprint"
            If lngParamNum > 1 Then ParamType = "Date" Else ParamType = "Long"
        Case "rptOutInsSummary"
            If lngParamNum = 3 Or lngParamNum = 4 Or lngParamNum = 7 Then ParamType = "Long" Else ParamType = "Date"
        Case "rptBilledContacts"
            If lngParamNum < 3 Then
                ParamType = "Date"
            ElseIf lngParamNum = 4 Then
                ParamType = "String"
            Else
                ParamType = "Long"
            End If
        Case "rptARAgingSummary"
            If lngParamNum = 2 Then
                ParamType = "Date"
            Else
                ParamType = "Long"
            End If
        Case "rptARPatientAgingDetail"
            If lngParamNum <= 4 Then
                ParamType = "Long"
            Else
                ParamType = "Date"
            End If
    End Select
End Function
Public Function getNextParam(Optional ByVal strRptParams As String = "", Optional ByVal blnReset As Boolean = False) As String
'-------------------------------------------------------------------------------
'Author: Eric Pena
'Date: 05/14/2001
'Description: This procedure returns the next paramater in the given delimited string
'Parameters: strRptParams - the delimited string to use (if blnReset is set)
'               blnReset - true to reset static parameter string, false otherwise
'Returns: The next parameter
'-------------------------------------------------------------------------------
'Revision History:
'-------------------------------------------------------------------------------
    Static strParams As String
    Dim lngLastPos As Long
    
    If blnReset Then strParams = strRptParams
    
    lngLastPos = InStr(strParams, ";")
    If lngLastPos < 1 Then
        If InStr(2, strParams, ";") > 0 Then
            'there are more params, this one is ''
            getNextParam = ""
        Else
            getNextParam = strParams
            strParams = ""
        End If
    Else
        getNextParam = Left(strParams, lngLastPos - 1)
    End If
    strParams = Right(strParams, Len(strParams) - lngLastPos)
End Function
