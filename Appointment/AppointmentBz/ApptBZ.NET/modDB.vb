Option Strict Off
Option Explicit On
Module modDB
	'-------------------------------------------------------------------------------
	'Module Name: modDB
	'Author: Dave Richkun
	'Date: 11/04/1999
	'Description: This module is intended to encapsulate generic database routines
	'             that may be used in any data-aware application.
	'-------------------------------------------------------------------------------
	'Revision History:
	'
	'-------------------------------------------------------------------------------
	
	
	Public Enum DataTypes
		typString = 1
		typNumber = 2
		typDate = 3
	End Enum
	
	
	Public Function ParseSQL(ByVal strValue As String) As String
		'-----------------------------------------------------------------------------------
		'Author: Dave Richkun
		'Date: 03/17/1998
		'Description: When inserting or updating text values that contain a single quote
		'             into many database tables. i.e. "Can't do this", if the single quote
		'             is not prefixed with another single quote, the SQL statement will fail.
		'             This procedure searches for single quotes in the strValue parameter
		'             and precedes each quote with an additional quote.  e.g. "Can''t do this"
		'Parameters:  strValue - The string that will be searched for single quotes
		'Returns:     The modified string value formatted with consecutive single quotes
		'             where required.
		'-----------------------------------------------------------------------------------
		'Revision History:
		'
		'-----------------------------------------------------------------------------------
		
		Dim intPos As Short
		Dim strCopy As String
		
		strCopy = strValue 'Copy parameter to a local variable
		
		intPos = InStr(1, strCopy, "'", 1) 'search for single quote
		While intPos <> 0
			strCopy = Mid(strCopy, 1, intPos) & "'" & Mid(strCopy, intPos + 1) ' append quote
			intPos = InStr(intPos + 2, strCopy, "'", 1) 'search for another single quote
		End While
		
		ParseSQL = strCopy 'return parsed string back to calling routine
		
	End Function
	
	
	Public Function IfNull(ByVal varValue As Object, ByVal strReplacement As String) As String
		'-----------------------------------------------------------------------------------
		'Author: Dave Richkun
		'Date: 03/17/1998
		'Description: Replaces a NULL value with a string value.
		'Parameters:  varValue - A variant data type that is checked for a value of NULL.
		'             strReplacement - The string that will replace the original value, if the
		'               value contains NULL.
		'Returns:     The strReplacement parameter only if varValue is identified as NULL,
		'             otherwise varValue is returned.
		'-----------------------------------------------------------------------------------
		'Revision History:
		'
		'-----------------------------------------------------------------------------------
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(varValue) Then
			IfNull = strReplacement
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			IfNull = varValue
		End If
		
	End Function
	
	
	
	Public Function IfNull2(ByVal varValue As Object, ByVal strReplacement As String, Optional ByVal lngDataType As DataTypes = 0) As String
		'-----------------------------------------------------------------------------------
		'Author: Dave Richkun
		'Date: 02/01/2000
		'Description: Replaces a NULL value with a string value.  This function was designed
		'               specifically to accomodate the various challenges dealing with data
		'               conversion from the Psyquel MS-Access database to the SQL Server
		'               database.
		'Parameters:  varValue - A variant data type that is checked for a value of NULL.
		'             strReplacement - The string that will replace the original value, if the
		'               value contains NULL.
		'             lngDataType - Identifies the data type of the value being checked.  This
		'               parameter determines if values are to be enclosed within single quotes.
		'Returns:     The strReplacement parameter only if varValue is identified as NULL,
		'             otherwise varValue is returned.
		'-----------------------------------------------------------------------------------
		'Revision History:
		'
		'-----------------------------------------------------------------------------------
		
		If lngDataType = DataTypes.typDate Then
			'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If varValue = 0 Then
				IfNull2 = strReplacement
				Exit Function
				'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf varValue < "01/01/1753" Then  'Enforce SmallDateTime data type restriction
				IfNull2 = strReplacement
				Exit Function
			End If
		End If
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(varValue) Or IsNothing(varValue) Then
			IfNull2 = strReplacement
		Else
			If lngDataType = DataTypes.typDate Or lngDataType = DataTypes.typString Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				IfNull2 = "'" & ParseSQL(varValue) & "'"
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				IfNull2 = varValue
			End If
		End If
		
	End Function
End Module