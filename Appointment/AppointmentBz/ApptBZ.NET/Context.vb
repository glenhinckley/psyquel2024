Option Strict Off
Option Explicit On
Module Context
	
	Private Declare Function GetComputerNameAPI Lib "kernel32"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	
	
	Public Sub CtxSetAbort()
		Dim ctx As System.EnterpriseServices.ContextUtil
		'UPGRADE_ISSUE: COMSVCSLib.AppServer  .GetObjectContext was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		ctx = GetObjectContext
		
		If Not (ctx Is Nothing) Then
			ctx.SetAbort()
		End If
		'UPGRADE_NOTE: Object ctx may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ctx = Nothing
	End Sub
	
	Public Sub CtxSetComplete()
		Dim ctx As System.EnterpriseServices.ContextUtil
		'UPGRADE_ISSUE: COMSVCSLib.AppServer  .GetObjectContext was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		ctx = GetObjectContext
		
		If Not (ctx Is Nothing) Then
			ctx.SetComplete()
		End If
		'UPGRADE_NOTE: Object ctx may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ctx = Nothing
	End Sub
	
	Public Function CtxCreateObject(ByVal sProgID As String) As Object
		Dim ctx As System.EnterpriseServices.ContextUtil
		'UPGRADE_ISSUE: COMSVCSLib.AppServer  .GetObjectContext was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		ctx = GetObjectContext
		
		If Not (ctx Is Nothing) Then
			'UPGRADE_ISSUE: COMSVCSLib.ObjectContext method ctx.CreateInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			CtxCreateObject = ctx.CreateInstance(sProgID)
			'UPGRADE_NOTE: Object ctx may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ctx = Nothing
		Else
			CtxCreateObject = CreateObject(sProgID)
		End If
	End Function
	
	'UPGRADE_NOTE: module was upgraded to module_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub CtxRaiseError(ByRef module_Renamed As String, ByRef functionName As String)
		' Save the error information before calling CtxSetAbort in case it has side effects
		Dim lErr As Integer
		Dim sErr As String
		lErr = Err.Number
		sErr = Err.Description
		
		CtxSetAbort()
		
		Err.Raise(lErr, SetErrSource(module_Renamed, functionName), sErr)
	End Sub
	
	'''Public Sub RaiseError(module As String, functionName As String)
	'''    Err.Raise Err.Number, SetErrSource(module, functionName), Err.Description
	'''End Sub
	
	
	Function GetComputerName() As String
		' Set or retrieve the name of the computer.
		Dim strBuffer As String
		Dim lngLen As Integer
		
		strBuffer = Space(255 + 1)
		lngLen = Len(strBuffer)
		If CBool(GetComputerNameAPI(strBuffer, lngLen)) Then
			GetComputerName = Left(strBuffer, lngLen)
		Else
			GetComputerName = ""
		End If
	End Function
	
	
	Function GetSource(ByVal pstrSource As String, ByVal pstrProc As String, ByVal pstrMod As String) As String
		Dim strFront As String
		Dim strBack As String
		
		strBack = pstrProc & "@" & GetComputerName()
		
		If Left(pstrMod, InStr(1, pstrMod, ":") - 1) = pstrSource Then
			strFront = pstrSource & Mid(pstrMod, InStr(1, pstrMod, ":"))
		Else
			strFront = "|" & pstrMod
		End If
		
		GetSource = strFront & strBack
	End Function
	
	Function SetErrSource(ByRef modName As String, ByRef procName As String) As String
		' Returns an error message like:  "[FMStocks_DB.Account] VerifyUser [on AHI version 5.21.176]"
		
		SetErrSource = Err.Source & "<br>" & "[" & modName & "]  " & procName & " [on " & GetComputerName() & " version " & GetVersionNumber() & "]"
	End Function
	
	
	Function GetVersionNumber() As String
		GetVersionNumber = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
	End Function
End Module