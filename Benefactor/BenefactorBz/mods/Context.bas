Attribute VB_Name = "Context"
Option Explicit
 
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" _
 (ByVal lpBuffer As String, nSize As Long) As Long
 

Public Sub CtxSetAbort()
    Dim ctx As ObjectContext
    Set ctx = GetObjectContext
    
    If Not (ctx Is Nothing) Then
        ctx.SetAbort
    End If
    Set ctx = Nothing
End Sub

Public Sub CtxSetComplete()
    Dim ctx As ObjectContext
    Set ctx = GetObjectContext
    
    If Not (ctx Is Nothing) Then
        ctx.SetComplete
    End If
    Set ctx = Nothing
End Sub

Public Function CtxCreateObject(ByVal sProgID As String) As Object
    Dim ctx As ObjectContext
    Set ctx = GetObjectContext
    
    If Not (ctx Is Nothing) Then
        Set CtxCreateObject = ctx.CreateInstance(sProgID)
        Set ctx = Nothing
    Else
        Set CtxCreateObject = CreateObject(sProgID)
    End If
End Function

Public Sub CtxRaiseError(module As String, functionName As String)
    ' Save the error information before calling CtxSetAbort in case it has side effects
    Dim lErr As Long
    Dim sErr As String
    lErr = VBA.Err.Number
    sErr = VBA.Err.Description
    
    CtxSetAbort
    
    Err.Raise lErr, SetErrSource(module, functionName), sErr
End Sub

'''Public Sub RaiseError(module As String, functionName As String)
'''    Err.Raise Err.Number, SetErrSource(module, functionName), Err.Description
'''End Sub


Function GetComputerName() As String
    ' Set or retrieve the name of the computer.
    Dim strBuffer As String
    Dim lngLen As Long
        
    strBuffer = Space(255 + 1)
    lngLen = Len(strBuffer)
    If CBool(GetComputerNameAPI(strBuffer, lngLen)) Then
        GetComputerName = Left$(strBuffer, lngLen)
    Else
        GetComputerName = ""
    End If
End Function


Function GetSource(ByVal pstrSource As String, ByVal pstrProc As String, ByVal pstrMod As String) As String
    Dim strFront As String
    Dim strBack As String
  
    strBack = pstrProc & "@" & GetComputerName()
  
    If Left$(pstrMod, (InStr(1, pstrMod, ":") - 1)) = pstrSource Then
        strFront = pstrSource & Mid$(pstrMod, InStr(1, pstrMod, ":"))
    Else
        strFront = "|" & pstrMod
    End If
  
    GetSource = strFront & strBack
End Function

Function SetErrSource(modName As String, procName As String) As String
    ' Returns an error message like:  "[FMStocks_DB.Account] VerifyUser [on AHI version 5.21.176]"
    
    SetErrSource = Err.Source & "<br>" & "[" & modName & "]  " & procName & _
            " [on " & GetComputerName() & " version " & GetVersionNumber() & "]"
End Function


Function GetVersionNumber() As String
    GetVersionNumber = App.Major & "." & App.Minor & "." & App.Revision
End Function



