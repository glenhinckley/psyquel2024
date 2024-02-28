VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBenefactorClinicDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'--------------------------------------------------------------------
'Class Name: CBenefactorClinicDB
'Date: 02/06/2000
'Author: Eric Pena
'Description:  MTS object designed to allow insertions into the tblBenefactorClinic table.
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------

Option Explicit
Private Const TABLE_NAME As String = "tblBenefactorClinic"
Private Const CLASS_NAME As String = "CBenefactorClinicDB"
Public Function FetchClinics(ByVal lngBenefactorID As Long) As Variant
'--------------------------------------------------------------------
'Date: 02/06/2000
'Author: Eric Pena
'Description:  Returns a variant array containing a list of all Clinics to which a Benefactor is associated.
'Parameters: lngBenefactorID - The ID of the Benefactor
'Returns: A variant array containing a list of all clinics to which a Benefactor is associated.
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    
    Dim cnnSQL As ADODB.Connection
    Dim rstSQL As ADODB.Recordset
    Dim strSQL As String
    Dim strList As String
    Dim intCTR As Integer
    Dim arrList() As String
    Dim varList As Variant

    On Error GoTo ErrTrap:

    

    'Prepare the SQL statement.
    strSQL = "SELECT "
    strSQL = strSQL & "fldClinicID "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "tblBenefactorClinic "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & "fldBenefactorID = " & lngBenefactorID
    
    'Instantiate and populate the Recordset.
    Set rstSQL = New ADODB.Recordset
    rstSQL.CursorLocation = adUseClient
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    Call rstSQL.Open(strSQL, cnnSQL, adOpenForwardOnly, adLockReadOnly, adCmdText + adAsyncFetch)
    
    'Disconnect the recordset, close the connection and return the recordset
    'to the calling environment.
    Set rstSQL.ActiveConnection = Nothing
    cnnSQL.Close
    Set cnnSQL = Nothing
    
    'set to null if there are no current associations (new Benefactor)
    If rstSQL.RecordCount = 0 Then
        FetchClinics = Null
        GoTo EMPTY_RECORDSET
    End If
    
    'Parse the recordset and build the comma-separated list.
    ReDim arrList(rstSQL.RecordCount - 1)
    For intCTR = 0 To rstSQL.RecordCount - 1
        arrList(intCTR) = rstSQL.Fields("fldClinicID").Value
        rstSQL.MoveNext
    Next intCTR

    varList = arrList
    FetchClinics = varList
    
EMPTY_RECORDSET:
    'Signal successful completion
    GetObjectContext.SetComplete
    Set cnnSQL = Nothing
    Set rstSQL = Nothing
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set rstSQL = Nothing
    varList = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Public Sub Insert(ByVal lngBenefactorID As Long, ByVal rstClinics As ADODB.Recordset, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 01/04/2000
'Author: Dave Richkun
'Description:  Inserts one or more rows into the tblBenefactorClinic table.
'Parameters: lngBenefactorID - ID of the Benefactor who is being associated to one or more clinics
'                   varClinics - Variant array containing the IDs of each clinic to which the Benefactor is associated.
'Returns: Null
'--------------------------------------------------------------------

    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim strSQL As String
    Dim intCTR As Integer
    Dim varCurrClinics As Variant
    Dim blnFound As Boolean
    
    On Error GoTo ErrTrap:
    
    
    
    'get list of currently associated clinics
    varCurrClinics = FetchClinics(lngBenefactorID)
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Instantiate and prepare the Command object.
    While Not rstClinics.EOF
        blnFound = False
        'if there are no current clinics, insert all
        If IsNull(varCurrClinics) Then
            GoTo NO_ASSOC
        End If
        'for each selected clinic, find out if it already exists
        For intCTR = 0 To UBound(varCurrClinics)
            If CInt(rstClinics.Fields("fldClinicID").Value) = CInt(varCurrClinics(intCTR)) Then
                blnFound = True
            End If
        Next intCTR
        'if it exists, make it enabled
        If blnFound Then
            Deleted False, lngBenefactorID, rstClinics.Fields("fldClinicID").Value, strUserName
        Else
NO_ASSOC:
            'add the association if it is not found
            Set cmdSQL = New ADODB.Command
            With cmdSQL
                .CommandType = adCmdText
                .CommandText = "uspInsBenefactorClinic"
                .CommandType = adCmdStoredProc
                .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngBenefactorID)
                .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , rstClinics.Fields("fldClinicID").Value)
                'TODO: Add the real parameter
                .Parameters.Append .CreateParameter("@LocalID", adVarChar, adParamInput, 20, Null)
                .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 20, strUserName)
                .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
            End With
            Set cmdSQL.ActiveConnection = cnnSQL
            cmdSQL.Execute , , adExecuteNoRecords
            'Check the ErrorNum parameter before deciding to commit the transaction
            If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
                GoTo ErrTrap
            End If
        End If
        rstClinics.MoveNext
    Wend
    
    'disable associations that were just unassociated by the user
    If Not IsNull(varCurrClinics) Then
        For intCTR = 0 To UBound(varCurrClinics)
            blnFound = False
            rstClinics.MoveFirst
            While Not rstClinics.EOF
                If CInt(varCurrClinics(intCTR)) = CInt(rstClinics.Fields("fldClinicID").Value) Then
                    blnFound = True
                End If
                rstClinics.MoveNext
            Wend
            If Not blnFound Then
                Deleted True, lngBenefactorID, varCurrClinics(intCTR), strUserName
            End If
        Next intCTR
    End If
    
    'Close the connection and free all resources
    cnnSQL.Close
EMPTY_ARRAY:
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    'Signal successful completion
    GetObjectContext.SetComplete
    Exit Sub
ErrTrap:
    'The Clinic array may be empty.
    If Err.Number = 9 Then
        GoTo EMPTY_ARRAY
    Else
        'Signal incompletion and raise the error to the calling environment.
        Set cmdSQL = Nothing
        Set cnnSQL = Nothing
        GetObjectContext.SetAbort
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub
Public Function IsEnabled(ByVal lngBenefactorID As Long, ByVal lngClinicID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 02/06/2000
'Author: Eric Pena
'Description:  Lets the caller know if a given Benefactor/clinic association is active
'Parameters: lngBenefactorID - The ID of the Benefactor
'                  lngClinicID - The ID of the Clinic
'Returns: T if the association is enabled/exists, F otherwise
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    
    Dim cnnSQL As ADODB.Connection
    Dim rstSQL As ADODB.Recordset
    Dim strSQL As String
    Dim strList As String
    Dim intCTR As Integer
    Dim arrList() As String
    Dim varList As Variant
    
    IsEnabled = False
    On Error GoTo ErrTrap:

    

    'Prepare the SQL statement.
    strSQL = "SELECT "
    strSQL = strSQL & "fldDisabledYN "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "tblBenefactorClinic "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & "fldBenefactorID = " & lngBenefactorID & " "
    strSQL = strSQL & "AND fldClinicID = " & lngClinicID
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Instantiate and populate the Recordset.
    Set rstSQL = New ADODB.Recordset
    rstSQL.CursorLocation = adUseClient
    Call rstSQL.Open(strSQL, cnnSQL, adOpenForwardOnly, adLockReadOnly, adCmdText + adAsyncFetch)
    
    'Disconnect the recordset, close the connection and return the recordset
    'to the calling environment.
    Set rstSQL.ActiveConnection = Nothing
    cnnSQL.Close
    Set cnnSQL = Nothing
    
    'set to null if there are no current associations (new Benefactor)
    If rstSQL.RecordCount = 0 Then
        GoTo EMPTY_RECORDSET
    End If
    If rstSQL.Fields("fldDisabledYN").Value = "N" Then
        IsEnabled = True
    End If
    
EMPTY_RECORDSET:
    'Signal successful completion
    GetObjectContext.SetComplete
    Set cnnSQL = Nothing
    Set rstSQL = Nothing
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub Deleted(ByVal blnDeleted As Boolean, ByVal lngBenefactorID As Long, ByVal lngClinicID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 3/01/2000
'Author: Eric Pena
'Description:  Flags a row in the tblBenefactor table marking the row as
'              deleted or undeleted.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngID - ID of the row in the table whose value will be
'               updated.
'            strUserName - Login name of the user responsible for
'               marking the row as deleted.
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim strSQL As String
    Dim blnFlag As String

    On Error GoTo ErrTrap:

    

    If blnDeleted = True Then
        blnFlag = "Y"
        'Prepare the SQL statement
        strSQL = "UPDATE "
        strSQL = strSQL & TABLE_NAME
        strSQL = strSQL & " SET "
        strSQL = strSQL & " fldDisabledYN = '" & blnFlag & "', "
        strSQL = strSQL & " fldDateDisabled = '" & Now & "', "
        strSQL = strSQL & " fldDisabledBy = '" & strUserName & "' "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "fldBenefactorID = " & lngBenefactorID & " AND "
        strSQL = strSQL & "fldClinicID = " & lngClinicID
        
    Else
        blnFlag = "N"
        'Prepare the SQL statement
        strSQL = "UPDATE "
        strSQL = strSQL & TABLE_NAME
        strSQL = strSQL & " SET "
        strSQL = strSQL & " fldDisabledYN = '" & blnFlag & "' "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "fldBenefactorID = " & lngBenefactorID & " AND "
        strSQL = strSQL & "fldClinicID = " & lngClinicID
    End If

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = strSQL
    cmdSQL.CommandType = adCmdText
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    'Close the connection and free all resources
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    'Signal successful completion
    GetObjectContext.SetComplete
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
