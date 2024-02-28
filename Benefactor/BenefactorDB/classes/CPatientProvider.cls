VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatientProviderDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CProviderClinicDB
'Date: 02/14/2000
'Author: Rick "Boom Boom" Segura
'Description:  MTS object designed to host methods associated with
'              data affecting the tblProviderClinic table.
'--------------------------------------------------------------------
' Revision History:
'   R001 03/24/2000 Richkun: Added DisableProviderClinics to auto-disabled
'         Provider-Clinics that may have been de-selected from a list of
'         formerly active clinics
'   R002 01/31/2001 Nichol: Added optional parameter blnIncludeDisabled to
'         FetchPatientProviders()
'   R003 Valentines Day 2001!  Nichol: Added method InsertWeb() for insertion
'        of pat-provider association from the web page--that is, allows the
'        association to be seen on the web site by both providers.  Also added
'        methods DeletedByAssoc() and UpdatePatProviderWebAssoc()
'   R004 03/30/2001 D. Nichol: Added DeletedByProvider()
'   R005 04/10/2001 Richkun: Added strUserName to Deleted() method
'   R006 02/08/2002 Richkun: Added method FetchPatientUserProviders()
'--------------------------------------------------------------------

Option Explicit

Private Const CLASS_NAME As String = "CPatientProviderDB"
Private Const TABLE_NAME As String = "tblPatientProvider"
Public Function Insert(ByVal lngPatientID As Long, _
                       ByVal lngProviderID As Long, _
                       ByVal strAddedBy As String) As Long
'--------------------------------------------------------------------
'Date: 05/16/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Creates a provider-patient association                '
'Parameters: lngPatientID - Patient ID                              '
'            lngProviderID - Provider ID                            '
'Returns: lngPatientProviderID                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
' -- 12/26/2000 : Nichol : changed return type to long              '
'--------------------------------------------------------------------
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    
    On Error GoTo ErrTrap:

    Set cmdSQL = New ADODB.Command
    
    With cmdSQL
        .CommandText = "uspInsPatientProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientProviderID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 30, strAddedBy)
        .Parameters.Append .CreateParameter("@WebAssocYN", adVarChar, adParamInput, 1, "N")
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute ', , adExecuteNoRecords
    
    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If

    Insert = cmdSQL.Parameters("@PatientProviderID").Value

    cnnSQL.Close
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    GetObjectContext.SetAbort
    Call RaiseError(Err, cmdSQL.Parameters("@ErrorNum").Value, CLASS_NAME)
End Function
Public Function InsertWeb(ByVal lngPatientID As Long, _
                       ByVal lngProviderID As Long, ByVal lngUserID As Long, _
                       ByVal strAddedBy As String) As Long
'--------------------------------------------------------------------
'Date: 02/14/2001                                                   '
'Author: Dave Nichol                                                '
'Description: Creates a provider-patient association                '
'Parameters: lngPatientID - Patient ID                              '
'            lngProviderID - Provider ID                            '
'            lngUserID - User ID of inserting user                  '
'            strAddedBy - Username of inserting user                '
'Returns: lngPatientProviderID                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    
    On Error GoTo ErrTrap:

    Set cmdSQL = New ADODB.Command
    
    With cmdSQL
        .CommandText = "uspInsWebPatientProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientProviderID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
        .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 30, strAddedBy)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute ', , adExecuteNoRecords
    
    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If

    InsertWeb = cmdSQL.Parameters("@PatientProviderID").Value

    cnnSQL.Close
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    GetObjectContext.SetAbort
    Call RaiseError(Err, cmdSQL.Parameters("@ErrorNum").Value, CLASS_NAME)
End Function
Public Sub Deleted(ByVal blnDeleted As Boolean, ByVal lngID As Long, _
                   ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 05/25/2000
'Author: Eric Pena
'Description:  Flags a row in the tbPatientProvider table marking the row as
'              deleted or undeleted.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngID - ID of the row in the table whose value will be
'               updated.
'            strUserName - User name of the user responsible for deleting the
'               Patient-Provider relationship R005
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
    Else
        blnFlag = "N"
    End If

    'Prepare the SQL statement
    strSQL = "UPDATE "
    strSQL = strSQL & TABLE_NAME
    strSQL = strSQL & " SET "
    strSQL = strSQL & " fldDisabledYN = '" & blnFlag & "', "
    strSQL = strSQL & " fldDateDisabled = '" & Now() & "', "
    strSQL = strSQL & " fldDisabledBy = '" & strUserName & "' " 'R005
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " fldPatientProviderID = " & lngID

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
Public Sub DeletedByAssoc(ByVal blnDeleted As Boolean, ByVal lngPatientID As Long, ByVal lngProviderID As Long)
'--------------------------------------------------------------------
'Date: 03/30/2001
'Author: David Nichol
'Description:  Flags a row in the tbPatientProvider table marking the row as
'              deleted or undeleted, by patientID and providerID.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngPatientID - ID of the patient.
'            lngProviderID - ID of the provider
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
    Else
        blnFlag = "N"
    End If

    'Prepare the SQL statement
    strSQL = "UPDATE "
    strSQL = strSQL & TABLE_NAME
    strSQL = strSQL & " SET "
    strSQL = strSQL & " fldDisabledYN = '" & blnFlag & "', "
    strSQL = strSQL & " fldDateDisabled = '" & Date & "' "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " fldPatientID = " & lngPatientID & " AND "
    strSQL = strSQL & " fldProviderID = " & lngProviderID

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
Public Sub DeletedByProvider(ByVal blnDeleted As Boolean, ByVal lngProviderID As Long, ByVal lngOwnerID)
'--------------------------------------------------------------------
'Date: 03/30/2001
'Author: Dave Nichol & Joshua Lockwood (and a little bit of Segura)
'Description:  Flags a row in the tbPatientProvider table marking the row as
'              deleted or undeleted, by providerID--deleting all the associations
'              where a given provider is the owner.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngProviderID - ID of the provider
'            lngOwnerID - ID of the owner of the patient.
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
    Else
        blnFlag = "N"
    End If

    'Prepare the SQL statement
    strSQL = "UPDATE "
    strSQL = strSQL & TABLE_NAME
    strSQL = strSQL & " SET "
    strSQL = strSQL & " fldDisabledYN = '" & blnFlag & "', "
    strSQL = strSQL & " fldDateDisabled = '" & Date & "' "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " fldProviderID = " & lngProviderID & " AND "
    strSQL = strSQL & " fldPatientID IN ( SELECT fldBenefactorID FROM tblBenefactor WHERE fldOwnerID = " & lngOwnerID & ")"

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

Public Function FetchPatientProviders(ByVal lngPatientID As Long, _
                                      Optional ByVal blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 05/25/2000
'Author: Eric Pena
'Description:  Retrieves all Providers associated with a specific patient.
'Parameters: lngPatientID - ID of the patient whose Provider records will be
'                 retrieved.
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command

    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    With cmd
        .CommandText = "uspSelPatientProviders"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = False, "N", "Y"))
    End With
    
    'Open the connection and execute the stored procedure
    Set cnn = New ADODB.Connection
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchPatientProviders = rst
    
    'Disconnect the recordset
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set cnn = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function FetchPatientUserProviders(ByVal lngPatientID As Long, _
                ByVal lngUserID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 02/08/2002
'Author: Dave Richkun
'Description:  Retrieves all Providers associated with a specific patient where
'               the provider is also serviced by the user (Provider/Office Manager).
'Parameters: lngPatientID - ID of the patient whose Provider records will be retrieved.
'            lngUserID - ID of the User (Provider/Office Manager)
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command

    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    With cmd
        .CommandText = "uspSelPatientUserProviders"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
    End With
    
    'Open the connection and execute the stored procedure
    Set cnn = New ADODB.Connection
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchPatientUserProviders = rst
    
    'Disconnect the recordset
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set cnn = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function FetchWebPatientProviders(ByVal lngPatientID As Long, _
                                      Optional ByVal blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 05/25/2000
'Author: Eric Pena
'Description:  Retrieves all Providers associated with a specific patient.
'Parameters: lngPatientID - ID of the patient whose Provider records will be
'                 retrieved.
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command

    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    With cmd
        .CommandText = "uspSelWebPatientProviders"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = False, "N", "Y"))
    End With
    
    'Open the connection and execute the stored procedure
    Set cnn = New ADODB.Connection
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchWebPatientProviders = rst
    
    'Disconnect the recordset
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set cnn = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function Exists(ByVal lngPatientID As Long, ByVal lngProviderID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 05/25/2000
'Author: Eric Pena
'Description:  Determines if the association identical to the parameters already exists in the table.
'Parameters:  all are values to be checked
'Returns: True if the assoc exists, false otherwise
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    
    Dim cnnSQL As ADODB.Connection
    Dim rstSQL As ADODB.Recordset
    Dim strSQL As String
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:

    

    'Prepare the SQL statement
    strSQL = "SELECT "
    strSQL = strSQL & "COUNT(*) AS TOTAL "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & TABLE_NAME
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " fldPatientID = " & lngPatientID & " AND "
    strSQL = strSQL & " fldProviderID = " & lngProviderID
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Instantiate and populate the Recordset.
    Set rstSQL = New ADODB.Recordset
    rstSQL.CursorLocation = adUseClient
    Call rstSQL.Open(strSQL, cnnSQL, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch)
    
    'Disconnect the recordset, close the connection and return the recordset
    'to the calling environment.
    Set rstSQL.ActiveConnection = Nothing

    'Close the connection and free all resources
    cnnSQL.Close
    Set cnnSQL = Nothing

    'Get the value from the Recordset then destroy it.
    If rstSQL.Fields("TOTAL").Value > 0 Then
        Exists = True
    Else
        Exists = False
    End If
    Set rstSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set rstSQL = Nothing
    Set cnnSQL = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub UpdatePatProviderWebAssoc(ByVal blnWebAssoc As Boolean, ByVal lngPatientID As Long, _
                                     ByVal lngProviderID As Long, ByVal lngUserID As Long)
'--------------------------------------------------------------------
'Date: 02/14/2001
'Author: David Nichol
'Description:  Updates tblPatientProvider.fldWebAssocYN for a patient/provider
'               pair utilizing a stored procedure.
'Parameters: Pretty self-explanatory.
'Returns: Null
'--------------------------------------------------------------------
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    
    With cmdSQL
        .CommandText = "uspUpdPatientProviderWebAssoc"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@WebAssocYN", adVarChar, adParamInput, 1, IIf(blnWebAssoc = True, "Y", "N"))
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords
    
    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@ErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If

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
    Call RaiseError(Err, cmdSQL.Parameters("@ErrorNum").Value, CLASS_NAME, cmdSQL.Parameters("@ErrorMsg").Value)
    
End Sub

