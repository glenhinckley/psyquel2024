VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatApptDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatApptDB                                             '
'Date: 08/25/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  MTS object designed to host methods associated with  '
'              data affecting the tblPatAppt table.                 '
'--------------------------------------------------------------------
' Revision History:
'  R001: Richkun 11/06/2001: Added FetchCheckInDetails() method
'  R002: Richkun 02/13/2002: Added FetchByApptPatient() method
'  R003: Richkun 02/15/2002: Added Reset() method
'--------------------------------------------------------------------

Option Explicit

Private Const CLASS_NAME As String = "CPatApptDB"
Private Const TABLE_NAME As String = "tblPatientAppt"

'--------------------------------------------------------------------
' Public Methods
'--------------------------------------------------------------------

Public Function Insert(ByVal lngApptID As Long, ByVal lngPatientID As Long) As Long
'--------------------------------------------------------------------
'Date: 07/13/2002
'Author: Dave Richkun
'Description:  Inserts a record into tblPatientAppt
'Parameters:  lngApptID - ID of parent appointment
'             lngPatientID - ID of patient
'Returns:  ID of row inserted on success, -1 on failure
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspInsPatAppt"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Assign the connection to the Command object and execute the stored procedure
    Call cnn.Open(CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    cmd.Execute ', , adExecuteNoRecords

    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmd.Parameters("@ErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
    
    Insert = cmd.Parameters("@ID").Value
    
    'Free resources
    Set cnn = Nothing
    Set cmd = Nothing
    
    ' Signal Completion
    GetObjectContext.SetComplete
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Function

Public Function Update(ByVal lngID As Long, ByVal lngApptID As Long, _
                       ByVal lngPatientID As Long)
'--------------------------------------------------------------------
'Date: 08/25/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Updates a row into the tblPatAppt table utilizing    '
'              a stored procedure.                                  '
'Parameters:  lngID - ID of the row in the table whose value will be'
'               updated.                                            '
'             lngApptID - The Appointment ID of the patient/appointment
'               association to which the record will be changed.    '
'             lngPatientID - The Patient ID of the patient/appointment
'               association to which the record will be changed.    '
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim prmID As ADODB.Parameter
    Dim prmDesc As ADODB.Parameter
    Dim prmDisabled As ADODB.Parameter
    Dim varParamErr As Variant

    On Error GoTo ErrTrap:

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    
    With cmdSQL
        .CommandText = "uspUpdPatAppt"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, , lngID)
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With

    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    'Check for errors
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        varParamErr = cmdSQL.Parameters("@SQLErrorNum").Value
        GoTo ErrTrap
    End If

    'Close the connection and free all resources
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Call RaiseError(Err, varParamErr, CLASS_NAME)
    End If
    
End Function

Public Function FetchByAppt(ByVal lngApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Retrieves records from tblPatAppt that are associated'
'               with the appointment whose ID matches the parameter '
'Parameters: lngApptID - ID of the appointment to find related      '
'               records for                                         '
'Returns: Recordset from tblPatientAppt                             '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelPatientApptByAppt"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByAppt = rst
    
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
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function FetchByID(ByVal lngPatApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 08/31/2000
'Author: Dave Richkun
'Description: Returns information associated with a patient appt record
'Parameters:  lngPatApptID - System ID of the record to return
'Returns:  Recordset
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim rstSQL As ADODB.Recordset
    Dim strSQL As String
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:

    Set rstSQL = New ADODB.Recordset
    Set cmdSQL = New ADODB.Command
    Set cnnSQL = New ADODB.Connection
    
    'Acquire the database connection.
    cnnSQL.Open (CONST_PSYQUEL_CNN)
    Set cmdSQL.ActiveConnection = cnnSQL
    
    'Create the parameter objects
    With cmdSQL
        .CommandText = "uspSelPatApptByID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatApptID", adInteger, adParamInput, , lngPatApptID)
    End With
    
    'Execute the stored procedure
    rstSQL.CursorLocation = adUseClient
    rstSQL.Open cmdSQL, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByID = rstSQL
    
    'Disconnect the recordset
    Set cmdSQL.ActiveConnection = Nothing
    Set cmdSQL = Nothing
    Set rstSQL.ActiveConnection = Nothing
    Set cnnSQL = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

    
Public Function FetchByApptPatient(ByVal lngApptID As Long, ByVal lngPatientID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 02/13/2002
'Author: Dave Richkun
'Description:  Retrieves appointment information using ApptID, PatientID
'Parameters: lngApptID - ID of the appointment
'            lngPatientID - ID of the patient
'Returns: Recordset
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelApptByApptIDPatient"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByApptPatient = rst
    
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
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
    
    
    
Public Function FetchMissingRec(ByVal lngApptID As Long, _
                                ByVal strINClause) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 08/31/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Retrieves a recordset of PatAppt records associated   '
'               with an appointment, but are not in the list given  '
'               passed in the second parameter                      '
'Parameters:  lngApptID - ID of the appointment being compared      '
'             strInClause - Concatenated string of PatAppt IDs that '
'               are to remain associated with the appointment       '
'Returns: Recordset of PatAppt records to be considered for deletion'
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnnSQL As ADODB.Connection
    Dim rstSQL As ADODB.Recordset
    Dim strSQL As String
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:

    'Prepare the SQL statement
    strSQL = "SELECT "
    strSQL = strSQL & " fldPatApptID, "
    strSQL = strSQL & " fldApptStatusID, "
    strSQL = strSQL & " fldPatientID "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & TABLE_NAME
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " fldApptID = " & lngApptID
    strSQL = strSQL & " AND "
    strSQL = strSQL & " fldPatApptID NOT IN ("
    strSQL = strSQL & strINClause & ")"
    
    'Instantiate the Recordset.
    Set rstSQL = New ADODB.Recordset
    rstSQL.CursorLocation = adUseClient
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    ' Populate the recordset
    Call rstSQL.Open(strSQL, cnnSQL, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch)
    
    Set FetchMissingRec = rstSQL
    
    'Disconnect the recordset, close the connection and return the recordset
    'to the calling environment.
    Set rstSQL.ActiveConnection = Nothing

    'Close the connection and free all resources
    cnnSQL.Close
    Set cnnSQL = Nothing
    Set rstSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function Delete(ByVal lngApptID As Long, ByVal lngPatientID)
'--------------------------------------------------------------------
'Date: 08/25/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Deletes an  record into tblPatAppt                   '
'Parameters:  lngApptID - System ID of appointment associated with  '
'               the patient.                                        '
'             lngPatientID - System ID of the patient associated    '
'               with the appointment.                               '
'Returns:  Null                                                     '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    Dim varParamErr As Variant
    
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspDelPatAppt"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
        
        'Assign the connection to the Command object and execute the stored procedure
        Call cnn.Open(CONST_PSYQUEL_CNN)
        Set .ActiveConnection = cnn
        .Execute ', , adExecuteNoRecords
        Set cnn = Nothing
        
        'Check the ErrorNum parameter before deciding to commit the transaction
        If .Parameters("@ErrorNum").Value <> 0 Then
            varParamErr = .Parameters("@ErrorNum").Value
            GoTo ErrTrap
        End If
        Set cmd = Nothing
    End With
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Call RaiseError(Err, varParamErr, CLASS_NAME)
    End If
End Function

Public Function Exists(ByVal lngApptID As Long, ByVal lngPatientID) As Boolean
'--------------------------------------------------------------------
'Date: 08/25/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Determines if an Appointment/Patient association      '
'               exists in tblPatAppointment                         '
'Parameters:  lngApptID - System ID of appointment associated with  '
'               the patient.                                        '
'             lngPatientID - System ID of the patient associated    '
'               with the appointment.                               '
'Returns:  TRUE if association is found, FALSE otherwise            '
'--------------------------------------------------------------------
'Revision History:                                                  '
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
    strSQL = strSQL & " fldApptID = " & lngApptID
    strSQL = strSQL & " AND "
    strSQL = strSQL & " fldPatientID = " & lngPatientID
    
    'Instantiate the Recordset.
    Set rstSQL = New ADODB.Recordset
    rstSQL.CursorLocation = adUseClient
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    ' Populate the recordset
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
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ChangeStatus(ByVal lngPatApptID As Long, ByVal lngApptStatusID As Long, _
        ByVal strNoShowFlag As String, ByVal dblCancelFee As Double, _
        ByVal strCancelExplain As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 09/01/2000
'Author: Dave Richkun
'Description:  Updates the status of a single Patient-Appointment record
'Parameters:  lngPatApptID - ID of the row in tblPatientAppt whose status value will be updated
'             lngApptStatusID - The value representing the updated appointment status
'             strNoShowFlag - If appointment status is 'No-Show', represents the user's decision
'                   to manage the No-Show.  May be one of:
'                    N - Do not Bill for No-Show
'                    P - Bill patient for No-Show
'                    I - Bill Insurance Company for No-Show
'             dblCancelFee - If appointment status is 'Cancelled', represents the amount
'                   charged to patient as Cancellation Fee
'             strCancelExplain - If appointment is 'Cancelled', represents the user's explanation
'                   for cancellation (assumed to be given by patient)
'             strUserName - User performng action
'Returns:  Null
'--------------------------------------------------------------------
'Revision History:
'  03/15/2001 Segura:  Added new parameter(strNoShowFlag)
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
  
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspUpdPatApptStatus"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatApptID", adInteger, adParamInput, , lngPatApptID)
        .Parameters.Append .CreateParameter("@ApptStatusID", adInteger, adParamInput, , lngApptStatusID)
        .Parameters.Append .CreateParameter("@NoShowFlag", adVarChar, adParamInput, 1, IIf(strNoShowFlag = "", Null, strNoShowFlag))
        .Parameters.Append .CreateParameter("@CancelFee", adCurrency, adParamInput, , dblCancelFee)
        .Parameters.Append .CreateParameter("@CancelExplain", adVarChar, adParamInput, 90, IIf(strCancelExplain = "", Null, strCancelExplain))
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
    End With
        
    'Assign the connection to the Command object and execute the stored procedure
    Call cnn.Open(CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    cmd.Execute ', , adExecuteNoRecords
    Set cnn = Nothing
    
    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmd.Parameters("@ErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
    Set cmd = Nothing

    GetObjectContext.SetComplete
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ChangeStatusByApptPat(ByVal lngApptID As Long, _
                                 ByVal lngPatientID As Long, _
                                 ByVal lngStatusID As Long, _
                                 ByVal strNoShowFlag As String, _
                                 ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 10/02/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Updates the value of the Status ID of a single       '
'               Patient/Appointment record                          '
'Parameters:  lngApptID - ID of the associated Appointment          '
'             lngPatientID - ID of the associated Patient           '
'             lngStatusID - The new value of the Status ID          '
'             strNoShowFlag -  NoShow Flag                          '
'             strUserName - User Login Name                         '
'Returns:   Nothing                                                 '
'--------------------------------------------------------------------
'Revision History:                                                  '
'  03/15/2001 Segura:  Added new parameter(strNoShowFlag)           '
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    Dim varParamErr As Variant
    
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspUpdPatApptStatusByApptPat"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@StatusID", adInteger, adParamInput, , lngStatusID)
        If strNoShowFlag > "" Then
            .Parameters.Append .CreateParameter("@NoShowFlag", adVarChar, adParamInput, 1, strNoShowFlag)
        Else
            .Parameters.Append .CreateParameter("@NoShowFlag", adVarChar, adParamInput, 1, Null)
        End If
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
        
        'Assign the connection to the Command object and execute the stored procedure
        Call cnn.Open(CONST_PSYQUEL_CNN)
        Set .ActiveConnection = cnn
        .Execute ', , adExecuteNoRecords
        Set cnn = Nothing
        
        'Check the ErrorNum parameter before deciding to commit the transaction
        If .Parameters("@ErrorNum").Value <> 0 Then
            varParamErr = .Parameters("@ErrorNum").Value
            GoTo ErrTrap
        End If
        Set cmd = Nothing
    End With
    
    GetObjectContext.SetComplete
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Call RaiseError(Err, varParamErr, CLASS_NAME)
    End If
End Sub

Public Sub DeleteByID(ByVal lngPatApptID As Long)
'--------------------------------------------------------------------
'Date: 10/02/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Physically deletes a record from tblPatientAppt      '
'Parameters:  lngPatApptID - ID of the record to delete             '
'Returns:   Nothing                                                 '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    Dim varParamErr As Variant
    
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspDelPatApptByID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatApptID", adInteger, adParamInput, , lngPatApptID)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
        
        'Assign the connection to the Command object and execute the stored procedure
        Call cnn.Open(CONST_PSYQUEL_CNN)
        Set .ActiveConnection = cnn
        .Execute ', , adExecuteNoRecords
        Set cnn = Nothing
        
        'Check the ErrorNum parameter before deciding to commit the transaction
        If .Parameters("@ErrorNum").Value <> 0 Then
            varParamErr = .Parameters("@ErrorNum").Value
            GoTo ErrTrap
        End If
        Set cmd = Nothing
    End With
    
    GetObjectContext.SetComplete
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Call RaiseError(Err, varParamErr, CLASS_NAME)
    End If
End Sub


Public Function FetchCheckInDetails(ByVal lngPatApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 11/06/2001
'Author: Dave Richkun
'Description:  Retrieves detailed Appointment information for a specific
'              Patient-Appointment ID
'Parameters: lngPatApptID - ID of the patient-appointment
'Returns: ADO Recordset
'--------------------------------------------------------------------
'Revision History:
'  R001: Created
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelCheckInDetails"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatApptID", adInteger, adParamInput, , lngPatApptID)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchCheckInDetails = rst
    
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
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Public Function UpdateUnbilledCounter(ByVal lngProviderID As Long, ByVal lngValue As Long)
'--------------------------------------------------------------------
'Date: 01/09/2002                                                  '
'Author: Eric Pena                                   '
'Description:  Updates the unbilled counter in tblProvider
'Parameters:  lngProviderID - Provider whose counter will be updated
'             lngValue - Number to inc(dec)rement by (+- 1)
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim prmID As ADODB.Parameter
    Dim prmDesc As ADODB.Parameter
    Dim prmDisabled As ADODB.Parameter
    Dim varParamErr As Variant

    On Error GoTo ErrTrap:

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    
    With cmdSQL
        .CommandText = "uspUpdCtr_Unbilled"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@Value", adInteger, adParamInput, , lngValue)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With

    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    'Check for errors
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        varParamErr = cmdSQL.Parameters("@SQLErrorNum").Value
        GoTo ErrTrap
    End If

    'Close the connection and free all resources
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Call RaiseError(Err, varParamErr, CLASS_NAME)
    End If
    
End Function


Public Function Reset(ByVal lngPatApptID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 02/15/2002
'Author: Dave Richkun
'Description:  Resets a patient appointment to a status of 'Attended' and
'              reverses any prior Cancellation and No-Show cancellations
'              applied against the patient.
'Parameters:  lngPatApptID - ID of the patient appointmnt
'             strUserName - User name of the user initiating the method
'Returns: Null
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    
    With cmdSQL
        .CommandText = "uspUpdResetAppt"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatApptID", adInteger, adParamInput, , lngPatApptID)
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With

    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    'Check for errors
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If

    'Close the connection and free all resources
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Call RaiseError(Err, cmdSQL.Parameters("@ErrorNum").Value, CLASS_NAME)
    End If
    
End Function




'--------------------------------------------------------------------
' PrivateMethods    +++++++++++++++++++++++++++++++++++++++++++++++++
'--------------------------------------------------------------------




