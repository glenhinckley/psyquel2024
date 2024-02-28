VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApptDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CApptDB                                                '
'Date: 08/25/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  MTS object designed to host methods associated with  '
'              data affecting the Appointment tables.               '
'--------------------------------------------------------------------
'Revision History:
'  R001: 06/14/2001 - Richkun: Added FetchUnBilledAppts()
'  R002: 08/08/2001 - Richkun: Added support for tblAppointment columns
'         fldReferPhy, fldReferPhyID
'  R003: 10/04/2001 - Richkun: Allowed for deleting of single appointment
'         within series of recurring appointments; Required significant
'         interface changes.
'  R004:  11/01/2001 - Richkun: Added FetchByApptIDs() method
'  R005:  11/06/2001 - Richkun: Added FetchByCheckInDate() method
'  R006:  11/30/2001 - Richkun: Added FetchConflictsByProvider() method
'  R006:  03/07/2002 - Richkun: Added FetchFutureRecurPatientApptByProvider() method
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CApptDB"

'--------------------------------------------------------------------
' Public Methods
'--------------------------------------------------------------------

Public Function FetchByProviderDateRange(ByVal lngProviderID As Long, _
       ByVal dteStartDate As Date, ByVal dteEndDate As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 09/01/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:   Retrieves a recordset of apointments from tblAppointment
'               that matches the criteria given in  the parameter arguments                             '
'Parameters:    lngProviderID - ID of Provider whose schedule is being sought                                    '
'               dteStartDate, dteEndDate - the date range that limits the
'               appointments being returned                  '
'Returns:   Recordset of appointments                               '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelAllApptByProviderDate"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@StartDate", adDBTimeStamp, adParamInput, , dteStartDate)
        .Parameters.Append .CreateParameter("@EndDate", adDBTimeStamp, adParamInput, , dteEndDate)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByProviderDateRange = rst
    
    'Disconnect the recordset and clean house
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

Public Function FindOpenApptTimeSlots(ByVal lngClinicID As Long, ByVal lngProviderID As Long, _
       ByVal dteStartDate As Date, ByVal dteEndDate As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 04/20/2023                                                   '
'Author: Duane C Orth                                               '
'Description:   Retrieves a recordset of providers who have an open '
'               time slot for a given clinic or provider            '
'Parameters:    lngClinicID - ID of Clinic                          '
'               dteStartDate, dteEndDate - the date range that limits the
'               appointments being returned                         '
'Returns:   Recordset of providers                                  '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelFindOpenTimeSlot"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , lngClinicID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@StartDate", adDBTimeStamp, adParamInput, , dteStartDate)
        .Parameters.Append .CreateParameter("@EndDate", adDBTimeStamp, adParamInput, , dteEndDate)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FindOpenApptTimeSlots = rst
    
    'Disconnect the recordset and clean house
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

Public Function FetchByCheckInDate(ByVal lngUserID As Long, ByVal dteArrivalDate As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 11/06/2001
'Author: Dave Richkun
'Description:   Returns a recordset of one-time and recurring patient appointments
'               that fall on or within a specific date.  This function was designed
'               to support the Check-In feature.
'Parameters:    lngUserID - ID of User retrieving the information
'               dteArrivalDate - The date on which the appointments occur i.e. the check-in date
'Returns:   Recordset of appointments                               '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelExpectedPatients"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
        .Parameters.Append .CreateParameter("@ArriveDate", adDBTimeStamp, adParamInput, , dteArrivalDate)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByCheckInDate = rst
    
    'Disconnect the recordset and clean house
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


Public Function FetchPatientApptByID(ByVal lngApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 07/12/2002
'Author: Dave Richkun
'Description:   Retrieves a recordset of detailed patient and plan information
'               associated with a scheduled patient appointment
'Parameters:    lngApptID - ID of the appointment
'Returns:   Recordset of patient and plan information associated with appointment
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelApptPatientsPlans"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchPatientApptByID = rst
    
    'Disconnect the recordset and clean house
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

Public Function FetchBlockApptByID(ByVal lngApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 07/13/2002
'Author: Dave Richkun
'Description: Retrieves a recordset of detailed information associated
'             with a scheduled block appointment
'Parameters: lngApptID - ID of the appointment
'Returns: Recordset of information associated with appointment
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelApptBlock"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchBlockApptByID = rst
    
    'Disconnect the recordset and clean house
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


Public Function FetchByApptIDs(ByVal strApptIDs As String) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 11/01/2001
'Author: Dave Richkun
'Description:   Retrieves a recordset of appointment information for one or more
'               appointments whose IDs are known.  This function was designed to
'               assist is displaying appointment summary information for conflicting
'               appointments.
'Parameters:    strApptIDs - A comma separated list of appointment IDs whose
'                   information is to be retrieved.
'Returns:  Recordset of appointment information                    '
'--------------------------------------------------------------------
'Revision History:
'  R004: Created
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelConflictingAppts"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptIDs", adVarChar, adParamInput, 512, strApptIDs)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByApptIDs = rst
    
    'Disconnect the recordset and clean house
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


Public Function InsertSingle(ByVal lngProviderID As Long, ByVal lngClinicID As Variant, _
                       ByVal lngCategoryID As Long, ByVal dteStartDateTime As Date, _
                       ByVal dteEndDateTime As Date, ByVal lngDuration As Long, _
                       ByVal strCPTCode As String, ByVal strDescription As String, _
                       ByVal strNote As String, ByVal strUserName As String) As Long
'--------------------------------------------------------------------
'Date: 07/13/2002
'Author: Dave Richkun
'Description:  Inserts an appointment record into tblAppointment via a stored procedure
'Parameters:  lngProviderID - ID of the provider whose schedule the appointment will appear
'             lngClinicID - ID of the place of service where the appointment is scheduled
'             lngCategoryID - ID of the category identifying if appointment is 'Patient' or 'Block'
'             dteStartDateTime - Start Date and Time of the appointment
'             dteEndDateTime - End Date and Time of the appointment
'             lngDuration - Appointment length in minutes
'             strCPTCode - The CPTCode associated with a patient appointment
'             strDescription - Short description of the appointment
'             strNote - A longer description pertaining to the appointment
'             strUserName - Name of user responsible for adding appointment
'Returns: ID of appointment ID on success, -1 on failure
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
        .CommandText = "uspInsAppointment"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , IIf(lngClinicID = 0, Null, lngClinicID))
        .Parameters.Append .CreateParameter("@CategoryID", adInteger, adParamInput, , lngCategoryID)
        .Parameters.Append .CreateParameter("@StartDateTime", adDBTimeStamp, adParamInput, , dteStartDateTime)
        .Parameters.Append .CreateParameter("@EndDateTime", adDBTimeStamp, adParamInput, , dteEndDateTime)
        .Parameters.Append .CreateParameter("@Duration", adInteger, adParamInput, , lngDuration)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, IIf(strCPTCode = "", Null, strCPTCode))
        .Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, 30, strDescription)
        .Parameters.Append .CreateParameter("@Note", adVarChar, adParamInput, 256, strNote)
        .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 30, strUserName)
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
    
    InsertSingle = cmd.Parameters("@ID").Value
    
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
    
End Function

Public Function Update(ByVal lngApptID As Long, ByVal lngProviderID As Long, _
        ByVal lngClinicID As Long, ByVal lngCategoryID As Long, ByVal dteStartDateTime As Date, _
        ByVal dteEndDateTime As Date, ByVal lngDuration As Long, ByVal strCPTCode As String, _
        ByVal strDescription As String, ByVal strNote As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 08/30/2000
'Author: Dave Richkun
'Description:  Updates an appointment record in tblAppointment via a stored procedure
'Parameters:
'lngProviderID lngApptID - System ID of the provide for whom the appointment is for.
'              lngProviderID - ID of the provider whose schedule the appointment will appear
'              lngClinicID - ID of the place of service where the appointment is scheduled
'              lngCategoryID - ID of the category identifying if appointment is 'Patient' or 'Block'
'              dteStartDateTime - Start Date and Time of the appointment
'              dteEndDateTime - End Date and Time of the appointment
'              lngDuration - Appointment length in minutes
'              strCPTCode - The CPTCode associated with a patient appointment
'              strDescription - Short description of the appointment
'              strNote - A longer description pertaining to the appointment
'              strUserName - Name of user responsible for adding appointment
'Returns:  Nothing                                                  '
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
        .CommandText = "uspUpdAppointment"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , IIf(lngClinicID = 0, Null, lngClinicID))
        .Parameters.Append .CreateParameter("@CategoryID", adInteger, adParamInput, , lngCategoryID)
        .Parameters.Append .CreateParameter("@StartDateTime", adDBTimeStamp, adParamInput, , dteStartDateTime)
        .Parameters.Append .CreateParameter("@EndDateTime", adDBTimeStamp, adParamInput, , dteEndDateTime)
        .Parameters.Append .CreateParameter("@Duration", adInteger, adParamInput, , lngDuration)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, IIf(strCPTCode = "", Null, strCPTCode))
        .Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, 30, strDescription)
        .Parameters.Append .CreateParameter("@Note", adVarChar, adParamInput, 256, strNote)
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
    
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)

End Function


''Public Sub UpdateRecurSingle(ByVal lngApptID As Long, ByVal lngProviderID As Long, _
''        ByVal lngClinicID As Long, ByVal lngCategoryID As Long, ByVal dtApptDate As Date, _
''        ByVal lngDuration As Long, ByVal strDescription As String, ByVal strNote As String, _
''        ByVal strUserName As String)
'''--------------------------------------------------------------------
'Date: 10/08/2001
'Author: Dave Richkun
'Description: Updates a single appointment within a recurring series.  The action taken
'             consists of inserting a row in the recurring appointment exception
'             table (tblRecurApptExc) so that the appointment does not appear on the
'             calendar in that date and time, and then inserting a regular single
'             appointment in the same time slot.
'Parameters:  lngApptID - ID of the Recurring appointment to be updated
'             lngProviderID - ID of Provider for whom the appointment is associated
'             varClinicID - ID of the place of service where appointment is scheduled (may be NULL)
'             lngCategoryID - ID of the category that classifies the appointment
'             dtApptDate - Date and Time of the appointment being altered
'             lngDuration - Appointment length in minutes
'             strDescription - Description of the appointment
'             strNote - Notes pertaining to the appointment or series of appointments
'             strUserName - UserName of person responsible for adding appointment(s) into the system
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'   R003: Created
'--------------------------------------------------------------------
''
''    Dim cmd As ADODB.Command
''    Dim cnn As ADODB.Connection
''
''    On Error GoTo ErrTrap:
''
''    Set cnn = New ADODB.Connection
''    Set cmd = New ADODB.Command
''
''    'Create the parameter objects
''    With cmd
''        .CommandText = "uspUpdSingleRecurAppt"
''        .CommandType = adCmdStoredProc
''        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
''        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
''        .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , IIf(lngClinicID > 0, lngClinicID, Null))
''        .Parameters.Append .CreateParameter("@CategoryID", adInteger, adParamInput, , lngCategoryID)
''        .Parameters.Append .CreateParameter("@ApptDate", adDBTimeStamp, adParamInput, , dtApptDate)
''        .Parameters.Append .CreateParameter("@Duration", adInteger, adParamInput, , lngDuration)
''        .Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, 30, strDescription)
''        .Parameters.Append .CreateParameter("@Note", adVarChar, adParamInput, 256, strNote)
''        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
''        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
''    End With
''
''    'Assign the connection to the Command object and execute the stored procedure
''    Call cnn.Open(CONST_PSYQUEL_CNN)
''    Set cmd.ActiveConnection = cnn
''    cmd.Execute ', , adExecuteNoRecords
''    Set cnn = Nothing
''
''    'Check the ErrorNum parameter before deciding to commit the transaction
''    If cmd.Parameters("@ErrorNum").Value <> 0 Then
''        GoTo ErrTrap
''    End If
''
''    Set cnn = Nothing
''    Set cmd = Nothing
''
''    GetObjectContext.SetComplete
''
''    Exit Sub
''ErrTrap:
''    'Signal incompletion and raise the error to the calling environment.
''    GetObjectContext.SetAbort
''    Set cnn = Nothing
''    Set cmd = Nothing
''    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
''
''End Sub


'Public Function FetchByProviderDateRangeHTML(ByVal lngProviderID As Long, _
                                         ByVal dteStartDate As Date, _
                                         ByVal dteEndDate As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 09/25/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:   Retrieves a recordset of apointments from           '
'               tblAppointment that matches the criteria given in   '
'               the parameter arguments                             '
'Parameters:    lngProviderID - ID of Provider whose schedule is    '
'                   being sought                                    '
'               dteStartDate, dteEndDate - the date range that limits
'                   the apointments being returned                  '
'Returns:   Recordset of appointments                               '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
'    Dim cnn As ADODB.Connection
'    Dim cmd As ADODB.Command
'    Dim rst As ADODB.Recordset
'
'    Dim SQLErrorNum As Long
'
'    On Error GoTo ErrTrap:
'
'    Set cmd = New ADODB.Command
'    Set cnn = New ADODB.Connection
'
'    'Acquire the database connection.
'    cnn.Open (CONST_PSYQUEL_CNN)
'    Set cmd.ActiveConnection = cnn
'
'
'    'Create the parameter objects
'    With cmd
'        .CommandText = "uspSelApptByProviderDate"
'        .CommandType = adCmdStoredProc
'        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
'        .Parameters.Append .CreateParameter("@StartDate", adDBTimeStamp, adParamInput, , dteStartDate)
'        .Parameters.Append .CreateParameter("@EndDate", adDBTimeStamp, adParamInput, , dteEndDate)
'        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
'    End With
'
'    'Execute the stored procedure
'    Set rst = New ADODB.Recordset
'    rst.CursorLocation = adUseClient
'    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
'
'    Set FetchByProviderDateRangeHTML = rst
'
'    'Disconnect the recordset and clean house
'    Set cmd.ActiveConnection = Nothing
'    Set cmd = Nothing
'    Set rst.ActiveConnection = Nothing
'    Set cnn = Nothing
'
'    'Signal successful completion
'    GetObjectContext.SetComplete
'
'    Exit Function
'
'ErrTrap:
'    'Signal incompletion and raise the error to the calling environment.
'    GetObjectContext.SetAbort
'    Set cnn = Nothing
'    Set cmd = Nothing
'    Set rst = Nothing
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function

Public Function FetchConflicts(ByVal lngProviderID As Long, _
                               ByVal dteStartDate As Date, _
                               ByVal dteEndDate As Date, _
                               ByVal lngApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 09/27/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:   Searches for existing appointments that conflict with
'                   the given appointment time range                '
'Parameters:    lngProviderID - ID of Provider whose schedule is    '
'                   being sought                                    '
'               dteStartDate, dteEndDate - the date range of the    '
'                   proposed new appointment                        '
'Returns:   Recordset of conflicting appointments                   '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelApptByConflict"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@StartDate", adDBTimeStamp, adParamInput, , dteStartDate)
        .Parameters.Append .CreateParameter("@EndDate", adDBTimeStamp, adParamInput, , dteEndDate)
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchConflicts = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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

Public Function FetchConflictsByProvider(ByVal lngProviderID As Long, _
    ByVal dteApptStartDateTime As Date, ByVal intDuration As Integer) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 11/30/2001
'Author: Dave Richkun
'Description: Searches for existing appointments that may conflict with
'             the passed appointment start date.  This method returns
'             conflicting appointments for regular non-recurring appointments
'             and recurring appointments.  **Note**: Those appointment
'             records returned that represent recurring appointments are
'             only 'potential' appointment confilicts.  It will still need
'             to be the responsibility of the business object to determine
'             if an actual conflict exists.
'Parameters:  lngProviderID - ID of Provider whose schedule is being searched
'             dteApptStartDateTime - Starting time of the appointment for which
'               conflict will be checked
'             intDuration - Duration of the appointment (in minutes)
'Returns:   Recordset of potentially conflicting appointments                   '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelApptConflictsByProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@ApptStartDate", adDBTimeStamp, adParamInput, , dteApptStartDateTime)
        .Parameters.Append .CreateParameter("@Duration", adInteger, adParamInput, , intDuration)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchConflictsByProvider = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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



Public Sub DeleteNonRecurring(ByVal lngApptID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 10/24/2001
'Author: Dave Richkun
'Description:  Deletes (disables) a non-recurring appointment in tblAppointment
'Parameters:  lngApptID - ID of the appointment to be deleted
'             strUserName - Name of user deleting the appointment
'Returns:  Nothing
'--------------------------------------------------------------------
'Revision History:
'  R003: Method name changed to DeleteSingle() from Delete()
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection

    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspDelAppointment"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientCount", adInteger, adParamReturnValue, , 0)
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
        
        'Assign the connection to the Command object and execute the stored procedure
        Call cnn.Open(CONST_PSYQUEL_CNN)
        Set .ActiveConnection = cnn
        .Execute ', , adExecuteNoRecords
        Set cnn = Nothing
        
        'Check the ErrorNum parameter before deciding to commit the transaction
        If .Parameters("@ErrorNum").Value <> 0 Then
            GoTo ErrTrap
        End If
        Set cmd = Nothing
    End With
    
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)

End Sub

Public Sub DeleteRecurSeries(ByVal lngApptID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 01/30/2001
'Author: Rick "Boom Boom" Segura
'Description:  Deletes an appointment in tblRecurAppt
'Parameters:  lngApptID - ID of the appointment to be deleted
'             strUserName - Name of user deleting the appointment
'Returns:  Null
'--------------------------------------------------------------------
'Revision History:
'   R003: Method name changed to DeleteRecurSeries() from DeleteRecur()
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection

    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspDelRecurAppt"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
        
        'Assign the connection to the Command object and execute the stored procedure
        Call cnn.Open(CONST_PSYQUEL_CNN)
        Set .ActiveConnection = cnn
        .Execute ', , adExecuteNoRecords
        Set cnn = Nothing
        
        'Check the ErrorNum parameter before deciding to commit the transaction
        If .Parameters("@ErrorNum").Value <> 0 Then
            GoTo ErrTrap
        End If
        Set cmd = Nothing
    End With
    
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)

End Sub

Public Function FetchProviderExceptions(ByVal lngProviderID As Long, _
    ByVal dtStartDateTime As Date, ByVal dtEndDateTime As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/05/2001
'Author: Dave Richkun
'Description: Returns recordset of exceptions made to a provider's recurring
'             appointments within a given date range.
'Parameters:  lngProviderID - ID of Provider
'             dteStartDate, dteEndDate - the date range in which to retrieve
'                   appointment exceptions
'Returns: Recordset of appointment exceptions
'--------------------------------------------------------------------
'Revision History:
'  R003: Created
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelRecurApptExc"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@StartDate", adDBTimeStamp, adParamInput, , dtStartDateTime)
        .Parameters.Append .CreateParameter("@EndDate", adDBTimeStamp, adParamInput, , dtEndDateTime)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchProviderExceptions = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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
   
   
Public Function FetchCheckInExceptions(ByVal lngUserID As Long, ByVal dtArrivalDate As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 11/06/2001
'Author: Dave Richkun
'Description: Returns recordset of exceptions made to recurring appointments on a
'             specific check-in date.  Results vary based on the role assigned to the
'             user.
'Parameters:  lngUserID - ID of User retrieving the result set
'             dtArrivalDate - the check-in date for which to retrieve appointment exceptions
'Returns: Recordset of appointment exceptions
'--------------------------------------------------------------------
'Revision History:
'  R003: Created
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelCheckInExceptions"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
        .Parameters.Append .CreateParameter("@StartDate", adDBTimeStamp, adParamInput, , dtArrivalDate)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchCheckInExceptions = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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
   
   

Public Function FetchECDetail(ByVal lngApptID As Long, _
                              ByVal lngPatientID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 11/07/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:   Retrieves a recordset of detailed information from  '
'               tblEncounterLog having the parameters values passed '
'Parameters:    lngApptID - ID of the appointment whose information '
'                   is being sought                                 '
'               lngPatientID - ID of the Patient whose information  '
'                   is being sought                                 '
'Returns:   Recordset of appointment information                    '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelApptECDetail"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchECDetail = rst
    
    'Disconnect the recordset and clean house
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
Public Function FetchEncByDOS(ByVal lngProviderID As Long, ByVal lngApptID As Long, ByVal dteDOS As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 03/29/2016                                                   '
'Author: DCO                                                        '
'Description:   Retrieves a recordset of detailed information from  '
'               tblEncounterLog having the parameters values passed '
'Parameters:    lngPatientID - ID of the Patient whose information  '
'                   is being sought                                 '
'               dteDOS - Date of Service                            '
'Returns:   Recordset of Encounter information                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelProviderEncountersByAppt"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@DOS", adDBTimeStamp, adParamInput, , dteDOS)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchEncByDOS = rst
    
    'Disconnect the recordset and clean house
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
Public Function InsertRecurAppt(ByVal lngProviderID As Long, ByVal lngClinicID As Long, _
                ByVal lngCategoryID As Long, ByVal dteStartDateTime As Date, _
                ByVal dteEndDateTime As Date, ByVal lngDuration As Long, _
                ByVal strCPTCode As String, ByVal strDescription As String, _
                ByVal strNote As String, ByVal strRecurPattern As String, _
                ByVal lngInterval As Long, ByVal lngDOWMask As String, _
                ByVal lngDOM As Long, ByVal lngWOM As Long, ByVal lngMOY As Long, _
                ByVal strUserName As String) As Long
'--------------------------------------------------------------------
'Date: 10/30/2001
'Author: Dave Richkun
'Description:  Inserts an appointment record into tblRecurAppt via a stored procedure                           '
'Parameters:  lngProviderID - ID of Provider whose calendar the recurring appointment is for
'             lngClinicID - If a recurring patient appointment, ID of Clinic where recurring appointment will take place
'             lngCategoryID - Category identifier of recurring appointment i.e. Patient or Block
'             dteStartDateTime - Start Date/Time of recurring appointment
'             dteEndDateTime - End Date/Time of recurring appointment
'             lngDuration - Duration (length in minutes) of recurring appointment
'             strCPTCode - If recurring appointment is a patient appointment, CPT code assigned to appointments'
'             strDescription - Brief description of recurring appointment
'             strNote - Extended description or additional information about recurring appointment
'             strRecurPattern - Single character identifying recurrance pattern of a recurring appointment
'                   D = Daily
'                   W = Weekly
'                   M = Monthly
'                   Y = Yearly
'             lngRecurInterval - The interval between recurring appointments i.e. this value would be 2 if appointment recurs every 2 days
'             lngRecurDOWMask - A bit-mapped number ANDed together representing individual or combined days of the week.
'                   1  = Sunday
'                   2  = Monday
'                   4  = Tuesday
'                   8  = Wednesday
'                   16 = Thursday
'                   32 = Friday
'                   64 = Saturday
'             lngRecurDOM - Day of the month when recurrance pattern is 'M'
'             lngRecurWOM - Week of the month when recurrance pattern is 'W'
'             lngRecurMOY - Month of the year when recurrance pattern is 'Y'
'             strUserName - Name of user responsible for creating the recurring appointment
'Returns:  ID of appointment on success, -1 on failure
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection

    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspInsRecurAppt"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , IIf(lngClinicID = 0, Null, lngClinicID))
        .Parameters.Append .CreateParameter("@CategoryID", adInteger, adParamInput, , lngCategoryID)
        .Parameters.Append .CreateParameter("@StartDateTime", adDBTimeStamp, adParamInput, , dteStartDateTime)
        .Parameters.Append .CreateParameter("@EndDateTime", adDBTimeStamp, adParamInput, , dteEndDateTime)
        .Parameters.Append .CreateParameter("@Duration", adInteger, adParamInput, , lngDuration)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, IIf(strCPTCode = "", Null, strCPTCode))
        .Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, 30, IIf(strDescription = "", Null, strDescription))
        .Parameters.Append .CreateParameter("@Note", adVarChar, adParamInput, 256, IIf(strNote = "", Null, strNote))
        .Parameters.Append .CreateParameter("@RecurPattern", adVarChar, adParamInput, 1, strRecurPattern)
        .Parameters.Append .CreateParameter("@Interval", adInteger, adParamInput, , lngInterval)
        .Parameters.Append .CreateParameter("@DOWMask", adInteger, adParamInput, , lngDOWMask)
        .Parameters.Append .CreateParameter("@DOM", adInteger, adParamInput, , lngDOM)
        .Parameters.Append .CreateParameter("@WOM", adInteger, adParamInput, , lngWOM)
        .Parameters.Append .CreateParameter("@MOY", adInteger, adParamInput, , lngMOY)
        .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 30, strUserName)
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
    
    InsertRecurAppt = cmd.Parameters("@ID").Value
    
    'Free resources
    Set cmd = Nothing
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
    
End Function

Public Sub InsertRecurApptExc(ByVal lngApptID As Long, ByVal dtApptDate As Date, _
                              ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 10/09/2001
'Author: Dave Richkun
'Description:  Inserts a record into tblRecurApptExc using a stored procedure                           '
'Parameters:  lngApptID - ID of the recurring appointment where an exception is being made
'             dtApptDate - The date of the appointment exception
'             strUserName - Name of user recording the exception
'Returns: Null
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
        .CommandText = "uspInsRecurApptExc"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ApptID", adInteger, adParamInput, , lngApptID)
        .Parameters.Append .CreateParameter("@ApptDate", adDBTimeStamp, adParamInput, , DateValue(dtApptDate))
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
    
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
    
End Sub

Public Sub UpdateRecurAppt(ByVal lngID As Long, ByVal dteStartDateTime As Date, _
                ByVal dteEndDateTime As Date, ByVal strCPTCode As String, ByVal lngDuration As Long, _
                ByVal lngProviderID As Long, ByVal lngClinicID As Long, _
                ByVal lngCategoryID As Long, ByVal strRecurPattern As String, _
                ByVal lngInterval As Long, ByVal lngDOWMask As Long, _
                ByVal lngDOM As Long, ByVal lngWOM As Long, ByVal lngMOY As Long, _
                ByVal strDescription As String, ByVal strNote As String, _
                ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 11/17/2000
'Author: Rick "Boom Boom" Segura
'Description:  Updates an appointment record in tblRecurAppt via a stored procedure                           '
'Parameters:  lngID - ID of record to be updated
'             dteStartDateTime - Start Date/Time of recurring appointment
'             dteEndDateTime - End Date/Time of recurring appointment
'             lngDuration - Duration(length) of appointment in minutes
'             lngProviderID - Provider ID
'             lngClinicID - Clinic ID
'             lngCategoryID _ Category ID
'             strRecurPattern - Recurring Appt Pattern
'             lngInterval - Pattern Interval value
'             strDOWMask - Day of Week Mask
'             lngDOM - Day of Month
'             lngWOM - Week Of Month
'             lngMOY - Month of Year
'             strDescription - Appt description
'             strNote - Appt note (extended description)
'             strUserName - Name of user adding appt
'Returns:  Appointment ID on success, -1 otherwise
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    Dim varSQLErr As Variant
    
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspUpdRecurAppt"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, , lngID)
        .Parameters.Append .CreateParameter("@StartDateTime", adDBTimeStamp, adParamInput, , dteStartDateTime)
        .Parameters.Append .CreateParameter("@EndDateTime", adDBTimeStamp, adParamInput, , dteEndDateTime)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, strCPTCode)
        .Parameters.Append .CreateParameter("@Duration", adInteger, adParamInput, , lngDuration)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , IIf(lngClinicID = 0, Null, lngClinicID))
        .Parameters.Append .CreateParameter("@CategoryID", adInteger, adParamInput, , lngCategoryID)
        .Parameters.Append .CreateParameter("@RecurPattern", adVarChar, adParamInput, 1, strRecurPattern)
        .Parameters.Append .CreateParameter("@Interval", adInteger, adParamInput, , lngInterval)
        .Parameters.Append .CreateParameter("@DOWMask", adInteger, adParamInput, , lngDOWMask)
        .Parameters.Append .CreateParameter("@DOM", adInteger, adParamInput, , lngDOM)
        .Parameters.Append .CreateParameter("@WOM", adInteger, adParamInput, , lngWOM)
        .Parameters.Append .CreateParameter("@MOY", adInteger, adParamInput, , lngMOY)
        .Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, 30, strDescription)
        .Parameters.Append .CreateParameter("@Note", adVarChar, adParamInput, 256, strNote)
        .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
    End With
        
    'Assign the connection to the Command object and execute the stored procedure
    Call cnn.Open(CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    cmd.Execute ', , adExecuteNoRecords
    Set cnn = Nothing

    'Check the ErrorNum parameter before deciding to commit the transaction
    varSQLErr = cmd.Parameters("@ErrorNum").Value
    If varSQLErr <> 0 Then
        GoTo ErrTrap
    End If
    
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, varSQLErr, CLASS_NAME)
    
End Sub

Public Function ExistsInTblAppt(ByVal lngApptID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 09/01/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:   Retrieves a recordset of detailed information for   '
'               the appontment having the ID given in the parameter '
'Parameters:    lngApptID - ID of the appointment whose information '
'                   is being sought                                 '
'Returns:   Recordset of appointment information                    '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelApptExists"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, , lngApptID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    ExistsInTblAppt = (rst.RecordCount > 0)
    
    'Disconnect the recordset and clean house
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

Public Function FetchFuturePatientApptByProvider(ByVal lngPatientID As Long, _
                 ByVal lngProviderID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 03/06/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:   Searches for non-attended future(including today)   '
'                   appointments for a provider/patient combination '
'Parameters:    lngPatientID - ID of patient whose appointments are '
'                   being sought                                    '
'               lngProviderID - ID of provider whose appointments are
'                   being sought                                    '
'Returns:   Recordset of appointments                               '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelPatientFutureApptByProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchFuturePatientApptByProvider = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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


Public Function FetchFutureRecurPatientApptByProvider(ByVal lngPatientID As Long, _
                 ByVal lngProviderID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 03/07/2002
'Author: Dave Richkun
'Description: Retrieves future recurring appointments for the passed patient
'             and provider
'Parameters:  lngPatientID - ID of patient
'             lngProviderID - ID of provider
'Returns:   Recordset of recur appointments
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelRecurPatientFutureApptByProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchFutureRecurPatientApptByProvider = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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


Public Function FetchFuturePatientApptByManager(ByVal lngPatientID As Long, _
                 ByVal lngUserID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 03/06/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:   Searches for non-attended future(including today)   '
'                   appointments for a patient with all provider for'
'                   the given office manager                        '
'Parameters:    lngPatientID - ID of patient whose appointments are '
'                   being sought                                    '
'               lngUserID - IID of manager asociated with the       '
'                   providers whose appointments are being sought   '
'Returns:   Recordset of  appointments                              '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelPatientFutureApptByManager"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ManagerID", adInteger, adParamInput, , lngUserID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchFuturePatientApptByManager = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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


Public Function FetchFutureRecurPatientApptByManager(ByVal lngPatientID As Long, _
                 ByVal lngUserID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 03/07/2002
'Author: Dave Richkun
'Description: Retrieves future recurring appointments for the passed patient
'             and office manager combination
'Parameters:  lngPatientID - ID of patient
'             lngUserID - ID of provider
'Returns:   Recordset of recur appointments
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelRecurPatientFutureApptByManager"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchFutureRecurPatientApptByManager = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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


Public Function FetchUnBilledAppts(ByVal lngUserID As Long) As ADODB.Recordset  'R001
'--------------------------------------------------------------------
'Date: 06/14/2001
'Author: Dave Richkun
'Description:   Retrieves unbilled appointments older than 3 days
'Parameters: lngUserID - ID of Provider or Office Manager retrieving records
'Returns:   Recordset of  appointments                              '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
  
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelUnBilledPastApptsByUser"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchUnBilledAppts = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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

Public Function FetchPendingAppts(ByVal lngUserID As Long) As ADODB.Recordset  'R001
'--------------------------------------------------------------------
'Date: 07/09/2007
'Author: Duane Orth
'Description:   Retrieves pending appointments
'Parameters: lngUserID - ID of Provider or Office Manager retrieving records
'Returns:   Recordset of  appointments                              '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
  
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelPendingApptsByUser"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
    End With
    
    'Execute the stored procedure
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchPendingAppts = rst
    
    'Disconnect the recordset and clean house
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst = Nothing
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


'--------------------------------------------------------------------
' PrivateMethods    +++++++++++++++++++++++++++++++++++++++++++++++++
'--------------------------------------------------------------------


