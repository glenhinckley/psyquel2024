VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApptBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CApptBZ                                                '
'Date: 08/25/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  MTS business object designed to call methods         '
'              associated with the Appointment classes.             '
'--------------------------------------------------------------------
' Revision History:
'  R001: 06/14/2001 Richkun: Added FetchUnBilledAppts()
'  R002: 08/08/2001 Richkun: Added support for tblAppointment columns
'        fldReferPhy, fldReferPhyID
'  R003: 10/04/2001 Richkun: Allowed for deleting, updating of single
'         appointment within series of recurring appointments; Required
'         significant interface changes.
'  R004: 10/19/2001 Richkun: Re-worked Insert, Update() methods
'  R005: 11/01/2001 Richkun: Added FetchByApptIDs() method
'  R006: 11/28/2001 Richkun: Added CloneInstance() method
'  R007: 12/24/2001 Richkun: Altered Update() method to support appointment
'          update and entry of patient payment details in one call
'  R008: 03/08/2002 Richkun: Added FetchFutureRecurPatientApptByProvider(),
'            FetchFutureRecurPatientApptByOffMgr() methods
'  R009: 05/13/2002 Richkun: Included support for fractional Unit values
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME = "CApptBZ"

Private Const START_TIME As String = "07:00:00 AM"

Private Const HEADING_T As Long = -1
Private Const OPEN_T As Long = 0
Private Const SCHEDULED_T As Long = 1

Private Const PATIENT_CAT As Long = 1
Private Const BLOCK_CAT As Long = 2

Private Const ATTENDED_CL As String = "Attended"
Private Const BLOCKED_CL As String = "Blocked"
Private Const CONFIRMED_CL As String = "Confirmed"
Private Const GROUP_CL As String = "Group"
Private Const HOLD_CL As String = "Held"
Private Const SCHEDULED_CL As String = "Scheduled"
Private Const NO_SHOW_CL As String = "NoShow"
Private Const CONFLICT_CL As String = "Conflict"
Private Const OPEN_CL As String = "Open"
Private Const TENATIVE_CL As String = "Tenative"
Private Const ROWHEADER_CL As String = "RowHeader"
Private Const PENDING_CL As String = "Pending"

Private Const ATTENDED_ST As Long = 3
Private Const CONFIRMED_ST As Long = 2
Private Const HOLD_ST As Long = 5
Private Const SCHEDULED_ST As Long = 1
Private Const NO_SHOW_ST As Long = 6
Private Const TENATIVE_ST As Long = 10
Private Const PENDING_ST As String = 11

Private Enum MsgType
    ApptCreateNoCert = 1
    ApptConfirmNoCert = 2
    ApptCancel = 3
    ApptDelete = 4
End Enum

Private Const APPT_TYPE_PATIENT As Long = 1
Private Const APPT_TYPE_BLOCK As Long = 2

Private Const APPT_STATUS_SCHEDULED As Integer = 1
Private Const APPT_STATUS_CONFIRMED As Integer = 2
Private Const APPT_STATUS_ATTENDED As Integer = 3
Private Const APPT_STATUS_CANCELLED As Integer = 4
Private Const APPT_STATUS_PENDING As Integer = 11

Public Enum typApptType
    apptTypeSingle = 1
    apptTypeRecurring = 2
End Enum

'--------------------------------------------------------------------
' Public Methods
'--------------------------------------------------------------------

Public Function FetchByProviderDateRange(ByVal lngProviderID As Long, _
     ByVal dteStartDate As Date, ByVal dteEndDate As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 08/30/2000
'Author: Rick "Boom Boom" Segura
'Description:  Fetches all appointments for the provider within a given date range                                    '
'Parameters:  lngProviderID - ID of the provider whose appointments are being retreived                                 '
'             dteStartDate - Start date of the search date range
'             dteEnsdDate - Start End of the search date range
'Returns:   Recordset of appointments
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set FetchByProviderDateRange = objAppt.FetchByProviderDateRange(lngProviderID, dteStartDate, dteEndDate)
    'Do not know they are Why adding an extra day to the search criteria disabled 05/08/2008
    'Set FetchByProviderDateRange = objAppt.FetchByProviderDateRange(lngProviderID, dteStartDate, DateAdd("d", 1, dteEndDate))

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
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
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set FindOpenApptTimeSlots = objAppt.FindOpenApptTimeSlots(lngClinicID, lngProviderID, dteStartDate, dteEndDate)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function FetchByCheckInDate(ByVal lngUserID As Long, ByVal dteArrivalDate As Date) As Variant
'--------------------------------------------------------------------
'Date: 11/06/2001
'Author: Dave Richkun
'Description:   Returns a 2-dimensional array of all patient appointments scheduled
'               for a specific user on the specified date.  Returns patient appointents
'               based on the User ID's role.  This function includes dates that fall on
'               the specified date that are part of recurring appointments.  This
'               function was designed to support the Check-In feature.
'Parameters:    lngUserID - ID of User retrieving the information
'               dteArrivalDate - The date on which the appointments occur i.e. the check-in date
'Returns:   Recordset of appointments                               '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CApptDB
    Dim rstAppt As ADODB.Recordset
    Dim rstExc As ADODB.Recordset
    Dim aryAppts() As Variant
    Dim aryRecurDates As Variant
    Dim intCtr As Integer
    Dim intCtr2 As Integer
    Dim intCtr3 As Integer
    Dim intApptCount As Integer
    Dim blnException As Boolean

    On Error GoTo ErrTrap:
    
    Set objAppt = CreateObject("ApptDB.CApptDB")
    Set rstAppt = objAppt.FetchByCheckInDate(lngUserID, dteArrivalDate)
    Set rstExc = objAppt.FetchCheckInExceptions(lngUserID, dteArrivalDate)
    
    If rstAppt.RecordCount > 0 Then
        ReDim aryAppts(4, rstAppt.RecordCount - 1)
    
        For intCtr = 0 To rstAppt.RecordCount - 1
            blnException = False
            'Get valid recurring dates if any
            If Trim(rstAppt.Fields("fldRecurPattern").Value) > "" Then
                aryRecurDates = GetRecurApptDates(dteArrivalDate, dteArrivalDate, _
                    rstAppt.Fields("fldStartDateTime").Value, rstAppt.Fields("fldEndDateTime").Value, _
                    rstAppt.Fields("fldDuration").Value, rstAppt.Fields("fldRecurPattern").Value, _
                    rstAppt.Fields("fldInterval").Value, rstAppt.Fields("fldDOWMask").Value, _
                    rstAppt.Fields("fldDOM").Value, rstAppt.Fields("fldWOM").Value, _
                    rstAppt.Fields("fldMOY").Value)
            
                'Look for recuring appointment date exceptions before including the date in the 'aryAppts' array
                If IsArray(aryRecurDates) Then
                    For intCtr2 = 0 To UBound(aryRecurDates)
                        blnException = False
                        If rstExc.RecordCount > 0 Then
                            rstExc.MoveFirst
                            For intCtr3 = 1 To rstExc.RecordCount
                                If CDate(rstExc.Fields("fldApptDate")) = CDate(aryRecurDates(intCtr2)) Then
                                    If rstExc.Fields("fldRecurApptID").Value = rstAppt.Fields("fldApptID") Then
                                        blnException = True
                                        Exit For
                                    End If
                                End If
                                If Not rstExc.EOF Then
                                    rstExc.MoveNext
                                End If
                            Next intCtr3
                        End If
                        
                    Next intCtr2
                Else
                    blnException = True
                End If
            End If
            
            If blnException = False Then
                aryAppts(0, intApptCount) = rstAppt.Fields("fldPatApptID").Value
                aryAppts(1, intApptCount) = rstAppt.Fields("fldPatientName").Value
                aryAppts(2, intApptCount) = rstAppt.Fields("fldProviderName").Value
                aryAppts(3, intApptCount) = rstAppt.Fields("fldStartDateTime").Value
                aryAppts(4, intApptCount) = rstAppt.Fields("fldDescription").Value
                
                intApptCount = intApptCount + 1
            End If
            
            rstAppt.MoveNext
        Next intCtr
    Else
        ReDim aryAppts(0, 0)
    End If
        
    'Shrink the array, if needed
    If IsArray(aryAppts) Then
        If IsEmpty(aryAppts(0, 0)) Then
            FetchByCheckInDate = Null
        Else
            ReDim Preserve aryAppts(4, intApptCount - 1)
            FetchByCheckInDate = aryAppts
        End If
    Else
        FetchByCheckInDate = Null
    End If
        
    ' Clean House
    Set rstAppt = Nothing
    Set rstExc = Nothing
    Set objAppt = Nothing
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Clean up
    Set rstAppt = Nothing
    Set objAppt = Nothing
    GetObjectContext.SetAbort
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
    
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    'Instantiate the Appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")
    Set FetchPatientApptByID = objAppt.FetchPatientApptByID(lngApptID)
    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
    
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
    
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    'Instantiate the Appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")
    Set FetchBlockApptByID = objAppt.FetchBlockApptByID(lngApptID)
    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
    
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
'  R005: Created
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    'Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")
    Set FetchByApptIDs = objAppt.FetchByApptIDs(strApptIDs)
    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function Insert(ByVal intRecurType As typApptType, ByVal lngProviderID As Long, _
       ByVal lngClinicID As Long, ByVal lngCategoryID As Long, ByVal dteStartDateTime As Date, _
       ByVal dteEndDateTime As Date, ByVal lngDuration As Long, ByVal strCPTCode As String, _
       ByVal strDescription As String, ByVal strNote As String, ByVal strUserName As String, _
       Optional ByVal varPatientArray As Variant, Optional ByVal strRecurPattern As String = "", _
       Optional ByVal lngRecurInterval As Long, Optional ByVal lngRecurDOWMask As Long, _
       Optional ByVal lngRecurDOM As Long, Optional ByVal lngRecurWOM As Long, _
       Optional ByVal lngRecurMOY As Long) As Long
'--------------------------------------------------------------------
'Date: 08/25/2000
'Author: Dave Richkun
'Description:  Inserts appointment information into appropriate tables  based on the type of appointment
'Parameters:  intRecurType - Enumerated parameter identifying appointment recurrence type
'             lngProviderID - ID of Provider for whom the appointment is scheduled
'             lngClinicID - ID of the Place of Service where appointment is scheduled (patient appointments only)
'             lngCategoryID - ID of Category identifying appointment type (Patient or Block)
'             dteStartDateTime - Start Date and Time of appointment
'             dteEndDateTime -  End Date and Time of appointment
'               For Single Appts: This value should be equal to dtStartDateTime plus Duration
'               For Recurring Appts:  This value should be the ending date of the recurrence and the ending time
'                   each appointment in this series.
'             lngDuration - Appointment length in minutes
'             strDescription - A short description of the appointment
'             strNote - Additional information about the appointment
'             strUserName - Login name of individual creating the appointment record
'             varPatientArray - This is a 2-D array containing one 'column' for each patient appointment.  The 'row'
'                   elements are defined as follows:
'                   0 = PatientAppointment ID
'                   1 = Patient ID
'                   2 = Appointment Status ID
'             strCPTCode - The CPT Code associated with a patient appointment
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
'Returns:  ID assigned to Appointment on success, otherwise -1
'--------------------------------------------------------------------
'Revision History:                                                  '
'
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptDB
    Dim objPatAppt As ApptBZ.CPatApptBZ
    Dim strErrMsg As String
    Dim lngApptID As Long
    Dim lngRecurApptID As Long
    Dim intCtr As Integer
    Dim aryDates As Variant
    Dim aryPatients As Variant
    
    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CApptDB")

    Select Case intRecurType
        Case 1 ' Single instance i.e. one-time, non-recurring appointment
            'Insert Core Appt Info
            lngApptID = objAppt.InsertSingle(lngProviderID, lngClinicID, _
                lngCategoryID, dteStartDateTime, dteEndDateTime, lngDuration, _
                strCPTCode, strDescription, strNote, strUserName)
    
            If lngApptID < 1 Then
                strErrMsg = "An error occured when inserting the core appointment information."
                GoTo ErrTrap:
            End If
    
            Insert = lngApptID
        Case 2 'Recurring appointment
            lngApptID = objAppt.InsertRecurAppt(lngProviderID, lngClinicID, lngCategoryID, _
                    dteStartDateTime, dteEndDateTime, lngDuration, strCPTCode, strDescription, _
                           strNote, strRecurPattern, lngRecurInterval, lngRecurDOWMask, _
                           lngRecurDOM, lngRecurWOM, lngRecurMOY, strUserName)

            Insert = lngApptID

            If lngApptID < 1 Then
                strErrMsg = "An error occured when inserting the recurring appointment information."
                GoTo ErrTrap:
            End If
    End Select

    'If this is a patient appointment, insert rows into tblPatientAppt
    If lngCategoryID = APPT_TYPE_PATIENT Then
        'Insert rows into tblPatientAppt
        Set objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
        For intCtr = 0 To UBound(varPatientArray, 1)
            If varPatientArray(intCtr, 1) < 1 Then
                strErrMsg = "An error occured when inserting the patient appointment information(Missing Patient ID)."
                GoTo ErrTrap:
            Else
                Call objPatAppt.Insert(lngApptID, varPatientArray(intCtr, 1))
            End If
        Next
        Set objPatAppt = Nothing
    End If

    'Clean up
    Set objAppt = Nothing
    Set objPatAppt = Nothing

    GetObjectContext.SetComplete

    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort

    Set objAppt = Nothing
    Set objPatAppt = Nothing

    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Public Function Update(ByVal lngApptID As Long, ByVal intRecurType As typApptType, ByVal lngProviderID As Long, _
       ByVal lngClinicID As Long, ByVal lngCategoryID As Long, ByVal dteStartDateTime As Date, _
       ByVal dteEndDateTime As Date, ByVal lngDuration As Long, ByVal strCPTCode As String, _
       ByVal strDescription As String, ByVal strNote As String, ByVal strUserName As String, _
       Optional ByVal varPatientArray As Variant, Optional ByVal strRecurPattern As String, _
       Optional ByVal lngRecurInterval As Long, Optional ByVal lngRecurDOWMask As Long, _
       Optional ByVal lngRecurDOM As Long, Optional ByVal lngRecurWOM As Long, _
       Optional ByVal lngRecurMOY As Long, Optional blnRecurInstance As Boolean = False) As Long
'----------------------------------------------------------------------------------------------------------
'Author: Dave Richkun
'Description:  Updates appointment information into appropriate tables based on the type of appointment
'Parameters:  lngApptID - ID of the appointment to update
'             intRecurType - Enumerated parameter identifying appointment recurrence type
'             lngProviderID - ID of Provider for whom the appointment is scheduled
'             lngClinicID - ID of the Place of Service where appointment is scheduled (patient appointments only)
'             lngCategoryID - ID of Category identifying appointment type (Patient or Block)
'             dteStartDateTime - Start Date and Time of appointment
'             dteEndDateTime -  End Date and Time of appointment
'               For Single Appts: This value should be equal to dtStartDateTime plus Duration
'               For Recurring Appts:  This value should be the ending date of the recurrence and the ending time
'                   each appointment in this series.
'             lngDuration - Appointment length in minutes
'             strDescription - A short description of the appointment
'             strNote - Additional information about the appointment
'             strUserName - Login name of individual creating the appointment record
'             varPatientArray - For patient appoinments, this is a 2-D array with the following elements:
'                   0 = PatientAppointment ID
'                   1 = Patient ID
'                   2 = Appointment Status ID
'             strCPTCode - The CPT Code associated with a patient appointment
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
'             blnRecurInstance - Boolean value identifying if the appointment we are updating
'                     is a single instance from a recurring patient appointment series.  Any changes made
'                     to an instance of a recurring patient appointment forces the appointment instance
'                     to become a new one-time patient, so that billing procedures can be properly applied
'                     to the instance.
'Returns:  ID assigned to Appointment on success, otherwise -1
'----------------------------------------------------------------------------------------------------------
' Revision History:
'   R007
'----------------------------------------------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptDB
    Dim objPatAppt As ApptBZ.CPatApptBZ
    Dim strErrMsg As String
    Dim intCtr As Integer
    Dim lngNewApptID As Long
    
    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CApptDB")

    'Check if this is an update to an instance of a recurring appointment
    If blnRecurInstance = True Then
        'Create a new one-time patient appointment to replace this instance.
        lngNewApptID = objAppt.InsertSingle(lngProviderID, lngClinicID, lngCategoryID, _
            dteStartDateTime, dteEndDateTime, lngDuration, strCPTCode, strDescription, _
            strNote, strUserName)
            
        'If this is a patient appointment, insert a new patient appointment record for each
        'patient in the appointment, replacing the appropriate element in the Patient Array
        'with the new PatApptID.
        If lngCategoryID = APPT_TYPE_PATIENT Then
            'Insert rows into tblPatientAppt
            Set objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
            For intCtr = 0 To UBound(varPatientArray, 1)
                varPatientArray(intCtr, 0) = objPatAppt.Insert(lngNewApptID, varPatientArray(intCtr, 1))
            Next
            Set objPatAppt = Nothing
        End If
        
        'Insert a row into tblRecurApptExc so the recurring appointment and the new one-time
        'appointment do not appear as conflicts.
        Call objAppt.InsertRecurApptExc(lngApptID, dteStartDateTime, strUserName)
        
        'Update the NEW appointment.  This effectively becomes a recursive call to the Update() method
        'The only difference is that the PatientAppt element of the Patient array must now contain the
        'new Patient Appointment IDs.
        Call Update(lngNewApptID, intRecurType, lngProviderID, lngClinicID, lngCategoryID, _
            dteStartDateTime, dteEndDateTime, lngDuration, strCPTCode, strDescription, strNote, _
            strUserName, varPatientArray, "", 0, 0, 0, 0, 0, False)
    
        GoTo CLEANUP:  'We are done.
    End If

    Select Case intRecurType
        Case 1 ' Single instance i.e. one-time
            'Update Core Appt Info
            Call objAppt.Update(lngApptID, lngProviderID, lngClinicID, _
                lngCategoryID, dteStartDateTime, dteEndDateTime, _
                lngDuration, strCPTCode, strDescription, strNote, strUserName)
        
        Case 2 'Recurring
            Call objAppt.UpdateRecurAppt(lngApptID, dteStartDateTime, dteEndDateTime, strCPTCode, lngDuration, _
                    lngProviderID, lngClinicID, lngCategoryID, strRecurPattern, lngRecurInterval, _
                    lngRecurDOWMask, lngRecurDOM, lngRecurWOM, lngRecurMOY, strDescription, _
                    strNote, strUserName)

            If lngApptID < 1 Then
                strErrMsg = "An error occured when updating the recurring appointment information."
                GoTo ErrTrap:
            End If
    End Select

    If lngCategoryID = APPT_TYPE_PATIENT Then
        'Update rows in tblPatientAppt
        Set objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
        For intCtr = 0 To UBound(varPatientArray, 1)
            'Do not update patient appointments already Billed (in the case of Group Appointments)
            If varPatientArray(intCtr, 2) <> "Y" Then
                If varPatientArray(intCtr, 2) <= 2 Or _
                   varPatientArray(intCtr, 2) = APPT_STATUS_CANCELLED Or _
                   varPatientArray(intCtr, 2) = APPT_STATUS_PENDING Then
                    'If Appt Status is 'Scheduled', 'Confirmed' or 'Cancelled', only a minimal update is required
                    Call objPatAppt.ChangeStatus(varPatientArray(intCtr, 0), varPatientArray(intCtr, 2), _
                    "N", 0, "", strUserName)
                End If
            End If
        Next
        Set objPatAppt = Nothing
    End If


CLEANUP:
    'Return ApptID to calling routine
    If lngNewApptID = 0 Then
        Update = lngApptID
    Else
        Update = lngNewApptID
    End If
    
    'Clean up
    Set objAppt = Nothing
    Set objPatAppt = Nothing

    GetObjectContext.SetComplete

    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort

    Set objAppt = Nothing
    Set objPatAppt = Nothing

    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Public Function CloneInstance(ByVal lngApptID As Long, ByVal dteStartDateTime As Date, _
        ByVal strUserName As String) As Long
'--------------------------------------------------------------------
'Date: 11/28/2001
'Author: Dave Richkun
'Description: Creates an single instance of a recurring appointment.  Before
'             alterations can be made to a recurring appointment, an instance
'             of the appointment is created.  All data changes occur on the instance
'             while the recurring appointment is left in tact.
'Parameters:  lngApptID - ID of the Recurring appointment from which the instance
'                   will be created.
'             dtStartDateTime - Start Date/Time of the appointment instance
'             strUserName - Name of the user creating the appointment instance
'Returns: ID of cloned appointment
'--------------------------------------------------------------------
'Revision History:
'  R006: Created
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CApptDB
    Dim objPatAppt As ApptBZ.CPatApptBZ
    Dim rstAppt As ADODB.Recordset
    Dim dteEndDateTime As Date
    Dim lngNewApptID As Long
    Dim lngPatApptID As Long
    Dim strErrMsg As String
    Dim intCtr As Integer
    Dim lngTempPatientID As Long

    On Error GoTo ErrTrap:
    
    'Fetch appointment information
    Set rstAppt = FetchPatientApptByID(lngApptID)
    If rstAppt.RecordCount = 0 Then
        strErrMsg = "Appointment ID not found."
        GoTo ErrTrap:
    End If
    dteStartDateTime = DateValue(dteStartDateTime) & " " & TimeValue(rstAppt.Fields("fldStartDateTime").Value)
    dteEndDateTime = DateAdd("n", rstAppt.Fields("fldDuration").Value, dteStartDateTime)

    Set objAppt = CreateObject("ApptDB.CApptDB")
    
    'Insert single instance of appointment, copying values from original.
    lngNewApptID = objAppt.InsertSingle(rstAppt.Fields("fldProviderID").Value, rstAppt.Fields("fldClinicID").Value, _
        rstAppt.Fields("fldCategoryID").Value, dteStartDateTime, dteEndDateTime, rstAppt.Fields("fldDuration").Value, _
        IfNull(rstAppt.Fields("fldCPTCode").Value, ""), IfNull(rstAppt.Fields("fldDescription").Value, ""), _
        IfNull(rstAppt.Fields("fldNote").Value, ""), strUserName)
        
    'If this is a patient appointment, insert a new patient appointment record for each
    'patient in the appointment, replacing the appropriate element in the Patient Array
    'with the new PatApptID.
    If rstAppt.Fields("fldCategoryID").Value = APPT_TYPE_PATIENT Then
        'Insert rows into tblPatientAppt.  The recordset will return multiple rows per patient if
        'the patient has more than one plan.  Ensure only one patient appointment record is inserted
        'per cloned appointment.
        Set objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
        lngTempPatientID = -1
        For intCtr = 1 To rstAppt.RecordCount
            If lngTempPatientID <> rstAppt.Fields("fldPatientID").Value Then
                lngPatApptID = objPatAppt.Insert(lngNewApptID, rstAppt.Fields("fldPatientID").Value)
                lngTempPatientID = rstAppt.Fields("fldPatientID").Value
            End If
            rstAppt.MoveNext
        Next
        Set objPatAppt = Nothing
    End If
    
    Set rstAppt = Nothing
    
    'Insert a row into tblRecurApptExc so the recurring appointment and the new one-time
    'appointment do not appear as conflicts.
    Call objAppt.InsertRecurApptExc(lngApptID, dteStartDateTime, strUserName)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    
    'Return new appointment ID
    CloneInstance = lngNewApptID

    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Set objPatAppt = Nothing
    Set rstAppt = Nothing
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Public Function ConflictExists(ByVal lngProviderID As Long, _
                               ByVal dteStartDateTime As Date, _
                               ByVal dteEndDateTime As Date, _
                               ByVal lngApptID As Long, _
                               Optional ByRef varConflicts As Variant) As Boolean
'--------------------------------------------------------------------
'Date: 09/27/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Determines if a scheduling conflict exists between   '
'               existing appointments and the given appointment     '
'               time range                                          '
'Parameters:  lngProviderID - ID of the provider whose appointments '
'               are being retreived                                 '
'             dteStartDateTime - Start date/time of the new appointment
'             dteEnsdDateTime - End date/time of the new appointment'
'Returns:   TRUE if conflict is detected, FALSE otherwise           '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptDB
    Dim rst As ADODB.Recordset

    On Error GoTo ErrTrap:

    ConflictExists = False ' Assume no conflicts

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set rst = objAppt.FetchConflicts(lngProviderID, dteStartDateTime, dteEndDateTime, lngApptID)

    If rst.RecordCount > 0 Then
        ConflictExists = True
    Else
        With rst
            Do While Not .EOF
                If .Fields("fldRecurPattern").Value > "" Then
                    ConflictExists = True
                    Exit Do
                End If
                .MoveNext
            Loop
        End With
    End If

    If Not IsMissing(varConflicts) Then
        Set varConflicts = rst
    End If

    Set rst = Nothing

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    Set rst = Nothing
    GetObjectContext.SetAbort

End Function

Public Function ConflictExistsRec(ByVal lngProviderID As Long, _
                               ByVal dteStartDateTime As Date, _
                               ByVal dteEndDateTime As Date, _
                               ByVal lngApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 01/19/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:                                            '
'Parameters:  lngProviderID - ID of the provider whose appointments '
'               are being retreived                                 '
'             dteStartDateTime - Start date/time of the new appointment
'             dteEnsdDateTime - End date/time of the new appointment'
'Returns:           '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set ConflictExistsRec = objAppt.FetchConflicts(lngProviderID, dteStartDateTime, dteEndDateTime, lngApptID)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort

End Function

Public Sub DeleteNonRecurring(ByVal lngApptID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 10/04/2001
'Author: Dave Richkun
'Description:  Deletes (disables) a non-recuring appointment
'Parameters:  lngApptID - ID of appointment to delete                  '
'             strUserName - Name of user deleting the appointment
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'  R003: Changed method name to DeleteNonRecurring() from Delete()
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    'Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")
    Call objAppt.DeleteNonRecurring(lngApptID, strUserName)
    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Sub

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort

End Sub

Public Sub DeleteRecurSeries(ByVal lngApptID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 10/04/2001
'Author: Dave Richkun
'Description:  Deletes the entire series of a recurring appointment
'Parameters:  lngApptID - ID of appointment to delete                  '
'             strUserName - Name of user deleting the appointment
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'  R003: Changed method name to DeleteRecurSeries() from Delete()
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    Call objAppt.DeleteRecurSeries(lngApptID, strUserName)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Sub

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort

End Sub


Public Sub DeleteRecurSingle(ByVal lngApptID As Long, ByVal dtApptDate As Date, _
            ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 10/04/2001
'Author: Dave Richkun
'Description:  Deletes a single appointment
'Parameters:  lngApptID - ID of appointment to delete
'             strUserName - Name of user deleting the appointment
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'  R003: Added method
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    Call objAppt.InsertRecurApptExc(lngApptID, dtApptDate, strUserName)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Sub

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort

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

    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set FetchProviderExceptions = objAppt.FetchProviderExceptions(lngProviderID, _
            dtStartDateTime, dtEndDateTime)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
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
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set FetchECDetail = objAppt.FetchECDetail(lngApptID, lngPatientID)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort

End Function
Public Function FetchEncByDOS(ByVal lngProviderID As Long, ByVal lngApptID As Long, ByVal dteDOS As Date) As ADODB.Recordset
   Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CApptDB")

    Set FetchEncByDOS = objAppt.FetchEncByDOS(lngProviderID, lngApptID, dteDOS)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort

End Function
Public Function GetRecurApptDates(ByVal dteSeekStartDate As Date, ByVal dteSeekEndDate As Date, _
        ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, _
        ByVal lngDuration As Long, ByVal strRecurPattern As String, _
        ByVal lngInterval As Long, ByVal lngDOWMask As Long, _
        ByVal lngDOM As Long, ByVal lngWOM As Long, _
        ByVal lngMOY As Long, Optional blnReturnString As Boolean = False) As Variant
'--------------------------------------------------------------------
'Date: 12/10/2001
'Author: Dave Richkun
'Description:   Returns array of recurring dates for a recurring appointment within
'                   a date range.  The date range may or may not be the entire span
'                   of the recurring appointment.
'Parameters:    dteSeekStartDate - Date on which the method should start looking for
'                   recurring appointments.
'               dteSeekEndDate -  Date on which the method should stop looking for
'                   recurring appointments.
'               dteStartDateTime - The Start DateTime of the recurring appointment as
'                   it is recorded in the database.
'               dteEndDateTime - The End DateTime of the recurring appointment as
'                   it is recorded in the database.
'               lngDuration - recur. appt. duration(minutes)        '
'               strRecurPattern - recur. appt. pattern              '
'               lngInterval - recur. appt. interval                 '
'               lngDOWMask - recur. appt. Day Of Week Mask          '
'               lngDOM - recur. appt. Day Of Month                  '
'               lngWOM - recur. appt. Week Of Month                 '
'               lngMOY - recur. appt. Month Of Year                 '
'Returns:   Array of applicable dates if found, NULL otherwise      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------

    Dim strDates As String
    Dim dteTemp As Date
    Dim dteMonth As Date
    Dim i As Long
    Dim lngNDOM As Long
    Dim lngOffset As Long
    Dim varAry As Variant

    On Error GoTo Err_Trap

    Select Case strRecurPattern
        Case "D" 'Daily Recurring Appointment'
            If lngDOWMask > 0 Then
                ' Recur Every Weekday   '
                ' OR                    '
                ' Recur Every Weekend   '
                dteTemp = DateValue(dteStartDateTime)
                ' Look for dates while currrent calculated date is  '
                ' prior to recur. appt end date and prior to seek   '
                ' end date                                          '
                While dteTemp <= DateValue(dteSeekEndDate) And dteTemp <= DateValue(dteEndDateTime)
                    If IsANDed((2 ^ (DatePart("w", dteTemp) - 1)), lngDOWMask) And _
                    dteTemp >= DateValue(dteSeekStartDate) Then
                        ConcatString CStr(dteTemp), strDates
                    End If
                    dteTemp = DateAdd("d", 1, dteTemp)
                Wend
            Else
                ' Recur Every X Day(s)  '
                dteTemp = DateValue(dteStartDateTime)
                ' Look for dates while currrent calculated date is  '
                ' prior to recur. appt end date and prior to seek   '
                ' end date                                          '
                While dteTemp <= DateValue(dteSeekEndDate) And dteTemp <= DateValue(dteEndDateTime)
                    If dteTemp >= DateValue(dteSeekStartDate) Then
                    ' applicable date found....log it!  '
                        ConcatString CStr(dteTemp), strDates
                    End If

                    dteTemp = DateAdd("d", lngInterval, dteTemp)
                Wend
            End If

        Case "W" ' -= Weekly Recurring Appointment =-'
            For i = vbSunday To vbSaturday
                ' If current vbDay is in the mask then...
                If IsANDed((2 ^ (i - 1)), lngDOWMask) Then
                    ' Determine the first date falling on the current vbDay
                    lngOffset = i - DatePart("w", dteStartDateTime)
                    If lngOffset < 0 Then
                        lngOffset = lngOffset + 7
                    End If

                    ' Now we do the dirty work and begin finding valid dates
                    dteTemp = DateValue(DateAdd("d", lngOffset, dteStartDateTime))
                    While dteTemp <= DateValue(dteSeekEndDate) And dteTemp <= DateValue(dteEndDateTime)
                        If dteTemp >= DateValue(dteSeekStartDate) Then
                            ConcatString CStr(dteTemp), strDates
                        End If
                        dteTemp = DateAdd("ww", lngInterval, dteTemp)
                    Wend
                End If
            Next

        Case "M" ' -= Monthly Recurring Appointment =-'
            If lngDOWMask Then
            'The [1st, 2nd, 3rd, 4th, Last] [Sun, Mon, Tue, Wed, Thu, Fri, Sat] of every Y month(s)
                dteMonth = CDate(DatePart("m", dteStartDateTime) & "/1/" & DatePart("yyyy", dteStartDateTime))
                lngNDOM = ((Log(lngDOWMask) / Log(2))) + 1
                dteTemp = DateValue(GetXDayOfMonth(dteMonth, lngNDOM, lngDOM))

                While dteTemp <= DateValue(dteSeekEndDate) And dteTemp <= DateValue(dteEndDateTime)
                    If dteTemp >= DateValue(dteStartDateTime) And dteTemp >= DateValue(dteSeekStartDate) Then
                        ConcatString CStr(dteTemp), strDates
                    End If
                    dteMonth = DateAdd("m", lngInterval, dteMonth)
                    dteTemp = GetXDayOfMonth(dteMonth, lngNDOM, lngDOM)
                Wend
            Else
            ' Day X of Every Y month(s)                                 '
            ' OR                                                        '
            ' Recur Every [1st, 2nd, 3rd, 4th] Day of Every Y month(s)  '
            ' OR                                                        '
            ' Recur Every LAST Day of Every Y month(s)                  '
                dteMonth = CDate(DatePart("m", dteStartDateTime) & "/1/" & DatePart("yyyy", dteStartDateTime))
                dteTemp = DateValue(GetDayOfMonth(dteMonth, lngDOM))

                While dteTemp <= DateValue(dteSeekEndDate) And dteTemp <= DateValue(dteEndDateTime)
                    If dteTemp >= DateValue(dteStartDateTime) And dteTemp >= DateValue(dteSeekStartDate) Then
                        ConcatString CStr(dteTemp), strDates
                    End If
                    dteMonth = DateAdd("m", lngInterval, dteMonth)
                    dteTemp = GetDayOfMonth(dteMonth, lngDOM)
                Wend
            End If
        Case "Y" ' -= Yearly Recurring Appointment =-'
            If lngDOWMask Then
            ' [1st, 2nd, 3rd, 4th, Last] [Sun, Mon, Tue, Wed, Thu, Fri, Sat] of '
            ' [Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec]      '
                dteMonth = CDate(DatePart("m", dteStartDateTime) & "/1/" & DatePart("yyyy", dteStartDateTime))
                lngNDOM = ((Log(lngDOWMask) / Log(2))) + 1
                dteTemp = GetXDayOfMonth(dteMonth, lngNDOM, lngDOM)

                While dteTemp <= DateValue(dteSeekEndDate) And dteTemp <= DateValue(dteEndDateTime)
                    If dteTemp >= DateValue(dteStartDateTime) And dteTemp >= DateValue(dteSeekStartDate) Then
                        ConcatString CStr(dteTemp), strDates
                    End If
                    dteMonth = DateAdd("yyyy", lngInterval, dteMonth)
                    dteTemp = GetXDayOfMonth(dteMonth, lngNDOM, lngDOM)
                Wend
            Else
            ' Every [1st, 2nd, 3rd, 4th, Last] Day                          '
            ' [Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec]  '
                dteMonth = CDate(DatePart("m", dteStartDateTime) & "/1/" & DatePart("yyyy", dteStartDateTime))
                dteTemp = GetDayOfMonth(dteMonth, lngDOM)

                While dteTemp <= DateValue(dteSeekEndDate) And dteTemp <= DateValue(dteEndDateTime)
                    If dteTemp >= DateValue(dteStartDateTime) And dteTemp >= DateValue(dteSeekStartDate) Then
                        ConcatString CStr(dteTemp), strDates
                    End If
                    dteMonth = DateAdd("yyyy", 1, dteMonth)
                    dteTemp = GetDayOfMonth(dteMonth, lngDOM)
                Wend
            End If

    End Select

    If Trim(strDates) > "" Then
        varAry = Split(strDates, ",")
        SortDates varAry

        If blnReturnString Then
            GetRecurApptDates = Join(varAry, ",")
        Else
            GetRecurApptDates = varAry
        End If
    Else
        GetRecurApptDates = Null
    End If

    Exit Function

Err_Trap:
    Err.Raise Err.Number, CLASS_NAME & ":GetRecurApptDates()", Err.Description
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
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set FetchFuturePatientApptByProvider = objAppt.FetchFuturePatientApptByProvider(lngPatientID, lngProviderID)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
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
'  R008: Created
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CApptDB
    Dim rst As ADODB.Recordset

    On Error GoTo ErrTrap:

    'Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")
    Set rst = objAppt.FetchFutureRecurPatientApptByProvider(lngPatientID, lngProviderID)
    Set objAppt = Nothing

    Set FetchFutureRecurPatientApptByProvider = rst

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
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
    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set FetchFuturePatientApptByManager = objAppt.FetchFuturePatientApptByManager(lngPatientID, lngUserID)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
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
'             lngUserID - ID of office manager
'Returns:   Recordset of recur appointments
'--------------------------------------------------------------------
'Revision History:
'  R008: Created
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CApptDB
    Dim rst As ADODB.Recordset

    On Error GoTo ErrTrap:

    'Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")
    Set rst = objAppt.FetchFutureRecurPatientApptByManager(lngPatientID, lngUserID)
    Set objAppt = Nothing

    Set FetchFutureRecurPatientApptByManager = rst

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function GetConflicts(ByVal lngApptID As Long, ByVal lngProviderID As Long, _
                ByVal varDates As Variant, ByVal intDuration As Integer) As Variant
'--------------------------------------------------------------------
'Date: 10/25/2001
'Author: Dave Richkun
'Description:  Returns a variant array of dates/times from Provider's calendar that
'              will conflict with appointment dates passed in the varDates parameter
'Parameters:   lngApptID - ID of appointment.  If value of 'zero' then new appointment
'                   is being made and the parameter is irrelevant; if the value is a
'                   positive number, we want to exclude the appointment from conflicting
'                   with itself.
'              lngProviderID - ID of Provider whose calendar is being searched
'              varDates - Single dimensional array of dates that will be checked for
'                   appointment conflicts.  Dates are expected to include the date
'                   and time portions i.e. 10/25/2001 8:30:00 AM
'              intDuration - The duration of the appointment dates in minutes
'Returns:      Array of conflicting dates if any, NULL otherwise
'--------------------------------------------------------------------

    On Error GoTo ErrTrap:

    Dim intCtr1 As Integer
    Dim intCtr2 As Integer
    Dim intCtr3 As Integer
    Dim objAppt As ApptDB.CApptDB
    Dim rstAppt As ADODB.Recordset
    Dim intNumConflicts As Integer
    Dim varConflict() As Variant
    Dim varTemp As Variant

    'Defensive programming
    If IsNull(varDates) Then
        Erase varConflict
        GetConflicts = Null
        Exit Function
    End If

    intNumConflicts = -1
    Set objAppt = CreateObject("ApptDB.CApptDB")
    For intCtr1 = 0 To UBound(varDates)
        'Check for conflicts, consider recurring appointments
        Set rstAppt = objAppt.FetchConflictsByProvider(lngProviderID, varDates(intCtr1), intDuration)
        For intCtr2 = 1 To rstAppt.RecordCount
            If rstAppt.Fields("fldRecurYN").Value = "N" Then
                'Exclude patient appts with patient count of zero - these appts have been cancelled
                If rstAppt.Fields("fldCategoryID").Value = 1 And rstAppt.Fields("PatientCount").Value > 0 Then
                    If rstAppt.Fields("fldApptID").Value <> lngApptID Then  'The appot can not conflict with itself.
                        'The appointment is conflicting
                        intNumConflicts = intNumConflicts + 1
                        ReDim Preserve varConflict(3, intNumConflicts)
                        varConflict(0, intNumConflicts) = rstAppt.Fields("fldApptID").Value
                        varConflict(1, intNumConflicts) = rstAppt.Fields("fldStartDateTime").Value
                        varConflict(2, intNumConflicts) = rstAppt.Fields("fldDuration").Value
                        If rstAppt.Fields("fldCategoryID").Value = 1 Then
                            varConflict(3, intNumConflicts) = rstAppt.Fields("PatientName").Value
                        Else
                            varConflict(3, intNumConflicts) = rstAppt.Fields("fldDescription").Value
                        End If
                    End If
                End If
            Else
                'The appointment may be conflicting.  Check the recurrance to determine if actually a conflict.
               varTemp = GetRecurApptDates(rstAppt.Fields("fldStartDateTime").Value, rstAppt.Fields("fldEndDateTime").Value, _
                    rstAppt.Fields("fldStartDateTime").Value, rstAppt.Fields("fldEndDateTime").Value, intDuration, _
                    rstAppt.Fields("fldRecurPattern").Value, rstAppt.Fields("fldInterval").Value, rstAppt.Fields("fldDOWMask").Value, _
                    rstAppt.Fields("fldDOM").Value, rstAppt.Fields("fldWOM").Value, rstAppt.Fields("fldMOY").Value)
                If IsArray(varTemp) Then
                    For intCtr3 = 0 To UBound(varTemp)
                        If DateValue(varTemp(intCtr3)) = DateValue(varDates(intCtr1)) Then
                            'The appointment is conflicting
                            intNumConflicts = intNumConflicts + 1
                            ReDim Preserve varConflict(3, intNumConflicts)
                            varConflict(0, intNumConflicts) = rstAppt.Fields("fldApptID").Value
                            varConflict(1, intNumConflicts) = varTemp(intCtr3) & " " & TimeValue(rstAppt.Fields("fldStartDateTime").Value)
                            varConflict(2, intNumConflicts) = rstAppt.Fields("fldDuration").Value
                            If rstAppt.Fields("fldCategoryID").Value = 1 Then
                                varConflict(3, intNumConflicts) = rstAppt.Fields("PatientName").Value
                            Else
                                varConflict(3, intNumConflicts) = rstAppt.Fields("fldDescription").Value
                            End If
                        End If
                    Next intCtr3
                    
                    Erase varTemp
                End If
            End If
        
            rstAppt.MoveNext
        Next intCtr2
        Set rstAppt = Nothing
    Next intCtr1
    Set objAppt = Nothing

    If intNumConflicts < 0 Then
        Erase varConflict
        GetConflicts = Null
    Else
        GetConflicts = varConflict
    End If

    Exit Function

ErrTrap:
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub FetchNextOpenProviderTimeSlot(ByVal lngProviderID As Long, _
                                              ByVal lngLength As Long, _
                                              ByVal dteSearch As Date, _
                                              ByRef varStartTime As Variant, _
                                              ByRef varEndTime As Variant)
'--------------------------------------------------------------------
'Date: 04/19/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Fetches a provider's next open time slot in relation '
'               the given date parameter and time length required   '
'Parameters:   lngProviderID - Provider ID                          '
'              lngLength - Minimum length of minutes required       '
'              dteSearch - Date/Time to begin looking for open slot '
'              dteStartTime - Start date/time value of open slot    '
'              dteendTime - End date/time value of open slot        '
'Returns:      Nothing                                              '
'--------------------------------------------------------------------
    Dim strStartTimes As String, strEndTimes As String
    Dim aryStartTimes As Variant, aryEndTimes As Variant
    Dim dteStartTime As Date, dteEndTime As Date
    Dim dteStartSearch As Date
    Dim dteEndDummy As Date
    Dim varTimer As Variant

    Dim aryRet As Variant

    Const MIN_TIME = " 07:00 AM"

    On Error GoTo Error_Handler

    dteStartTime = -1
    dteEndTime = -1
    varTimer = Now()

TRY_AGAIN:
    If DateDiff("s", varTimer, Now()) > 6 Then
        dteEndTime = DateValue(dteSearch)
    Else
        If DatePart("h", dteSearch) < 7 Then
            dteStartSearch = CDate(DateValue(dteSearch) & MIN_TIME)
        Else
            dteStartSearch = dteSearch
        End If

        ' Retrieve array
        aryRet = GetClosedTimeSlotArray(lngProviderID, dteSearch)

        ' Sort array(s)
        If Not IsNull(aryRet) Then
            SortClosedDates aryRet(0), aryRet(1)
        End If
        ' Now analyze array
        AnalyzeArray aryRet, dteStartSearch, lngLength, dteStartTime, dteEndTime

        If dteStartTime < 0 Then
            dteSearch = DateAdd("d", 1, CDate(DateValue(dteSearch)))
            GoTo TRY_AGAIN
        End If
   End If

    varStartTime = dteStartTime
    varEndTime = dteEndTime

    GetObjectContext.SetComplete
    Exit Sub

Error_Handler:
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'--------------------------------------------------------------------
' PrivateMethods    +++++++++++++++++++++++++++++++++++++++++++++++++
'--------------------------------------------------------------------

Private Function AnalyzeArray(ByVal varDates As Variant, _
                              ByVal dteStart As Date, _
                              ByVal lngLength As Long, _
                              ByRef dteStartTime As Date, _
                              ByRef dteEndTime As Date)

'--------------------------------------------------------------------
'Date: 04/19/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Analyzes an array of appointment and searches for    '
'               the first available open slot time with a length of '
'               at least the parameter given(lngLength)             '
'Parameters:    varDates - array of appointments                    '
'               dteStart - date to begin looking for open slot      '
'               lngLength - minimum length open slot has to be      '
'               dteStartTime - Start date/time of open slot         '
'               dteEndTime - End date/time of open slot             '
'Returns:      Nothing                                              '
'--------------------------------------------------------------------
    Const MIN_TIME = " 7:00 AM"
    Const MAX_TIME = " 10:00 PM"

    Dim intCnt As Integer
    Dim dteCurrentTime As Date
    Dim dteMaxTime As Date
    Dim aryStart As Variant, aryEnd As Variant
    Dim dteEOD As Date

    dteEOD = CDate(DateValue(dteStart) & MAX_TIME)
    If dteStart >= dteEOD Then Exit Function

    If IsNull(varDates) Then
    ' A null values represents an open day
        dteStartTime = CDate(DateValue(dteStart) & MIN_TIME)
        dteEndTime = CDate(DateValue(dteStart) & MAX_TIME)
    Else

        aryStart = varDates(0)
        aryEnd = varDates(1)
        dteMaxTime = CDate(DateValue(dteStart) & MAX_TIME)

        For intCnt = 0 To UBound(aryStart)

            If aryStart(intCnt) >= dteStart Then
                If DateDiff("n", dteStart, aryStart(intCnt)) >= lngLength Then
                ' We got a winner, report it and move on
                    dteStartTime = dteStart
                    dteEndTime = CDate(aryStart(intCnt))
                    Exit For
                Else
                ' Go Fish
                    dteStart = MaxDate(CDate(aryEnd(intCnt)), dteStart)
                End If
            Else
            ' Go Fish
                dteStart = MaxDate(CDate(aryEnd(intCnt)), dteStart)
            End If
        Next
    End If

    ' Final check
    If dteStartTime < 0 Then
        If dteEOD >= dteStart Then
            If DateDiff("n", dteStart, dteEOD) >= lngLength Then
            ' We got a winner, report it and move on
                dteStartTime = dteStart
                dteEndTime = dteEOD
            End If
        End If
    End If

End Function


Private Function MaxDate(ByVal dteDate1 As Date, dteDate2 As Date) As Date
'--------------------------------------------------------------------
'Date: 04/19/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Compares two dates values and returns the greater value
'Parameters:   dteDate1, dteDate2  - dates to be compared           '
'Returns:      Larger Date of the two parameters                     '
'--------------------------------------------------------------------
    If dteDate2 > dteDate1 Then
        MaxDate = dteDate2
    Else
        MaxDate = dteDate1
    End If
End Function

Private Function GetClosedTimeSlotArray(ByVal lngProviderID As Long, _
                                     ByVal dteSearch As Date) _
                                     As Variant
'--------------------------------------------------------------------
'Date: 04/19/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Obtains a provider's schedule and builds an appointment
'               array based on the acquired information             '
'Parameters:    lngProviderID - Provider ID                         '
'               dteSearch - Date to obtain information for          '
'Returns:      Array of date/time ranges                            '
'--------------------------------------------------------------------

    Dim rstDaySchedule As ADODB.Recordset
    Dim strStartTimes As String, strEndTimes As String
    Dim aryStartTimes As Variant, aryEndTimes As Variant
    Dim aryRet(1) As Variant

    Set rstDaySchedule = FetchByProviderDateRange(lngProviderID, dteSearch, dteSearch)

    With rstDaySchedule
        While Not .EOF
            If .Fields("fldRecurPattern") > "" Then

                If Not IsNull(GetRecurApptDates(DateValue(dteSearch), DateValue(dteSearch), _
                .Fields("fldStartDateTime").Value, .Fields("fldEndDateTime").Value, _
                .Fields("fldDuration").Value, .Fields("fldRecurPattern").Value, _
                .Fields("fldInterval").Value, .Fields("fldDOWmask").Value, _
                .Fields("fldDOM").Value, .Fields("fldWOM").Value, _
                .Fields("fldMOY").Value, True)) Then

                    If strStartTimes > "" Then
                        strStartTimes = strStartTimes & "," & DateValue(dteSearch) & " " & TimeValue(.Fields("fldStartDateTime").Value)
                        strEndTimes = strEndTimes & "," & DateValue(dteSearch) & " " & TimeValue(.Fields("fldEndDateTime").Value)
                    Else
                        strStartTimes = DateValue(dteSearch) & " " & TimeValue(.Fields("fldStartDateTime").Value)
                        strEndTimes = DateValue(dteSearch) & " " & TimeValue(.Fields("fldEndDateTime").Value)
                    End If

                End If
            Else
                If strStartTimes > "" Then
                    strStartTimes = strStartTimes & "," & .Fields("fldStartDateTime").Value
                    strEndTimes = strEndTimes & "," & .Fields("fldEndDateTime").Value
                Else
                    strStartTimes = .Fields("fldStartDateTime").Value
                    strEndTimes = .Fields("fldEndDateTime").Value
                End If
            End If
            .MoveNext
       Wend
    End With
    Set rstDaySchedule = Nothing

    If strStartTimes > "" Then
        aryStartTimes = Split(strStartTimes, ",")
        aryEndTimes = Split(strEndTimes, ",")
        aryRet(0) = aryStartTimes
        aryRet(1) = aryEndTimes
        GetClosedTimeSlotArray = aryRet
    Else
        GetClosedTimeSlotArray = Null
    End If

    Exit Function

Error_Handler:
    Set rstDaySchedule = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub SortClosedDates(ByRef aryStart As Variant, ByRef aryEnd As Variant)
'--------------------------------------------------------------------
'Date: 04/19/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Sorts a 2D array of dates                            '
'Parameters:   Array of dates                                       '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim dteTemp As Date

    Dim i As Long
    Dim j As Long

    For j = (UBound(aryStart) - 1) To 0 Step -1

        For i = 0 To j
            ' Sort by Start Dates
            If CDate(aryStart(i)) > CDate(aryStart(i + 1)) Then
                dteTemp = aryStart(i)
                aryStart(i) = aryStart(i + 1)
                aryStart(i + 1) = dteTemp
                dteTemp = aryEnd(i)
                aryEnd(i) = aryEnd(i + 1)
                aryEnd(i + 1) = dteTemp

            ElseIf CDate(aryStart(i)) = CDate(aryStart(i + 1)) Then
                If CDate(aryEnd(i)) > CDate(aryEnd(i + 1)) Then
                ' Then sort by End Date
                    dteTemp = aryStart(i)
                    aryStart(i) = aryStart(i + 1)
                    aryStart(i + 1) = dteTemp
                    dteTemp = aryEnd(i)
                    aryEnd(i) = aryEnd(i + 1)
                    aryEnd(i + 1) = dteTemp
                End If
            End If
        Next

    Next

End Sub

Private Function GetCachedDates(ByVal lngApptID As Long, ByVal dteStartDateTime As Date, _
                      ByVal dteEndDateTime As Date, ByVal lngDuration As Long, _
                      ByVal strRecurPattern As String, ByVal lngInterval As Long, _
                      ByVal lngDOWMask As Long, ByVal lngDOM As Long, _
                      ByVal lngWOM As Long, ByVal lngMOY As Long, _
                      ByRef dicAppts As Dictionary) As Variant
'--------------------------------------------------------------------
'Date: 04/16/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Retrieves calculated dates in a Hit/Miss basis       '
'Parameters:   Values to be calculated(if required)                 '
'Returns:      Array of dates                                       '
'--------------------------------------------------------------------

    If Not dicAppts.Exists(CStr(lngApptID)) Then
        dicAppts.Add CStr(lngApptID), GetRecurApptDates(dteStartDateTime, dteEndDateTime, _
                        dteStartDateTime, dteEndDateTime, lngDuration, _
                        strRecurPattern, lngInterval, lngDOWMask, lngDOM, _
                        lngWOM, lngMOY)
    End If
    GetCachedDates = dicAppts.Item(CStr(lngApptID))
End Function

Private Function ValidateSingle(ByVal lngProviderID As Long, ByVal lngClinicID As Long, _
        ByVal lngCategoryID As Long, ByVal dteStartDateTime As Date, _
        ByVal dteEndDateTime As Date, ByVal lngDuration As Long, ByVal strNote As String, _
        ByVal strUserName As String, ByVal varPatientArray As Variant, _
        ByVal strCPTCode As String, ByRef strErrMessage) As Boolean
'--------------------------------------------------------------------
'Date: 08/25/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Validates data before allowing inserts and update                '
'Parameters:  The values to be checked.                             '
'Returns:   True if all data is valid, False otherwise              '
'--------------------------------------------------------------------

    Dim intCtr As Integer

    ' Check provider ID
    If lngProviderID < 1 Then
        strErrMessage = "Invalid Provider ID passed."
        ValidateSingle = False
        Exit Function
    End If

    ' Check category ID
    If lngCategoryID <= 0 Then
        strErrMessage = "Invalid Appointment Category ID passed."
        ValidateSingle = False
        Exit Function
    End If

    ' Check for a valid date range
    If dteStartDateTime >= dteEndDateTime Then
        strErrMessage = "Start Date/Time cannot be greater than End Date/Time."
        ValidateSingle = False
        Exit Function
    End If

    If lngCategoryID = APPT_TYPE_PATIENT Then
        'Check clinic ID
        If (lngClinicID <= 0) Then
            strErrMessage = "Invalid Clinic ID passed."
            ValidateSingle = False
            Exit Function
        End If
    
        'If a CPT code is given, make sure it is valid
        If Len(Trim(strCPTCode)) > 0 Then
            Dim objCPTCode As ListBz.CCPTCodeBz
            Set objCPTCode = CreateObject("ListBz.CCPTCodeBz")
            If Not objCPTCode.Exists(Trim(strCPTCode)) Then
                strErrMessage = "Invalid CPT Code passed."
                ValidateSingle = False
                Set objCPTCode = Nothing
                Exit Function
            End If
            Set objCPTCode = Nothing
        End If
    
        'Check for a valid patient array and array values
        If Not IsAnArray(varPatientArray) Then
            strErrMessage = "No valid Patient ID's were passed."
            ValidateSingle = False
            Exit Function
        Else
            For intCtr = 0 To UBound(varPatientArray, 1)
                If varPatientArray(intCtr, 1) < 1 Then
                    strErrMessage = "At least 1 invalid Patient ID was passed."
                    ValidateSingle = False
                    Exit Function
                End If
            Next
        End If
    End If

    ' Check for a user name
    If Trim(strUserName) = "" Then
        strErrMessage = "User name is required"
        ValidateSingle = False
        Exit Function
    End If

    ' Check for valid start/end times and duration combination
    ' (startime + duration) = endtime
    If DateDiff("n", dteStartDateTime, dteEndDateTime) <> lngDuration Then
        strErrMessage = "Miscalculated Start/End Time and Duration values passed."
        ValidateSingle = False
        Exit Function
    End If

    'If we get here, all is well...
    ValidateSingle = True

End Function

Private Function ValidateRecur(ByVal varPatientArray As Variant, ByVal strCPTCode As String, _
            ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, _
            ByVal lngDuration As Long, ByVal lngProviderID As Long, _
            ByVal lngClinicID As Long, ByVal lngCategoryID As Long, _
            ByVal strRecurPattern As String, ByVal lngInterval As Long, _
            ByVal lngDOWMask As Long, ByVal lngDOM As Long, ByVal lngWOM As Long, _
            ByVal lngMOY As Long, ByVal strUserName As String, _
            ByRef strErrMessage As String) As Boolean

    Dim intCtr As Integer
    
    ValidateRecur = False ' Assume Failure
    
    'Check provider ID
    If lngProviderID < 1 Then
        strErrMessage = "Invalid Provider ID passed."
        Exit Function
    End If

    'Check category ID
    If lngCategoryID < 1 Then
        strErrMessage = "Invalid Appointment Category ID passed."
        Exit Function
    End If

    'Check for a valid date range
    If dteStartDateTime >= dteEndDateTime Then
        strErrMessage = "Invalid date range passed."
        Exit Function
    End If

    'Check for a valid duration
    If lngDuration <= 0 Then
        strErrMessage = "Invalid duration passed."
        Exit Function
    End If
    
    'Check for a valid interval
    If lngInterval <= 0 Then
        strErrMessage = "Invalid interval passed."
        Exit Function
    End If
    
    If lngCategoryID = APPT_TYPE_PATIENT Then
        'Check clinic ID
        If lngClinicID < 1 Then
            strErrMessage = "Invalid Clinic ID passed."
            Exit Function
        End If

        'If a CPT code is given, make sure it is valid
        If Len(Trim(strCPTCode)) > 0 Then
            Dim objCPTCode As ListBz.CCPTCodeBz
            Set objCPTCode = CreateObject("ListBz.CCPTCodeBz")
            If Not objCPTCode.Exists(Trim(strCPTCode)) Then
                strErrMessage = "Invalid CPT Code passed."
                Set objCPTCode = Nothing
                Exit Function
            End If
            Set objCPTCode = Nothing
        End If
    
        'Check for a valid patient array and array values
        If Not IsAnArray(varPatientArray) Then
            strErrMessage = "No valid Patient ID's were passed."
            Exit Function
        Else
            For intCtr = 0 To UBound(varPatientArray, 1)
                If varPatientArray(intCtr, 1) < 1 Then
                    strErrMessage = "At least 1 invalid Patient ID was passed."
                    Exit Function
                End If
            Next
        End If
    End If
    
    Select Case strRecurPattern
        Case "D" 'Daily

        Case "W" 'Weekly
            If lngDOWMask < 1 Then
                strErrMessage = "Invalid day-of-week mask passed."
                Exit Function
            End If

        Case "M" 'Monthly
            ' Check for a valid Day-Of-Month value
            If lngDOM < 0 Or lngDOM > 31 Then
                strErrMessage = "Invalid Day-of-Month passed."
                Exit Function
            End If

            ' Check for a valid Week-Of-Month value
            If lngWOM < 0 Or lngWOM > 4 Then
                strErrMessage = "Invalid Week-of-Month passed."
                Exit Function
            End If

        Case "Y" 'Yearly


        Case Else
            strErrMessage = "Invalid recurring pattern passed."
        Exit Function
    End Select

    ' Check for a user name
    If Trim(strUserName) = "" Then
        strErrMessage = "Current username is required"
        Exit Function
    End If

    ValidateRecur = True ' Everything is ok
    
End Function



Private Sub UpdateSinglePatAppt(ByVal lngApptID As Long, _
                                ByVal varPatApptInfo)
'--------------------------------------------------------------------
'Date: 08/30/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Validates business rules before updating appoinment   '
'               records for an appointment having the ID given in the
'               parameter listing                                   '
'Parameters:  lngApptID - ID of the appointment to be updated       '
'             varPatientApptInfo - For patient appoinments, this is '
'               a 2-D array with elements as follows:               '
'                   0 = PatientAppointment ID (0)                   '
'                   1 = Patient ID                                  '
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim lngID As Long
    Dim i As Integer
    Dim strErrMsg As String
    Dim objPatAppt As ApptBZ.CPatApptBZ
    Dim rst As ADODB.Recordset

    On Error GoTo Error_Handler

    Set objPatAppt = CreateObject("ApptBZ.CPatApptBZ")

    For i = 0 To UBound(varPatApptInfo, 1)

        If varPatApptInfo(i, 0) = 0 Then
            ' This a new pat/appt rec ....  simple insert
            lngID = objPatAppt.Insert(lngApptID, varPatApptInfo(i, 1))

            ' Check for DB error
            If lngID < 1 Then
                strErrMsg = "An error occured while trying to add a Patient/Appointment record."
            End If

        Else
            ' Existing rec with a possible status change
            Set rst = objPatAppt.FetchByID(varPatApptInfo(i, 0))
            Call objPatAppt.Update(varPatApptInfo(i, 0), lngApptID, varPatApptInfo(i, 1))
            Set rst = Nothing

        End If
    Next

    Set objPatAppt = Nothing
    Exit Sub
Error_Handler:
    Set objPatAppt = Nothing
    Set rst = Nothing

    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Sub DeleteBatchPatAppt(ByVal lngApptID As Long, _
                                ByVal varPatApptInfo)
'--------------------------------------------------------------------
'Date: 09/05/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Validates business rules before deleting              '
'               Patient/Appointment records                         '
'Parameters:  lngApptID - ID of the associated appointment record   '
'             varPatientApptInfo - For patient appoinments, this is '
'               a 2-D array with elements as follows:               '
'                   0 = PatientAppointment ID (0)                   '
'                   1 = Patient ID                                  '
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim i As Integer
    Dim objPatAppt As ApptDB.CPatApptDB
    Dim strErrMsg As String
    Dim strINClause
    Dim rst As ADODB.Recordset

    On Error GoTo Error_Handler

    Set objPatAppt = CreateObject("ApptDB.CPatApptDB")

    ' Build the IN clause parameter
    ' The IN clause represents all patient/appointment records that
    ' are to remain associated with the appointment record
    For i = 0 To UBound(varPatApptInfo, 1)
        If varPatApptInfo(i, 0) > 0 Then
            If Trim(strINClause) > "" Then
                strINClause = strINClause & ", " & varPatApptInfo(i, 0)
            Else
                strINClause = varPatApptInfo(i, 0)
            End If
        End If
    Next


    If Trim(strINClause) > "" Then
        Set rst = objPatAppt.FetchMissingRec(lngApptID, strINClause)

        While Not rst.EOF
            ' Delete existin patient/appointment records if they have not
            ' been marked as attended
            If rst.Fields("fldApptStatusID") <> 3 Then
                Call objPatAppt.Delete(lngApptID, rst.Fields("fldPatientID"))
            Else
                strErrMsg = "Cannot modify an Attended apointment."
                GoTo Error_Handler
            End If

            rst.MoveNext
        Wend
        Set rst = Nothing
    End If

    Set objPatAppt = Nothing
    Exit Sub

Error_Handler:
    Set objPatAppt = Nothing
    Set rst = Nothing
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Sub ConcatString(ByVal strAdd As String, ByRef strAddTo As String, Optional strDel As String = ",")
'--------------------------------------------------------------------
'Date: 01/16/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Concatenates one string to another, delimiting tokens '
'               as necessary                                        '
'Parameters:  strAdd - string to concatenate to other string        '
'             strAddTo - string concatenation is performed on       '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    If Len(Trim(strAddTo)) Then
        strAddTo = strAddTo & strDel & strAdd
    Else
        strAddTo = strAdd
    End If
End Sub

Private Function IsANDed(ByVal lngVal As Long, _
                         ByVal lngMask As Long) As Boolean
'--------------------------------------------------------------------
'Date: 01/16/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Determines if a value is ANDed within a mask          '
'Parameters:  lngVal - value being sought                           '
'             lngMask - masked value being searched                 '
'Returns:    True if value is ANDed, False otherwise                '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    IsANDed = ((lngVal And lngMask) = lngVal)
End Function

Private Sub SortDates(ByRef varArray As Variant)
'--------------------------------------------------------------------
'Date: 01/16/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Orders the date values of an array in chronological   '
'               order                                               '
'Parameters:  varArray - array to be ordered                        '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim dteTemp As Date
    Dim i As Long
    Dim j As Long

    ' Implements Bubble Sort algorithm
    ' BAH!  I know it's girly, but it works.
    ' I'm still l337
    If Not IsNull(varArray) Then
        If IsArray(varArray) Then
            For j = (UBound(varArray) - 1) To 0 Step -1

                For i = 0 To j
                    If CDate(varArray(i)) > CDate(varArray(i + 1)) Then
                        dteTemp = varArray(i)
                        varArray(i) = varArray(i + 1)
                        varArray(i + 1) = dteTemp
                    End If
                Next

            Next
        End If
    End If

End Sub

Private Function DaysInMonth(ByVal dteDate As Date) As Integer
'--------------------------------------------------------------------
'Date: 01/16/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Calculates the number of days in a given month        '
'Parameters:  dteDate - Date containing the month being evaluated   '
'Returns:     The number of days in the given month                 '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    DaysInMonth = DateSerial(Year(dteDate), _
                    Month(dteDate) + 1, 1) - _
                    DateSerial(Year(dteDate), _
                    Month(dteDate), 1)
End Function

Private Function GetDayOfMonth(ByVal dteDate As Date, _
                               ByVal lngDOM As Long) As Date
'--------------------------------------------------------------------
'Date: 01/17/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Calculates the Xth or Last day of a given month       '
'Parameters:  dteDate - Date to perform calculation on              '
'             lngDOM - Day Of Month to use for calculation(-1 = last)
'Returns:     adjusted date                                         '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim lngNumDays As Long
    Dim lngDOM2Use As Long
    Dim lngPDOM As Long

    ' Do not get an invalid date.  eg. 02/31/2002
    lngNumDays = CLng(DaysInMonth(dteDate))
    lngPDOM = IIf((lngDOM > 0), lngDOM, 31)

    lngDOM2Use = IIf((lngPDOM > lngNumDays), lngNumDays, lngPDOM)

    GetDayOfMonth = CDate(DatePart("m", dteDate) & "/" & lngDOM2Use & "/" & DatePart("yyyy", dteDate))

End Function

Private Function GetXDayOfMonth(ByVal dteDate As Date, ByVal lngDay As Long, _
                                ByVal lngPlace As Long) As Date
'--------------------------------------------------------------------
'Date: 01/17/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Calculates the Xth or Last "week" day of a given month'
'Parameters:  dteDate - Date to perform calculation on              '
'             lngDay - Day to use (vbSunday(1) thru vbSaturday(7))  '
'             lngPlace - ordinal value of day to calculate          '
'             (1st, 2nd, 3rd, 4th, Last(-1))                        '
'Returns:     adjusted date                                         '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim lngPPlace As Long
    Dim dteTemp As Date
    Dim lngOffset As Long

    lngPPlace = IIf((lngPlace > 4), 4, lngPlace)

    If lngPPlace > 0 Then
    ' 1st, 2nd, 3rd, 4th    '
        lngOffset = lngDay - DatePart("w", dteDate)
        If lngOffset < 0 Then lngOffset = lngOffset + 7

        dteTemp = DateAdd("d", lngOffset, dteDate)
        GetXDayOfMonth = DateAdd("ww", (lngPPlace - 1), dteTemp)
    Else
    ' Last                  '
        dteTemp = CDate(DatePart("m", dteDate) & "/" & DaysInMonth(dteDate) & "/" & DatePart("yyyy", dteDate))
        lngOffset = lngDay - DatePart("w", dteTemp)
        If lngOffset > 0 Then lngOffset = lngOffset - 7
        GetXDayOfMonth = DateAdd("d", lngOffset, dteTemp)
    End If
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

    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set FetchUnBilledAppts = objAppt.FetchUnBilledAppts(lngUserID)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function FetchPendingAppts(ByVal lngUserID As Long) As ADODB.Recordset  'R001
'--------------------------------------------------------------------
'Date: 07/09/2007
'Author: Duane C Orth
'Description:   Retrieves pending appointments older than 3 days
'Parameters: lngUserID - ID of Provider or Office Manager retrieving records
'Returns:   Recordset of  appointments                              '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CApptDB

    On Error GoTo ErrTrap:

    ' Instantiate the appt object
    Set objAppt = CreateObject("ApptDB.CApptDB")

    ' Populate the recordset
    Set FetchPendingAppts = objAppt.FetchPendingAppts(lngUserID)

    Set objAppt = Nothing

    GetObjectContext.SetComplete
    Exit Function

ErrTrap:
    Set objAppt = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub SendMessage(ByVal intMsgType As MsgType, ByVal lngProviderID As Long, _
        ByVal lngPatientID As Long, ByVal strUserName As String, Optional ByVal dteStartDateTime As Date, _
        Optional ByVal intRecurType As Integer, Optional ByVal strCancelReason As String = "")
'--------------------------------------------------------------------
'Date: 10/25/2001
'Author: Dave Richkun
'Description: Sends notification messages to Providers and Office Managers
'             about events pertaining to Non-certified, Cancelled, and
'             Deleted appointments.
'Parameters: intMsgType - Enumerated value identifying the type of message that will be sent
'            lngProviderID - ID of the Provider who will receive messages and whose Office managers may also receive messages
'            lngPatientID - ID of patient to whom the message applies
'            strUserName - Name of user triggering the event that initiated message sending
'            dteStartDateTime - Optional parameter expected when intMsgType = 'NoCert' and when
'                   intMsgType = 'ApptCancel'.  When 'NoCert' this parameter contains value identifying
'                   starting date of a series of recurring apppointments.  When 'ApptCancel' this
'                   parameter contains the date of the cancelled appointment.
'            intRecurType - Optional parameter expected when intMsgType = 'NoCert' - contains value identifying
'                   appointment type i.e. one-time or recurring
'Returns: Null
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------

    Dim objMsg As ListBz.CMsgBz
    Dim objBFact As BenefactorBz.CBenefactorBz
    Dim objUser As ClinicBz.CUserBz
    Dim rstUser As ADODB.Recordset
    Dim rstBFact As ADODB.Recordset
    Dim blnProvider As Boolean
    Dim strPatientName As String
    Dim strProviderName As String
    Dim strMsg As String

    On Error GoTo ErrTrap:

    Set objUser = CreateObject("ClinicBz.CUserBz")

    Set rstUser = objUser.FetchDetail(lngProviderID, blnProvider)
    strProviderName = rstUser.Fields("fldFirstName").Value & " " & rstUser.Fields("fldLastName").Value
    Set rstUser = Nothing
    
    'Build recordset of Office Managers associated with Provider.  They too will be notified.
    Set rstUser = objUser.FetchManagersByProvider(lngProviderID)
    
    Set objMsg = CreateObject("ListBz.CMsgBz")
    Set objBFact = CreateObject("BenefactorBz.CBenefactorBz")
        
    Set rstBFact = objBFact.FetchByID(lngPatientID)
    strPatientName = rstBFact.Fields("fldFirst").Value & " " & rstBFact.Fields("fldLast").Value
    Set rstBFact = Nothing

    Select Case intMsgType
        Case ApptCreateNoCert
            If intRecurType = 3 Then
                strMsg = "One or more recurring appointments were made for " & strPatientName & " starting "
            Else
                strMsg = "An appointment was made for " & strPatientName & " "
            End If
            strMsg = strMsg & Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
            strMsg = strMsg & " without primary certification."

            'Notify the Provider
            Call objMsg.Insert(strMsg, strUserName, lngProviderID, , "N")

            'Notify all Office Managers
            If intRecurType = 3 Then
                strMsg = "One or more recurring appointments were made for " & strPatientName & " with " & strProviderName & " starting at "
            Else
                strMsg = "An appointment was made for " & strPatientName & " with " & strProviderName & " at "
            End If
            strMsg = strMsg & Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
            strMsg = strMsg & " without primary certification."

            If Not rstUser.BOF And Not rstUser.EOF Then
                rstUser.MoveFirst
                While Not rstUser.EOF
                    If rstUser.Fields("fldDisabledYN").Value = "N" Then
                        Call objMsg.Insert(strMsg, strUserName, rstUser.Fields("fldUserID").Value, , "N")
                    End If
                    rstUser.MoveNext
                Wend
            End If
    
        Case ApptConfirmNoCert
            strMsg = "Your appointment with " & strPatientName & " was confirmed for "
            strMsg = strMsg & Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
            strMsg = strMsg & " without primary certification.  (" & strUserName & ")"

            'Notify the Provider
            Call objMsg.Insert(strMsg, strUserName, lngProviderID, , "N")

            'Notify all Office Managers
            strMsg = "An appointment with " & strPatientName & " for " & strProviderName _
            & " was confirmed for " & Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
            strMsg = strMsg & " without primary certification.  (" & strUserName & ")"
            
            If Not rstUser.BOF And Not rstUser.EOF Then
                rstUser.MoveFirst
                While Not rstUser.EOF
                    If rstUser.Fields("fldDisabledYN").Value = "N" Then
                        Call objMsg.Insert(strMsg, strUserName, rstUser.Fields("fldUserID").Value, , "N")
                    End If
                    rstUser.MoveNext
                Wend
            End If
    
    Case ApptCancel
        strMsg = "Your appointment on " & Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
        strMsg = strMsg & " with " & strPatientName & " has been cancelled. (" & strUserName
        If Len(strCancelReason) > 1 Then strMsg = strMsg & ": " & strCancelReason
        strMsg = strMsg & ")"

        'Notify the Provider
        'Call objMsg.Insert(strMsg, strUserName, lngProviderID, , "C")

        'Notify all Office Managers
        If Not rstUser.BOF And rstUser.EOF Then
            rstUser.MoveFirst
            While Not rstUser.EOF
                If rstUser.Fields("fldDisabledYN").Value = "N" Then
                    Call objMsg.Insert(strMsg, strUserName, rstUser.Fields("fldUserID").Value, , "C")
                End If
                rstUser.MoveNext
            Wend
        End If

        Case ApptDelete
            strMsg = "Your appointment on " & Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
            strMsg = strMsg & " with " & strPatientName & " was deleted. (" & strUserName & ")"
    
            'Notify the Provider
            Call objMsg.Insert(strMsg, strUserName, lngProviderID, , "C")
    
            'Notify all Office Managers
            If Not rstUser.BOF And rstUser.EOF Then
                rstUser.MoveFirst
                While Not rstUser.EOF
                    If rstUser.Fields("fldDisabledYN").Value = "N" Then
                        Call objMsg.Insert(strMsg, strUserName, rstUser.Fields("fldUserID").Value, , "C")
                    End If
                    rstUser.MoveNext
                Wend
            End If
    End Select

    Set rstUser = Nothing
    Set objMsg = Nothing
    Set objBFact = Nothing
    Set objUser = Nothing

    Exit Sub
    
ErrTrap:
    Set objMsg = Nothing
    Set objBFact = Nothing
    Set objUser = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub





