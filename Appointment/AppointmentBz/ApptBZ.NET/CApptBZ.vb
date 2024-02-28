Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("CApptBZ_NET.CApptBZ")> Public Class CApptBZ
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
	
	Private Const CLASS_NAME As String = "CApptBZ"
	
	Private Const START_TIME As String = "07:00:00 AM"
	
	Private Const HEADING_T As Integer = -1
	Private Const OPEN_T As Integer = 0
	Private Const SCHEDULED_T As Integer = 1
	
	Private Const PATIENT_CAT As Integer = 1
	Private Const BLOCK_CAT As Integer = 2
	
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
	
	Private Const ATTENDED_ST As Integer = 3
	Private Const CONFIRMED_ST As Integer = 2
	Private Const HOLD_ST As Integer = 5
	Private Const SCHEDULED_ST As Integer = 1
	Private Const NO_SHOW_ST As Integer = 6
	Private Const TENATIVE_ST As Integer = 10
	Private Const PENDING_ST As String = "11"
	
	Private Enum MsgType
		ApptCreateNoCert = 1
		ApptConfirmNoCert = 2
		ApptCancel = 3
		ApptDelete = 4
	End Enum
	
	Private Const APPT_TYPE_PATIENT As Integer = 1
	Private Const APPT_TYPE_BLOCK As Integer = 2
	
	Private Const APPT_STATUS_SCHEDULED As Short = 1
	Private Const APPT_STATUS_CONFIRMED As Short = 2
	Private Const APPT_STATUS_ATTENDED As Short = 3
	Private Const APPT_STATUS_CANCELLED As Short = 4
	Private Const APPT_STATUS_PENDING As Short = 11
	
	Public Enum typApptType
		apptTypeSingle = 1
		apptTypeRecurring = 2
	End Enum
	
	'--------------------------------------------------------------------
	' Public Methods
	'--------------------------------------------------------------------
	
	Public Function FetchByProviderDateRange(ByVal lngProviderID As Integer, ByVal dteStartDate As Date, ByVal dteEndDate As Date) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchByProviderDateRange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchByProviderDateRange = objAppt.FetchByProviderDateRange(lngProviderID, dteStartDate, dteEndDate)
		'Do not know they are Why adding an extra day to the search criteria disabled 05/08/2008
		'Set FetchByProviderDateRange = objAppt.FetchByProviderDateRange(lngProviderID, dteStartDate, DateAdd("d", 1, dteEndDate))
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	Public Function FindOpenApptTimeSlots(ByVal lngClinicID As Integer, ByVal lngProviderID As Integer, ByVal dteStartDate As Date, ByVal dteEndDate As Date) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FindOpenApptTimeSlots. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FindOpenApptTimeSlots = objAppt.FindOpenApptTimeSlots(lngClinicID, lngProviderID, dteStartDate, dteEndDate)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	Public Function FetchByCheckInDate(ByVal lngUserID As Integer, ByVal dteArrivalDate As Date) As Object
		Dim ApptDB As Object
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
		Dim aryAppts() As Object
		Dim aryRecurDates As Object
		Dim intCtr As Short
		Dim intCtr2 As Short
		Dim intCtr3 As Short
		Dim intApptCount As Short
		Dim blnException As Boolean
		
		On Error GoTo ErrTrap
		
		objAppt = CreateObject("ApptDB.CApptDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchByCheckInDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rstAppt = objAppt.FetchByCheckInDate(lngUserID, dteArrivalDate)
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchCheckInExceptions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rstExc = objAppt.FetchCheckInExceptions(lngUserID, dteArrivalDate)
		
		If rstAppt.RecordCount > 0 Then
			ReDim aryAppts(4, rstAppt.RecordCount - 1)
			
			For intCtr = 0 To rstAppt.RecordCount - 1
				blnException = False
				'Get valid recurring dates if any
				If Trim(rstAppt.Fields("fldRecurPattern").Value) > "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object GetRecurApptDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryRecurDates = GetRecurApptDates(dteArrivalDate, dteArrivalDate, rstAppt.Fields("fldStartDateTime").Value, rstAppt.Fields("fldEndDateTime").Value, rstAppt.Fields("fldDuration").Value, rstAppt.Fields("fldRecurPattern").Value, rstAppt.Fields("fldInterval").Value, rstAppt.Fields("fldDOWMask").Value, rstAppt.Fields("fldDOM").Value, rstAppt.Fields("fldWOM").Value, rstAppt.Fields("fldMOY").Value)
					
					'Look for recuring appointment date exceptions before including the date in the 'aryAppts' array
					If IsArray(aryRecurDates) Then
						For intCtr2 = 0 To UBound(aryRecurDates)
							blnException = False
							If rstExc.RecordCount > 0 Then
								rstExc.MoveFirst()
								For intCtr3 = 1 To rstExc.RecordCount
									'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If CDate(rstExc.Fields("fldApptDate").Value) = CDate(aryRecurDates(intCtr2)) Then
										If rstExc.Fields("fldRecurApptID").Value = rstAppt.Fields("fldApptID").Value Then
											blnException = True
											Exit For
										End If
									End If
									If Not rstExc.EOF Then
										rstExc.MoveNext()
									End If
								Next intCtr3
							End If
							
						Next intCtr2
					Else
						blnException = True
					End If
				End If
				
				If blnException = False Then
					'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(0, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryAppts(0, intApptCount) = rstAppt.Fields("fldPatApptID").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(1, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryAppts(1, intApptCount) = rstAppt.Fields("fldPatientName").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(2, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryAppts(2, intApptCount) = rstAppt.Fields("fldProviderName").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(3, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryAppts(3, intApptCount) = rstAppt.Fields("fldStartDateTime").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(4, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryAppts(4, intApptCount) = rstAppt.Fields("fldDescription").Value
					
					intApptCount = intApptCount + 1
				End If
				
				rstAppt.MoveNext()
			Next intCtr
		Else
			ReDim aryAppts(0, 0)
		End If
		
		'Shrink the array, if needed
		If IsArray(aryAppts) Then
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNothing(aryAppts(0, 0)) Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object FetchByCheckInDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FetchByCheckInDate = System.DBNull.Value
			Else
				ReDim Preserve aryAppts(4, intApptCount - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object FetchByCheckInDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FetchByCheckInDate = VB6.CopyArray(aryAppts)
			End If
		Else
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FetchByCheckInDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FetchByCheckInDate = System.DBNull.Value
		End If
		
		' Clean House
		'UPGRADE_NOTE: Object rstAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAppt = Nothing
		'UPGRADE_NOTE: Object rstExc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstExc = Nothing
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Function
		
ErrTrap: 
		'Clean up
		'UPGRADE_NOTE: Object rstAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAppt = Nothing
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
		
	End Function
	
	Public Function FetchPatientApptByID(ByVal lngApptID As Integer) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		'Instantiate the Appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchPatientApptByID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchPatientApptByID = objAppt.FetchPatientApptByID(lngApptID)
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	
	Public Function FetchBlockApptByID(ByVal lngApptID As Integer) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		'Instantiate the Appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchBlockApptByID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchBlockApptByID = objAppt.FetchBlockApptByID(lngApptID)
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	
	Public Function FetchByApptIDs(ByVal strApptIDs As String) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		'Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchByApptIDs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchByApptIDs = objAppt.FetchByApptIDs(strApptIDs)
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
		
	End Function
	
	Public Function Insert(ByVal intRecurType As typApptType, ByVal lngProviderID As Integer, ByVal lngClinicID As Integer, ByVal lngCategoryID As Integer, ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, ByVal lngDuration As Integer, ByVal strCPTCode As String, ByVal strDescription As String, ByVal strNote As String, ByVal strUserName As String, Optional ByVal varPatientArray As Object = Nothing, Optional ByVal strRecurPattern As String = "", Optional ByVal lngRecurInterval As Integer = 0, Optional ByVal lngRecurDOWMask As Integer = 0, Optional ByVal lngRecurDOM As Integer = 0, Optional ByVal lngRecurWOM As Integer = 0, Optional ByVal lngRecurMOY As Integer = 0) As Integer
		Dim ApptDB As Object
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
		Dim lngApptID As Integer
		Dim lngRecurApptID As Integer
		Dim intCtr As Short
		Dim aryDates As Object
		Dim aryPatients As Object
		
		On Error GoTo ErrTrap
		
		objAppt = CreateObject("ApptDB.CApptDB")
		
		Select Case intRecurType
			Case 1 ' Single instance i.e. one-time, non-recurring appointment
				'Insert Core Appt Info
				'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.InsertSingle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngApptID = objAppt.InsertSingle(lngProviderID, lngClinicID, lngCategoryID, dteStartDateTime, dteEndDateTime, lngDuration, strCPTCode, strDescription, strNote, strUserName)
				
				If lngApptID < 1 Then
					strErrMsg = "An error occured when inserting the core appointment information."
					GoTo ErrTrap
				End If
				
				Insert = lngApptID
			Case 2 'Recurring appointment
				'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.InsertRecurAppt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngApptID = objAppt.InsertRecurAppt(lngProviderID, lngClinicID, lngCategoryID, dteStartDateTime, dteEndDateTime, lngDuration, strCPTCode, strDescription, strNote, strRecurPattern, lngRecurInterval, lngRecurDOWMask, lngRecurDOM, lngRecurWOM, lngRecurMOY, strUserName)
				
				Insert = lngApptID
				
				If lngApptID < 1 Then
					strErrMsg = "An error occured when inserting the recurring appointment information."
					GoTo ErrTrap
				End If
		End Select
		
		'If this is a patient appointment, insert rows into tblPatientAppt
		If lngCategoryID = APPT_TYPE_PATIENT Then
			'Insert rows into tblPatientAppt
			objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
			For intCtr = 0 To UBound(varPatientArray, 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(intCtr, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If varPatientArray(intCtr, 1) < 1 Then
					strErrMsg = "An error occured when inserting the patient appointment information(Missing Patient ID)."
					GoTo ErrTrap
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call objPatAppt.Insert(lngApptID, varPatientArray(intCtr, 1))
				End If
			Next 
			'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objPatAppt = Nothing
		End If
		
		'Clean up
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Function
		
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.
		System.EnterpriseServices.ContextUtil.SetAbort()
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMsg)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
		
	End Function
	
	Public Function Update(ByVal lngApptID As Integer, ByVal intRecurType As typApptType, ByVal lngProviderID As Integer, ByVal lngClinicID As Integer, ByVal lngCategoryID As Integer, ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, ByVal lngDuration As Integer, ByVal strCPTCode As String, ByVal strDescription As String, ByVal strNote As String, ByVal strUserName As String, Optional ByVal varPatientArray As Object = Nothing, Optional ByVal strRecurPattern As String = "", Optional ByVal lngRecurInterval As Integer = 0, Optional ByVal lngRecurDOWMask As Integer = 0, Optional ByVal lngRecurDOM As Integer = 0, Optional ByVal lngRecurWOM As Integer = 0, Optional ByVal lngRecurMOY As Integer = 0, Optional ByRef blnRecurInstance As Boolean = False) As Integer
		Dim ApptDB As Object
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
		Dim intCtr As Short
		Dim lngNewApptID As Integer
		
		On Error GoTo ErrTrap
		
		objAppt = CreateObject("ApptDB.CApptDB")
		
		'Check if this is an update to an instance of a recurring appointment
		If blnRecurInstance = True Then
			'Create a new one-time patient appointment to replace this instance.
			'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.InsertSingle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngNewApptID = objAppt.InsertSingle(lngProviderID, lngClinicID, lngCategoryID, dteStartDateTime, dteEndDateTime, lngDuration, strCPTCode, strDescription, strNote, strUserName)
			
			'If this is a patient appointment, insert a new patient appointment record for each
			'patient in the appointment, replacing the appropriate element in the Patient Array
			'with the new PatApptID.
			If lngCategoryID = APPT_TYPE_PATIENT Then
				'Insert rows into tblPatientAppt
				objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
				For intCtr = 0 To UBound(varPatientArray, 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varPatientArray(intCtr, 0) = objPatAppt.Insert(lngNewApptID, varPatientArray(intCtr, 1))
				Next 
				'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objPatAppt = Nothing
			End If
			
			'Insert a row into tblRecurApptExc so the recurring appointment and the new one-time
			'appointment do not appear as conflicts.
			'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.InsertRecurApptExc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call objAppt.InsertRecurApptExc(lngApptID, dteStartDateTime, strUserName)
			
			'Update the NEW appointment.  This effectively becomes a recursive call to the Update() method
			'The only difference is that the PatientAppt element of the Patient array must now contain the
			'new Patient Appointment IDs.
			Call Update(lngNewApptID, intRecurType, lngProviderID, lngClinicID, lngCategoryID, dteStartDateTime, dteEndDateTime, lngDuration, strCPTCode, strDescription, strNote, strUserName, varPatientArray, "", 0, 0, 0, 0, 0, False)
			
			GoTo CLEANUP 'We are done.
		End If
		
		Select Case intRecurType
			Case 1 ' Single instance i.e. one-time
				'Update Core Appt Info
				'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call objAppt.Update(lngApptID, lngProviderID, lngClinicID, lngCategoryID, dteStartDateTime, dteEndDateTime, lngDuration, strCPTCode, strDescription, strNote, strUserName)
				
			Case 2 'Recurring
				'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.UpdateRecurAppt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call objAppt.UpdateRecurAppt(lngApptID, dteStartDateTime, dteEndDateTime, strCPTCode, lngDuration, lngProviderID, lngClinicID, lngCategoryID, strRecurPattern, lngRecurInterval, lngRecurDOWMask, lngRecurDOM, lngRecurWOM, lngRecurMOY, strDescription, strNote, strUserName)
				
				If lngApptID < 1 Then
					strErrMsg = "An error occured when updating the recurring appointment information."
					GoTo ErrTrap
				End If
		End Select
		
		If lngCategoryID = APPT_TYPE_PATIENT Then
			'Update rows in tblPatientAppt
			objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
			For intCtr = 0 To UBound(varPatientArray, 1)
				'Do not update patient appointments already Billed (in the case of Group Appointments)
				'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(intCtr, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If varPatientArray(intCtr, 2) <> "Y" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(intCtr, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If varPatientArray(intCtr, 2) <= 2 Or varPatientArray(intCtr, 2) = APPT_STATUS_CANCELLED Or varPatientArray(intCtr, 2) = APPT_STATUS_PENDING Then
						'If Appt Status is 'Scheduled', 'Confirmed' or 'Cancelled', only a minimal update is required
						'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(intCtr, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call objPatAppt.ChangeStatus(varPatientArray(intCtr, 0), varPatientArray(intCtr, 2), "N", 0, "", strUserName)
					End If
				End If
			Next 
			'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objPatAppt = Nothing
		End If
		
		
CLEANUP: 
		'Return ApptID to calling routine
		If lngNewApptID = 0 Then
			Update = lngApptID
		Else
			Update = lngNewApptID
		End If
		
		'Clean up
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Function
		
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.
		System.EnterpriseServices.ContextUtil.SetAbort()
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMsg)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
		
	End Function
	
	Public Function CloneInstance(ByVal lngApptID As Integer, ByVal dteStartDateTime As Date, ByVal strUserName As String) As Integer
		Dim ApptDB As Object
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
		Dim lngNewApptID As Integer
		Dim lngPatApptID As Integer
		Dim strErrMsg As String
		Dim intCtr As Short
		Dim lngTempPatientID As Integer
		
		On Error GoTo ErrTrap
		
		'Fetch appointment information
		rstAppt = FetchPatientApptByID(lngApptID)
		If rstAppt.RecordCount = 0 Then
			strErrMsg = "Appointment ID not found."
			GoTo ErrTrap
		End If
		'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dteStartDateTime = CDate(DateValue(CStr(dteStartDateTime)) & " " & TimeValue(rstAppt.Fields("fldStartDateTime").Value))
		dteEndDateTime = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, rstAppt.Fields("fldDuration").Value, dteStartDateTime)
		
		objAppt = CreateObject("ApptDB.CApptDB")
		
		'Insert single instance of appointment, copying values from original.
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.InsertSingle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngNewApptID = objAppt.InsertSingle(rstAppt.Fields("fldProviderID").Value, rstAppt.Fields("fldClinicID").Value, rstAppt.Fields("fldCategoryID").Value, dteStartDateTime, dteEndDateTime, rstAppt.Fields("fldDuration").Value, IfNull(rstAppt.Fields("fldCPTCode").Value, ""), IfNull(rstAppt.Fields("fldDescription").Value, ""), IfNull(rstAppt.Fields("fldNote").Value, ""), strUserName)
		
		'If this is a patient appointment, insert a new patient appointment record for each
		'patient in the appointment, replacing the appropriate element in the Patient Array
		'with the new PatApptID.
		If rstAppt.Fields("fldCategoryID").Value = APPT_TYPE_PATIENT Then
			'Insert rows into tblPatientAppt.  The recordset will return multiple rows per patient if
			'the patient has more than one plan.  Ensure only one patient appointment record is inserted
			'per cloned appointment.
			objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
			lngTempPatientID = -1
			For intCtr = 1 To rstAppt.RecordCount
				If lngTempPatientID <> rstAppt.Fields("fldPatientID").Value Then
					lngPatApptID = objPatAppt.Insert(lngNewApptID, rstAppt.Fields("fldPatientID").Value)
					lngTempPatientID = rstAppt.Fields("fldPatientID").Value
				End If
				rstAppt.MoveNext()
			Next 
			'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objPatAppt = Nothing
		End If
		
		'UPGRADE_NOTE: Object rstAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAppt = Nothing
		
		'Insert a row into tblRecurApptExc so the recurring appointment and the new one-time
		'appointment do not appear as conflicts.
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.InsertRecurApptExc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call objAppt.InsertRecurApptExc(lngApptID, dteStartDateTime, strUserName)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		'Return new appointment ID
		CloneInstance = lngNewApptID
		
		Exit Function
		
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.
		System.EnterpriseServices.ContextUtil.SetAbort()
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		'UPGRADE_NOTE: Object rstAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAppt = Nothing
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMsg)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
		
	End Function
	
	Public Function ConflictExists(ByVal lngProviderID As Integer, ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, ByVal lngApptID As Integer, Optional ByRef varConflicts As Object = Nothing) As Boolean
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		ConflictExists = False ' Assume no conflicts
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchConflicts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rst = objAppt.FetchConflicts(lngProviderID, dteStartDateTime, dteEndDateTime, lngApptID)
		
		If rst.RecordCount > 0 Then
			ConflictExists = True
		Else
			With rst
				Do While Not .EOF
					If .Fields("fldRecurPattern").Value > "" Then
						ConflictExists = True
						Exit Do
					End If
					.MoveNext()
				Loop 
			End With
		End If
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If Not IsNothing(varConflicts) Then
			varConflicts = rst
		End If
		
		'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rst = Nothing
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rst = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
	End Function
	
	Public Function ConflictExistsRec(ByVal lngProviderID As Integer, ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, ByVal lngApptID As Integer) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchConflicts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ConflictExistsRec = objAppt.FetchConflicts(lngProviderID, dteStartDateTime, dteEndDateTime, lngApptID)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
	End Function
	
	Public Sub DeleteNonRecurring(ByVal lngApptID As Integer, ByVal strUserName As String)
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		'Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.DeleteNonRecurring. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call objAppt.DeleteNonRecurring(lngApptID, strUserName)
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Sub
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
	End Sub
	
	Public Sub DeleteRecurSeries(ByVal lngApptID As Integer, ByVal strUserName As String)
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.DeleteRecurSeries. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call objAppt.DeleteRecurSeries(lngApptID, strUserName)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Sub
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
	End Sub
	
	
	Public Sub DeleteRecurSingle(ByVal lngApptID As Integer, ByVal dtApptDate As Date, ByVal strUserName As String)
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.InsertRecurApptExc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call objAppt.InsertRecurApptExc(lngApptID, dtApptDate, strUserName)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Sub
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
	End Sub
	
	
	Public Function FetchProviderExceptions(ByVal lngProviderID As Integer, ByVal dtStartDateTime As Date, ByVal dtEndDateTime As Date) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchProviderExceptions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchProviderExceptions = objAppt.FetchProviderExceptions(lngProviderID, dtStartDateTime, dtEndDateTime)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	
	Public Function FetchECDetail(ByVal lngApptID As Integer, ByVal lngPatientID As Integer) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchECDetail. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchECDetail = objAppt.FetchECDetail(lngApptID, lngPatientID)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
	End Function
	Public Function FetchEncByDOS(ByVal lngProviderID As Integer, ByVal lngApptID As Integer, ByVal dteDOS As Date) As ADODB.Recordset
		Dim ApptDB As Object
		Dim objAppt As ApptDB.CApptDB
		
		On Error GoTo ErrTrap
		
		objAppt = CreateObject("ApptDB.CApptDB")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchEncByDOS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchEncByDOS = objAppt.FetchEncByDOS(lngProviderID, lngApptID, dteDOS)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		
	End Function
	Public Function GetRecurApptDates(ByVal dteSeekStartDate As Date, ByVal dteSeekEndDate As Date, ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, ByVal lngDuration As Integer, ByVal strRecurPattern As String, ByVal lngInterval As Integer, ByVal lngDOWMask As Integer, ByVal lngDOM As Integer, ByVal lngWOM As Integer, ByVal lngMOY As Integer, Optional ByRef blnReturnString As Boolean = False) As Object
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
		Dim i As Integer
		Dim lngNDOM As Integer
		Dim lngOffset As Integer
		Dim varAry As Object
		
		On Error GoTo Err_Trap
		
		Select Case strRecurPattern
			Case "D" 'Daily Recurring Appointment'
				If lngDOWMask > 0 Then
					' Recur Every Weekday   '
					' OR                    '
					' Recur Every Weekend   '
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					dteTemp = DateValue(CStr(dteStartDateTime))
					' Look for dates while currrent calculated date is  '
					' prior to recur. appt end date and prior to seek   '
					' end date                                          '
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					While dteTemp <= DateValue(CStr(dteSeekEndDate)) And dteTemp <= DateValue(CStr(dteEndDateTime))
						'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If IsANDed(2 ^ (DatePart(Microsoft.VisualBasic.DateInterval.WeekDay, dteTemp) - 1), lngDOWMask) And dteTemp >= DateValue(CStr(dteSeekStartDate)) Then
							ConcatString(CStr(dteTemp), strDates)
						End If
						dteTemp = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, dteTemp)
					End While
				Else
					' Recur Every X Day(s)  '
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					dteTemp = DateValue(CStr(dteStartDateTime))
					' Look for dates while currrent calculated date is  '
					' prior to recur. appt end date and prior to seek   '
					' end date                                          '
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					While dteTemp <= DateValue(CStr(dteSeekEndDate)) And dteTemp <= DateValue(CStr(dteEndDateTime))
						'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If dteTemp >= DateValue(CStr(dteSeekStartDate)) Then
							' applicable date found....log it!  '
							ConcatString(CStr(dteTemp), strDates)
						End If
						
						dteTemp = DateAdd(Microsoft.VisualBasic.DateInterval.Day, lngInterval, dteTemp)
					End While
				End If
				
			Case "W" ' -= Weekly Recurring Appointment =-'
				For i = FirstDayOfWeek.Sunday To FirstDayOfWeek.Saturday
					' If current vbDay is in the mask then...
					If IsANDed(2 ^ (i - 1), lngDOWMask) Then
						' Determine the first date falling on the current vbDay
						lngOffset = i - DatePart(Microsoft.VisualBasic.DateInterval.WeekDay, dteStartDateTime)
						If lngOffset < 0 Then
							lngOffset = lngOffset + 7
						End If
						
						' Now we do the dirty work and begin finding valid dates
						'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						dteTemp = DateValue(CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, lngOffset, dteStartDateTime)))
						'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						While dteTemp <= DateValue(CStr(dteSeekEndDate)) And dteTemp <= DateValue(CStr(dteEndDateTime))
							'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							If dteTemp >= DateValue(CStr(dteSeekStartDate)) Then
								ConcatString(CStr(dteTemp), strDates)
							End If
							dteTemp = DateAdd(Microsoft.VisualBasic.DateInterval.WeekOfYear, lngInterval, dteTemp)
						End While
					End If
				Next 
				
			Case "M" ' -= Monthly Recurring Appointment =-'
				If lngDOWMask Then
					'The [1st, 2nd, 3rd, 4th, Last] [Sun, Mon, Tue, Wed, Thu, Fri, Sat] of every Y month(s)
					dteMonth = CDate(DatePart(Microsoft.VisualBasic.DateInterval.Month, dteStartDateTime) & "/1/" & DatePart(Microsoft.VisualBasic.DateInterval.Year, dteStartDateTime))
					lngNDOM = ((System.Math.Log(lngDOWMask) / System.Math.Log(2))) + 1
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					dteTemp = DateValue(CStr(GetXDayOfMonth(dteMonth, lngNDOM, lngDOM)))
					
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					While dteTemp <= DateValue(CStr(dteSeekEndDate)) And dteTemp <= DateValue(CStr(dteEndDateTime))
						'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If dteTemp >= DateValue(CStr(dteStartDateTime)) And dteTemp >= DateValue(CStr(dteSeekStartDate)) Then
							ConcatString(CStr(dteTemp), strDates)
						End If
						dteMonth = DateAdd(Microsoft.VisualBasic.DateInterval.Month, lngInterval, dteMonth)
						dteTemp = GetXDayOfMonth(dteMonth, lngNDOM, lngDOM)
					End While
				Else
					' Day X of Every Y month(s)                                 '
					' OR                                                        '
					' Recur Every [1st, 2nd, 3rd, 4th] Day of Every Y month(s)  '
					' OR                                                        '
					' Recur Every LAST Day of Every Y month(s)                  '
					dteMonth = CDate(DatePart(Microsoft.VisualBasic.DateInterval.Month, dteStartDateTime) & "/1/" & DatePart(Microsoft.VisualBasic.DateInterval.Year, dteStartDateTime))
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					dteTemp = DateValue(CStr(GetDayOfMonth(dteMonth, lngDOM)))
					
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					While dteTemp <= DateValue(CStr(dteSeekEndDate)) And dteTemp <= DateValue(CStr(dteEndDateTime))
						'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If dteTemp >= DateValue(CStr(dteStartDateTime)) And dteTemp >= DateValue(CStr(dteSeekStartDate)) Then
							ConcatString(CStr(dteTemp), strDates)
						End If
						dteMonth = DateAdd(Microsoft.VisualBasic.DateInterval.Month, lngInterval, dteMonth)
						dteTemp = GetDayOfMonth(dteMonth, lngDOM)
					End While
				End If
			Case "Y" ' -= Yearly Recurring Appointment =-'
				If lngDOWMask Then
					' [1st, 2nd, 3rd, 4th, Last] [Sun, Mon, Tue, Wed, Thu, Fri, Sat] of '
					' [Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec]      '
					dteMonth = CDate(DatePart(Microsoft.VisualBasic.DateInterval.Month, dteStartDateTime) & "/1/" & DatePart(Microsoft.VisualBasic.DateInterval.Year, dteStartDateTime))
					lngNDOM = ((System.Math.Log(lngDOWMask) / System.Math.Log(2))) + 1
					dteTemp = GetXDayOfMonth(dteMonth, lngNDOM, lngDOM)
					
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					While dteTemp <= DateValue(CStr(dteSeekEndDate)) And dteTemp <= DateValue(CStr(dteEndDateTime))
						'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If dteTemp >= DateValue(CStr(dteStartDateTime)) And dteTemp >= DateValue(CStr(dteSeekStartDate)) Then
							ConcatString(CStr(dteTemp), strDates)
						End If
						dteMonth = DateAdd(Microsoft.VisualBasic.DateInterval.Year, lngInterval, dteMonth)
						dteTemp = GetXDayOfMonth(dteMonth, lngNDOM, lngDOM)
					End While
				Else
					' Every [1st, 2nd, 3rd, 4th, Last] Day                          '
					' [Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec]  '
					dteMonth = CDate(DatePart(Microsoft.VisualBasic.DateInterval.Month, dteStartDateTime) & "/1/" & DatePart(Microsoft.VisualBasic.DateInterval.Year, dteStartDateTime))
					dteTemp = GetDayOfMonth(dteMonth, lngDOM)
					
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					While dteTemp <= DateValue(CStr(dteSeekEndDate)) And dteTemp <= DateValue(CStr(dteEndDateTime))
						'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If dteTemp >= DateValue(CStr(dteStartDateTime)) And dteTemp >= DateValue(CStr(dteSeekStartDate)) Then
							ConcatString(CStr(dteTemp), strDates)
						End If
						dteMonth = DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, dteMonth)
						dteTemp = GetDayOfMonth(dteMonth, lngDOM)
					End While
				End If
				
		End Select
		
		If Trim(strDates) > "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object varAry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varAry = Split(strDates, ",")
			SortDates(varAry)
			
			If blnReturnString Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varAry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetRecurApptDates = Join(varAry, ",")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object varAry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object GetRecurApptDates. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetRecurApptDates = varAry
			End If
		Else
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetRecurApptDates. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetRecurApptDates = System.DBNull.Value
		End If
		
		Exit Function
		
Err_Trap: 
		Err.Raise(Err.Number, CLASS_NAME & ":GetRecurApptDates()", Err.Description)
	End Function
	
	Public Function FetchFuturePatientApptByProvider(ByVal lngPatientID As Integer, ByVal lngProviderID As Integer) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchFuturePatientApptByProvider. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchFuturePatientApptByProvider = objAppt.FetchFuturePatientApptByProvider(lngPatientID, lngProviderID)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	Public Function FetchFutureRecurPatientApptByProvider(ByVal lngPatientID As Integer, ByVal lngProviderID As Integer) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		'Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchFutureRecurPatientApptByProvider. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rst = objAppt.FetchFutureRecurPatientApptByProvider(lngPatientID, lngProviderID)
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		FetchFutureRecurPatientApptByProvider = rst
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
		
	End Function
	
	
	Public Function FetchFuturePatientApptByManager(ByVal lngPatientID As Integer, ByVal lngUserID As Integer) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchFuturePatientApptByManager. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchFuturePatientApptByManager = objAppt.FetchFuturePatientApptByManager(lngPatientID, lngUserID)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	
	Public Function FetchFutureRecurPatientApptByManager(ByVal lngPatientID As Integer, ByVal lngUserID As Integer) As ADODB.Recordset
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		'Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchFutureRecurPatientApptByManager. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rst = objAppt.FetchFutureRecurPatientApptByManager(lngPatientID, lngUserID)
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		FetchFutureRecurPatientApptByManager = rst
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
		
	End Function
	
	
	Public Function GetConflicts(ByVal lngApptID As Integer, ByVal lngProviderID As Integer, ByVal varDates As Object, ByVal intDuration As Short) As Object
		Dim ApptDB As Object
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
		
		On Error GoTo ErrTrap
		
		Dim intCtr1 As Short
		Dim intCtr2 As Short
		Dim intCtr3 As Short
		Dim objAppt As ApptDB.CApptDB
		Dim rstAppt As ADODB.Recordset
		Dim intNumConflicts As Short
		Dim varConflict() As Object
		Dim varTemp As Object
		
		'Defensive programming
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(varDates) Then
			Erase varConflict
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetConflicts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetConflicts = System.DBNull.Value
			Exit Function
		End If
		
		intNumConflicts = -1
		objAppt = CreateObject("ApptDB.CApptDB")
		For intCtr1 = 0 To UBound(varDates)
			'Check for conflicts, consider recurring appointments
			'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchConflictsByProvider. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rstAppt = objAppt.FetchConflictsByProvider(lngProviderID, varDates(intCtr1), intDuration)
			For intCtr2 = 1 To rstAppt.RecordCount
				If rstAppt.Fields("fldRecurYN").Value = "N" Then
					'Exclude patient appts with patient count of zero - these appts have been cancelled
					If rstAppt.Fields("fldCategoryID").Value = 1 And rstAppt.Fields("PatientCount").Value > 0 Then
						If rstAppt.Fields("fldApptID").Value <> lngApptID Then 'The appot can not conflict with itself.
							'The appointment is conflicting
							intNumConflicts = intNumConflicts + 1
							ReDim Preserve varConflict(3, intNumConflicts)
							'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(0, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							varConflict(0, intNumConflicts) = rstAppt.Fields("fldApptID").Value
							'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(1, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							varConflict(1, intNumConflicts) = rstAppt.Fields("fldStartDateTime").Value
							'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(2, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							varConflict(2, intNumConflicts) = rstAppt.Fields("fldDuration").Value
							If rstAppt.Fields("fldCategoryID").Value = 1 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(3, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								varConflict(3, intNumConflicts) = rstAppt.Fields("PatientName").Value
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(3, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								varConflict(3, intNumConflicts) = rstAppt.Fields("fldDescription").Value
							End If
						End If
					End If
				Else
					'The appointment may be conflicting.  Check the recurrance to determine if actually a conflict.
					'UPGRADE_WARNING: Couldn't resolve default property of object GetRecurApptDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varTemp = GetRecurApptDates(rstAppt.Fields("fldStartDateTime").Value, rstAppt.Fields("fldEndDateTime").Value, rstAppt.Fields("fldStartDateTime").Value, rstAppt.Fields("fldEndDateTime").Value, intDuration, rstAppt.Fields("fldRecurPattern").Value, rstAppt.Fields("fldInterval").Value, rstAppt.Fields("fldDOWMask").Value, rstAppt.Fields("fldDOM").Value, rstAppt.Fields("fldWOM").Value, rstAppt.Fields("fldMOY").Value)
					If IsArray(varTemp) Then
						For intCtr3 = 0 To UBound(varTemp)
							'UPGRADE_WARNING: Couldn't resolve default property of object varDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							'UPGRADE_WARNING: Couldn't resolve default property of object varTemp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If DateValue(varTemp(intCtr3)) = DateValue(varDates(intCtr1)) Then
								'The appointment is conflicting
								intNumConflicts = intNumConflicts + 1
								ReDim Preserve varConflict(3, intNumConflicts)
								'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(0, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								varConflict(0, intNumConflicts) = rstAppt.Fields("fldApptID").Value
								'UPGRADE_WARNING: Couldn't resolve default property of object varTemp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(1, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								varConflict(1, intNumConflicts) = varTemp(intCtr3) & " " & TimeValue(rstAppt.Fields("fldStartDateTime").Value)
								'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(2, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								varConflict(2, intNumConflicts) = rstAppt.Fields("fldDuration").Value
								If rstAppt.Fields("fldCategoryID").Value = 1 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(3, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									varConflict(3, intNumConflicts) = rstAppt.Fields("PatientName").Value
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object varConflict(3, intNumConflicts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									varConflict(3, intNumConflicts) = rstAppt.Fields("fldDescription").Value
								End If
							End If
						Next intCtr3
						
						Erase varTemp
					End If
				End If
				
				rstAppt.MoveNext()
			Next intCtr2
			'UPGRADE_NOTE: Object rstAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstAppt = Nothing
		Next intCtr1
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		If intNumConflicts < 0 Then
			Erase varConflict
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetConflicts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetConflicts = System.DBNull.Value
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetConflicts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetConflicts = VB6.CopyArray(varConflict)
		End If
		
		Exit Function
		
ErrTrap: 
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
		
	End Function
	
	Public Sub FetchNextOpenProviderTimeSlot(ByVal lngProviderID As Integer, ByVal lngLength As Integer, ByVal dteSearch As Date, ByRef varStartTime As Object, ByRef varEndTime As Object)
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
		Dim strStartTimes, strEndTimes As String
		Dim aryStartTimes, aryEndTimes As Object
		Dim dteStartTime, dteEndTime As Date
		Dim dteStartSearch As Date
		Dim dteEndDummy As Date
		Dim varTimer As Object
		
		Dim aryRet As Object
		
		Const MIN_TIME As String = " 07:00 AM"
		
		On Error GoTo Error_Handler
		
		dteStartTime = System.Date.FromOADate(-1)
		dteEndTime = System.Date.FromOADate(-1)
		'UPGRADE_WARNING: Couldn't resolve default property of object varTimer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varTimer = Now
		
TRY_AGAIN: 
		'UPGRADE_WARNING: Couldn't resolve default property of object varTimer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If DateDiff(Microsoft.VisualBasic.DateInterval.Second, varTimer, Now) > 6 Then
			'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			dteEndTime = DateValue(CStr(dteSearch))
		Else
			If DatePart(Microsoft.VisualBasic.DateInterval.Hour, dteSearch) < 7 Then
				'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				dteStartSearch = CDate(DateValue(CStr(dteSearch)) & MIN_TIME)
			Else
				dteStartSearch = dteSearch
			End If
			
			' Retrieve array
			'UPGRADE_WARNING: Couldn't resolve default property of object GetClosedTimeSlotArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object aryRet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aryRet = GetClosedTimeSlotArray(lngProviderID, dteSearch)
			
			' Sort array(s)
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(aryRet) Then
				SortClosedDates(aryRet(0), aryRet(1))
			End If
			' Now analyze array
			AnalyzeArray(aryRet, dteStartSearch, lngLength, dteStartTime, dteEndTime)
			
			If dteStartTime < System.Date.FromOADate(0) Then
				'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				dteSearch = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(DateValue(CStr(dteSearch))))
				GoTo TRY_AGAIN
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varStartTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varStartTime = dteStartTime
		'UPGRADE_WARNING: Couldn't resolve default property of object varEndTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varEndTime = dteEndTime
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Sub
		
Error_Handler: 
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Sub
	
	'--------------------------------------------------------------------
	' PrivateMethods    +++++++++++++++++++++++++++++++++++++++++++++++++
	'--------------------------------------------------------------------
	
	Private Function AnalyzeArray(ByVal varDates As Object, ByVal dteStart As Date, ByVal lngLength As Integer, ByRef dteStartTime As Date, ByRef dteEndTime As Date) As Object
		
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
		Const MIN_TIME As String = " 7:00 AM"
		Const MAX_TIME As String = " 10:00 PM"
		
		Dim intCnt As Short
		Dim dteCurrentTime As Date
		Dim dteMaxTime As Date
		Dim aryStart, aryEnd As Object
		Dim dteEOD As Date
		
		'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dteEOD = CDate(DateValue(CStr(dteStart)) & MAX_TIME)
		If dteStart >= dteEOD Then Exit Function
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(varDates) Then
			' A null values represents an open day
			'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			dteStartTime = CDate(DateValue(CStr(dteStart)) & MIN_TIME)
			'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			dteEndTime = CDate(DateValue(CStr(dteStart)) & MAX_TIME)
		Else
			
			'UPGRADE_WARNING: Couldn't resolve default property of object varDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object aryStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aryStart = varDates(0)
			'UPGRADE_WARNING: Couldn't resolve default property of object varDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aryEnd = varDates(1)
			'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			dteMaxTime = CDate(DateValue(CStr(dteStart)) & MAX_TIME)
			
			For intCnt = 0 To UBound(aryStart)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(intCnt). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If aryStart(intCnt) >= dteStart Then
					'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
					If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, dteStart, aryStart(intCnt)) >= lngLength Then
						' We got a winner, report it and move on
						dteStartTime = dteStart
						'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						dteEndTime = CDate(aryStart(intCnt))
						Exit For
					Else
						' Go Fish
						'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						dteStart = MaxDate(CDate(aryEnd(intCnt)), dteStart)
					End If
				Else
					' Go Fish
					'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					dteStart = MaxDate(CDate(aryEnd(intCnt)), dteStart)
				End If
			Next 
		End If
		
		' Final check
		If dteStartTime < System.Date.FromOADate(0) Then
			If dteEOD >= dteStart Then
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, dteStart, dteEOD) >= lngLength Then
					' We got a winner, report it and move on
					dteStartTime = dteStart
					dteEndTime = dteEOD
				End If
			End If
		End If
		
	End Function
	
	
	Private Function MaxDate(ByVal dteDate1 As Date, ByRef dteDate2 As Date) As Date
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
	
	Private Function GetClosedTimeSlotArray(ByVal lngProviderID As Integer, ByVal dteSearch As Date) As Object
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
		Dim strStartTimes, strEndTimes As String
		Dim aryStartTimes, aryEndTimes As Object
		Dim aryRet(1) As Object
		
		rstDaySchedule = FetchByProviderDateRange(lngProviderID, dteSearch, dteSearch)
		
		With rstDaySchedule
			While Not .EOF
				If .Fields("fldRecurPattern").Value > "" Then
					
					'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(GetRecurApptDates(DateValue(CStr(dteSearch)), DateValue(CStr(dteSearch)), .Fields("fldStartDateTime").Value, .Fields("fldEndDateTime").Value, .Fields("fldDuration").Value, .Fields("fldRecurPattern").Value, .Fields("fldInterval").Value, .Fields("fldDOWmask").Value, .Fields("fldDOM").Value, .Fields("fldWOM").Value, .Fields("fldMOY").Value, True)) Then
						
						If strStartTimes > "" Then
							'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							strStartTimes = strStartTimes & "," & DateValue(CStr(dteSearch)) & " " & TimeValue(.Fields("fldStartDateTime").Value)
							'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							strEndTimes = strEndTimes & "," & DateValue(CStr(dteSearch)) & " " & TimeValue(.Fields("fldEndDateTime").Value)
						Else
							'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							strStartTimes = DateValue(CStr(dteSearch)) & " " & TimeValue(.Fields("fldStartDateTime").Value)
							'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							strEndTimes = DateValue(CStr(dteSearch)) & " " & TimeValue(.Fields("fldEndDateTime").Value)
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
				.MoveNext()
			End While
		End With
		'UPGRADE_NOTE: Object rstDaySchedule may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstDaySchedule = Nothing
		
		If strStartTimes > "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object aryStartTimes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aryStartTimes = Split(strStartTimes, ",")
			'UPGRADE_WARNING: Couldn't resolve default property of object aryEndTimes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aryEndTimes = Split(strEndTimes, ",")
			'UPGRADE_WARNING: Couldn't resolve default property of object aryStartTimes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object aryRet(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aryRet(0) = aryStartTimes
			'UPGRADE_WARNING: Couldn't resolve default property of object aryEndTimes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object aryRet(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aryRet(1) = aryEndTimes
			'UPGRADE_WARNING: Couldn't resolve default property of object GetClosedTimeSlotArray. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetClosedTimeSlotArray = VB6.CopyArray(aryRet)
		Else
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetClosedTimeSlotArray. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetClosedTimeSlotArray = System.DBNull.Value
		End If
		
		Exit Function
		
Error_Handler: 
		'UPGRADE_NOTE: Object rstDaySchedule may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstDaySchedule = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	Private Sub SortClosedDates(ByRef aryStart As Object, ByRef aryEnd As Object)
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
		
		Dim i As Integer
		Dim j As Integer
		
		For j = (UBound(aryStart) - 1) To 0 Step -1
			
			For i = 0 To j
				' Sort by Start Dates
				'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CDate(aryStart(i)) > CDate(aryStart(i + 1)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					dteTemp = aryStart(i)
					'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryStart(i) = aryStart(i + 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryStart(i + 1) = dteTemp
					'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					dteTemp = aryEnd(i)
					'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryEnd(i) = aryEnd(i + 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryEnd(i + 1) = dteTemp
					
					'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf CDate(aryStart(i)) = CDate(aryStart(i + 1)) Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CDate(aryEnd(i)) > CDate(aryEnd(i + 1)) Then
						' Then sort by End Date
						'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						dteTemp = aryStart(i)
						'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryStart(i) = aryStart(i + 1)
						'UPGRADE_WARNING: Couldn't resolve default property of object aryStart(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryStart(i + 1) = dteTemp
						'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						dteTemp = aryEnd(i)
						'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryEnd(i) = aryEnd(i + 1)
						'UPGRADE_WARNING: Couldn't resolve default property of object aryEnd(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryEnd(i + 1) = dteTemp
					End If
				End If
			Next 
			
		Next 
		
	End Sub
	
	Private Function GetCachedDates(ByVal lngApptID As Integer, ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, ByVal lngDuration As Integer, ByVal strRecurPattern As String, ByVal lngInterval As Integer, ByVal lngDOWMask As Integer, ByVal lngDOM As Integer, ByVal lngWOM As Integer, ByVal lngMOY As Integer, ByRef dicAppts As Scripting.Dictionary) As Object
		'--------------------------------------------------------------------
		'Date: 04/16/2001                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Retrieves calculated dates in a Hit/Miss basis       '
		'Parameters:   Values to be calculated(if required)                 '
		'Returns:      Array of dates                                       '
		'--------------------------------------------------------------------
		
		If Not dicAppts.Exists(CStr(lngApptID)) Then
			dicAppts.Add(CStr(lngApptID), GetRecurApptDates(dteStartDateTime, dteEndDateTime, dteStartDateTime, dteEndDateTime, lngDuration, strRecurPattern, lngInterval, lngDOWMask, lngDOM, lngWOM, lngMOY))
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object dicAppts.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetCachedDates. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCachedDates = dicAppts.Item(CStr(lngApptID))
	End Function
	
	Private Function ValidateSingle(ByVal lngProviderID As Integer, ByVal lngClinicID As Integer, ByVal lngCategoryID As Integer, ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, ByVal lngDuration As Integer, ByVal strNote As String, ByVal strUserName As String, ByVal varPatientArray As Object, ByVal strCPTCode As String, ByRef strErrMessage As Object) As Boolean
		Dim ListBz As Object
		'--------------------------------------------------------------------
		'Date: 08/25/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Validates data before allowing inserts and update                '
		'Parameters:  The values to be checked.                             '
		'Returns:   True if all data is valid, False otherwise              '
		'--------------------------------------------------------------------
		
		Dim intCtr As Short
		
		' Check provider ID
		If lngProviderID < 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strErrMessage = "Invalid Provider ID passed."
			ValidateSingle = False
			Exit Function
		End If
		
		' Check category ID
		If lngCategoryID <= 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strErrMessage = "Invalid Appointment Category ID passed."
			ValidateSingle = False
			Exit Function
		End If
		
		' Check for a valid date range
		If dteStartDateTime >= dteEndDateTime Then
			'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strErrMessage = "Start Date/Time cannot be greater than End Date/Time."
			ValidateSingle = False
			Exit Function
		End If
		
		Dim objCPTCode As ListBz.CCPTCodeBz
		If lngCategoryID = APPT_TYPE_PATIENT Then
			'Check clinic ID
			If (lngClinicID <= 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strErrMessage = "Invalid Clinic ID passed."
				ValidateSingle = False
				Exit Function
			End If
			
			'If a CPT code is given, make sure it is valid
			If Len(Trim(strCPTCode)) > 0 Then
				objCPTCode = CreateObject("ListBz.CCPTCodeBz")
				'UPGRADE_WARNING: Couldn't resolve default property of object objCPTCode.Exists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not objCPTCode.Exists(Trim(strCPTCode)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strErrMessage = "Invalid CPT Code passed."
					ValidateSingle = False
					'UPGRADE_NOTE: Object objCPTCode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objCPTCode = Nothing
					Exit Function
				End If
				'UPGRADE_NOTE: Object objCPTCode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objCPTCode = Nothing
			End If
			
			'Check for a valid patient array and array values
			If Not IsAnArray(varPatientArray) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strErrMessage = "No valid Patient ID's were passed."
				ValidateSingle = False
				Exit Function
			Else
				For intCtr = 0 To UBound(varPatientArray, 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(intCtr, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If varPatientArray(intCtr, 1) < 1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strErrMessage = "At least 1 invalid Patient ID was passed."
						ValidateSingle = False
						Exit Function
					End If
				Next 
			End If
		End If
		
		' Check for a user name
		If Trim(strUserName) = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strErrMessage = "User name is required"
			ValidateSingle = False
			Exit Function
		End If
		
		' Check for valid start/end times and duration combination
		' (startime + duration) = endtime
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, dteStartDateTime, dteEndDateTime) <> lngDuration Then
			'UPGRADE_WARNING: Couldn't resolve default property of object strErrMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strErrMessage = "Miscalculated Start/End Time and Duration values passed."
			ValidateSingle = False
			Exit Function
		End If
		
		'If we get here, all is well...
		ValidateSingle = True
		
	End Function
	
	Private Function ValidateRecur(ByVal varPatientArray As Object, ByVal strCPTCode As String, ByVal dteStartDateTime As Date, ByVal dteEndDateTime As Date, ByVal lngDuration As Integer, ByVal lngProviderID As Integer, ByVal lngClinicID As Integer, ByVal lngCategoryID As Integer, ByVal strRecurPattern As String, ByVal lngInterval As Integer, ByVal lngDOWMask As Integer, ByVal lngDOM As Integer, ByVal lngWOM As Integer, ByVal lngMOY As Integer, ByVal strUserName As String, ByRef strErrMessage As String) As Boolean
		Dim ListBz As Object
		
		Dim intCtr As Short
		
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
		
		Dim objCPTCode As ListBz.CCPTCodeBz
		If lngCategoryID = APPT_TYPE_PATIENT Then
			'Check clinic ID
			If lngClinicID < 1 Then
				strErrMessage = "Invalid Clinic ID passed."
				Exit Function
			End If
			
			'If a CPT code is given, make sure it is valid
			If Len(Trim(strCPTCode)) > 0 Then
				objCPTCode = CreateObject("ListBz.CCPTCodeBz")
				'UPGRADE_WARNING: Couldn't resolve default property of object objCPTCode.Exists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not objCPTCode.Exists(Trim(strCPTCode)) Then
					strErrMessage = "Invalid CPT Code passed."
					'UPGRADE_NOTE: Object objCPTCode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objCPTCode = Nothing
					Exit Function
				End If
				'UPGRADE_NOTE: Object objCPTCode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objCPTCode = Nothing
			End If
			
			'Check for a valid patient array and array values
			If Not IsAnArray(varPatientArray) Then
				strErrMessage = "No valid Patient ID's were passed."
				Exit Function
			Else
				For intCtr = 0 To UBound(varPatientArray, 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object varPatientArray(intCtr, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	
	
	Private Sub UpdateSinglePatAppt(ByVal lngApptID As Integer, ByVal varPatApptInfo As Object)
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
		Dim lngID As Integer
		Dim i As Short
		Dim strErrMsg As String
		Dim objPatAppt As ApptBZ.CPatApptBZ
		Dim rst As ADODB.Recordset
		
		On Error GoTo Error_Handler
		
		objPatAppt = CreateObject("ApptBZ.CPatApptBZ")
		
		For i = 0 To UBound(varPatApptInfo, 1)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object varPatApptInfo(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If varPatApptInfo(i, 0) = 0 Then
				' This a new pat/appt rec ....  simple insert
				'UPGRADE_WARNING: Couldn't resolve default property of object varPatApptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngID = objPatAppt.Insert(lngApptID, varPatApptInfo(i, 1))
				
				' Check for DB error
				If lngID < 1 Then
					strErrMsg = "An error occured while trying to add a Patient/Appointment record."
				End If
				
			Else
				' Existing rec with a possible status change
				'UPGRADE_WARNING: Couldn't resolve default property of object varPatApptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rst = objPatAppt.FetchByID(varPatApptInfo(i, 0))
				'UPGRADE_WARNING: Couldn't resolve default property of object varPatApptInfo(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object varPatApptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call objPatAppt.Update(varPatApptInfo(i, 0), lngApptID, varPatApptInfo(i, 1))
				'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rst = Nothing
				
			End If
		Next 
		
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		Exit Sub
Error_Handler: 
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rst = Nothing
		
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMsg)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
	End Sub
	
	Private Sub DeleteBatchPatAppt(ByVal lngApptID As Integer, ByVal varPatApptInfo As Object)
		Dim ApptDB As Object
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
		Dim i As Short
		Dim objPatAppt As ApptDB.CPatApptDB
		Dim strErrMsg As String
		Dim strINClause As Object
		Dim rst As ADODB.Recordset
		
		On Error GoTo Error_Handler
		
		objPatAppt = CreateObject("ApptDB.CPatApptDB")
		
		' Build the IN clause parameter
		' The IN clause represents all patient/appointment records that
		' are to remain associated with the appointment record
		For i = 0 To UBound(varPatApptInfo, 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object varPatApptInfo(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If varPatApptInfo(i, 0) > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object strINClause. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Trim(strINClause) > "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object varPatApptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object strINClause. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strINClause = strINClause & ", " & varPatApptInfo(i, 0)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object varPatApptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object strINClause. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strINClause = varPatApptInfo(i, 0)
				End If
			End If
		Next 
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strINClause. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Trim(strINClause) > "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object objPatAppt.FetchMissingRec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rst = objPatAppt.FetchMissingRec(lngApptID, strINClause)
			
			While Not rst.EOF
				' Delete existin patient/appointment records if they have not
				' been marked as attended
				If rst.Fields("fldApptStatusID").Value <> 3 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object objPatAppt.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call objPatAppt.Delete(lngApptID, rst.Fields("fldPatientID"))
				Else
					strErrMsg = "Cannot modify an Attended apointment."
					GoTo Error_Handler
				End If
				
				rst.MoveNext()
			End While
			'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rst = Nothing
		End If
		
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		Exit Sub
		
Error_Handler: 
		'UPGRADE_NOTE: Object objPatAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPatAppt = Nothing
		'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rst = Nothing
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMsg)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
	End Sub
	
	Private Sub ConcatString(ByVal strAdd As String, ByRef strAddTo As String, Optional ByRef strDel As String = ",")
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
	
	Private Function IsANDed(ByVal lngVal As Integer, ByVal lngMask As Integer) As Boolean
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
	
	Private Sub SortDates(ByRef varArray As Object)
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
		Dim i As Integer
		Dim j As Integer
		
		' Implements Bubble Sort algorithm
		' BAH!  I know it's girly, but it works.
		' I'm still l337
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(varArray) Then
			If IsArray(varArray) Then
				For j = (UBound(varArray) - 1) To 0 Step -1
					
					For i = 0 To j
						'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If CDate(varArray(i)) > CDate(varArray(i + 1)) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							dteTemp = varArray(i)
							'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							varArray(i) = varArray(i + 1)
							'UPGRADE_WARNING: Couldn't resolve default property of object varArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							varArray(i + 1) = dteTemp
						End If
					Next 
					
				Next 
			End If
		End If
		
	End Sub
	
	Private Function DaysInMonth(ByVal dteDate As Date) As Short
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
		DaysInMonth = DateSerial(Year(dteDate), Month(dteDate) + 1, 1).ToOADate - DateSerial(Year(dteDate), Month(dteDate), 1).ToOADate
	End Function
	
	Private Function GetDayOfMonth(ByVal dteDate As Date, ByVal lngDOM As Integer) As Date
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
		Dim lngNumDays As Integer
		Dim lngDOM2Use As Integer
		Dim lngPDOM As Integer
		
		' Do not get an invalid date.  eg. 02/31/2002
		lngNumDays = CInt(DaysInMonth(dteDate))
		lngPDOM = IIf((lngDOM > 0), lngDOM, 31)
		
		lngDOM2Use = IIf((lngPDOM > lngNumDays), lngNumDays, lngPDOM)
		
		GetDayOfMonth = CDate(DatePart(Microsoft.VisualBasic.DateInterval.Month, dteDate) & "/" & lngDOM2Use & "/" & DatePart(Microsoft.VisualBasic.DateInterval.Year, dteDate))
		
	End Function
	
	Private Function GetXDayOfMonth(ByVal dteDate As Date, ByVal lngDay As Integer, ByVal lngPlace As Integer) As Date
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
		Dim lngPPlace As Integer
		Dim dteTemp As Date
		Dim lngOffset As Integer
		
		lngPPlace = IIf((lngPlace > 4), 4, lngPlace)
		
		If lngPPlace > 0 Then
			' 1st, 2nd, 3rd, 4th    '
			lngOffset = lngDay - DatePart(Microsoft.VisualBasic.DateInterval.WeekDay, dteDate)
			If lngOffset < 0 Then lngOffset = lngOffset + 7
			
			dteTemp = DateAdd(Microsoft.VisualBasic.DateInterval.Day, lngOffset, dteDate)
			GetXDayOfMonth = DateAdd(Microsoft.VisualBasic.DateInterval.WeekOfYear, lngPPlace - 1, dteTemp)
		Else
			' Last                  '
			dteTemp = CDate(DatePart(Microsoft.VisualBasic.DateInterval.Month, dteDate) & "/" & DaysInMonth(dteDate) & "/" & DatePart(Microsoft.VisualBasic.DateInterval.Year, dteDate))
			lngOffset = lngDay - DatePart(Microsoft.VisualBasic.DateInterval.WeekDay, dteTemp)
			If lngOffset > 0 Then lngOffset = lngOffset - 7
			GetXDayOfMonth = DateAdd(Microsoft.VisualBasic.DateInterval.Day, lngOffset, dteTemp)
		End If
	End Function
	
	
	Public Function FetchUnBilledAppts(ByVal lngUserID As Integer) As ADODB.Recordset
		Dim ApptDB As Object 'R001
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchUnBilledAppts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchUnBilledAppts = objAppt.FetchUnBilledAppts(lngUserID)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	Public Function FetchPendingAppts(ByVal lngUserID As Integer) As ADODB.Recordset
		Dim ApptDB As Object 'R001
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
		
		On Error GoTo ErrTrap
		
		' Instantiate the appt object
		objAppt = CreateObject("ApptDB.CApptDB")
		
		' Populate the recordset
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.FetchPendingAppts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FetchPendingAppts = objAppt.FetchPendingAppts(lngUserID)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
ErrTrap: 
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	Private Sub SendMessage(ByVal intMsgType As MsgType, ByVal lngProviderID As Integer, ByVal lngPatientID As Integer, ByVal strUserName As String, Optional ByVal dteStartDateTime As Date = #12:00:00 AM#, Optional ByVal intRecurType As Short = 0, Optional ByVal strCancelReason As String = "")
		Dim ClinicBz As Object
		Dim BenefactorBz As Object
		Dim ListBz As Object
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
		
		On Error GoTo ErrTrap
		
		objUser = CreateObject("ClinicBz.CUserBz")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objUser.FetchDetail. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rstUser = objUser.FetchDetail(lngProviderID, blnProvider)
		strProviderName = rstUser.Fields("fldFirstName").Value & " " & rstUser.Fields("fldLastName").Value
		'UPGRADE_NOTE: Object rstUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstUser = Nothing
		
		'Build recordset of Office Managers associated with Provider.  They too will be notified.
		'UPGRADE_WARNING: Couldn't resolve default property of object objUser.FetchManagersByProvider. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rstUser = objUser.FetchManagersByProvider(lngProviderID)
		
		objMsg = CreateObject("ListBz.CMsgBz")
		objBFact = CreateObject("BenefactorBz.CBenefactorBz")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objBFact.FetchByID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rstBFact = objBFact.FetchByID(lngPatientID)
		strPatientName = rstBFact.Fields("fldFirst").Value & " " & rstBFact.Fields("fldLast").Value
		'UPGRADE_NOTE: Object rstBFact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstBFact = Nothing
		
		Select Case intMsgType
			Case MsgType.ApptCreateNoCert
				If intRecurType = 3 Then
					strMsg = "One or more recurring appointments were made for " & strPatientName & " starting "
				Else
					strMsg = "An appointment was made for " & strPatientName & " "
				End If
				strMsg = strMsg & VB6.Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
				strMsg = strMsg & " without primary certification."
				
				'Notify the Provider
				'UPGRADE_WARNING: Couldn't resolve default property of object objMsg.Insert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call objMsg.Insert(strMsg, strUserName, lngProviderID,  , "N")
				
				'Notify all Office Managers
				If intRecurType = 3 Then
					strMsg = "One or more recurring appointments were made for " & strPatientName & " with " & strProviderName & " starting at "
				Else
					strMsg = "An appointment was made for " & strPatientName & " with " & strProviderName & " at "
				End If
				strMsg = strMsg & VB6.Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
				strMsg = strMsg & " without primary certification."
				
				If Not rstUser.BOF And Not rstUser.EOF Then
					rstUser.MoveFirst()
					While Not rstUser.EOF
						If rstUser.Fields("fldDisabledYN").Value = "N" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object objMsg.Insert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Call objMsg.Insert(strMsg, strUserName, rstUser.Fields("fldUserID").Value,  , "N")
						End If
						rstUser.MoveNext()
					End While
				End If
				
			Case MsgType.ApptConfirmNoCert
				strMsg = "Your appointment with " & strPatientName & " was confirmed for "
				strMsg = strMsg & VB6.Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
				strMsg = strMsg & " without primary certification.  (" & strUserName & ")"
				
				'Notify the Provider
				'UPGRADE_WARNING: Couldn't resolve default property of object objMsg.Insert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call objMsg.Insert(strMsg, strUserName, lngProviderID,  , "N")
				
				'Notify all Office Managers
				strMsg = "An appointment with " & strPatientName & " for " & strProviderName & " was confirmed for " & VB6.Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
				strMsg = strMsg & " without primary certification.  (" & strUserName & ")"
				
				If Not rstUser.BOF And Not rstUser.EOF Then
					rstUser.MoveFirst()
					While Not rstUser.EOF
						If rstUser.Fields("fldDisabledYN").Value = "N" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object objMsg.Insert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Call objMsg.Insert(strMsg, strUserName, rstUser.Fields("fldUserID").Value,  , "N")
						End If
						rstUser.MoveNext()
					End While
				End If
				
			Case MsgType.ApptCancel
				strMsg = "Your appointment on " & VB6.Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
				strMsg = strMsg & " with " & strPatientName & " has been cancelled. (" & strUserName
				If Len(strCancelReason) > 1 Then strMsg = strMsg & ": " & strCancelReason
				strMsg = strMsg & ")"
				
				'Notify the Provider
				'Call objMsg.Insert(strMsg, strUserName, lngProviderID, , "C")
				
				'Notify all Office Managers
				If Not rstUser.BOF And rstUser.EOF Then
					rstUser.MoveFirst()
					While Not rstUser.EOF
						If rstUser.Fields("fldDisabledYN").Value = "N" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object objMsg.Insert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Call objMsg.Insert(strMsg, strUserName, rstUser.Fields("fldUserID").Value,  , "C")
						End If
						rstUser.MoveNext()
					End While
				End If
				
			Case MsgType.ApptDelete
				strMsg = "Your appointment on " & VB6.Format(dteStartDateTime, "mmm dd, yyyy hh:mm AMPM")
				strMsg = strMsg & " with " & strPatientName & " was deleted. (" & strUserName & ")"
				
				'Notify the Provider
				'UPGRADE_WARNING: Couldn't resolve default property of object objMsg.Insert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call objMsg.Insert(strMsg, strUserName, lngProviderID,  , "C")
				
				'Notify all Office Managers
				If Not rstUser.BOF And rstUser.EOF Then
					rstUser.MoveFirst()
					While Not rstUser.EOF
						If rstUser.Fields("fldDisabledYN").Value = "N" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object objMsg.Insert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Call objMsg.Insert(strMsg, strUserName, rstUser.Fields("fldUserID").Value,  , "C")
						End If
						rstUser.MoveNext()
					End While
				End If
		End Select
		
		'UPGRADE_NOTE: Object rstUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstUser = Nothing
		'UPGRADE_NOTE: Object objMsg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objMsg = Nothing
		'UPGRADE_NOTE: Object objBFact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objBFact = Nothing
		'UPGRADE_NOTE: Object objUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objUser = Nothing
		
		Exit Sub
		
ErrTrap: 
		'UPGRADE_NOTE: Object objMsg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objMsg = Nothing
		'UPGRADE_NOTE: Object objBFact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objBFact = Nothing
		'UPGRADE_NOTE: Object objUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objUser = Nothing
		Err.Raise(Err.Number, Err.Source, Err.Description)
		
	End Sub
End Class