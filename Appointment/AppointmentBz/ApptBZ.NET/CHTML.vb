Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("CHTMLBZ_NET.CHTMLBZ")> Public Class CHTMLBZ
	'--------------------------------------------------------------------
	'Class Name: CHTMLBZ                                            '
	'Date: 11/28/2000                                                   '
	'Author: Rick "Boom Boom" Segura                                    '
	'Description:  MTS business object designed to call methods         '
	'              associated with the CHTMLTest class.                 '
	'--------------------------------------------------------------------
	'Revision History:
	'  R001: Richkun 10/03/2001 - Ensured recurring appointments were accurately
	'           identified in appointment cells.
	'--------------------------------------------------------------------
	
	
	Private Const CLASS_NAME As String = "CHTMLBZ"
	
	Private Const START_TIME As String = "07:00:00 AM"
	Private Const END_TIME As String = "10:00:00 PM"
	
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
	Private Const ROWHEADER_LEFT_CL As String = "RowHeaderLeft"
	Private Const ROWHEADER_RIGHT_CL As String = "RowHeaderRIGHT"
	Private Const PENDING_CL As String = "Pending"
	
	Private Const ATTENDED_ST As Integer = 3
	Private Const CONFIRMED_ST As Integer = 2
	Private Const HOLD_ST As Integer = 5
	Private Const SCHEDULED_ST As Integer = 1
	Private Const NO_SHOW_ST As Integer = 6
	Private Const TENATIVE_ST As Integer = 10
	Private Const PENDING_ST As Integer = 11
	
	Public Structure udtApptCell
		Dim ID As String
		Dim ParentOffset As Integer
		Dim Type As Integer
		Dim ApptList As String
		Dim Text As String
		Dim RowSpan As Integer
		'UPGRADE_NOTE: Class was upgraded to Class_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Class_Renamed As String
		Dim Tag As String
		Dim DateTime As String
		Dim Row As Short
		Dim Recur As String 'R001
	End Structure
	
	
	'--------------------------------------------------------------------
	' Public Functions  '
	'--------------------------------------------------------------------
	
	Public Function BuildSchedTable(ByVal lngProviderID As Integer, ByVal dteStartDate As Date, ByVal dteEndDate As Date, Optional ByRef strStartTime As String = "", Optional ByRef strEndTime As String = "", Optional ByVal lngInterval As Integer = 15) As String
		'--------------------------------------------------------------------
		'Date: 11/28/2000
		'Author: Rick "Boom Boom" Segura
		'Description:  Generates the HTML code that produces the web calendar
		'Parameters: lngProviderID - ID of provider whose schedule is being produced                                            '
		'            lngClinicID - ID of clinic pertaining to schedule
		'            dteStartDate - Start date for the schedule
		'            dteEndDate - End date for the schedule
		'Returns: String of HTML code
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		
		Dim strHTML As String
		Dim dteStart As Date
		Dim dteEnd As Date
		Dim varSchedObj() As udtApptCell
		Dim strErrMessage As String
		Dim lngUBound1 As Integer
		Dim lngUBound2 As Integer
		Dim i, j As Integer
		Dim lngRowCount As Integer
		Dim strPStartTime As String
		
		On Error GoTo Err_Trap
		
		If strStartTime > "" And strEndTime > "" Then
			strPStartTime = strStartTime
		Else
			strStartTime = START_TIME
			strEndTime = END_TIME
			strPStartTime = START_TIME
			lngRowCount = (60 - 1)
		End If
		
		Select Case lngInterval
			Case 5
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 12) - 1
			Case 10
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 6) - 1
			Case 15
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 4) - 1
			Case 20
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 3) - 1
			Case 30
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 2) - 1
			Case 60
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 1) - 1
		End Select
		
		If dteStartDate > dteEndDate Then
			dteStart = dteEndDate
			dteEnd = dteStartDate
		Else
			dteStart = dteStartDate
			dteEnd = dteEndDate
		End If
		
		' Initialize the schedule array
		'UPGRADE_WARNING: Couldn't resolve default property of object InitializeSchedObject(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varSchedObj = InitializeSchedObject(dteStart, dteEnd, strPStartTime, lngRowCount, lngInterval)
		
		If IsArray(varSchedObj) Then
			lngUBound1 = UBound(varSchedObj, 1)
			lngUBound2 = UBound(varSchedObj, 2)
			
			' Populate the schedule array
			Call PopulateSchedObject(lngProviderID, dteStart, dteEnd, varSchedObj, strPStartTime, lngInterval)
			
			For j = 0 To lngUBound2
				strHTML = strHTML & "<TR>"
				For i = 0 To lngUBound1
					
					Select Case varSchedObj(i, j).Type
						
						Case HEADING_T
							strHTML = strHTML & "<TD "
							strHTML = strHTML & "ALIGN=" & IIf(i = 0, "'RIGHT' ", "'LEFT' ")
							strHTML = strHTML & "WIDTH='62' "
							strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
							strHTML = strHTML & "class='" & varSchedObj(i, j).Class_Renamed & "' >"
							If (i = 0) Then
								strHTML = strHTML & "<A name=" & TimeValue(CStr(CDate(varSchedObj(i, j).DateTime))).ToOADate() & "></A>"
							End If
							strHTML = strHTML & HDateTime(varSchedObj(i, j).DateTime)
							strHTML = strHTML & "</TD>" & vbCrLf
							
						Case OPEN_T
							'strHTML = strHTML & "<TD "
							'strHTML = strHTML & "<TD onClick=""appt('" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "')"" "
							strHTML = strHTML & "<TD onClick=""appt('" & varSchedObj(i, j).ID & ";" & CDate(varSchedObj(i, j).DateTime).ToOADate() & ";" & varSchedObj(i, j).Class_Renamed & ";" & lngProviderID & "')"" "
							strHTML = strHTML & "WIDTH='89' "
							strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
							strHTML = strHTML & "id='" & CDate(varSchedObj(i, j).DateTime).ToOADate() & "' "
							strHTML = strHTML & "time='" & CDate(varSchedObj(i, j).DateTime).ToOADate() & "' " 'R001
							strHTML = strHTML & "class='" & varSchedObj(i, j).Class_Renamed & "' "
							strHTML = strHTML & "title='" & varSchedObj(i, j).Class_Renamed & "' "
							strHTML = strHTML & ">&nbsp;</TD>" & vbCrLf
							'strHTML = strHTML & "><div class='hand' onClick=""appt('" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "')"">&nbsp;</div></TD>" & vbCrLf
							
						Case SCHEDULED_T
							If varSchedObj(i, j).ParentOffset = 0 Then
								'strHTML = strHTML & "<TD "
								strHTML = strHTML & "<TD onClick=""appt('" & varSchedObj(i, j).ID & ";" & CDate(varSchedObj(i, j).DateTime).ToOADate() & ";" & varSchedObj(i, j).Class_Renamed & ";" & lngProviderID & "')"" "
								strHTML = strHTML & "WIDTH='89' "
								If varSchedObj(i, j).RowSpan > 1 Then
									strHTML = strHTML & "rowspan=" & varSchedObj(i, j).RowSpan & " "
								End If
								strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
								strHTML = strHTML & "id='" & varSchedObj(i, j).ID & "' "
								strHTML = strHTML & "time='" & CDate(varSchedObj(i, j).DateTime).ToOADate() & "' " 'R001
								strHTML = strHTML & "recur='" & varSchedObj(i, j).Recur & "' " 'R001
								strHTML = strHTML & "class='" & varSchedObj(i, j).Class_Renamed & "' "
								strHTML = strHTML & "title='" & varSchedObj(i, j).Class_Renamed & "' "
								strHTML = strHTML & ">" & varSchedObj(i, j).Text & "</TD>" & vbCrLf
								'strHTML = strHTML & "><div class='hand' onClick=""appt('" & varSchedObj(i, j).ID & "')"">" & varSchedObj(i, j).Text & "</div></TD>" & vbCrLf
							End If
							
					End Select
					
				Next 
				strHTML = strHTML & "</TR>" & vbCrLf
			Next 
			
		Else
			strErrMessage = "An error occured while retreiving the schedule."
			GoTo Err_Trap
		End If
		
		BuildSchedTable = strHTML
		
		Erase varSchedObj
		
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
Err_Trap: 
		'Signal incompletion and raise the error to the calling environment.
		
		System.EnterpriseServices.ContextUtil.SetAbort()
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMessage)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
	End Function
	
	Public Function BuildDayAtAGlance(ByVal lngUserID As Integer, ByVal dteDate As Date, ByVal lngClinicID As Integer, Optional ByVal strStartTime As String = "7:00 AM", Optional ByVal strEndTime As String = "10:00 PM", Optional ByVal lngInterval As Integer = 15) As String
		Dim ClinicBz As Object
		Dim ProviderBz As Object
		'--------------------------------------------------------------------
		'Date: 07/30/2001                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Generates the HTML code that produces the Day At A   '
		'               Glance web table                                    '
		'Parameters: lngUserID - the schedule for the proveiders associated '
		'               with the person with is ID will be shown            '
		'            dteDate - Day to retrieve for schedule                 '
		'            strStartTime - Start Time for Schedule                 '
		'            strEndTime - End Time for Schedule                     '
		'Returns: String of HTML code                                       '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		Dim objProvider As ProviderBz.CProviderBZ
		Dim objClinic As ClinicBz.CUserBz
		Dim rstProviders As ADODB.Recordset
		Dim strHTML As Object
		Dim varSchedObj() As udtApptCell
		Dim intTimeUBound As Object
		Dim intCnt As Short
		Dim strErrMessage As String
		Dim lngUBound1 As Integer
		Dim i As Object
		Dim j As Object
		
		On Error GoTo ErrTrap
		
		
		Select Case lngInterval
			Case 5
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				'UPGRADE_WARNING: Couldn't resolve default property of object intTimeUBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intTimeUBound = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 12) - 1
			Case 10
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				'UPGRADE_WARNING: Couldn't resolve default property of object intTimeUBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intTimeUBound = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 6) - 1
			Case 15
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				'UPGRADE_WARNING: Couldn't resolve default property of object intTimeUBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intTimeUBound = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 4) - 1
			Case 20
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				'UPGRADE_WARNING: Couldn't resolve default property of object intTimeUBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intTimeUBound = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 3) - 1
			Case 30
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				'UPGRADE_WARNING: Couldn't resolve default property of object intTimeUBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intTimeUBound = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 2) - 1
			Case 60
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				'UPGRADE_WARNING: Couldn't resolve default property of object intTimeUBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				intTimeUBound = (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60 * 1) - 1
		End Select
		
		If lngClinicID > 0 Then
			objProvider = CreateObject("ProviderBz.CProviderBz")
			'UPGRADE_WARNING: Couldn't resolve default property of object objProvider.FetchProvidersByClinic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rstProviders = objProvider.FetchProvidersByClinic(lngClinicID)
			'UPGRADE_NOTE: Object objProvider may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objProvider = Nothing
		Else
			objClinic = CreateObject("ClinicBz.CUserBz")
			'UPGRADE_WARNING: Couldn't resolve default property of object objClinic.FetchProviders. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rstProviders = objClinic.FetchProviders(lngUserID)
			'UPGRADE_NOTE: Object objClinic may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objClinic = Nothing
		End If
		
		With rstProviders
			.MoveLast()
			.MoveFirst()
			
			Dim aryProviders(.RecordCount - 1, 3) As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object intTimeUBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object InitializeDayObject(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varSchedObj = InitializeDayObject(dteDate, strStartTime, intTimeUBound, .RecordCount, lngInterval)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			i = 0
			While Not .EOF
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aryProviders(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				aryProviders(i, 0) = .Fields("fldUserID").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aryProviders(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				aryProviders(i, 1) = .Fields("fldFirstName").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aryProviders(i, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				aryProviders(i, 2) = .Fields("fldMI").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aryProviders(i, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				aryProviders(i, 3) = .Fields("fldLastName").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aryProviders(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call PopulateDaySchedObject(aryProviders(i, 0), dteDate, i + 1, varSchedObj, strStartTime, lngInterval)
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				i = i + 1
				.MoveNext()
			End While
			
		End With
		'UPGRADE_NOTE: Object rstProviders may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstProviders = Nothing
		
		If IsArray(varSchedObj) Then
			
			lngUBound1 = UBound(varSchedObj, 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object intTimeUBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For j = 0 To intTimeUBound
				'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strHTML = strHTML & "<TR>"
				For i = 0 To lngUBound1
					
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Select Case varSchedObj(i, j).Type
						
						Case HEADING_T
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "<TD "
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "ALIGN=" & IIf(i = 0, "'RIGHT' ", "'LEFT' ")
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "WIDTH='65' "
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "class='" & varSchedObj(i, j).Class_Renamed & "' >"
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If (i = 0) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "<A name=" & TimeValue(CStr(CDate(varSchedObj(i, j).DateTime))).ToOADate() & "></A>"
							End If
							' old        If (i = 0) Then
							'                strHTML = strHTML & "<A name=" & CDbl(TimeValue(CDate(varSchedObj(i, j).ID))) & "></A>"
							'            End If
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & HDateTime(varSchedObj(i, j).DateTime)
							'old        strHTML = strHTML & HDateTime(varSchedObj(i, j).ID)
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "</TD>" & vbCrLf
							
						Case OPEN_T
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object aryProviders(i - 1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "<TD onClick=""appt('" & varSchedObj(i, j).ID & ";" & CDate(varSchedObj(i, j).DateTime).ToOADate() & ";" & varSchedObj(i, j).Class_Renamed & ";" & aryProviders(i - 1, 0) & "')"" "
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "WIDTH='90' "
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "id='" & CDate(varSchedObj(i, j).DateTime).ToOADate() & "' "
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "time='" & CDate(varSchedObj(i, j).DateTime).ToOADate() & "' " 'R001
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "class='" & varSchedObj(i, j).Class_Renamed & "' "
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & "title='" & varSchedObj(i, j).Class_Renamed & "' "
							'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strHTML = strHTML & ">&nbsp;</TD>" & vbCrLf
						Case SCHEDULED_T
							'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If varSchedObj(i, j).ParentOffset = 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object aryProviders(i - 1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "<TD onClick=""appt('" & varSchedObj(i, j).ID & ";" & CDate(varSchedObj(i, j).DateTime).ToOADate() & ";" & varSchedObj(i, j).Class_Renamed & ";" & aryProviders(i - 1, 0) & "')"" "
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "WIDTH='90' "
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If varSchedObj(i, j).RowSpan > 1 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									strHTML = strHTML & "rowspan=" & varSchedObj(i, j).RowSpan & " "
								End If
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "id='" & varSchedObj(i, j).ID & "' "
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "time='" & CDate(varSchedObj(i, j).DateTime).ToOADate() & "' " 'R001
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "recur='" & varSchedObj(i, j).Recur & "' " 'R001
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "class='" & varSchedObj(i, j).Class_Renamed & "' "
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & "title='" & varSchedObj(i, j).Class_Renamed & "' "
								'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strHTML = strHTML & ">" & varSchedObj(i, j).Text & "</TD>" & vbCrLf
							End If
					End Select
					
				Next 
				'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strHTML = strHTML & "</TR>" & vbCrLf
			Next 
			
		Else
			strErrMessage = "An error occured while retreiving the schedule."
			GoTo ErrTrap
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strHTML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BuildDayAtAGlance = strHTML
		
		Erase varSchedObj
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.
		
		System.EnterpriseServices.ContextUtil.SetAbort()
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMessage)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
	End Function
	Public Function FetchOpenTimeSlots(ByVal lngProviderID As Integer, ByVal dteStartDate As Date, ByVal dteEndDate As Date, Optional ByRef strStartTime As String = "", Optional ByRef strEndTime As String = "", Optional ByVal lngInterval As Integer = 15) As Object
		'--------------------------------------------------------------------
		'Date: 02/10/2005
		'Author: Duane C Orth
		'Description:  Generates the HTML code that produces the web calendar
		'Parameters: lngProviderID - ID of provider whose schedule is being produced                                            '
		'            lngClinicID - ID of clinic pertaining to schedule
		'            dteStartDate - Start date for the schedule
		'            dteEndDate - End date for the schedule
		'Returns: String of HTML code
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		
		Dim strHTML As String
		Dim dteStart As Date
		Dim dteEnd As Date
		Dim varSchedObj() As udtApptCell
		Dim aryAppts() As Object
		Dim intApptCount As Integer
		Dim strErrMessage As String
		Dim lngUBound1 As Integer
		Dim lngUBound2 As Integer
		Dim i, j As Integer
		Dim lngRowCount As Integer
		Dim strPStartTime As String
		
		On Error GoTo Err_Trap
		
		If strStartTime > "" And strEndTime > "" Then
			strPStartTime = strStartTime
		Else
			strStartTime = START_TIME
			strEndTime = END_TIME
			strPStartTime = START_TIME
			lngRowCount = (60 - 1)
		End If
		
		Select Case lngInterval
			Case 5
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = ((DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60) * 12) - 1
			Case 10
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = ((DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60) * 6) - 1
			Case 15
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = ((DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60) * 4) - 1
			Case 20
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = ((DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60) * 3) - 1
			Case 30
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = ((DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60) * 2) - 1
			Case 60
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				lngRowCount = ((DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(strStartTime), CDate(strEndTime)) / 60) * 1) - 1
		End Select
		
		If dteStartDate > dteEndDate Then
			dteStart = dteEndDate
			dteEnd = dteStartDate
		Else
			dteStart = dteStartDate
			dteEnd = dteEndDate
		End If
		
		' Initialize the schedule array
		'UPGRADE_WARNING: Couldn't resolve default property of object InitializeSchedObject(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varSchedObj = InitializeSchedObject(dteStart, dteEnd, strPStartTime, lngRowCount, lngInterval)
		
		If IsArray(varSchedObj) Then
			lngUBound1 = UBound(varSchedObj, 1)
			lngUBound2 = UBound(varSchedObj, 2)
			ReDim aryAppts(5, lngUBound1 * lngUBound2)
			
			' Populate the schedule array
			Call PopulateSchedObject(lngProviderID, dteStart, dteEnd, varSchedObj, strPStartTime, lngInterval)
			
			For i = 0 To lngUBound1
				For j = 0 To lngUBound2
					
					If varSchedObj(i, j).Type = OPEN_T Then
						'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(0, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryAppts(0, intApptCount) = CDate(varSchedObj(i, j).DateTime).ToOADate()
						'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(1, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryAppts(1, intApptCount) = HDateTime(varSchedObj(i, j).DateTime)
						'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(2, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryAppts(2, intApptCount) = VB6.Format(CDate(varSchedObj(i, j).DateTime).ToOADate(), "Short Date")
						'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(3, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryAppts(3, intApptCount) = varSchedObj(i, j).Text
						'UPGRADE_WARNING: Couldn't resolve default property of object aryAppts(4, intApptCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryAppts(4, intApptCount) = varSchedObj(i, j).Class_Renamed
						intApptCount = intApptCount + 1
					End If
					
				Next 
			Next 
			
		Else
			strErrMessage = "An error occured while retreiving the schedule."
			GoTo Err_Trap
		End If
		
		'Shrink the array, if needed
		If IsArray(aryAppts) Then
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNothing(aryAppts) Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object FetchOpenTimeSlots. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FetchOpenTimeSlots = System.DBNull.Value
			Else
				ReDim Preserve aryAppts(5, intApptCount)
				'UPGRADE_WARNING: Couldn't resolve default property of object FetchOpenTimeSlots. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FetchOpenTimeSlots = VB6.CopyArray(aryAppts)
			End If
		Else
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FetchOpenTimeSlots. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FetchOpenTimeSlots = System.DBNull.Value
		End If
		
		Erase varSchedObj
		System.EnterpriseServices.ContextUtil.SetComplete()
		Exit Function
		
Err_Trap: 
		'Signal incompletion and raise the error to the calling environment.
		
		System.EnterpriseServices.ContextUtil.SetAbort()
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMessage)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
	End Function
	
	'--------------------------------------------------------------------
	' Private Functions
	'--------------------------------------------------------------------
	
	Private Function InitializeSchedObject(ByVal dteStartDate As Date, ByVal dteEndDate As Date, ByVal strStartTime As String, ByVal lngUBound2 As Integer, ByVal lngInterval As Integer) As Object
		'--------------------------------------------------------------------
		'Date: 11/28/2000
		'Author: Rick "Boom Boom" Segura
		'Description:  Creates and initializes the schedule object(array)
		'Parameters: dteStartDate - Start date for the schedule
		'            dteEndDate - End date for the schedule
		'            strStartTime - The starting time of the calendar (as a string)
		'            lngUBound2 - Value identifying the 2nd dimension of the
		'               2-dimensional calendar array (times).
		'Returns: Array of udtApptCells
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		
		Dim lngNumDays As Integer
		Dim arySched() As udtApptCell
		Dim lngUBound1 As Integer
		Dim i, j As Integer
		Dim dteDay As Date
		Dim dteCell As Date
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		lngNumDays = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dteStartDate, dteEndDate) + 1
		lngUBound1 = lngNumDays + 1
		
		ReDim arySched(lngUBound1, lngUBound2)
		
		'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dteDay = CDate(DateValue(CStr(dteStartDate)) & " " & strStartTime)
		
		'Initialize
		For i = 0 To lngUBound1 - 1
			dteCell = dteDay
			
			For j = 0 To lngUBound2
				If i = 0 Then
					'Populate both RowHeaders simultaneously
					arySched(i, j).Type = HEADING_T
					arySched(i, j).Class_Renamed = ROWHEADER_LEFT_CL
					
					arySched(lngUBound1, j).Type = HEADING_T
					arySched(lngUBound1, j).Class_Renamed = ROWHEADER_RIGHT_CL
					arySched(lngUBound1, j).DateTime = VB6.Format(dteCell, "mm-dd-yyyy hh:nn:00 AM/PM")
					arySched(lngUBound1, j).Row = lngUBound1
				Else
					arySched(i, j).ID = CStr(0) 'Used to hold Appt IDs     'was CStr(dteCell)
					arySched(i, j).Class_Renamed = OPEN_CL
				End If
				arySched(i, j).DateTime = VB6.Format(dteCell, "mm-dd-yyyy hh:nn:00 AM/PM")
				arySched(i, j).Row = j
				
				dteCell = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, lngInterval, dteCell)
			Next 
			
			If i > 0 Then ' Don't increment heading
				dteDay = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, dteDay)
			End If
		Next 
		
		' Return the Schedule object
		'UPGRADE_WARNING: Couldn't resolve default property of object InitializeSchedObject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InitializeSchedObject = VB6.CopyArray(arySched)
		
	End Function
	
	Private Function InitializeDayObject(ByVal dteDate As Date, ByVal strStartTime As String, ByVal lngTimeUBound As Integer, ByVal lngNumProviders As Integer, ByVal lngInterval As Integer) As Object
		'--------------------------------------------------------------------
		'Date: 11/28/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Creates and initializes the schedule object(array)   '
		'Parameters: dteDate - Day for the schedule                         '
		'            strStartTime - StartTime for schedule                  '
		'            lngTimeUBound - Upper Bound for time range             '
		'Returns: Array of udtApptCells                                     '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		
		Dim lngNumDays As Integer
		Dim arySched() As udtApptCell
		Dim lngUBound1 As Integer
		Dim i, j As Integer
		Dim dteDay As Date
		Dim dteCell As Date
		
		lngUBound1 = lngNumProviders + 1
		
		ReDim arySched(lngUBound1, lngTimeUBound)
		
		'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dteDay = CDate(DateValue(CStr(dteDate)) & " " & strStartTime)
		
		' Initialize
		For i = 0 To lngUBound1
			dteCell = dteDay
			
			For j = 0 To lngTimeUBound
				arySched(i, j).ID = CStr(0) 'Used to hold Appt IDs     'was CStr(dteCell)
				arySched(i, j).Class_Renamed = OPEN_CL
				' Default cell type is OPEN_T (0)
				If i = 0 Then
					arySched(i, j).Type = HEADING_T
					arySched(i, j).Class_Renamed = ROWHEADER_LEFT_CL
				End If
				If i = lngUBound1 Then
					arySched(i, j).Type = HEADING_T
					arySched(i, j).Class_Renamed = ROWHEADER_RIGHT_CL
				End If
				arySched(i, j).DateTime = VB6.Format(dteCell, "mm-dd-yyyy hh:nn:00 AM/PM")
				arySched(i, j).Row = j
				dteCell = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, lngInterval, dteCell)
			Next 
			
			'  If i > 0 Then  ' Don't increment heading
			'      dteDay = DateAdd("d", 1, dteDay)
			'  End If
		Next 
		
		' Return the Schedule object
		'UPGRADE_WARNING: Couldn't resolve default property of object InitializeDayObject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InitializeDayObject = VB6.CopyArray(arySched)
	End Function
	
	Private Sub PopulateSchedObject(ByVal lngProviderID As Integer, ByVal dteStartDate As Date, ByVal dteEndDate As Date, ByRef arySched() As udtApptCell, ByVal strApptStartTime As String, ByVal lngInterval As Integer)
		'--------------------------------------------------------------------
		'Date: 11/29/2000
		'Author: Rick "Boom Boom" Segura
		'Description:  Driver for populating the Schedule Object
		'Parameters: lngProviderID - ID of provider whose schedule is being produced                                            '
		'            dteStartDate - Start date for the schedule
		'            dteEndDate - End date for the schedule
		'            arySched - 2-dimensional array of UDTs that will be populated
		'            strApptStartTime - Starting time of the calendar in string format
		'Returns: Null
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		Dim objAppt As ApptBZ.CApptBZ
		Dim rstAppt As ADODB.Recordset
		Dim rstExc As ADODB.Recordset
		Dim rstEnc As ADODB.Recordset
		Dim lngUBound1, lngUBound2 As Integer
		Dim lngDayCell As Integer
		Dim lngTimeCell As Integer
		Dim lngOffset As Integer
		Dim lngRowSpan As Integer
		Dim i As Integer
		Dim intCtr As Short
		Dim lngApptStatusID As Integer
		Dim aryRecurDates As Object
		Dim lngTimeBound As Object
		Dim blnRecur As Boolean
		Dim blnException As Boolean
		
		On Error GoTo ErrTrap
		
		objAppt = CreateObject("ApptBZ.CApptBZ")
		rstAppt = objAppt.FetchByProviderDateRange(lngProviderID, dteStartDate, dteEndDate)
		rstExc = objAppt.FetchProviderExceptions(lngProviderID, dteStartDate, dteEndDate)
		'UPGRADE_WARNING: Couldn't resolve default property of object lngTimeBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngTimeBound = UBound(arySched, 2)
		
		While Not rstAppt.EOF
			lngApptStatusID = rstAppt.Fields("fldApptStatus").Value
			lngTimeCell = GetTimeCell(rstAppt.Fields("fldStartDateTime").Value, strApptStartTime, lngInterval)
			'UPGRADE_WARNING: Couldn't resolve default property of object lngTimeBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If lngTimeCell > lngTimeBound Then GoTo SkipIt ' Appointment is beyond range
			
			If lngTimeCell < 0 Then
				lngRowSpan = CShort(rstAppt.Fields("fldDuration").Value / lngInterval)
				If (lngTimeCell + lngRowSpan) < 0 Then
					GoTo SkipIt
				Else
					lngRowSpan = lngTimeCell + lngRowSpan
					lngTimeCell = 0
					'UPGRADE_WARNING: Couldn't resolve default property of object lngTimeBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngRowSpan = Min(lngRowSpan, lngTimeBound - lngTimeCell + 1)
				End If
			Else
				' Ensure span does not go beyond schedule range
				'UPGRADE_WARNING: Couldn't resolve default property of object lngTimeBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRowSpan = Min(CShort(rstAppt.Fields("fldDuration").Value / lngInterval), lngTimeBound - lngTimeCell + 1)
			End If
			
			'Get valid recurring dates if any
			If Trim(rstAppt.Fields("fldRecurPattern").Value) > "" Then
				blnRecur = True 'R001
				'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.GetRecurApptDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				aryRecurDates = objAppt.GetRecurApptDates(dteStartDate, dteEndDate, rstAppt.Fields("fldStartDateTime").Value, rstAppt.Fields("fldEndDateTime").Value, rstAppt.Fields("fldDuration").Value, rstAppt.Fields("fldRecurPattern").Value, rstAppt.Fields("fldInterval").Value, rstAppt.Fields("fldDOWMask").Value, rstAppt.Fields("fldDOM").Value, rstAppt.Fields("fldWOM").Value, rstAppt.Fields("fldMOY").Value)
			Else
				blnRecur = False 'R001
				'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				aryRecurDates = New Object(){rstAppt.Fields("fldStartDateTime").Value}
			End If
			
			'Fill cell for every valid date
			If IsArray(aryRecurDates) Then
				For i = 0 To UBound(aryRecurDates)
					lngApptStatusID = rstAppt.Fields("fldApptStatus").Value
					blnException = False
					'Take out recurring appointment exceptions
					If rstAppt.Fields("fldRecurPattern").Value > "" Then
						If rstExc.RecordCount > 0 Then
							rstExc.MoveFirst()
							For intCtr = 1 To rstExc.RecordCount
								'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If CDate(rstExc.Fields("fldApptDate").Value) = CDate(aryRecurDates(i)) Then
									If rstExc.Fields("fldRecurApptID").Value = rstAppt.Fields("fldApptID").Value Then
										blnException = True
										Exit For
									End If
								End If
								If Not rstExc.EOF Then
									rstExc.MoveNext()
								End If
							Next 
						End If
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngDayCell = GetDayCell(CDate(aryRecurDates(i)), dteStartDate)
					
					If blnException = False Then
						'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rstEnc = objAppt.FetchEncByDOS(lngProviderID, rstAppt.Fields("fldApptID").Value, CDate(aryRecurDates(i)))
						If rstEnc.RecordCount > 0 Then
							lngApptStatusID = 3
						End If
						'UPGRADE_NOTE: Object rstEnc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						rstEnc = Nothing
						
						Call FillCell(rstAppt.Fields("fldApptID").Value, rstAppt.Fields("PatientCount").Value, lngRowSpan, rstAppt.Fields("fldCategoryID").Value, rstAppt.Fields("fldApptTitle").Value, Mid(NNs(rstAppt.Fields("fldNote").Value), 1, 25), lngApptStatusID, lngDayCell, lngTimeCell, blnRecur, arySched)
					End If
				Next 
			End If
SkipIt: 
			rstAppt.MoveNext()
		End While
		
		' Clean House
		'UPGRADE_NOTE: Object rstAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAppt = Nothing
		'UPGRADE_NOTE: Object rstExc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstExc = Nothing
		'UPGRADE_NOTE: Object rstEnc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstEnc = Nothing
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Sub
		
ErrTrap: 
		' Clean House
		'UPGRADE_NOTE: Object rstAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAppt = Nothing
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object rstExc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstExc = Nothing
		'UPGRADE_NOTE: Object rstEnc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstEnc = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Sub
	
	Private Sub PopulateDaySchedObject(ByVal lngProviderID As Integer, ByVal dteDate As Date, ByVal intCol As Short, ByRef arySched() As udtApptCell, ByVal strStartTime As String, ByVal lngInterval As Integer)
		'--------------------------------------------------------------------
		'Date: 07/30/2001                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Driver for populating the Schedule Object            '
		'Parameters: lngProviderID - ID of provider whose schedule is being '
		'               produced                                            '
		'            lngClinicID - ID of clinic pertaining to schedule      '
		'            dteStartDate - Start date for the schedule             '
		'            dteEndDate - End date for the schedule                 '
		'            arySched - Schedule array                              '
		'Returns: Nothing                                                   '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		Dim rst As ADODB.Recordset
		Dim objAppt As ApptBZ.CApptBZ
		Dim rstExc As ADODB.Recordset
		Dim lngUBound1, lngUBound2 As Integer
		Dim lngDayCell As Integer
		Dim lngTimeCell As Integer
		Dim lngOffset As Integer
		Dim lngRowSpan As Integer
		Dim i As Integer
		Dim intCtr As Short
		Dim aryRecurDates As Object
		Dim lngTimeBound As Object
		Dim blnRecur As Boolean 'R001
		Dim blnException As Boolean
		
		On Error GoTo Err_Trap
		
		' Populate the recordset
		objAppt = CreateObject("ApptBZ.CApptBZ")
		rst = objAppt.FetchByProviderDateRange(lngProviderID, dteDate, dteDate)
		'   Set rst = objAppt.FetchByProviderDateRange(lngProviderID, dteDate, DateAdd("d", -1, CDate(dteDate)))
		rstExc = objAppt.FetchProviderExceptions(lngProviderID, dteDate, dteDate)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object lngTimeBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngTimeBound = UBound(arySched, 2)
		
		With rst
			While Not .EOF
				lngTimeCell = GetTimeCell(.Fields("fldStartDateTime").Value, strStartTime, lngInterval)
				'UPGRADE_WARNING: Couldn't resolve default property of object lngTimeBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If lngTimeCell > lngTimeBound Then GoTo SkipIt ' Appointment is beyond range
				
				If lngTimeCell < 0 Then
					lngRowSpan = CShort(.Fields("fldDuration").Value / lngInterval)
					If (lngTimeCell + lngRowSpan) < 0 Then
						GoTo SkipIt
					Else
						lngRowSpan = lngTimeCell + lngRowSpan
						lngTimeCell = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object lngTimeBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRowSpan = Min(lngRowSpan, lngTimeBound - lngTimeCell + 1)
					End If
				Else
					' Ensure span does not go beyond schedule range
					'UPGRADE_WARNING: Couldn't resolve default property of object lngTimeBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngRowSpan = Min(CShort(.Fields("fldDuration").Value / lngInterval), lngTimeBound - lngTimeCell + 1)
				End If
				
				' Get valid recurring dates if any
				If Trim(.Fields("fldRecurPattern").Value) > "" Then
					blnRecur = True 'R001
					'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.GetRecurApptDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryRecurDates = objAppt.GetRecurApptDates(dteDate, dteDate, .Fields("fldStartDateTime").Value, .Fields("fldEndDateTime").Value, .Fields("fldDuration").Value, .Fields("fldRecurPattern").Value, .Fields("fldInterval").Value, .Fields("fldDOWMask").Value, .Fields("fldDOM").Value, .Fields("fldWOM").Value, .Fields("fldMOY").Value)
				Else
					blnRecur = False 'R001
					'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					aryRecurDates = New Object(){.Fields("fldStartDateTime").Value}
				End If
				
				'Fill cell for every valid date
				If IsArray(aryRecurDates) Then
					For i = 0 To UBound(aryRecurDates)
						blnException = False
						'Take out recurring appointment exceptions
						If .Fields("fldRecurPattern").Value > "" Then
							If rstExc.RecordCount > 0 Then
								rstExc.MoveFirst()
								For intCtr = 1 To rstExc.RecordCount
									'UPGRADE_WARNING: Couldn't resolve default property of object aryRecurDates(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If CDate(rstExc.Fields("fldApptDate").Value) = CDate(aryRecurDates(i)) Then
										If rstExc.Fields("fldRecurApptID").Value = .Fields("fldApptID").Value Then
											blnException = True
											Exit For
										End If
									End If
									If Not rstExc.EOF Then
										rstExc.MoveNext()
									End If
								Next 
							End If
						End If
						
						lngDayCell = intCol
						' lngDayCell = GetDayCell(CDate(aryRecurDates(i)), dteDate)
						
						If blnException = False Then
							Call FillCell(.Fields("fldApptID").Value, .Fields("PatientCount").Value, lngRowSpan, .Fields("fldCategoryID").Value, NNs(.Fields("fldApptTitle").Value), Mid(NNs(.Fields("fldNote").Value), 1, 25), .Fields("fldApptStatus").Value, lngDayCell, lngTimeCell, blnRecur, arySched)
						End If
					Next 
				End If
				
SkipIt: 
				.MoveNext()
			End While
		End With
		
		' Clean House
		'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rst = Nothing
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object rstExc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstExc = Nothing
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Sub
		
Err_Trap: 
		' Clean House
		'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rst = Nothing
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object rstExc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstExc = Nothing
		System.EnterpriseServices.ContextUtil.SetAbort()
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Sub
	
	Private Sub FillCell(ByVal lngApptID As Integer, ByVal lngPatientCount As Integer, ByVal intRowSpan As Short, ByVal lngCategoryID As Integer, ByVal strApptTitle As String, ByVal strNote As String, ByVal lngApptStatusID As Integer, ByVal lngDayCell As Integer, ByVal lngTimeCell As Integer, ByVal blnRecur As Boolean, ByRef arySched() As udtApptCell)
		'--------------------------------------------------------------------
		'Date: 01/22/2001                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Appropriatley sets cell values based on the given    '
		'              parameters and schedule attributes                   '
		'Parameters: all required fields                                    '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		Const BR As String = "<BR>-----<BR>"
		
		Dim i As Short
		Dim intNewRowSpan As Short
		Dim lngNewTimeCell As Integer
		Dim strApptList As String
		Dim strTextList As String
		Dim varApptList As Object
		Dim varTextList As Object
		Dim udtCell As udtApptCell
		Dim lngTimeBound As Integer
		
		' Look for conflicts
		intNewRowSpan = intRowSpan
		lngNewTimeCell = lngTimeCell
		
		'UPGRADE_WARNING: Couldn't resolve default property of object udtCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		udtCell = arySched(lngDayCell, lngNewTimeCell)
		
		If udtCell.ParentOffset Then
			lngNewTimeCell = lngNewTimeCell - udtCell.ParentOffset
			intNewRowSpan = Max(intNewRowSpan + udtCell.ParentOffset, arySched(lngDayCell, lngNewTimeCell).RowSpan)
		End If
		
		' Make sure our rowspan stays inbounds
		If (lngNewTimeCell + intNewRowSpan - 1) > UBound(arySched, 2) Then
			intNewRowSpan = UBound(arySched, 2) - lngNewTimeCell
		End If
		
		' Next look for appointments that are overlapping
		For i = 0 To (intNewRowSpan - 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object udtCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			udtCell = arySched(lngDayCell, lngNewTimeCell + i)
			
			If udtCell.ParentOffset = 0 And udtCell.Type = SCHEDULED_T Then
				StrCatDel(strApptList, udtCell.ApptList)
				StrCatDel(strTextList, udtCell.Text, BR)
				intNewRowSpan = Max(intNewRowSpan, udtCell.RowSpan + i)
			End If
		Next 
		
		' Set "conflict" arrays
		If strApptList > "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object varApptList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varApptList = Split(strApptList, ",")
			'UPGRADE_WARNING: Couldn't resolve default property of object varTextList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varTextList = Split(strTextList, BR)
			'strTextList = Join(varTextList, BR)
		Else
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varApptList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varApptList = System.DBNull.Value
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varTextList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varTextList = System.DBNull.Value
		End If
		
		' Make sure our rowspan stays inbounds -- DR why is this occuring twice in this procedure?
		If (lngNewTimeCell + intNewRowSpan - 1) > UBound(arySched, 2) Then
			intNewRowSpan = UBound(arySched, 2) - lngNewTimeCell
		End If
		
		If IsArray(varApptList) Then
			' Conflict Exists
			For i = 0 To (intNewRowSpan - 1)
				' Populate array values
				With arySched(lngDayCell, lngNewTimeCell + i)
					.ID = strApptList & "," & lngApptID
					.ParentOffset = i
					.Type = SCHEDULED_T
					.ApptList = .ID
					.RowSpan = intNewRowSpan - i
					.Class_Renamed = CONFLICT_CL
					.Tag = ""
					.Text = strTextList & BR
					If blnRecur = True Then 'R001
						.Recur = "Y"
					Else
						.Recur = "N"
					End If
					Select Case lngCategoryID
						Case PATIENT_CAT
							If lngPatientCount > 1 Then
								.Text = .Text & "Group(" & lngPatientCount & ")"
							Else
								.Text = .Text & strApptTitle
							End If
							If strNote > "" Then
								.Text = .Text & "<BR>" & FormatText(strNote)
							End If
							
						Case BLOCK_CAT
							If strNote > "" Then
								.Text = .Text & FormatText(strNote)
							Else
								.Text = .Text & "Block"
							End If
					End Select
					
				End With
			Next 
		Else
			
			For i = 0 To (intRowSpan - 1)
				' Populate array values
				With arySched(lngDayCell, lngTimeCell + i)
					.ApptList = CStr(lngApptID)
					.ParentOffset = i
					.Type = SCHEDULED_T
					.RowSpan = intRowSpan - i
					If blnRecur = True Then 'R001
						.Recur = "Y"
					Else
						.Recur = "N"
					End If
					
					Select Case lngCategoryID
						Case PATIENT_CAT
							.ID = CStr(lngApptID)
							
							If lngPatientCount > 1 Then
								.Text = "Group(" & lngPatientCount & ")"
							Else
								.Text = strApptTitle
							End If
							
							If strNote > "" Then
								.Text = .Text & "<BR>" & FormatText(strNote)
							End If
							
							Select Case lngApptStatusID
								Case ATTENDED_ST
									.Class_Renamed = ATTENDED_CL
								Case CONFIRMED_ST
									If lngPatientCount > 1 Then
										.Class_Renamed = GROUP_CL
									Else
										.Class_Renamed = CONFIRMED_CL
									End If
								Case SCHEDULED_ST
									If lngPatientCount > 1 Then
										.Class_Renamed = GROUP_CL
									Else
										.Class_Renamed = SCHEDULED_CL
									End If
								Case HOLD_ST
									.Text = "Hold"
									.Class_Renamed = HOLD_CL
								Case NO_SHOW_ST
									.Class_Renamed = NO_SHOW_CL
								Case TENATIVE_ST
									.Class_Renamed = TENATIVE_CL
								Case PENDING_ST
									.Class_Renamed = PENDING_CL
							End Select
							
						Case BLOCK_CAT
							.ID = CStr(lngApptID)
							If strNote > "" Then
								.Text = FormatText(strNote)
							Else
								.Text = "Block"
							End If
							.Class_Renamed = BLOCKED_CL
							.Tag = "S"
					End Select
				End With
			Next 
		End If
		
	End Sub
	
	Private Function GetDayCell(ByVal dteSeek As Date, ByVal dteStart As Date) As Integer
		'--------------------------------------------------------------------
		'Date: 11/29/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Maps a date to Day column on the schedule object     '
		'Parameters: dteSeek - date to be translated                        '
		'            dteStart - relative start date                         '
		'Returns: Long                                                      '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		GetDayCell = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, dteStart, dteSeek) + 1)
		
	End Function
	
	Private Function GetTimeCell(ByVal dteSeek As Date, ByVal strStartTime As Object, ByVal lngInterval As Integer) As Integer
		'--------------------------------------------------------------------
		'Date: 11/29/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Maps a date(time) to time row on the schedule object '
		'Parameters: dteSeek - date to be translated                        '
		'Returns: Long                                                      '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		'UPGRADE_WARNING: Couldn't resolve default property of object strStartTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		GetTimeCell = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Minute, strStartTime, TimeValue(CStr(dteSeek))) / lngInterval)
		
	End Function
	
	Private Function FDateTime(ByVal strDate As String) As String
		'--------------------------------------------------------------------
		'Date: 01/23/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Formats time/date string for a scheduled cell        '
		'Parameters: strDate - date to be formatted                         '
		'Returns: formatted time/date string                                '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		FDateTime = Replace(Replace(strDate, ":00 ", ""), " ", "|")
	End Function
	
	Private Function HDateTime(ByVal strDate As String) As String
		'--------------------------------------------------------------------
		'Date: 01/23/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Formats time/date string for a heading cell          '
		'Parameters: strDate - date to be formatted                         '
		'Returns: formatted time/date string                                '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		HDateTime = Replace(FormatDateTime(CDate(strDate), DateFormat.LongTime), ":00 ", " ")
	End Function
	
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function FormatText(ByVal str_Renamed As String) As String
		'--------------------------------------------------------------------
		'Date: 07/16/2001                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Formats text that could potentially harm HTML        '
		'Parameters: str - String to be formatted                           '
		'Returns: formatted string                                          '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		FormatText = Replace(Replace(str_Renamed, "<", "&lt;"), ">", "&gt;")
	End Function
End Class