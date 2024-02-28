VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTMLBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Option Explicit

Private Const CLASS_NAME As String = "CHTMLBZ"

Private Const START_TIME As String = "07:00:00 AM"
Private Const END_TIME As String = "10:00:00 PM"

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
Private Const ROWHEADER_LEFT_CL As String = "RowHeaderLeft"
Private Const ROWHEADER_RIGHT_CL As String = "RowHeaderRIGHT"
Private Const PENDING_CL As String = "Pending"

Private Const ATTENDED_ST As Long = 3
Private Const CONFIRMED_ST As Long = 2
Private Const HOLD_ST As Long = 5
Private Const SCHEDULED_ST As Long = 1
Private Const NO_SHOW_ST As Long = 6
Private Const TENATIVE_ST As Long = 10
Private Const PENDING_ST As Long = 11

Public Type udtApptCell
    ID As String
    ParentOffset As Long
    Type As Long
    ApptList As String
    Text As String
    RowSpan As Long
    Class As String
    Tag As String
    DateTime As String
    Row As Integer
    Recur As String 'R001
End Type


'--------------------------------------------------------------------
' Public Functions  '
'--------------------------------------------------------------------

Public Function BuildSchedTable(ByVal lngProviderID As Long, ByVal dteStartDate As Date, _
                                 ByVal dteEndDate As Date, Optional strStartTime As String = "", _
                                 Optional strEndTime As String = "", _
                                 Optional ByVal lngInterval As Long = 15) As String
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
    Dim lngUBound1 As Long
    Dim lngUBound2 As Long
    Dim i As Long, j As Long
    Dim lngRowCount As Long
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
            lngRowCount = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 12) - 1
        Case 10
            lngRowCount = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 6) - 1
        Case 15
            lngRowCount = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 4) - 1
        Case 20
            lngRowCount = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 3) - 1
        Case 30
            lngRowCount = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 2) - 1
        Case 60
            lngRowCount = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 1) - 1
    End Select
        
    If dteStartDate > dteEndDate Then
        dteStart = dteEndDate
        dteEnd = dteStartDate
    Else
        dteStart = dteStartDate
        dteEnd = dteEndDate
    End If
    
    ' Initialize the schedule array
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
                        strHTML = strHTML & "class='" & varSchedObj(i, j).Class & "' >"
                        If (i = 0) Then
                            strHTML = strHTML & "<A name=" & CDbl(TimeValue(CDate(varSchedObj(i, j).DateTime))) & "></A>"
                        End If
                        strHTML = strHTML & HDateTime(varSchedObj(i, j).DateTime)
                        strHTML = strHTML & "</TD>" & vbCrLf
            
                    Case OPEN_T
                        'strHTML = strHTML & "<TD "
                        'strHTML = strHTML & "<TD onClick=""appt('" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "')"" "
                        strHTML = strHTML & "<TD onClick=""appt('" & varSchedObj(i, j).ID & ";" & CDbl(CDate(varSchedObj(i, j).DateTime)) & ";" & varSchedObj(i, j).Class & ";" & lngProviderID & "')"" "
                        strHTML = strHTML & "WIDTH='89' "
                        strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
                        strHTML = strHTML & "id='" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "' "
                        strHTML = strHTML & "time='" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "' " 'R001
                        strHTML = strHTML & "class='" & varSchedObj(i, j).Class & "' "
                        strHTML = strHTML & "title='" & varSchedObj(i, j).Class & "' "
                        strHTML = strHTML & ">&nbsp;</TD>" & vbCrLf
                        'strHTML = strHTML & "><div class='hand' onClick=""appt('" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "')"">&nbsp;</div></TD>" & vbCrLf
                        
                    Case SCHEDULED_T
                        If varSchedObj(i, j).ParentOffset = 0 Then
                            'strHTML = strHTML & "<TD "
                            strHTML = strHTML & "<TD onClick=""appt('" & varSchedObj(i, j).ID & ";" & CDbl(CDate(varSchedObj(i, j).DateTime)) & ";" & varSchedObj(i, j).Class & ";" & lngProviderID & "')"" "
                            strHTML = strHTML & "WIDTH='89' "
                            If varSchedObj(i, j).RowSpan > 1 Then
                                strHTML = strHTML & "rowspan=" & varSchedObj(i, j).RowSpan & " "
                            End If
                            strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
                            strHTML = strHTML & "id='" & varSchedObj(i, j).ID & "' "
                            strHTML = strHTML & "time='" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "' " 'R001
                            strHTML = strHTML & "recur='" & varSchedObj(i, j).Recur & "' " 'R001
                            strHTML = strHTML & "class='" & varSchedObj(i, j).Class & "' "
                            strHTML = strHTML & "title='" & varSchedObj(i, j).Class & "' "
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
    
    GetObjectContext.SetComplete
    Exit Function
    
Err_Trap:
    'Signal incompletion and raise the error to the calling environment.

    GetObjectContext.SetAbort
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Public Function BuildDayAtAGlance(ByVal lngUserID As Long, ByVal dteDate As Date, _
                                 ByVal lngClinicID As Long, _
                                 Optional ByVal strStartTime As String = "7:00 AM", _
                                 Optional ByVal strEndTime As String = "10:00 PM", _
                                 Optional ByVal lngInterval As Long = 15) As String
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
    Dim strHTML
    Dim varSchedObj() As udtApptCell
    Dim intTimeUBound
    Dim intCnt As Integer
    Dim strErrMessage As String
    Dim lngUBound1 As Long
    Dim i
    Dim j
    
    On Error GoTo ErrTrap
    
    
    Select Case lngInterval
        Case 5
            intTimeUBound = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 12) - 1
        Case 10
            intTimeUBound = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 6) - 1
        Case 15
            intTimeUBound = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 4) - 1
        Case 20
            intTimeUBound = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 3) - 1
        Case 30
            intTimeUBound = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 2) - 1
        Case 60
            intTimeUBound = (DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60 * 1) - 1
    End Select
    
    If lngClinicID > 0 Then
        Set objProvider = CreateObject("ProviderBz.CProviderBz")
        Set rstProviders = objProvider.FetchProvidersByClinic(lngClinicID)
        Set objProvider = Nothing
    Else
        Set objClinic = CreateObject("ClinicBz.CUserBz")
        Set rstProviders = objClinic.FetchProviders(lngUserID)
        Set objClinic = Nothing
    End If
    
    With rstProviders
        .MoveLast
        .MoveFirst
        
        ReDim aryProviders(.RecordCount - 1, 3)
        varSchedObj = InitializeDayObject(dteDate, strStartTime, intTimeUBound, .RecordCount, lngInterval)
        
        i = 0
        While Not .EOF
            aryProviders(i, 0) = .Fields("fldUserID").Value
            aryProviders(i, 1) = .Fields("fldFirstName").Value
            aryProviders(i, 2) = .Fields("fldMI").Value
            aryProviders(i, 3) = .Fields("fldLastName").Value
            Call PopulateDaySchedObject(aryProviders(i, 0), dteDate, i + 1, varSchedObj, strStartTime, lngInterval)
            i = i + 1
            .MoveNext
        Wend

    End With
    Set rstProviders = Nothing
    
    If IsArray(varSchedObj) Then
         
        lngUBound1 = UBound(varSchedObj, 1)
        For j = 0 To intTimeUBound
            strHTML = strHTML & "<TR>"
            For i = 0 To lngUBound1
                
                Select Case varSchedObj(i, j).Type
                
                    Case HEADING_T
                        strHTML = strHTML & "<TD "
                        strHTML = strHTML & "ALIGN=" & IIf(i = 0, "'RIGHT' ", "'LEFT' ")
                        strHTML = strHTML & "WIDTH='65' "
                        strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
                        strHTML = strHTML & "class='" & varSchedObj(i, j).Class & "' >"
                        If (i = 0) Then
                            strHTML = strHTML & "<A name=" & CDbl(TimeValue(CDate(varSchedObj(i, j).DateTime))) & "></A>"
                        End If
            ' old        If (i = 0) Then
            '                strHTML = strHTML & "<A name=" & CDbl(TimeValue(CDate(varSchedObj(i, j).ID))) & "></A>"
            '            End If
                        strHTML = strHTML & HDateTime(varSchedObj(i, j).DateTime)
            'old        strHTML = strHTML & HDateTime(varSchedObj(i, j).ID)
                        strHTML = strHTML & "</TD>" & vbCrLf
            
                    Case OPEN_T
                        strHTML = strHTML & "<TD onClick=""appt('" & varSchedObj(i, j).ID & ";" & CDbl(CDate(varSchedObj(i, j).DateTime)) & ";" & varSchedObj(i, j).Class & ";" & aryProviders(i - 1, 0) & "')"" "
                        strHTML = strHTML & "WIDTH='90' "
                        strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
                        strHTML = strHTML & "id='" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "' "
                        strHTML = strHTML & "time='" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "' " 'R001
                        strHTML = strHTML & "class='" & varSchedObj(i, j).Class & "' "
                        strHTML = strHTML & "title='" & varSchedObj(i, j).Class & "' "
                        strHTML = strHTML & ">&nbsp;</TD>" & vbCrLf
                    Case SCHEDULED_T
                        If varSchedObj(i, j).ParentOffset = 0 Then
                            strHTML = strHTML & "<TD onClick=""appt('" & varSchedObj(i, j).ID & ";" & CDbl(CDate(varSchedObj(i, j).DateTime)) & ";" & varSchedObj(i, j).Class & ";" & aryProviders(i - 1, 0) & "')"" "
                            strHTML = strHTML & "WIDTH='90' "
                            If varSchedObj(i, j).RowSpan > 1 Then
                                strHTML = strHTML & "rowspan=" & varSchedObj(i, j).RowSpan & " "
                            End If
                            strHTML = strHTML & "row='" & varSchedObj(i, j).Row & "' "
                            strHTML = strHTML & "id='" & varSchedObj(i, j).ID & "' "
                            strHTML = strHTML & "time='" & CDbl(CDate(varSchedObj(i, j).DateTime)) & "' " 'R001
                            strHTML = strHTML & "recur='" & varSchedObj(i, j).Recur & "' " 'R001
                            strHTML = strHTML & "class='" & varSchedObj(i, j).Class & "' "
                            strHTML = strHTML & "title='" & varSchedObj(i, j).Class & "' "
                            strHTML = strHTML & ">" & varSchedObj(i, j).Text & "</TD>" & vbCrLf
                        End If
                End Select
            
            Next
            strHTML = strHTML & "</TR>" & vbCrLf
        Next
        
    Else
        strErrMessage = "An error occured while retreiving the schedule."
        GoTo ErrTrap:
    End If
    
    BuildDayAtAGlance = strHTML
    
    Erase varSchedObj
    GetObjectContext.SetComplete
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.

    GetObjectContext.SetAbort
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function
Public Function FetchOpenTimeSlots(ByVal lngProviderID As Long, ByVal dteStartDate As Date, _
                                 ByVal dteEndDate As Date, Optional strStartTime As String = "", _
                                 Optional strEndTime As String = "", _
                                 Optional ByVal lngInterval As Long = 15) As Variant
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
    Dim aryAppts() As Variant, intApptCount As Long
    Dim strErrMessage As String
    Dim lngUBound1 As Long
    Dim lngUBound2 As Long
    Dim i As Long, j As Long
    Dim lngRowCount As Long
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
            lngRowCount = ((DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60) * 12) - 1
        Case 10
            lngRowCount = ((DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60) * 6) - 1
        Case 15
            lngRowCount = ((DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60) * 4) - 1
        Case 20
            lngRowCount = ((DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60) * 3) - 1
        Case 30
            lngRowCount = ((DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60) * 2) - 1
        Case 60
            lngRowCount = ((DateDiff("n", CDate(strStartTime), CDate(strEndTime)) / 60) * 1) - 1
    End Select
    
    If dteStartDate > dteEndDate Then
        dteStart = dteEndDate
        dteEnd = dteStartDate
    Else
        dteStart = dteStartDate
        dteEnd = dteEndDate
    End If
    
    ' Initialize the schedule array
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
                        aryAppts(0, intApptCount) = CDbl(CDate(varSchedObj(i, j).DateTime))
                        aryAppts(1, intApptCount) = HDateTime(varSchedObj(i, j).DateTime)
                        aryAppts(2, intApptCount) = Format(CDbl(CDate(varSchedObj(i, j).DateTime)), "Short Date")
                        aryAppts(3, intApptCount) = varSchedObj(i, j).Text
                        aryAppts(4, intApptCount) = varSchedObj(i, j).Class
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
        If IsEmpty(aryAppts) Then
            FetchOpenTimeSlots = Null
        Else
            ReDim Preserve aryAppts(5, intApptCount)
            FetchOpenTimeSlots = aryAppts
        End If
    Else
        FetchOpenTimeSlots = Null
    End If

    Erase varSchedObj
    GetObjectContext.SetComplete
    Exit Function
    
Err_Trap:
    'Signal incompletion and raise the error to the calling environment.

    GetObjectContext.SetAbort
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

'--------------------------------------------------------------------
' Private Functions
'--------------------------------------------------------------------

Private Function InitializeSchedObject(ByVal dteStartDate As Date, _
                                       ByVal dteEndDate As Date, ByVal strStartTime As String, _
                                       ByVal lngUBound2 As Long, _
                                       ByVal lngInterval As Long) As Variant
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

    Dim lngNumDays As Long
    Dim arySched() As udtApptCell
    Dim lngUBound1 As Long
    Dim i As Long, j As Long
    Dim dteDay As Date
    Dim dteCell As Date
    
    lngNumDays = DateDiff("d", dteStartDate, dteEndDate) + 1
    lngUBound1 = lngNumDays + 1
    
    ReDim arySched(lngUBound1, lngUBound2)
    
    dteDay = CDate(DateValue(dteStartDate) & " " & strStartTime)
    
    'Initialize
    For i = 0 To lngUBound1 - 1
        dteCell = dteDay
        
        For j = 0 To lngUBound2
            If i = 0 Then
                'Populate both RowHeaders simultaneously
                arySched(i, j).Type = HEADING_T
                arySched(i, j).Class = ROWHEADER_LEFT_CL
                
                arySched(lngUBound1, j).Type = HEADING_T
                arySched(lngUBound1, j).Class = ROWHEADER_RIGHT_CL
                arySched(lngUBound1, j).DateTime = Format(dteCell, "mm-dd-yyyy hh:nn:00 AM/PM")
                arySched(lngUBound1, j).Row = lngUBound1
            Else
                arySched(i, j).ID = 0 'Used to hold Appt IDs     'was CStr(dteCell)
                arySched(i, j).Class = OPEN_CL
            End If
            arySched(i, j).DateTime = Format(dteCell, "mm-dd-yyyy hh:nn:00 AM/PM")
            arySched(i, j).Row = j
            
            dteCell = DateAdd("n", lngInterval, dteCell)
        Next
        
        If i > 0 Then  ' Don't increment heading
            dteDay = DateAdd("d", 1, dteDay)
        End If
    Next
    
    ' Return the Schedule object
    InitializeSchedObject = arySched
    
End Function

Private Function InitializeDayObject(ByVal dteDate As Date, _
                                       ByVal strStartTime As String, _
                                       ByVal lngTimeUBound As Long, _
                                       ByVal lngNumProviders As Long, _
                                       ByVal lngInterval As Long) As Variant
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

    Dim lngNumDays As Long
    Dim arySched() As udtApptCell
    Dim lngUBound1 As Long
    Dim i As Long, j As Long
    Dim dteDay As Date
    Dim dteCell As Date
    
    lngUBound1 = lngNumProviders + 1
    
    ReDim arySched(lngUBound1, lngTimeUBound)
    
    dteDay = CDate(DateValue(dteDate) & " " & strStartTime)
    
    ' Initialize
    For i = 0 To lngUBound1
        dteCell = dteDay
        
        For j = 0 To lngTimeUBound
            arySched(i, j).ID = 0 'Used to hold Appt IDs     'was CStr(dteCell)
            arySched(i, j).Class = OPEN_CL
            ' Default cell type is OPEN_T (0)
            If i = 0 Then
                arySched(i, j).Type = HEADING_T
                arySched(i, j).Class = ROWHEADER_LEFT_CL
            End If
            If i = lngUBound1 Then
                arySched(i, j).Type = HEADING_T
                arySched(i, j).Class = ROWHEADER_RIGHT_CL
            End If
            arySched(i, j).DateTime = Format(dteCell, "mm-dd-yyyy hh:nn:00 AM/PM")
            arySched(i, j).Row = j
            dteCell = DateAdd("n", lngInterval, dteCell)
        Next
        
      '  If i > 0 Then  ' Don't increment heading
      '      dteDay = DateAdd("d", 1, dteDay)
      '  End If
    Next
    
    ' Return the Schedule object
    InitializeDayObject = arySched
End Function

Private Sub PopulateSchedObject(ByVal lngProviderID As Long, _
        ByVal dteStartDate As Date, ByVal dteEndDate As Date, _
        ByRef arySched() As udtApptCell, ByVal strApptStartTime As String, _
        ByVal lngInterval As Long)
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
    Dim lngUBound1 As Long, lngUBound2 As Long
    Dim lngDayCell As Long
    Dim lngTimeCell As Long
    Dim lngOffset As Long
    Dim lngRowSpan As Long
    Dim i As Long
    Dim intCtr As Integer
    Dim lngApptStatusID As Long
    Dim aryRecurDates As Variant
    Dim lngTimeBound
    Dim blnRecur As Boolean
    Dim blnException As Boolean
    
    On Error GoTo ErrTrap:
    
    Set objAppt = CreateObject("ApptBZ.CApptBZ")
    Set rstAppt = objAppt.FetchByProviderDateRange(lngProviderID, dteStartDate, dteEndDate)
    Set rstExc = objAppt.FetchProviderExceptions(lngProviderID, dteStartDate, dteEndDate)
    lngTimeBound = UBound(arySched, 2)
    
    While Not rstAppt.EOF
        lngApptStatusID = rstAppt.Fields("fldApptStatus").Value
        lngTimeCell = GetTimeCell(rstAppt.Fields("fldStartDateTime").Value, strApptStartTime, lngInterval)
        If lngTimeCell > lngTimeBound Then GoTo SkipIt  ' Appointment is beyond range
        
        If lngTimeCell < 0 Then
            lngRowSpan = CInt(rstAppt.Fields("fldDuration").Value / lngInterval)
            If (lngTimeCell + lngRowSpan) < 0 Then
                GoTo SkipIt
            Else
                lngRowSpan = lngTimeCell + lngRowSpan
                lngTimeCell = 0
                lngRowSpan = Min(lngRowSpan, lngTimeBound - lngTimeCell + 1)
            End If
        Else
            ' Ensure span does not go beyond schedule range
            lngRowSpan = Min(CInt(rstAppt.Fields("fldDuration").Value / lngInterval), lngTimeBound - lngTimeCell + 1)
        End If
        
        'Get valid recurring dates if any
        If Trim(rstAppt.Fields("fldRecurPattern").Value) > "" Then
            blnRecur = True 'R001
            aryRecurDates = objAppt.GetRecurApptDates(dteStartDate, dteEndDate, _
                            rstAppt.Fields("fldStartDateTime").Value, rstAppt.Fields("fldEndDateTime").Value, _
                            rstAppt.Fields("fldDuration").Value, rstAppt.Fields("fldRecurPattern").Value, _
                            rstAppt.Fields("fldInterval").Value, rstAppt.Fields("fldDOWMask").Value, _
                            rstAppt.Fields("fldDOM").Value, rstAppt.Fields("fldWOM").Value, _
                            rstAppt.Fields("fldMOY").Value)
        Else
            blnRecur = False 'R001
            aryRecurDates = Array(rstAppt.Fields("fldStartDateTime").Value)
        End If
        
        'Fill cell for every valid date
        If IsArray(aryRecurDates) Then
            For i = 0 To UBound(aryRecurDates)
                lngApptStatusID = rstAppt.Fields("fldApptStatus").Value
                blnException = False
                'Take out recurring appointment exceptions
                If rstAppt.Fields("fldRecurPattern").Value > "" Then
                    If rstExc.RecordCount > 0 Then
                        rstExc.MoveFirst
                        For intCtr = 1 To rstExc.RecordCount
                            If CDate(rstExc.Fields("fldApptDate")) = CDate(aryRecurDates(i)) Then
                                If rstExc.Fields("fldRecurApptID").Value = rstAppt.Fields("fldApptID") Then
                                    blnException = True
                                    Exit For
                                End If
                            End If
                            If Not rstExc.EOF Then
                                rstExc.MoveNext
                            End If
                        Next
                    End If
                End If
                
                lngDayCell = GetDayCell(CDate(aryRecurDates(i)), dteStartDate)
                
                If blnException = False Then
                    Set rstEnc = objAppt.FetchEncByDOS(lngProviderID, rstAppt.Fields("fldApptID"), CDate(aryRecurDates(i)))
                    If rstEnc.RecordCount > 0 Then
                        lngApptStatusID = 3
                    End If
                    Set rstEnc = Nothing
                
                    Call FillCell(rstAppt.Fields("fldApptID").Value, rstAppt.Fields("PatientCount").Value, _
                        lngRowSpan, rstAppt.Fields("fldCategoryID").Value, rstAppt.Fields("fldApptTitle").Value, _
                        Mid(NNs(rstAppt.Fields("fldNote").Value), 1, 25), lngApptStatusID, _
                        lngDayCell, lngTimeCell, blnRecur, arySched)
                End If
            Next
        End If
SkipIt:
        rstAppt.MoveNext
    Wend
        
    ' Clean House
    Set rstAppt = Nothing
    Set rstExc = Nothing
    Set rstEnc = Nothing
    Set objAppt = Nothing
    GetObjectContext.SetComplete
    
    Exit Sub
    
ErrTrap:
    ' Clean House
    Set rstAppt = Nothing
    Set objAppt = Nothing
    Set rstExc = Nothing
    Set rstEnc = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub PopulateDaySchedObject(ByVal lngProviderID As Long, _
                                ByVal dteDate As Date, _
                                ByVal intCol As Integer, _
                                ByRef arySched() As udtApptCell, _
                                ByVal strStartTime As String, _
                                ByVal lngInterval As Long)
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
    Dim lngUBound1 As Long, lngUBound2 As Long
    Dim lngDayCell As Long
    Dim lngTimeCell As Long
    Dim lngOffset As Long
    Dim lngRowSpan As Long
    Dim i As Long
    Dim intCtr As Integer
    Dim aryRecurDates As Variant
    Dim lngTimeBound
    Dim blnRecur As Boolean 'R001
    Dim blnException As Boolean
    
    On Error GoTo Err_Trap
    
  ' Populate the recordset
    Set objAppt = CreateObject("ApptBZ.CApptBZ")
    Set rst = objAppt.FetchByProviderDateRange(lngProviderID, dteDate, dteDate)
 '   Set rst = objAppt.FetchByProviderDateRange(lngProviderID, dteDate, DateAdd("d", -1, CDate(dteDate)))
    Set rstExc = objAppt.FetchProviderExceptions(lngProviderID, dteDate, dteDate)
    
    lngTimeBound = UBound(arySched, 2)
    
    With rst
        While Not .EOF
            lngTimeCell = GetTimeCell(.Fields("fldStartDateTime").Value, strStartTime, lngInterval)
            If lngTimeCell > lngTimeBound Then GoTo SkipIt  ' Appointment is beyond range
            
            If lngTimeCell < 0 Then
                lngRowSpan = CInt(.Fields("fldDuration").Value / lngInterval)
                If (lngTimeCell + lngRowSpan) < 0 Then
                    GoTo SkipIt
                Else
                    lngRowSpan = lngTimeCell + lngRowSpan
                    lngTimeCell = 0
                    lngRowSpan = Min(lngRowSpan, lngTimeBound - lngTimeCell + 1)
                End If
            Else
                ' Ensure span does not go beyond schedule range
                lngRowSpan = Min(CInt(.Fields("fldDuration").Value / lngInterval), lngTimeBound - lngTimeCell + 1)
            End If
            
            ' Get valid recurring dates if any
            If Trim(.Fields("fldRecurPattern").Value) > "" Then
                blnRecur = True 'R001
                aryRecurDates = objAppt.GetRecurApptDates(dteDate, dteDate, _
                                .Fields("fldStartDateTime").Value, .Fields("fldEndDateTime").Value, _
                                .Fields("fldDuration").Value, .Fields("fldRecurPattern").Value, _
                                .Fields("fldInterval").Value, .Fields("fldDOWMask").Value, _
                                .Fields("fldDOM").Value, .Fields("fldWOM").Value, _
                                .Fields("fldMOY").Value)
            Else
                blnRecur = False 'R001
                aryRecurDates = Array(.Fields("fldStartDateTime").Value)
            End If
 
          'Fill cell for every valid date
            If IsArray(aryRecurDates) Then
               For i = 0 To UBound(aryRecurDates)
                blnException = False
                'Take out recurring appointment exceptions
                If .Fields("fldRecurPattern").Value > "" Then
                    If rstExc.RecordCount > 0 Then
                        rstExc.MoveFirst
                        For intCtr = 1 To rstExc.RecordCount
                            If CDate(rstExc.Fields("fldApptDate")) = CDate(aryRecurDates(i)) Then
                                If rstExc.Fields("fldRecurApptID").Value = .Fields("fldApptID") Then
                                    blnException = True
                                    Exit For
                                End If
                            End If
                            If Not rstExc.EOF Then
                                rstExc.MoveNext
                            End If
                        Next
                    End If
                End If
                
                lngDayCell = intCol
               ' lngDayCell = GetDayCell(CDate(aryRecurDates(i)), dteDate)
                
                If blnException = False Then
                    Call FillCell(.Fields("fldApptID").Value, .Fields("PatientCount").Value, _
                        lngRowSpan, .Fields("fldCategoryID").Value, NNs(.Fields("fldApptTitle").Value), _
                        Mid(NNs(.Fields("fldNote").Value), 1, 25), .Fields("fldApptStatus").Value, _
                        lngDayCell, lngTimeCell, blnRecur, arySched)
                End If
               Next
            End If
            
SkipIt:
            .MoveNext
        Wend
    End With
    
    ' Clean House
    Set rst = Nothing
    Set objAppt = Nothing
    Set rstExc = Nothing
    GetObjectContext.SetComplete
    
    Exit Sub
    
Err_Trap:
    ' Clean House
    Set rst = Nothing
    Set objAppt = Nothing
    Set rstExc = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub FillCell(ByVal lngApptID As Long, ByVal lngPatientCount As Long, _
                     ByVal intRowSpan As Integer, ByVal lngCategoryID As Long, _
                     ByVal strApptTitle As String, ByVal strNote As String, _
                     ByVal lngApptStatusID As Long, ByVal lngDayCell As Long, _
                     ByVal lngTimeCell As Long, ByVal blnRecur As Boolean, _
                     ByRef arySched() As udtApptCell)
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
    
    Dim i As Integer
    Dim intNewRowSpan As Integer
    Dim lngNewTimeCell As Long
    Dim strApptList As String
    Dim strTextList As String
    Dim varApptList As Variant
    Dim varTextList As Variant
    Dim udtCell As udtApptCell
    Dim lngTimeBound As Long
    
    ' Look for conflicts
    intNewRowSpan = intRowSpan
    lngNewTimeCell = lngTimeCell
    
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
        udtCell = arySched(lngDayCell, lngNewTimeCell + i)

        If udtCell.ParentOffset = 0 And udtCell.Type = SCHEDULED_T Then
            StrCatDel strApptList, udtCell.ApptList
            StrCatDel strTextList, udtCell.Text, BR
            intNewRowSpan = Max(intNewRowSpan, (udtCell.RowSpan + i))
        End If
    Next

    ' Set "conflict" arrays
    If strApptList > "" Then
        varApptList = Split(strApptList, ",")
        varTextList = Split(strTextList, BR)
        'strTextList = Join(varTextList, BR)
    Else
        varApptList = Null
        varTextList = Null
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
                .Class = CONFLICT_CL
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
                .ApptList = lngApptID
                .ParentOffset = i
                .Type = SCHEDULED_T
                .RowSpan = intRowSpan - i
                If blnRecur = True Then  'R001
                    .Recur = "Y"
                Else
                    .Recur = "N"
                End If
                
                Select Case lngCategoryID
                    Case PATIENT_CAT
                        .ID = lngApptID
                        
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
                                .Class = ATTENDED_CL
                            Case CONFIRMED_ST
                                If lngPatientCount > 1 Then
                                    .Class = GROUP_CL
                                Else
                                    .Class = CONFIRMED_CL
                                End If
                            Case SCHEDULED_ST
                                If lngPatientCount > 1 Then
                                    .Class = GROUP_CL
                                Else
                                    .Class = SCHEDULED_CL
                                End If
                            Case HOLD_ST
                                .Text = "Hold"
                                .Class = HOLD_CL
                            Case NO_SHOW_ST
                                .Class = NO_SHOW_CL
                            Case TENATIVE_ST
                                .Class = TENATIVE_CL
                            Case PENDING_ST
                                .Class = PENDING_CL
                        End Select
                                         
                    Case BLOCK_CAT
                        .ID = lngApptID
                        If strNote > "" Then
                            .Text = FormatText(strNote)
                        Else
                            .Text = "Block"
                        End If
                        .Class = BLOCKED_CL
                        .Tag = "S"
                End Select
            End With
        Next
    End If

End Sub

Private Function GetDayCell(ByVal dteSeek As Date, ByVal dteStart As Date) As Long
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
    GetDayCell = CLng(DateDiff("d", dteStart, dteSeek) + 1)
    
End Function

Private Function GetTimeCell(ByVal dteSeek As Date, ByVal strStartTime, _
        ByVal lngInterval As Long) As Long
'--------------------------------------------------------------------
'Date: 11/29/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Maps a date(time) to time row on the schedule object '
'Parameters: dteSeek - date to be translated                        '
'Returns: Long                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    GetTimeCell = CLng(DateDiff("n", strStartTime, TimeValue(dteSeek)) / lngInterval)

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
    HDateTime = Replace(FormatDateTime(strDate, vbLongTime), ":00 ", " ")
End Function

Private Function FormatText(ByVal str As String) As String
'--------------------------------------------------------------------
'Date: 07/16/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Formats text that could potentially harm HTML        '
'Parameters: str - String to be formatted                           '
'Returns: formatted string                                          '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    FormatText = Replace(Replace(str, "<", "&lt;"), ">", "&gt;")
End Function





