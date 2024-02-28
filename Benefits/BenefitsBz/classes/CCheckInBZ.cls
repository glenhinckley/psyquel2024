VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheckInBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CCheckInBZ                                             '
'Date: 11/16/2000                                                   '
'Author: Chris Dereadt                                              '
'Description:  MTS business object designed to call methods         '
'              associated with the Check in classes.                '
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME = "CCheckInBZ"

'--------------------------------------------------------------------
' Public Methods    +++++++++++++++++++++++++++++++++++++++++++++++++
'--------------------------------------------------------------------

Public Function Fetch(ByVal lngID As Long, _
                      ByVal dteStartDate As Date, _
                      ByVal dteEndDate As Date) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 11/16/2000                                                   '
'Author: Chris Dereadt                                              '
'Description: Fetches all information about appointment check in    '
'Parameters:  lngID - ID of the provider / office manager           '
'             dteStartDate - Start date of the search date range    '
'             dteEnsdDate - Start End of the search date range      '
'Returns:   ADODB Recordset of Check in info                        '
'--------------------------------------------------------------------
    Dim objCheckIn As BenefitsDB.CCheckInDB
    Dim rst As ADODB.Recordset
    
    On Error GoTo Err_Handler

    ' Instantiate the appt object
    Set objCheckIn = CreateObject("BenefitsDB.CCheckInDB")
    
    ' Populate the recordset
    Set rst = objCheckIn.Fetch(lngID, dteStartDate, dteEndDate)
    
    Set Fetch = rst
    
    Set rst = Nothing
    
    Set objCheckIn = Nothing
    
    GetObjectContext.SetComplete
    Exit Function
    
Err_Handler:
    Set objCheckIn = Nothing
    Set rst = Nothing
    GetObjectContext.SetAbort

End Function

Public Function FetchAmountOwed(ByVal lngPatRPPlanID As Long, ByVal lngPatientID As Long, ByVal lngRPID As Long, _
                                ByVal lngPlanID As Long, ByVal lngProviderID As Long, ByVal strCPTCode As String, _
                                ByVal strCPTCode2 As String, ByVal lngProviderCredential As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 01/15/2001                                                   '
'Author: Morgan Travis                                              '
'Description: Fetches all information amount owed                   '
'Parameters:        '
'Returns:   ADODB Recordset of Check in info                        '
'--------------------------------------------------------------------
    Dim objCheckIn As BenefitsDB.CCheckInDB
    Dim rst As ADODB.Recordset
    
    On Error GoTo Err_Handler

    ' Instantiate the appt object
    Set objCheckIn = CreateObject("BenefitsDB.CCheckInDB")
    
    ' Populate the recordset
    Set rst = objCheckIn.FetchAmountOwed(lngPatRPPlanID, lngPatientID, lngRPID, lngPlanID, lngProviderID, strCPTCode, strCPTCode2, lngProviderCredential)
    
    Set FetchAmountOwed = rst
    
    Set rst = Nothing
    
    Set objCheckIn = Nothing
    
    GetObjectContext.SetComplete
    Exit Function
    
Err_Handler:
    Set objCheckIn = Nothing
    Set rst = Nothing
    GetObjectContext.SetAbort

End Function

Public Function FetchCopayInfo(ByVal lngPatientID As Long, ByVal lngProviderID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 02/12/2001                                                   '
'Author: Chris Dereadt                                              '
'Description:   Fetches Copay information                           '
'Parameters:  lngPatientID - Id for the patient                     '
'Returns:   ADODB Recordset of Check in info                        '
'--------------------------------------------------------------------
    Dim objCheckIn As BenefitsDB.CCheckInDB
    Dim rst As ADODB.Recordset
    
    On Error GoTo Err_Handler

    ' Instantiate the appt object
    Set objCheckIn = CreateObject("BenefitsDB.CCheckInDB")
    
    ' Populate the recordset
    Set rst = objCheckIn.FetchCopayInfo(lngPatientID, lngProviderID)
    
    Set FetchCopayInfo = rst
    
    Set rst = Nothing
    Set objCheckIn = Nothing
    
    GetObjectContext.SetComplete
    Exit Function
    
Err_Handler:
    Set objCheckIn = Nothing
    Set rst = Nothing
    GetObjectContext.SetAbort
    
    Err.Raise Err.Number, CLASS_NAME, Err.Description

End Function

Public Function FetchCopayData(ByVal lngProviderID As Long, ByVal dteStartDate As Date, _
                      ByVal dteEndDate As Date, ByVal lngPatientID As Long) As Collection
'--------------------------------------------------------------------
'Date: 03/08/2001                                                   '
'Author: Chris Dereadt                                              '
'Description:   Fetches Copay information                           '
'Parameters:  lngPatientID - Id for the patient                     '
'Returns:   ADODB Recordset of Check in info                        '
'--------------------------------------------------------------------
    Dim objCheckInData As BenefitsDB.CCheckInDB
    Dim col As Collection
    
    On Error GoTo Err_Handler

    ' Instantiate the appt object
    Set objCheckInData = CreateObject("BenefitsDB.CCheckInDB")
    
    ' Populate the recordset
    Set col = objCheckInData.FetchCopayData(lngProviderID, dteStartDate, dteEndDate, lngPatientID)
    
    Set FetchCopayData = col
    
    GetObjectContext.SetComplete
    Set objCheckInData = Nothing
    Set col = Nothing
    
    Exit Function
    
Err_Handler:
    Set objCheckInData = Nothing
    'Set col = Nothing
    GetObjectContext.SetAbort
    
    Err.Raise Err.Number, CLASS_NAME, Err.Description

End Function





