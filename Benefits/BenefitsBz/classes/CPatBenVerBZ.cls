VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatBenVerBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatBenVerBZ
'Date: 10/26/2000
'Author: Morgan Travis
'Description:  MTS object designed to host methods associated with
'               data affecting Patient Benefits Verification table.
'--------------------------------------------------------------------
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CPatBenVerBZ"


Public Function Insert(ByVal lngPatRPPlanID As Long, _
                       ByVal strUser As String, ByVal dteDateTime As Date, _
                       ByVal strContact As String, ByVal strNotes As String)
'--------------------------------------------------------------------
'Date: 10/26/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblPatientBenefitVerfication table utilizing
'              a stored procedure.
'Parameters: lngPatRPPlanID - Patient Resposible Party Plan ID number to insert into table
'            strUser - System User that will be inserted into the table.
'            dteDateTime - The date/time of the entry that will be inserted into the table.
'            strContact - The Name of the person contacted about the verification
'            strNotes - Notes about the verification
'--------------------------------------------------------------------

    Dim objPatBenVer As BenefitsDB.CPatBEnVerDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    Set objPatBenVer = CreateObject("BenefitsDB.CPatBenVerDB")
    If VerifyData(strContact, strErrMsg) Then
        Call objPatBenVer.Insert(lngPatRPPlanID, strUser, dteDateTime, strContact, strNotes)
    Else
        GoTo ErrTrap
    End If
    Set objPatBenVer = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBenVer = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError + 512 + 16, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If

End Function

Public Function FetchByPatRPPlan(ByVal lngPatRPPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/26/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblPatientBenefitVerfication table.     '
'Parameters: lngProviderID - Provider ID number to fetch from table '
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    
    Dim objPatBenVer As BenefitsDB.CPatBEnVerDB
    Dim strErrMsg As String
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objPatBenVer = CreateObject("BenefitsDB.CPatBenVerDB")
    Set rstSQL = objPatBenVer.FetchByPatRPPlan(lngPatRPPlanID)
      
    Set FetchByPatRPPlan = rstSQL
    Set objPatBenVer = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBenVer = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description

End Function



Public Function FetchPendingVerif(ByVal lngUserID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/26/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblPatientBenefitVerfication table.     '
'Parameters: lngUserID - Provider ID number to fetch from table '
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    
    Dim objPatBenVer As BenefitsDB.CPatBEnVerDB
    Dim strErrMsg As String
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objPatBenVer = CreateObject("BenefitsDB.CPatBenVerDB")
    Set rstSQL = objPatBenVer.FetchPendingVerif(lngUserID)
      
    Set FetchPendingVerif = rstSQL
    Set objPatBenVer = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBenVer = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description

End Function

Private Function VerifyData(ByVal strContact As String, ByRef strErrMsg As String) As Boolean
    
    VerifyData = False
    If strContact = "" Then
        strErrMsg = "Contact Not Valid"
        Exit Function
    End If
        
    VerifyData = True

End Function



