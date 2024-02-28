VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBenefactorDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CBenefactorDB
'Date: 3/01/2000
'Author: Eric Pena
'Description:  MTS object designed to host methods associated with
'              data affecting the tblBenefactor table.
'--------------------------------------------------------------------
' Revision History:
' R001 04/03/2000 Richkun: Removed code that auto-associates Clinics to
'           Benefactors.  This should be performed when the Encounter is
'           logged, or when an appointment is made.
' R002 07/11/2000 Richkun: Added FetchByID method
' R003 10/13/2000 Richkun: Removed code that disabled Pat-RP-Plan records
'           from Deleted() method.
' R004 11/02/2000 Richkun: Added lngProviderID to Deleted() method
' R005 01/02/2001 Nichol: Added fldOwnerID to select list in Fetch() method,
'           and to the interfaces of the Insert method
' R006 01/09/2001 Richkun: Added fldPhysFax to Insert() abd Update() methods
' R007 04/18/2001 Richkun - Added support for HomePhMsg and WorkPhMsg columns
' R008 05/30/2002 Richkun - Added FetchArchiveCandidates() method
' R009 07/26/2002 Richkun - Added FetchPatientBalanceByUser() method
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CBenefactorDB"
Private Const TABLE_NAME As String = "tblBenefactor"

Public Function Fetch(Optional blnIncludeDisabled As Boolean = False, Optional ByVal strWhere As String, Optional ByVal strOrderBy As String) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/01/2000
'Author: Eric Pena
'Description:  Retrieves records from the tblBenefactor table.
'Parameters: blnIncludeDisabled - Optional parameter that identifies if
'               records flagged as 'Disabled' or 'De-activated' are to be
'               included in the record set. The default value is False.
'            strWhere - Optional 'Where' clause of the SQL statement by
'               which records will be filtered.  If the parameter is not
'               supplied, all records will be retrieved.
'             strOrderBy - Optional 'Order By' clause in which retrieved
'               records will be sorted.  The default is set to the
'               'fldLast' and 'fldFirst' columns
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
   
    Dim cnnSQL As ADODB.Connection
    Dim rstSQL As ADODB.Recordset
    Dim strDfltWhere As String
    Dim strSQL As String

    On Error GoTo ErrTrap:

    'Prepare the SQL statement.
    strSQL = "SELECT "
    strSQL = strSQL & TABLE_NAME & ".* "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & TABLE_NAME & " "
    
    'Determine whether to include 'disabled' records'
    If blnIncludeDisabled = False Then
        strDfltWhere = " WHERE fldDisabledYN = 'N' "
    End If
    
    'Attach user's where clause if applicable
    strSQL = strSQL & strDfltWhere
    
    If Trim(strWhere > "") Then If blnIncludeDisabled Then strSQL = strSQL & " WHERE " & strWhere & " " Else strSQL = strSQL & " AND " & strWhere & " "
    
    If Trim(strOrderBy) > "" Then
        strSQL = strSQL & " ORDER BY "
        strSQL = strSQL & strOrderBy
    Else
        strSQL = strSQL & " ORDER BY "
        strSQL = strSQL & "fldLast, fldFirst "
    End If


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
    
    cnnSQL.Close
    Set cnnSQL = Nothing
    Set Fetch = rstSQL
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cnnSQL = Nothing
    Set rstSQL = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function FetchByID(ByVal lngBenefactorID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 07/11/2000
'Author: Dave Richkun
'Description:  Retrieves generic information for a specific benefactor
'Parameters: lngBenefactorID - ID of the Benefactor to retrieve information for
'Returns: Recordset
'--------------------------------------------------------------------

     Dim rst As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     On Error GoTo ErrTrap:
     
     Set rst = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     
    'Create the parameter objects
     With cmd
         .CommandText = "uspSelBenefactorByID"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@BenefactorID", adInteger, adParamInput, , lngBenefactorID)
     End With
    
    'Open the connection and execute the stored procedure
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
     rst.CursorLocation = adUseClient
     rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
     
     Set FetchByID = rst
    
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
     Set cmd = Nothing
     Set cnn = Nothing
     Set rst = Nothing
     Call RaiseError(Err, Err.Number, CLASS_NAME)
End Function

Public Function Insert(ByVal strLast As String, ByVal strMaiden As String, ByVal strMI As String, _
                    ByVal strFirst As String, ByVal strNickName As String, ByVal strAddress1 As String, ByVal strAddress2 As String, _
                    ByVal strCity As String, ByVal strState As String, ByVal strCountry As String, ByVal strZip As String, ByVal strCounty As String, _
                    ByVal strHomePhoneDesc As String, ByVal strHomePhone As String, _
                    ByVal strMobilePhoneDesc As String, ByVal strMobilePhone As String, _
                    ByVal strWorkPhoneDesc As String, ByVal strWorkPhone As String, ByVal strWorkExt As String, _
                    ByVal strEmail As String, _
                    ByVal dteDOB As Date, ByVal strSex As String, ByVal strSSN As String, ByVal strDLNum As String, ByVal varCC As Variant, _
                    ByVal lngMarital As Long, ByVal lngEthnicity As Long, ByVal lngHHIncome As Long, _
                    ByVal lngEmployment As Long, ByVal strEmployer As String, ByVal varPCP As Variant, ByVal varREF As Variant, _
                    ByVal strEmgName As String, ByVal strEmgRelat As String, ByVal strEmgPhone As String, _
                    ByVal strGuardianName As String, ByVal strGuardianRelat As String, ByVal strGuardianPhone As String, _
                    ByVal strOtherPhone1 As String, ByVal strOtherPhoneDesc1 As String, ByVal strOtherPhone2 As String, ByVal strOtherPhoneDesc2 As String, ByVal strOtherPhone3 As String, ByVal strOtherPhoneDesc3 As String, _
                    ByVal strNotes As String, ByVal strInternalID As String, ByVal lngClinicID As Long, _
                    ByVal dteAdmitDate As Date, ByVal strUserName As String, ByVal strPassword As String, ByRef varFlags As Variant, _
                    ByVal strAddedBy As String, Optional ByVal lngOwnerID As Long) As Long
'--------------------------------------------------------------------
'Date: 3/01/2000
'Author: Eric Pena
'Description:  Inserts a row into the tblBenefactor table utilizing
'              a stored procedure.
'Parameters: Each parameter identifies the column value that will be
'              inserted.
'            varPCP -
'            0) PhysLast
'            1) PhysFirst
'            2) PhysMI
'            3) strPhysTitle
'            4) strPhysician
'            5) strPhysNPI
'            6) PhysID
'            7) PhysPhone
'            8) PhysFax
'Returns: ID (Primary Key) of the row inserted.
'--------------------------------------------------------------------
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim intCount As Integer
    Dim lngID As Long
    Dim arrFlags() As Variant
    ReDim arrFlags(10)
    Dim arrPcp() As Variant
    ReDim arrPcp(10)
    Dim arrRef() As Variant
    ReDim arrRef(10)
    Dim arrCC() As Variant
    ReDim arrCC(1, 4)
    
    On Error GoTo ErrTrap:
    
    If Not IsArray(varFlags) Or IsEmpty(varFlags) Then
        arrFlags(0) = ""
        arrFlags(1) = ""
        arrFlags(2) = ""
        arrFlags(3) = ""
        arrFlags(4) = ""
        arrFlags(5) = ""
        arrFlags(6) = ""
        arrFlags(7) = ""
        arrFlags(8) = ""
    Else
        arrFlags(0) = IfNull(varFlags(0), "N")
        arrFlags(1) = IfNull(varFlags(1), "N")
        arrFlags(2) = IfNull(varFlags(2), "N")
        arrFlags(3) = IfNull(varFlags(3), "N")
        arrFlags(4) = IfNull(varFlags(4), "N")
        arrFlags(5) = IfNull(varFlags(5), "N")
        arrFlags(6) = IfNull(varFlags(6), "N")
        arrFlags(7) = IfNull(varFlags(7), "N")
        arrFlags(8) = IfNull(varFlags(8), "N")
    End If
    
    If Not IsArray(varPCP) Or IsEmpty(varPCP) Then
        arrPcp(0) = ""
        arrPcp(1) = ""
        arrPcp(2) = ""
        arrPcp(3) = ""
        arrPcp(4) = ""
        arrPcp(5) = ""
        arrPcp(6) = ""
        arrPcp(7) = ""
        arrPcp(8) = ""
    Else
        arrPcp(0) = IfNull(varPCP(0), "")
        arrPcp(1) = IfNull(varPCP(1), "")
        arrPcp(2) = IfNull(varPCP(2), "")
        arrPcp(3) = IfNull(varPCP(3), "")
        arrPcp(4) = IfNull(varPCP(4), "")
        arrPcp(5) = IfNull(varPCP(5), "")
        arrPcp(6) = IfNull(varPCP(6), "")
        arrPcp(7) = IfNull(varPCP(7), "")
        arrPcp(8) = IfNull(varPCP(8), "")
    End If
    
    If Not IsArray(varREF) Or IsEmpty(varREF) Then
        arrRef(0) = ""
        arrRef(1) = ""
        arrRef(2) = ""
        arrRef(3) = ""
        arrRef(4) = ""
        arrRef(5) = ""
        arrRef(6) = "2"
        arrRef(7) = 0
        arrRef(8) = ""
    Else
        arrRef(0) = IfNull(varREF(0), "")
        arrRef(1) = IfNull(varREF(1), "")
        arrRef(2) = IfNull(varREF(2), "")
        arrRef(3) = IfNull(varREF(3), "")
        arrRef(4) = IfNull(varREF(4), "")
        arrRef(5) = IfNull(varREF(5), "")
        arrRef(6) = IfNull(varREF(6), "2")
        arrRef(7) = IfNull(varREF(7), 0)
        arrRef(8) = IfNull(varREF(8), "")
    End If
        
'       arrCC(1, 3) Credit Card Info
    If Not IsArray(varCC) Or IsEmpty(varCC) Then
        arrCC(0, 0) = ""
        arrCC(0, 1) = ""
        arrCC(0, 2) = 0
        arrCC(0, 3) = 0
        arrCC(0, 4) = ""
        arrCC(1, 0) = ""
        arrCC(1, 1) = ""
        arrCC(1, 2) = 0
        arrCC(1, 3) = 0
        arrCC(1, 4) = ""
    Else
        arrCC(0, 0) = IfNull(varCC(0, 0), "")
        arrCC(0, 1) = IfNull(varCC(0, 1), "")
        arrCC(0, 2) = IIf(Not IsNumeric(IfNull(varCC(0, 2), 0)), 0, IfNull(varCC(0, 2), 0))
        arrCC(0, 3) = IIf(Not IsNumeric(IfNull(varCC(0, 3), 0)), 0, IfNull(varCC(0, 3), 0))
        arrCC(0, 4) = IfNull(varCC(0, 4), "")
        arrCC(1, 0) = IfNull(varCC(1, 0), "")
        arrCC(1, 1) = IfNull(varCC(1, 1), "")
        arrCC(1, 2) = IIf(Not IsNumeric(IfNull(varCC(1, 2), 0)), 0, IfNull(varCC(1, 2), 0))
        arrCC(1, 3) = IIf(Not IsNumeric(IfNull(varCC(1, 3), 0)), 0, IfNull(varCC(1, 3), 0))
        arrCC(1, 4) = IfNull(varCC(1, 4), "")
    End If

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    'Create the parameter objects
    With cmdSQL
        .CommandText = "uspInsBenefactor"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@BenefactorID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@OwnerID", adInteger, adParamInput, , IIf(lngOwnerID <= 0, Null, lngOwnerID))
        .Parameters.Append .CreateParameter("@Last", adVarChar, adParamInput, 20, strLast)
        .Parameters.Append .CreateParameter("@Maiden", adVarChar, adParamInput, 20, strMaiden)
        .Parameters.Append .CreateParameter("@MI", adVarChar, adParamInput, 1, strMI)
        .Parameters.Append .CreateParameter("@First", adVarChar, adParamInput, 20, strFirst)
        .Parameters.Append .CreateParameter("@NickName", adVarChar, adParamInput, 20, strNickName)
        .Parameters.Append .CreateParameter("@Address1", adVarChar, adParamInput, 60, strAddress1)
        .Parameters.Append .CreateParameter("@Address2", adVarChar, adParamInput, 60, strAddress2)
        .Parameters.Append .CreateParameter("@City", adVarChar, adParamInput, 40, strCity)
        .Parameters.Append .CreateParameter("@County", adVarChar, adParamInput, 40, strCounty)
        .Parameters.Append .CreateParameter("@State", adVarChar, adParamInput, 2, strState)
        .Parameters.Append .CreateParameter("@Country", adVarChar, adParamInput, 3, strCountry)
        .Parameters.Append .CreateParameter("@Zip", adVarChar, adParamInput, 11, strZip)
        .Parameters.Append .CreateParameter("@HomePhoneDesc", adVarChar, adParamInput, 20, strHomePhoneDesc)
        .Parameters.Append .CreateParameter("@HomePhone", adVarChar, adParamInput, 15, NumbersOnly(strHomePhone))
        .Parameters.Append .CreateParameter("@HomePhMsgYN", adVarChar, adParamInput, 1, arrFlags(0))
        .Parameters.Append .CreateParameter("@MobilePhoneDesc", adVarChar, adParamInput, 20, strMobilePhoneDesc)
        .Parameters.Append .CreateParameter("@MobilePhone", adVarChar, adParamInput, 15, NumbersOnly(strMobilePhone))
        .Parameters.Append .CreateParameter("@MobilePhMsgYN", adVarChar, adParamInput, 1, arrFlags(1))
        .Parameters.Append .CreateParameter("@TextYN", adVarChar, adParamInput, 1, arrFlags(2))
        .Parameters.Append .CreateParameter("@WorkPhoneDesc", adVarChar, adParamInput, 20, strWorkPhoneDesc)
        .Parameters.Append .CreateParameter("@WorkPhone", adVarChar, adParamInput, 15, NumbersOnly(strWorkPhone))
        .Parameters.Append .CreateParameter("@WorkExt", adVarChar, adParamInput, 6, strWorkExt)
        .Parameters.Append .CreateParameter("@WorkPhMsgYN", adVarChar, adParamInput, 1, arrFlags(3))
        .Parameters.Append .CreateParameter("@Email", adVarChar, adParamInput, 60, strEmail)
        .Parameters.Append .CreateParameter("@EmailYN", adVarChar, adParamInput, 1, arrFlags(4))
        .Parameters.Append .CreateParameter("@PrintPatientStmtYN", adVarChar, adParamInput, 1, arrFlags(5))
        .Parameters.Append .CreateParameter("@ApptReminderYN", adVarChar, adParamInput, 1, arrFlags(6))
        .Parameters.Append .CreateParameter("@NoteTrackingYN", adVarChar, adParamInput, 1, arrFlags(7))
        .Parameters.Append .CreateParameter("@DOB", adDBTimeStamp, adParamInput, , IIf(dteDOB = #12:00:00 AM#, Null, dteDOB))
        .Parameters.Append .CreateParameter("@Sex", adVarChar, adParamInput, 1, IIf(strSex = "", Null, strSex))
        .Parameters.Append .CreateParameter("@SSN", adVarChar, adParamInput, 12, strSSN)
        .Parameters.Append .CreateParameter("@DLNum", adVarChar, adParamInput, 20, strDLNum)
        .Parameters.Append .CreateParameter("@PriCCardNbr", adVarChar, adParamInput, 50, arrCC(0, 0))
        .Parameters.Append .CreateParameter("@PriCCardSecurity", adVarChar, adParamInput, 5, arrCC(0, 1))
        .Parameters.Append .CreateParameter("@PriCCardMth", adInteger, adParamInput, , IIf(arrCC(0, 2) <= 0, Null, arrCC(0, 2)))
        .Parameters.Append .CreateParameter("@PriCCardYr", adInteger, adParamInput, , IIf(arrCC(0, 3) <= 0, Null, arrCC(0, 3)))
        .Parameters.Append .CreateParameter("@PriCCVaultID", adVarChar, adParamInput, 50, arrCC(0, 4))
        .Parameters.Append .CreateParameter("@SecCCardNbr", adVarChar, adParamInput, 50, arrCC(1, 0))
        .Parameters.Append .CreateParameter("@SecCCardSecurity", adVarChar, adParamInput, 5, arrCC(1, 1))
        .Parameters.Append .CreateParameter("@SecCCardMth", adInteger, adParamInput, , IIf(arrCC(1, 2) <= 0, Null, arrCC(1, 2)))
        .Parameters.Append .CreateParameter("@SecCCardYr", adInteger, adParamInput, , IIf(arrCC(1, 3) <= 0, Null, arrCC(1, 3)))
        .Parameters.Append .CreateParameter("@SecCCVaultID", adVarChar, adParamInput, 50, arrCC(1, 4))
        .Parameters.Append .CreateParameter("@MaritalStatusID", adInteger, adParamInput, , IIf(lngMarital <= 0, Null, lngMarital))
        .Parameters.Append .CreateParameter("@EthnicityID", adInteger, adParamInput, , IIf(lngEthnicity <= 0, Null, lngEthnicity))
        .Parameters.Append .CreateParameter("@EmploymentID", adInteger, adParamInput, , IIf(lngEmployment <= 0, Null, lngEmployment))
        .Parameters.Append .CreateParameter("@HHIncomeID", adInteger, adParamInput, , IIf(lngHHIncome <= 0, 1, lngHHIncome))
        .Parameters.Append .CreateParameter("@Employer", adVarChar, adParamInput, 35, strEmployer)
        .Parameters.Append .CreateParameter("@PhyLast", adVarChar, adParamInput, 20, arrPcp(0))
        .Parameters.Append .CreateParameter("@PhyFirst", adVarChar, adParamInput, 20, arrPcp(1))
        .Parameters.Append .CreateParameter("@PhyMI", adVarChar, adParamInput, 1, arrPcp(2))
        .Parameters.Append .CreateParameter("@PhyTitle", adVarChar, adParamInput, 10, arrPcp(3))
        .Parameters.Append .CreateParameter("@PhyName", adVarChar, adParamInput, 30, arrPcp(4))
        .Parameters.Append .CreateParameter("@PhyNPI", adVarChar, adParamInput, 10, arrPcp(5))
        .Parameters.Append .CreateParameter("@PhyID", adVarChar, adParamInput, 20, arrPcp(6))
        .Parameters.Append .CreateParameter("@PhyPhone", adVarChar, adParamInput, 15, NumbersOnly(arrPcp(7)))
        .Parameters.Append .CreateParameter("@PhyFax", adVarChar, adParamInput, 15, NumbersOnly(arrPcp(8)))
        .Parameters.Append .CreateParameter("@ReferLast", adVarChar, adParamInput, 20, arrRef(0))
        .Parameters.Append .CreateParameter("@ReferFirst", adVarChar, adParamInput, 20, arrRef(1))
        .Parameters.Append .CreateParameter("@ReferMI", adVarChar, adParamInput, 1, arrRef(2))
        .Parameters.Append .CreateParameter("@ReferTitle", adVarChar, adParamInput, 10, arrRef(3))
        .Parameters.Append .CreateParameter("@ReferID", adInteger, adParamInput, , IIf(arrRef(7) <= 0, Null, arrRef(7)))
        .Parameters.Append .CreateParameter("@ReferBy", adVarChar, adParamInput, 200, arrRef(4))
        .Parameters.Append .CreateParameter("@ReferNPI", adVarChar, adParamInput, 10, arrRef(5))
        .Parameters.Append .CreateParameter("@ReferType", adVarChar, adParamInput, 1, arrRef(6))
        .Parameters.Append .CreateParameter("@EmgName", adVarChar, adParamInput, 30, strEmgName)
        .Parameters.Append .CreateParameter("@EmgRelat", adVarChar, adParamInput, 20, strEmgRelat)
        .Parameters.Append .CreateParameter("@EmgPhone", adVarChar, adParamInput, 15, NumbersOnly(strEmgPhone))
        .Parameters.Append .CreateParameter("@GuardianName", adVarChar, adParamInput, 60, strGuardianName)
        .Parameters.Append .CreateParameter("@GuardianRelat", adVarChar, adParamInput, 20, strGuardianRelat)
        .Parameters.Append .CreateParameter("@GuardianPhone", adVarChar, adParamInput, 15, NumbersOnly(strGuardianPhone))
        .Parameters.Append .CreateParameter("@OtherPhone1", adVarChar, adParamInput, 15, NumbersOnly(strOtherPhone1))
        .Parameters.Append .CreateParameter("@OtherPhoneDesc1", adVarChar, adParamInput, 50, strOtherPhoneDesc1)
        .Parameters.Append .CreateParameter("@OtherPhone2", adVarChar, adParamInput, 15, NumbersOnly(strOtherPhone2))
        .Parameters.Append .CreateParameter("@OtherPhoneDesc2", adVarChar, adParamInput, 50, strOtherPhoneDesc2)
        .Parameters.Append .CreateParameter("@OtherPhone3", adVarChar, adParamInput, 15, NumbersOnly(strOtherPhone3))
        .Parameters.Append .CreateParameter("@OtherPhoneDesc3", adVarChar, adParamInput, 50, strOtherPhoneDesc3)
        .Parameters.Append .CreateParameter("@Notes", adVarChar, adParamInput, 2500, strNotes)
        .Parameters.Append .CreateParameter("@InternalID", adVarChar, adParamInput, 25, strInternalID)
        .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , IIf(lngClinicID <= 0, Null, lngClinicID))
        .Parameters.Append .CreateParameter("@AdmitDate", adDBTimeStamp, adParamInput, , IIf(dteAdmitDate = #12:00:00 AM#, Null, dteAdmitDate))
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, 20, strPassword)
        .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 30, strAddedBy)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute ', , adExecuteNoRecords

    lngID = cmdSQL.Parameters("@BenefactorID").Value
    
    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
        
    'Close the connection
    cnnSQL.Close
    
    'free all resources
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Insert = lngID
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    GetObjectContext.SetAbort
    Call RaiseError(Err, cmdSQL.Parameters("@SQLErrorNum").Value, CLASS_NAME)
    
End Function

Public Sub Update(ByVal lngID As Long, ByVal strLast As String, ByVal strMaiden As String, ByVal strMI As String, _
                    ByVal strFirst As String, ByVal strNickName As String, ByVal strAddress1 As String, ByVal strAddress2 As String, _
                    ByVal strCity As String, ByVal strState As String, ByVal strCountry As String, ByVal strZip As String, ByVal strCounty As String, _
                    ByVal strHomePhoneDesc As String, ByVal strHomePhone As String, _
                    ByVal strMobilePhoneDesc As String, ByVal strMobilePhone As String, _
                    ByVal strWorkPhoneDesc As String, ByVal strWorkPhone As String, ByVal strWorkExt As String, _
                    ByVal strEmail As String, _
                    ByVal dteDOB As Date, ByVal strSex As String, ByVal strSSN As String, ByVal strDLNum As String, ByVal varCC As Variant, _
                    ByVal lngMarital As Long, ByVal lngEthnicity As Long, ByVal lngHHIncome As Long, _
                    ByVal lngEmployment As Long, ByVal strEmployer As String, ByVal varPCP As Variant, ByVal varREF As Variant, _
                    ByVal strEmgName As String, ByVal strEmgRelat As String, ByVal strEmgPhone As String, _
                    ByVal strGuardianName As String, ByVal strGuardianRelat As String, ByVal strGuardianPhone As String, _
                    ByVal strOtherPhone1 As String, ByVal strOtherPhoneDesc1 As String, _
                    ByVal strOtherPhone2 As String, ByVal strOtherPhoneDesc2 As String, _
                    ByVal strOtherPhone3 As String, ByVal strOtherPhoneDesc3 As String, _
                    ByVal strNotes As String, ByVal strInternalID As String, ByVal lngClinicID As Long, _
                    ByVal dteAdmitDate As Date, ByVal strUserName As String, ByVal strPassword As String, _
                    ByRef varFlags As Variant, ByVal strUpdatedBy)
'--------------------------------------------------------------------
'Date: 3/01/2000
'Author: Eric Pena
'Description:  Updates a row in the tblBenefactor table utilizing
'              a stored procedure.
'Parameters: Each parameter identifies the column value that will be
'              inserted.
'Returns: Null
'--------------------------------------------------------------------
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    
    Dim arrFlags() As Variant
    ReDim arrFlags(10)
    Dim arrPcp() As Variant
    ReDim arrPcp(10)
    Dim arrRef() As Variant
    ReDim arrRef(10)
    Dim arrCC() As Variant
    ReDim arrCC(1, 4)
    
    On Error GoTo ErrTrap:
    
    If Not IsArray(varFlags) Or IsEmpty(varFlags) Then
        arrFlags(0) = ""
        arrFlags(1) = ""
        arrFlags(2) = ""
        arrFlags(3) = ""
        arrFlags(4) = ""
        arrFlags(5) = ""
        arrFlags(6) = ""
        arrFlags(7) = ""
        arrFlags(8) = ""
    Else
        arrFlags(0) = IfNull(varFlags(0), "N")
        arrFlags(1) = IfNull(varFlags(1), "N")
        arrFlags(2) = IfNull(varFlags(2), "N")
        arrFlags(3) = IfNull(varFlags(3), "N")
        arrFlags(4) = IfNull(varFlags(4), "N")
        arrFlags(5) = IfNull(varFlags(5), "N")
        arrFlags(6) = IfNull(varFlags(6), "N")
        arrFlags(7) = IfNull(varFlags(7), "N")
        arrFlags(8) = IfNull(varFlags(8), "N")
    End If
    
    If Not IsArray(varPCP) Or IsEmpty(varPCP) Then
        arrPcp(0) = ""
        arrPcp(1) = ""
        arrPcp(2) = ""
        arrPcp(3) = ""
        arrPcp(4) = ""
        arrPcp(5) = ""
        arrPcp(6) = ""
        arrPcp(7) = ""
        arrPcp(8) = ""
    Else
        arrPcp(0) = IfNull(varPCP(0), "")
        arrPcp(1) = IfNull(varPCP(1), "")
        arrPcp(2) = IfNull(varPCP(2), "")
        arrPcp(3) = IfNull(varPCP(3), "")
        arrPcp(4) = IfNull(varPCP(4), "")
        arrPcp(5) = IfNull(varPCP(5), "")
        arrPcp(6) = IfNull(varPCP(6), "")
        arrPcp(7) = IfNull(varPCP(7), "")
        arrPcp(8) = IfNull(varPCP(8), "")
    End If
    
    If Not IsArray(varREF) Or IsEmpty(varREF) Then
        arrRef(0) = ""
        arrRef(1) = ""
        arrRef(2) = ""
        arrRef(3) = ""
        arrRef(4) = ""
        arrRef(5) = ""
        arrRef(6) = "2"
        arrRef(7) = 0
        arrRef(8) = ""
    Else
        arrRef(0) = IfNull(varREF(0), "")
        arrRef(1) = IfNull(varREF(1), "")
        arrRef(2) = IfNull(varREF(2), "")
        arrRef(3) = IfNull(varREF(3), "")
        arrRef(4) = IfNull(varREF(4), "")
        arrRef(5) = IfNull(varREF(5), "")
        arrRef(6) = IfNull(varREF(6), "2")
        arrRef(7) = IfNull(varREF(7), 0)
        arrRef(8) = IfNull(varREF(8), "")
    End If
    
'       arrCC(1, 4) Credit Card Info
    If Not IsArray(varCC) Or IsEmpty(varCC) Then
        arrCC(0, 0) = ""
        arrCC(0, 1) = ""
        arrCC(0, 2) = 0
        arrCC(0, 3) = 0
        arrCC(0, 4) = ""
        arrCC(1, 0) = ""
        arrCC(1, 1) = ""
        arrCC(1, 2) = 0
        arrCC(1, 3) = 0
        arrCC(1, 4) = ""
    Else
        arrCC(0, 0) = IfNull(varCC(0, 0), "")
        arrCC(0, 1) = IfNull(varCC(0, 1), "")
        arrCC(0, 2) = IIf(Not IsNumeric(IfNull(varCC(0, 2), 0)), 0, IfNull(varCC(0, 2), 0))
        arrCC(0, 3) = IIf(Not IsNumeric(IfNull(varCC(0, 3), 0)), 0, IfNull(varCC(0, 3), 0))
        arrCC(0, 4) = IfNull(varCC(0, 4), "")
        arrCC(1, 0) = IfNull(varCC(1, 0), "")
        arrCC(1, 1) = IfNull(varCC(1, 1), "")
        arrCC(1, 2) = IIf(Not IsNumeric(IfNull(varCC(1, 2), 0)), 0, IfNull(varCC(1, 2), 0))
        arrCC(1, 3) = IIf(Not IsNumeric(IfNull(varCC(1, 3), 0)), 0, IfNull(varCC(1, 3), 0))
        arrCC(1, 4) = IfNull(varCC(1, 4), "")
    End If
      
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    With cmdSQL
        .CommandText = "uspUpdBenefactor"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, , lngID)
        .Parameters.Append .CreateParameter("@Last", adVarChar, adParamInput, 20, strLast)
        .Parameters.Append .CreateParameter("@Maiden", adVarChar, adParamInput, 20, strMaiden)
        .Parameters.Append .CreateParameter("@MI", adVarChar, adParamInput, 1, strMI)
        .Parameters.Append .CreateParameter("@First", adVarChar, adParamInput, 20, strFirst)
        .Parameters.Append .CreateParameter("@NickName", adVarChar, adParamInput, 20, strNickName)
        .Parameters.Append .CreateParameter("@Address1", adVarChar, adParamInput, 60, strAddress1)
        .Parameters.Append .CreateParameter("@Address2", adVarChar, adParamInput, 60, strAddress2)
        .Parameters.Append .CreateParameter("@City", adVarChar, adParamInput, 40, strCity)
        .Parameters.Append .CreateParameter("@County", adVarChar, adParamInput, 40, strCounty)
        .Parameters.Append .CreateParameter("@State", adVarChar, adParamInput, 2, strState)
        .Parameters.Append .CreateParameter("@Country", adVarChar, adParamInput, 3, strCountry)
        .Parameters.Append .CreateParameter("@Zip", adVarChar, adParamInput, 11, strZip)
        .Parameters.Append .CreateParameter("@HomePhoneDesc", adVarChar, adParamInput, 20, strHomePhoneDesc)
        .Parameters.Append .CreateParameter("@HomePhone", adVarChar, adParamInput, 15, NumbersOnly(strHomePhone))
        .Parameters.Append .CreateParameter("@HomePhMsgYN", adVarChar, adParamInput, 1, arrFlags(0))
        .Parameters.Append .CreateParameter("@MobilePhoneDesc", adVarChar, adParamInput, 20, strMobilePhoneDesc)
        .Parameters.Append .CreateParameter("@MobilePhone", adVarChar, adParamInput, 15, NumbersOnly(strMobilePhone))
        .Parameters.Append .CreateParameter("@MobilePhMsgYN", adVarChar, adParamInput, 1, arrFlags(1))
        .Parameters.Append .CreateParameter("@TextYN", adVarChar, adParamInput, 1, arrFlags(2))
        .Parameters.Append .CreateParameter("@WorkPhoneDesc", adVarChar, adParamInput, 20, strWorkPhoneDesc)
        .Parameters.Append .CreateParameter("@WorkPhone", adVarChar, adParamInput, 15, NumbersOnly(strWorkPhone))
        .Parameters.Append .CreateParameter("@WorkExt", adVarChar, adParamInput, 6, strWorkExt)
        .Parameters.Append .CreateParameter("@WorkPhMsgYN", adVarChar, adParamInput, 1, arrFlags(3))
        .Parameters.Append .CreateParameter("@Email", adVarChar, adParamInput, 60, strEmail)
        .Parameters.Append .CreateParameter("@EmailYN", adVarChar, adParamInput, 1, arrFlags(4))
        .Parameters.Append .CreateParameter("@PrintPatientStmtYN", adVarChar, adParamInput, 1, arrFlags(5))
        .Parameters.Append .CreateParameter("@ApptReminderYN", adVarChar, adParamInput, 1, arrFlags(6))
        .Parameters.Append .CreateParameter("@NoteTrackingYN", adVarChar, adParamInput, 1, arrFlags(7))
        .Parameters.Append .CreateParameter("@DOB", adDBTimeStamp, adParamInput, , IIf(dteDOB = #12:00:00 AM#, Null, dteDOB))
        .Parameters.Append .CreateParameter("@Sex", adVarChar, adParamInput, 1, IIf(strSex = "", Null, strSex))
        .Parameters.Append .CreateParameter("@SSN", adVarChar, adParamInput, 12, strSSN)
        .Parameters.Append .CreateParameter("@DLNum", adVarChar, adParamInput, 20, strDLNum)
        .Parameters.Append .CreateParameter("@PriCCardNbr", adVarChar, adParamInput, 50, arrCC(0, 0))
        .Parameters.Append .CreateParameter("@PriCCardSecurity", adVarChar, adParamInput, 5, arrCC(0, 1))
        .Parameters.Append .CreateParameter("@PriCCardMth", adInteger, adParamInput, , IIf(arrCC(0, 2) <= 0, Null, arrCC(0, 2)))
        .Parameters.Append .CreateParameter("@PriCCardYr", adInteger, adParamInput, , IIf(arrCC(0, 3) <= 0, Null, arrCC(0, 3)))
        .Parameters.Append .CreateParameter("@PriCCVaultID", adVarChar, adParamInput, 50, arrCC(0, 4))
        .Parameters.Append .CreateParameter("@SecCCardNbr", adVarChar, adParamInput, 50, arrCC(1, 0))
        .Parameters.Append .CreateParameter("@SecCCardSecurity", adVarChar, adParamInput, 5, arrCC(1, 1))
        .Parameters.Append .CreateParameter("@SecCCardMth", adInteger, adParamInput, , IIf(arrCC(1, 2) <= 0, Null, arrCC(1, 2)))
        .Parameters.Append .CreateParameter("@SecCCardYr", adInteger, adParamInput, , IIf(arrCC(1, 3) <= 0, Null, arrCC(1, 3)))
        .Parameters.Append .CreateParameter("@SecCCVaultID", adVarChar, adParamInput, 50, arrCC(1, 4))
        .Parameters.Append .CreateParameter("@MaritalStatusID", adInteger, adParamInput, , IIf(lngMarital <= 0, Null, lngMarital))
        .Parameters.Append .CreateParameter("@EthnicityID", adInteger, adParamInput, , IIf(lngEthnicity <= 0, Null, lngEthnicity))
        .Parameters.Append .CreateParameter("@EmploymentID", adInteger, adParamInput, , IIf(lngEmployment <= 0, Null, lngEmployment))
        .Parameters.Append .CreateParameter("@HHIncomeID", adInteger, adParamInput, , IIf(lngHHIncome <= 0, 1, lngHHIncome))
        .Parameters.Append .CreateParameter("@Employer", adVarChar, adParamInput, 35, strEmployer)
        .Parameters.Append .CreateParameter("@PhyLast", adVarChar, adParamInput, 20, arrPcp(0))
        .Parameters.Append .CreateParameter("@PhyFirst", adVarChar, adParamInput, 20, arrPcp(1))
        .Parameters.Append .CreateParameter("@PhyMI", adVarChar, adParamInput, 1, arrPcp(2))
        .Parameters.Append .CreateParameter("@PhyTitle", adVarChar, adParamInput, 10, arrPcp(3))
        .Parameters.Append .CreateParameter("@PhyName", adVarChar, adParamInput, 30, arrPcp(4))
        .Parameters.Append .CreateParameter("@PhyNPI", adVarChar, adParamInput, 10, arrPcp(5))
        .Parameters.Append .CreateParameter("@PhyID", adVarChar, adParamInput, 20, arrPcp(6))
        .Parameters.Append .CreateParameter("@PhyPhone", adVarChar, adParamInput, 15, NumbersOnly(arrPcp(7)))
        .Parameters.Append .CreateParameter("@PhyFax", adVarChar, adParamInput, 15, NumbersOnly(arrPcp(8)))
        .Parameters.Append .CreateParameter("@ReferLast", adVarChar, adParamInput, 20, arrRef(0))
        .Parameters.Append .CreateParameter("@ReferFirst", adVarChar, adParamInput, 20, arrRef(1))
        .Parameters.Append .CreateParameter("@ReferMI", adVarChar, adParamInput, 1, arrRef(2))
        .Parameters.Append .CreateParameter("@ReferTitle", adVarChar, adParamInput, 10, arrRef(3))
        .Parameters.Append .CreateParameter("@ReferID", adInteger, adParamInput, , IIf(arrRef(7) <= 0, Null, arrRef(7)))
        .Parameters.Append .CreateParameter("@ReferBy", adVarChar, adParamInput, 200, arrRef(4))
        .Parameters.Append .CreateParameter("@ReferNPI", adVarChar, adParamInput, 10, arrRef(5))
        .Parameters.Append .CreateParameter("@ReferType", adVarChar, adParamInput, 1, arrRef(6))
        .Parameters.Append .CreateParameter("@EmgName", adVarChar, adParamInput, 30, strEmgName)
        .Parameters.Append .CreateParameter("@EmgRelat", adVarChar, adParamInput, 20, strEmgRelat)
        .Parameters.Append .CreateParameter("@EmgPhone", adVarChar, adParamInput, 15, NumbersOnly(strEmgPhone))
        .Parameters.Append .CreateParameter("@GuardianName", adVarChar, adParamInput, 60, strGuardianName)
        .Parameters.Append .CreateParameter("@GuardianRelat", adVarChar, adParamInput, 20, strGuardianRelat)
        .Parameters.Append .CreateParameter("@GuardianPhone", adVarChar, adParamInput, 15, NumbersOnly(strGuardianPhone))
        .Parameters.Append .CreateParameter("@OtherPhone1", adVarChar, adParamInput, 15, NumbersOnly(strOtherPhone1))
        .Parameters.Append .CreateParameter("@OtherPhoneDesc1", adVarChar, adParamInput, 50, strOtherPhoneDesc1)
        .Parameters.Append .CreateParameter("@OtherPhone2", adVarChar, adParamInput, 15, NumbersOnly(strOtherPhone2))
        .Parameters.Append .CreateParameter("@OtherPhoneDesc2", adVarChar, adParamInput, 50, strOtherPhoneDesc2)
        .Parameters.Append .CreateParameter("@OtherPhone3", adVarChar, adParamInput, 15, NumbersOnly(strOtherPhone3))
        .Parameters.Append .CreateParameter("@OtherPhoneDesc3", adVarChar, adParamInput, 50, strOtherPhoneDesc3)
        .Parameters.Append .CreateParameter("@Notes", adVarChar, adParamInput, 2500, strNotes)
        .Parameters.Append .CreateParameter("@InternalID", adVarChar, adParamInput, 25, strInternalID)
        .Parameters.Append .CreateParameter("@ClinicID", adInteger, adParamInput, , IIf(lngClinicID <= 0, Null, lngClinicID))
        .Parameters.Append .CreateParameter("@AdmitDate", adDBTimeStamp, adParamInput, , IIf(dteAdmitDate = #12:00:00 AM#, Null, dteAdmitDate))
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, 20, strPassword)
        .Parameters.Append .CreateParameter("@UpdatedBy", adVarChar, adParamInput, 30, strUpdatedBy)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute ', , adExecuteNoRecords
    
    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
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
    Call RaiseError(Err, cmdSQL.Parameters("@SQLErrorNum").Value, CLASS_NAME, cmdSQL.Parameters("@ErrorMsg").Value)
    
End Sub

Public Sub Deleted(ByVal blnDeleted As Boolean, ByVal lngPatientID As Long, _
                   ByVal lngProviderID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 3/01/2000
'Author: Eric Pena
'Description:  Flags a row in the tblBenefactor table marking the row as
'              deleted or undeleted.
'--------------------------------------------------------------------
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngPatientID - ID of the row in the Benefactor table whose
'               value will be  updated.
'            lngProviderID - ID of the Provider who wishes the patient to be
'               archived/restored.  This parameter is needed because patients
'               may be associated with more than one provider.  'R004
'            strUserName - Login name of the user responsible for
'               marking the row as deleted.
'Returns: Null
'--------------------------------------------------------------------
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim strSQL As String
    Dim strSQL2 As String
    Dim strSQL3 As String
    Dim blnFlag As String

    On Error GoTo ErrTrap:

    If blnDeleted = True Then
        blnFlag = "Y"
        'Prepare the SQL statement
        strSQL = "UPDATE "
        strSQL = strSQL & "tblPatientProvider"
        strSQL = strSQL & " SET "
        strSQL = strSQL & " fldDisabledYN = '" & blnFlag & "', "
        strSQL = strSQL & " fldDateDisabled = '" & Now & "', "
        strSQL = strSQL & " fldDisabledBy = '" & strUserName & "' "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "fldPatientID = " & lngPatientID & " AND "
        strSQL = strSQL & "fldProviderID = " & lngProviderID
        
        'Disable Benefactor-clinic associations
        strSQL2 = "UPDATE tblBenefactorClinic SET "
        strSQL2 = strSQL2 & " fldDisabledYN = '" & blnFlag & "', "
        strSQL2 = strSQL2 & " fldDateDisabled = '" & Now & "', "
        strSQL2 = strSQL2 & " fldDisabledBy = '" & strUserName & "' "
        strSQL2 = strSQL2 & "WHERE "
        strSQL2 = strSQL2 & "fldBenefactorID = " & lngPatientID
        
        'Commented out following lines: DR - R003
        'Disable Benefactor-rp-plan associations
'        strSQL3 = "UPDATE tblPatRPPlan SET "
'        strSQL3 = strSQL3 & " fldDisabledYN = '" & blnFlag & "', "
'        strSQL3 = strSQL3 & " fldDateDisabled = '" & Now & "', "
'        strSQL3 = strSQL3 & " fldDisabledBy = '" & strUserName & "' "
'        strSQL3 = strSQL3 & "WHERE "
'        strSQL3 = strSQL3 & "fldPatientID = " & lngID & " OR "
'        strSQL3 = strSQL3 & "fldRPID = " & lngID
    Else
        blnFlag = "N"
        
        'Prepare the SQL statement
        strSQL = "UPDATE "
        strSQL = strSQL & "tblPatientProvider"
        strSQL = strSQL & " SET "
        strSQL = strSQL & " fldDisabledYN = '" & blnFlag & "' "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "fldPatientID = " & lngPatientID & " AND "
        strSQL = strSQL & "fldProviderID = " & lngProviderID
        
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

    If blnDeleted = True Then
        cmdSQL.CommandText = strSQL2
        cmdSQL.Execute , , adExecuteNoRecords
        'cmdSQL.CommandText = strSQL3  - Removed R003
        'cmdSQL.Execute , , adExecuteNoRecords - Removed R003
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
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function Search(ByVal strLast As String, ByVal strMaiden As String, ByVal strMI As String, ByVal strFirst As String, _
                    ByVal strNickName As String, ByVal strAddress1 As String, ByVal strAddress2 As String, ByVal strCity As String, _
                    ByVal strState As String, ByVal strZip As String, ByVal strHomePhone As String, ByVal strWorkPhone As String, _
                    ByVal strWorkExt As String, ByVal strEmail As String, ByVal strSex As String, ByVal strSSN As String, ByVal strDLNum As String, _
                    ByVal strEmployer As String, ByVal strPhysician As String, ByVal strPsyPhone As String, ByVal strEmgName As String, _
                    ByVal strEmgRelat As String, ByVal strEmgPhone As String, ByVal strNotes As String) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/01/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure based by parameters inputted.
'Parameters: Each parameter identifies the column value that will be searched
'Returns: ADORst
'--------------------------------------------------------------------
    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    
    On Error GoTo ErrTrap:
    
    Set rs = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
       
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelBenefactor_UDF"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@Last", adVarChar, adParamInput, 20, strLast)
        .Parameters.Append .CreateParameter("@Maiden", adVarChar, adParamInput, 20, strMaiden)
        .Parameters.Append .CreateParameter("@MI", adVarChar, adParamInput, 1, strMI)
        .Parameters.Append .CreateParameter("@First", adVarChar, adParamInput, 20, strFirst)
        .Parameters.Append .CreateParameter("@NickName", adVarChar, adParamInput, 20, strNickName)
        .Parameters.Append .CreateParameter("@Address1", adVarChar, adParamInput, 60, strAddress1)
        .Parameters.Append .CreateParameter("@Address2", adVarChar, adParamInput, 60, strAddress2)
        .Parameters.Append .CreateParameter("@City", adVarChar, adParamInput, 40, strCity)
        .Parameters.Append .CreateParameter("@State", adVarChar, adParamInput, 2, strState)
        .Parameters.Append .CreateParameter("@Zip", adVarChar, adParamInput, 12, strZip)
        .Parameters.Append .CreateParameter("@HomePhone", adVarChar, adParamInput, 15, strHomePhone)
        .Parameters.Append .CreateParameter("@WorkPhone", adVarChar, adParamInput, 15, strWorkPhone)
        .Parameters.Append .CreateParameter("@WorkExt", adVarChar, adParamInput, 6, strWorkExt)
        .Parameters.Append .CreateParameter("@Email", adVarChar, adParamInput, 60, strEmail)
        .Parameters.Append .CreateParameter("@Sex", adVarChar, adParamInput, 1, strSex)
        .Parameters.Append .CreateParameter("@SSN", adVarChar, adParamInput, 12, strSSN)
        .Parameters.Append .CreateParameter("@DLNum", adVarChar, adParamInput, 20, strDLNum)
        .Parameters.Append .CreateParameter("@Employer", adVarChar, adParamInput, 35, strEmployer)
        .Parameters.Append .CreateParameter("@PhyName", adVarChar, adParamInput, 30, strPhysician)
        .Parameters.Append .CreateParameter("@PhyPhone", adVarChar, adParamInput, 15, strPsyPhone)
        .Parameters.Append .CreateParameter("@EmgName", adVarChar, adParamInput, 30, strEmgName)
        .Parameters.Append .CreateParameter("@EmgRelat", adVarChar, adParamInput, 20, strEmgRelat)
        .Parameters.Append .CreateParameter("@EmgPhone", adVarChar, adParamInput, 15, strEmgPhone)
        .Parameters.Append .CreateParameter("@Notes", adVarChar, adParamInput, 255, strNotes)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Execute the stored procedure
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set Search = rs
    
    'Disconnect the recordset
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rs.ActiveConnection = Nothing
    Set cnn = Nothing
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmd = Nothing
    Set cnn = Nothing
    Set rs = Nothing
    GetObjectContext.SetAbort
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
End Function
 Public Function FetchByProviderLName(ByVal intProvider As Long, ByVal strLast As String, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
     Dim rs As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     
     On Error GoTo ErrTrap:
     
     Set rs = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    'Create the parameter objects
     With cmd
         .CommandText = "uspSearchProviderPatientsByLName"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , intProvider)
         .Parameters.Append .CreateParameter("@LName", adVarChar, adParamInput, 30, strLast)
         .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = True, "Y", "N"))
     End With
    
    'Execute the stored procedure
     rs.CursorLocation = adUseClient
     rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
     Set FetchByProviderLName = rs
    
    'Disconnect the recordset
     Set cmd.ActiveConnection = Nothing
     Set cmd = Nothing
     Set rs.ActiveConnection = Nothing
     Set cnn = Nothing
    
    'Signal successful completion
     GetObjectContext.SetComplete
    
     Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmd = Nothing
    Set cnn = Nothing
    Set rs = Nothing
    GetObjectContext.SetAbort
    Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function
 Public Function FetchByProviderFName(ByVal intProvider As Long, ByVal strFirst As String, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
     Dim rs As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     
     On Error GoTo ErrTrap:
     
     Set rs = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    'Create the parameter objects
     With cmd
         .CommandText = "uspSearchProviderPatientsByFName"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , intProvider)
         .Parameters.Append .CreateParameter("@FName", adVarChar, adParamInput, 30, strFirst)
         .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = True, "Y", "N"))
     End With
    
    'Execute the stored procedure
     rs.CursorLocation = adUseClient
     rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
     Set FetchByProviderFName = rs
    
    'Disconnect the recordset
     Set cmd.ActiveConnection = Nothing
     Set cmd = Nothing
     Set rs.ActiveConnection = Nothing
     Set cnn = Nothing
    'Signal successful completion
     GetObjectContext.SetComplete
    
     Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmd = Nothing
    Set cnn = Nothing
    Set rs = Nothing
     GetObjectContext.SetAbort
     Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function
 Public Function FetchByProviderLNameFName(ByVal intProvider As Long, ByVal strLast As String, ByVal strFirst As String, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
     Dim rs As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     
     On Error GoTo ErrTrap:
     
     Set rs = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
     With cmd
         .CommandText = "uspSearchProviderPatientsByLNameFName"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , intProvider)
         .Parameters.Append .CreateParameter("@LName", adVarChar, adParamInput, 30, strLast)
         .Parameters.Append .CreateParameter("@FName", adVarChar, adParamInput, 30, strFirst)
         .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = True, "Y", "N"))
     End With
    
    'Execute the stored procedure
     rs.CursorLocation = adUseClient
     rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
     Set FetchByProviderLNameFName = rs
    
    'Disconnect the recordset
     Set cmd.ActiveConnection = Nothing
     Set cmd = Nothing
     Set rs.ActiveConnection = Nothing
     Set cnn = Nothing
    
    'Signal successful completion
     GetObjectContext.SetComplete
    
     Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmd = Nothing
    Set cnn = Nothing
    Set rs = Nothing
     GetObjectContext.SetAbort
     Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function
 
 Public Function FetchByProvider(ByVal intProvider As Long, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
     Dim rs As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     
     On Error GoTo ErrTrap:
     
     Set rs = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
     With cmd
         .CommandText = "uspSelProviderPatientsByName"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , intProvider)
         .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = True, "Y", "N"))
     End With
    
    'Execute the stored procedure
     rs.CursorLocation = adUseClient
     rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
     Set FetchByProvider = rs
    
    'Disconnect the recordset
     Set cmd.ActiveConnection = Nothing
     Set cmd = Nothing
     Set rs.ActiveConnection = Nothing
     Set cnn = Nothing
    
    'Signal successful completion
     GetObjectContext.SetComplete
     Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmd = Nothing
    Set cnn = Nothing
    Set rs = Nothing
     GetObjectContext.SetAbort
     Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function
 Public Function FetchRPByProvider(ByVal intProvider As Long, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 03/26/2005
'Author: Duane C Orth
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
     Dim rs As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     
     On Error GoTo ErrTrap:
     
     Set rs = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
     With cmd
         .CommandText = "uspSelRPByProvider"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , intProvider)
         .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = True, "Y", "N"))
     End With
    
    'Execute the stored procedure
     rs.CursorLocation = adUseClient
     rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
     Set FetchRPByProvider = rs
    
    'Disconnect the recordset
     Set cmd.ActiveConnection = Nothing
     Set cmd = Nothing
     Set rs.ActiveConnection = Nothing
     Set cnn = Nothing
    
    'Signal successful completion
     GetObjectContext.SetComplete
     Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmd = Nothing
    Set cnn = Nothing
    Set rs = Nothing
     GetObjectContext.SetAbort
     Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function
 Public Function FetchBySSNLName(ByVal strSSN As String, ByVal strLast As String, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
     Dim rs As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     On Error GoTo ErrTrap:
     
     
     Set rs = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
     With cmd
         .CommandText = "uspSearchBenefactorByLNameSSN"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@Last", adVarChar, adParamInput, 30, strLast)
         .Parameters.Append .CreateParameter("@SSN", adVarChar, adParamInput, 11, strSSN)
         .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = True, "Y", "N"))
     End With
    
    'Execute the stored procedure
     rs.CursorLocation = adUseClient
     rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
     Set FetchBySSNLName = rs
    
    'Disconnect the recordset
     Set cmd.ActiveConnection = Nothing
     Set cmd = Nothing
     Set rs.ActiveConnection = Nothing
     Set cnn = Nothing
    
    'Signal successful completion
     GetObjectContext.SetComplete
     Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmd = Nothing
    Set cnn = Nothing
    Set rs = Nothing
     GetObjectContext.SetAbort
     Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function
 
 Public Function FetchBySSNProvider(ByVal strSSN As String, ByVal lngProviderID As Long, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 5/11/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
     Dim rs As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     
     On Error GoTo ErrTrap:
     
     Set rs = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
     With cmd
         .CommandText = "uspSearchBenefactorBySSNProvider"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@SSN", adVarChar, adParamInput, 11, strSSN)
         .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
         .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, IIf(blnIncludeDisabled = True, "Y", "N"))
     End With
    
    'Execute the stored procedure
     rs.CursorLocation = adUseClient
     rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
     Set FetchBySSNProvider = rs
    
    'Disconnect the recordset
     Set cmd.ActiveConnection = Nothing
     Set cmd = Nothing
     Set rs.ActiveConnection = Nothing
     Set cnn = Nothing
    
    'Signal successful completion
     GetObjectContext.SetComplete
     Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
     GetObjectContext.SetAbort
    Set cmd = Nothing
    Set cnn = Nothing
    Set rs = Nothing
     Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function

Public Function FetchArchiveCandidates(ByVal lngUserID As Long, ByVal dtLastVisitDate As Date)
'--------------------------------------------------------------------
'Date: 5/30/2002
'Author: Dave Richkun
'Description:  Returns a list of patients and providers who have not visited with
'              each other after the LastVisitDate parameter
'Parameters: lngUserID - ID of user executing the method
'            dtLastVisitDate - The 'cut-off' date for which to retrieve archive candidates
'Returns: Recordset of archive candidates
'--------------------------------------------------------------------
     Dim rst As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     On Error GoTo ErrTrap:
     
     Set rst = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
     With cmd
         .CommandText = "uspSelPatientArchiveCandidates"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
         .Parameters.Append .CreateParameter("@LastVisitDate", adDBTimeStamp, adParamInput, , dtLastVisitDate)
     End With
    
    'Execute the stored procedure
     rst.CursorLocation = adUseClient
     rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
     Set FetchArchiveCandidates = rst
    
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
    Set cmd = Nothing
    Set cnn = Nothing
    Set rst = Nothing
    Call RaiseError(Err, Err.Number, CLASS_NAME)

End Function

Public Function FetchPatientBalanceByUser(ByVal lngPatientID As Long, ByVal lngUserID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 07/26/2002
'Author: Dave Richkun
'Description:  Returns the total amount charged and total amount paid for a
'              patient, regardless of the Responsibile party charged or paid.
'              The amounts returned are based on the user performing the method i.e.
'              amounts may vary based on whether the user is an office manager or a
'              provider.
'Parameters: lngPatientID - ID of patient whose balances are being returned
'            lngUserID - ID of user executing the method
'Returns: Recordset of archive candidates
'--------------------------------------------------------------------
     Dim rst As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     On Error GoTo ErrTrap:
     
     Set rst = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
     With cmd
         .CommandText = "uspSelPatBalByUser"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
         .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
     End With
    
    'Execute the stored procedure
     rst.CursorLocation = adUseClient
     rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
     Set FetchPatientBalanceByUser = rst
    
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
    Set cmd = Nothing
    Set cnn = Nothing
    Set rst = Nothing
    Call RaiseError(Err, Err.Number, CLASS_NAME)

End Function


Public Sub UpdatePatientCardNumber(ByVal lngPatientID As Long, ByVal lngPlanID As Long, _
                                   ByVal strCardNumber As String)
'--------------------------------------------------------------------
'Date: 10/15/2002
'Author: Dave Richkun
'Description:  Updates the card number for a patient plan
'Parameters: lngPatientID - ID of patient whose card number is being updated
'            lngPlanID - ID of plan whose card number is being updated
'            strCardNumber - Card number
'Returns: Null
'--------------------------------------------------------------------
                                   
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    
    With cmdSQL
        .CommandText = "uspUpdPatientCardNumber"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@PlanID", adInteger, adParamInput, , lngPlanID)
        .Parameters.Append .CreateParameter("@CardNumber", adVarChar, adParamInput, 30, strCardNumber)
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
Public Sub UpdatePatientPassword(ByVal lngPatientID As Long, _
                                 ByVal strUserName, _
                                 ByVal strNewPassword)
'--------------------------------------------------------------------
'Date: 04/17/2000
'Author: Dave Richkun
'Description:  Updates the fldLastPWChange value with the date passed.
'Parameters: strUserName - Username of the user whose password was changed
'            dtChanged - The date to record in the fldLastPWChange column
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    
    'Prepare stored procedure parameters
    With cmdSQL
        .CommandText = "uspUpdPatPassword"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, 30, strNewPassword)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
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
    GetObjectContext.SetAbort
    Call RaiseError(Err, cmdSQL.Parameters("@SQLErrorNum").Value, CLASS_NAME)

End Sub


Public Sub CombinePatients(ByVal lngFromPatID As Long, _
                           ByVal lngToPatID As Long, _
                           ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 03/07/2014
'Author: Duane C Orth
'Description:  Combines two patient records.. Deletes the from patient.
'Parameters: strUserName - Username of the user whose password was changed
'            dtChanged - The date to record in the fldLastPWChange column
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    
    'Prepare stored procedure parameters
    With cmdSQL
        .CommandText = "uspCombinePatients"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@FromPatientID", adInteger, adParamInput, , lngFromPatID)
        .Parameters.Append .CreateParameter("@ToPatientID", adInteger, adParamInput, , lngToPatID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
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
    GetObjectContext.SetAbort
    Call RaiseError(Err, cmdSQL.Parameters("@SQLErrorNum").Value, CLASS_NAME)

End Sub


Public Sub UpdateCC(ByVal lngPatientID As Long, ByRef varCC As Variant)
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim strSQL As String
    Dim arrCC() As Variant
    ReDim arrCC(1, 4)
    
    On Error GoTo ErrTrap:

'       arrCC(1, 4) Credit Card Info
    If Not IsArray(varCC) Or IsEmpty(varCC) Then
        arrCC(0, 0) = ""
        arrCC(0, 1) = ""
        arrCC(0, 2) = 0
        arrCC(0, 3) = 0
        arrCC(0, 4) = ""
        arrCC(1, 0) = ""
        arrCC(1, 1) = ""
        arrCC(1, 2) = 0
        arrCC(1, 3) = 0
        arrCC(1, 4) = ""
    Else
        arrCC(0, 0) = IfNull(varCC(0, 0), "")
        arrCC(0, 1) = IfNull(varCC(0, 1), "")
        arrCC(0, 2) = IIf(Not IsNumeric(IfNull(varCC(0, 2), 0)), 0, IfNull(varCC(0, 2), 0))
        arrCC(0, 3) = IIf(Not IsNumeric(IfNull(varCC(0, 3), 0)), 0, IfNull(varCC(0, 3), 0))
        arrCC(0, 4) = IfNull(varCC(0, 4), "")
        arrCC(1, 0) = IfNull(varCC(1, 0), "")
        arrCC(1, 1) = IfNull(varCC(1, 1), "")
        arrCC(1, 2) = IIf(Not IsNumeric(IfNull(varCC(1, 2), 0)), 0, IfNull(varCC(1, 2), 0))
        arrCC(1, 3) = IIf(Not IsNumeric(IfNull(varCC(1, 3), 0)), 0, IfNull(varCC(1, 3), 0))
        arrCC(1, 4) = IfNull(varCC(1, 4), "")
    End If


 'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    With cmdSQL
        .CommandText = "uspUpdPatientCCNumber"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@PriCCardNbr", adVarChar, adParamInput, 50, arrCC(0, 0))
        .Parameters.Append .CreateParameter("@PriCCardSecurity", adVarChar, adParamInput, 5, arrCC(0, 1))
        .Parameters.Append .CreateParameter("@PriCCardMth", adInteger, adParamInput, , IIf(arrCC(0, 2) <= 0, Null, arrCC(0, 2)))
        .Parameters.Append .CreateParameter("@PriCCardYr", adInteger, adParamInput, , IIf(arrCC(0, 3) <= 0, Null, arrCC(0, 3)))
        .Parameters.Append .CreateParameter("@PriCCVaultID", adVarChar, adParamInput, 50, arrCC(0, 4))
        .Parameters.Append .CreateParameter("@SecCCardNbr", adVarChar, adParamInput, 50, arrCC(1, 0))
        .Parameters.Append .CreateParameter("@SecCCardSecurity", adVarChar, adParamInput, 5, arrCC(1, 1))
        .Parameters.Append .CreateParameter("@SecCCardMth", adInteger, adParamInput, , IIf(arrCC(1, 2) <= 0, Null, arrCC(1, 2)))
        .Parameters.Append .CreateParameter("@SecCCardYr", adInteger, adParamInput, , IIf(arrCC(1, 3) <= 0, Null, arrCC(1, 3)))
        .Parameters.Append .CreateParameter("@SecCCVaultID", adVarChar, adParamInput, 50, arrCC(1, 4))
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
        'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute ', , adExecuteNoRecords
    
    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
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
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function ValidatePatient(ByVal lngOwnerID As Long, ByVal strUserName As String, ByVal strPassword As String) As ADODB.Recordset

     Dim rst As ADODB.Recordset
     Dim cmd As ADODB.Command
     Dim cnn As ADODB.Connection
     
     On Error GoTo ErrTrap:
     
     Set rst = New ADODB.Recordset
     Set cmd = New ADODB.Command
     Set cnn = New ADODB.Connection
     
    'Create the parameter objects
     With cmd
         .CommandText = "uspValidatePatient"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@OwnerID", adInteger, adParamInput, , lngOwnerID)
         .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
         .Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, 20, strPassword)
     End With
    
    'Open the connection and execute the stored procedure
     cnn.Open (CONST_PSYQUEL_CNN)
     Set cmd.ActiveConnection = cnn
     rst.CursorLocation = adUseClient
     rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
     
     Set ValidatePatient = rst
    
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
     Set cmd = Nothing
     Set cnn = Nothing
     Set rst = Nothing
     Call RaiseError(Err, Err.Number, CLASS_NAME)
End Function



