VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBenefactorBz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CBenefactorBz
'Date: 01/18/2000
'Author: Eric Pena
'Description:  MTS business object designed to call methods associated
'              with the CBenefactorDB class.
'--------------------------------------------------------------------
' Revision History:
'   R001: 11/02/2000 Richkun - Added lngProviderID parameter to Deleted() method.
'   R002: 01/09/2001 Richkun - Added strPhysFax parameter to Insert() abd Update() methods
'   R003: 03/13/2001 Nichol - Modified VerifyData() to use FormatZipCode()
'   R004: 04/18/2001 Richkun - Added support for HomePhMsg and WorkPhMsg columns
'   R005: 05/30/2003 Richkun - Added FetchArchiveCandidates() method
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CBenefactorBz"
Private Const TABLE_NAME As String = "tblBenefactor"

Public Function Fetch(Optional blnIncludeDisabled As Boolean = False, Optional ByVal strWhere As String, Optional ByVal strOrderBy As String) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 01/29/2000
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
'               'fldClinic' column
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB

    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set Fetch = objBenefactor.Fetch(blnIncludeDisabled, strWhere, strOrderBy)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function FetchByID(ByVal lngBenefactorID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 07/11/2000
'Author: Dave Richkun
'Description:  Returns demographic information about a particular Benefactor
'Parameters: lngBenefactorID - ID of the Benefactor to retrieve information for
'Returns: Recordset
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchByID = objBenefactor.FetchByID(lngBenefactorID)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objBenefactor = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function Insert(ByVal strLast As String, ByVal strMaiden As String, ByVal strMI As String, _
                    ByVal strFirst As String, ByVal strNickName As String, ByVal strAddress1 As String, ByVal strAddress2 As String, _
                    ByVal strCity As String, ByVal strState As String, ByVal strCountry As String, ByVal strZip As String, ByVal strCounty As String, _
                    ByVal strHomePhoneDesc As String, ByVal strHomePhone As String, _
                    ByVal strMobilePhoneDesc As String, ByVal strMobilePhone As String, _
                    ByVal strWorkPhoneDesc As String, ByVal strWorkPhone As String, ByVal strWorkExt As String, _
                    ByVal strEmail As String, _
                    ByVal dteDOB As Date, ByVal strSex As String, ByVal strSSN As String, _
                    ByVal strDLNum As String, ByRef varCC As Variant, ByVal lngMarital As Long, ByVal lngEthnicity As Long, ByVal lngHHIncome As Long, _
                    ByVal lngEmployment As Long, ByVal strEmployer As String, ByRef varPCP As Variant, ByRef varREF As Variant, _
                    ByVal strEmgName As String, ByVal strEmgRelat As String, ByVal strEmgPhone As String, _
                    ByVal strGuardianName As String, ByVal strGuardianRelat As String, ByVal strGuardianPhone As String, _
                    ByVal strOtherPhone1 As String, ByVal strOtherPhoneDesc1 As String, ByVal strOtherPhone2 As String, ByVal strOtherPhoneDesc2 As String, ByVal strOtherPhone3 As String, ByVal strOtherPhoneDesc3 As String, _
                    ByVal strNotes As String, ByVal strInternalID As String, ByVal lngClinicID As Long, _
                    ByVal dteAdmitDate As Date, ByVal strUserName As String, ByVal strPassword As String, ByRef varFlags As Variant, _
                    ByVal strAddedBy As String, ByVal lngUserID As Long, Optional ByVal lngOwnerID As Long = 0) As Long
'--------------------------------------------------------------------
'Date: 01/29/2000
'Author: Eric Pena
'Description:  Inserts a row into the tblBenefactor table utilizing
'              a stored procedure.
'Parameters: Each parameter identifies the column value that will be
'              inserted.
'Returns: ID (Primary Key) of the row inserted.
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    Dim strErrMessage As String
    
    On Error GoTo ErrTrap:
    
    If Trim(strCountry) = "USA" Then strZip = NumbersOnly(strZip)
    If Not VerifyData(0, strLast, strFirst, strHomePhone, strMobilePhone, strWorkPhone, strCountry, strZip, strSSN, varPCP, strEmgPhone, dteDOB, strEmail, strErrMessage) Then GoTo ErrTrap
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Insert = objBenefactor.Insert(strLast, strMaiden, strMI, strFirst, strNickName, strAddress1, strAddress2, _
            strCity, strState, strCountry, strZip, strCounty, strHomePhoneDesc, NumbersOnly(strHomePhone), strMobilePhoneDesc, NumbersOnly(strMobilePhone), _
            strWorkPhoneDesc, NumbersOnly(strWorkPhone), strWorkExt, strEmail, dteDOB, strSex, NumbersOnly(strSSN), strDLNum, varCC, _
            lngMarital, lngEthnicity, lngHHIncome, lngEmployment, strEmployer, varPCP, varREF, _
            strEmgName, strEmgRelat, NumbersOnly(strEmgPhone), strGuardianName, strGuardianRelat, NumbersOnly(strGuardianPhone), _
            NumbersOnly(strOtherPhone1), strOtherPhoneDesc1, NumbersOnly(strOtherPhone2), strOtherPhoneDesc2, NumbersOnly(strOtherPhone3), strOtherPhoneDesc3, _
            strNotes, strInternalID, lngClinicID, dteAdmitDate, strUserName, strPassword, varFlags, strAddedBy, lngOwnerID)
    Set objBenefactor = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Public Sub Update(ByVal lngID As Long, ByVal strLast As String, ByVal strMaiden As String, ByVal strMI As String, _
                    ByVal strFirst As String, ByVal strNickName As String, ByVal strAddress1 As String, ByVal strAddress2 As String, _
                    ByVal strCity As String, ByVal strState As String, ByVal strCountry As String, ByVal strZip As String, ByVal strCounty As String, _
                    ByVal strHomePhoneDesc As String, ByVal strHomePhone As String, _
                    ByVal strMobilePhoneDesc As String, ByVal strMobilePhone As String, _
                    ByVal strWorkPhoneDesc As String, ByVal strWorkPhone As String, ByVal strWorkExt As String, _
                    ByVal strEmail As String, _
                    ByVal dteDOB As Date, ByVal strSex As String, ByVal strSSN As String, _
                    ByVal strDLNum As String, ByRef varCC As Variant, ByVal lngMarital As Long, ByVal lngEthnicity As Long, ByVal lngHHIncome As Long, _
                    ByVal lngEmployment As Long, ByVal strEmployer As String, ByRef varPCP As Variant, ByRef varREF As Variant, _
                    ByVal strEmgName As String, ByVal strEmgRelat As String, ByVal strEmgPhone As String, _
                    ByVal strGuardianName As String, ByVal strGuardianRelat As String, ByVal strGuardianPhone As String, _
                    ByVal strOtherPhone1 As String, ByVal strOtherPhoneDesc1 As String, ByVal strOtherPhone2 As String, ByVal strOtherPhoneDesc2 As String, ByVal strOtherPhone3 As String, ByVal strOtherPhoneDesc3 As String, _
                    ByVal strNotes As String, ByVal strInternalID As String, ByVal lngClinicID As Long, _
                    ByVal dteAdmitDate As Date, ByVal strUserName As String, ByVal strPassword As String, ByRef varFlags As Variant, _
                    ByVal strUpdatedBy As String)
'--------------------------------------------------------------------
'Date: 01/29/2000
'Author: Eric Pena
'Description:  Updates a row in the tblBenefactor table utilizing
'              a stored procedure.
'Parameters: Each parameter identifies the column value that will be
'              inserted.
'Returns: Null
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    Dim strErrMessage As String

    On Error GoTo ErrTrap:
    
    If Trim(strCountry) = "USA" Then strZip = NumbersOnly(strZip)
    If Not VerifyData(lngID, strLast, strFirst, strHomePhone, strMobilePhone, strWorkPhone, strCountry, strZip, strSSN, varPCP, strEmgPhone, dteDOB, strEmail, strErrMessage) Then GoTo ErrTrap
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Call objBenefactor.Update(lngID, strLast, strMaiden, strMI, strFirst, strNickName, strAddress1, strAddress2, _
            strCity, strState, strCountry, strZip, strCounty, strHomePhoneDesc, NumbersOnly(strHomePhone), strMobilePhoneDesc, NumbersOnly(strMobilePhone), _
            strWorkPhoneDesc, NumbersOnly(strWorkPhone), strWorkExt, strEmail, dteDOB, strSex, strSSN, strDLNum, varCC, lngMarital, lngEthnicity, lngHHIncome, _
            lngEmployment, strEmployer, varPCP, varREF, _
            strEmgName, strEmgRelat, NumbersOnly(strEmgPhone), strGuardianName, strGuardianRelat, NumbersOnly(strGuardianPhone), _
            NumbersOnly(strOtherPhone1), strOtherPhoneDesc1, NumbersOnly(strOtherPhone2), strOtherPhoneDesc2, NumbersOnly(strOtherPhone3), strOtherPhoneDesc3, _
            strNotes, strInternalID, lngClinicID, dteAdmitDate, strUserName, strPassword, varFlags, strUpdatedBy)

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Set objBenefactor = Nothing
    
    Exit Sub

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    GetObjectContext.SetAbort
    Set objBenefactor = Nothing
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

Public Sub Deleted(ByVal blnDeleted As Boolean, ByVal lngPatientID As Long, _
                   ByVal lngProviderID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 01/29/2000
'Author: Eric Pena
'Description:  Flags a row in the tblBenefactor table marking the row as
'              deleted or undeleted.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngPatientID - ID of the row in the Benefactor table whose
'               value will be  updated.
'            lngProviderID - ID of the Provider who wishes the patient to be
'               archived/restored.  This parameter is needed because patients
'               may be associated with more than one provider.  'R001
'            strUserName - Login name of the user responsible for
'               marking the row as deleted.
'Returns: Null
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
   
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    
    'TODO: Check for active claims
    Call objBenefactor.Deleted(blnDeleted, lngPatientID, lngProviderID, _
                strUserName)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function Search(ByVal strLast As String, ByVal strMaiden As String, ByVal strMI As String, ByVal strFirst As String, _
                    ByVal strNickName As String, ByVal strAddress1 As String, ByVal strAddress2 As String, ByVal strCity As String, _
                    ByVal strState As String, ByVal strZip As String, ByVal strHomePhone As String, ByVal strWorkPhone As String, _
                    ByVal strWorkExt As String, ByVal strEmail As String, ByVal strSex As String, ByVal strSSN As String, ByVal strDLNum As String, _
                    ByVal strEmployer As String, ByVal strPhysician As String, ByVal strPhysPhone As String, ByVal strEmgName As String, _
                    ByVal strEmgRelat As String, ByVal strEmgPhone As String, ByVal strNotes As String) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 01/29/2000
'Author: Eric Pena
'Description:  Retrieves records from the tblBenefactor table.
'Parameters: The parameters contain values the returned records will contain
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set Search = objBenefactor.Search(strLast, strMaiden, strMI, strFirst, strNickName, strAddress1, strAddress2, strCity, _
                     strState, strZip, strHomePhone, strWorkPhone, strWorkExt, strEmail, strSex, strSSN, strDLNum, _
                     strEmployer, strPhysician, strPhysPhone, strEmgName, strEmgRelat, strEmgPhone, strNotes)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function VerifyData(ByVal lngID As Long, ByVal strLast As String, ByVal strFirst As String, _
                            ByRef strHomePhone As String, ByRef strMobilePhone As String, ByRef strWorkPhone As String, ByRef strCountry As String, _
                            ByRef strZip As String, ByRef strSSN As String, ByRef varPCP As Variant, ByRef strEmgPhone As String, _
                            ByVal dteDOB As Date, ByRef strEmail As String, ByRef strErrMessage As String) As Boolean
'--------------------------------------------------------------------
'Date: 01/29/2000
'Author: Eric Pena
'Description:  Verifies all required data has been provided by the user.
'Parameters:  The values to be checked.
'Returns: Boolean value identifying if all criteria has been satisfied.
'--------------------------------------------------------------------
    If Trim(strLast) = "" Then
        strErrMessage = "Last Name is required."
        VerifyData = False
        Exit Function
    End If
    If Trim(strFirst) = "" Then
        strErrMessage = "First Name is required."
        VerifyData = False
        Exit Function
    End If

    strHomePhone = NumbersOnly(strHomePhone)
    strWorkPhone = NumbersOnly(strWorkPhone)
    strMobilePhone = NumbersOnly(strMobilePhone)
    If strCountry = "USA" Then strZip = NumbersOnly(strZip)
    strSSN = NumbersOnly(strSSN)
  '  strPhysPhone = NumbersOnly(strPhysPhone)
    strEmgPhone = NumbersOnly(strEmgPhone)
    
    If Trim(strHomePhone) = "" And Trim(strMobilePhone) = "" And Trim(strWorkPhone) = "" Then
        strErrMessage = "A phone number is required."
        VerifyData = False
        Exit Function
    End If
    
    If (Len(strHomePhone) <> 10 And Len(strHomePhone) <> 0) Then
        strErrMessage = "Home Phone Number is invalid."
        VerifyData = False
        Exit Function
    End If
    
    If (Len(strMobilePhone) <> 10 And Len(strMobilePhone) <> 0) Then
        strErrMessage = "Mobile Phone Number is invalid."
        VerifyData = False
        Exit Function
    End If
    
    If (Len(strWorkPhone) <> 10 And Len(strWorkPhone) <> 0) Then
        strErrMessage = "Work Phone Number is invalid."
        VerifyData = False
        Exit Function
    End If
    
'    If (Len(strPhysPhone) <> 10 And Len(strPhysPhone) <> 0) Then
'        strErrMessage = "Psysician Phone Number is invalid."
'        VerifyData = False
'        Exit Function
'    End If
    
    If (Len(strEmgPhone) <> 10 And Len(strEmgPhone) <> 0) Then
        strErrMessage = "Emergency Phone Number is invalid."
        VerifyData = False
        Exit Function
    End If
    
    If (strCountry = "USA" And Len(strZip) <> 5 And Len(strZip) <> 9 And Len(strZip) <> 0) Then
        strErrMessage = "Zip Code is invalid."
        VerifyData = False
        Exit Function
    End If
        
    If (Len(strSSN) <> 9 And Len(strSSN) <> 0) Then
        strErrMessage = "Social Security Number is invalid."
        VerifyData = False
        Exit Function
    End If
    
    If IsDate(dteDOB) Then
        If dteDOB > Date Then
            strErrMessage = "Date of Birth cannot be in the future."
            VerifyData = False
            Exit Function
        End If
    End If
    
    If strCountry = "USA" Then strZip = FormatZipCode(strZip)
    
    If Len(strEmail) > 0 And (InStr(1, strEmail, "@") = 0 _
        Or InStr(1, strEmail, ".") = 0 _
        Or InStr(1, strEmail, ",") > 0 _
        Or InStr(1, strEmail, "(") > 0 _
        Or InStr(1, strEmail, ")") > 0 _
        Or InStr(1, strEmail, ":") > 0) Then
        strErrMessage = "Email Address is invalid."
        VerifyData = False
        Exit Function
    End If
    
    'Parse varPCP.
    If Not IsArray(varPCP) Then
        strErrMessage = "varPCP is not an array"
        VerifyData = False
        Exit Function
    End If
    
    If IsEmpty(varPCP) Then
        strErrMessage = "varPCP is Empty"
        VerifyData = False
        Exit Function
    End If
        
    'If we get here, all is well...
    VerifyData = True
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
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchByProviderLName = objBenefactor.FetchByProviderLName(intProvider, strLast, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

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
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchByProviderFName = objBenefactor.FetchByProviderFName(intProvider, strFirst, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
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
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchByProviderLNameFName = objBenefactor.FetchByProviderLNameFName(intProvider, strLast, strFirst, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

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
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchByProvider = objBenefactor.FetchByProvider(intProvider, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function FetchRPByProvider(ByVal intProvider As Long, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchRPByProvider = objBenefactor.FetchRPByProvider(intProvider, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
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
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchBySSNLName = objBenefactor.FetchBySSNLName(strSSN, strLast, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function FetchBySSNProvider(ByVal strSSN As String, ByVal lngProviderID As Long, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   All other parameters represent column values to search for
'Returns: ADORst
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchBySSNProvider = objBenefactor.FetchBySSNProvider(strSSN, lngProviderID, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

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

    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchArchiveCandidates = objBenefactor.FetchArchiveCandidates(lngUserID, dtLastVisitDate)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

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

    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set FetchPatientBalanceByUser = objBenefactor.FetchPatientBalanceByUser(lngPatientID, lngUserID)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

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

    Dim objBenefactor As BenefactorDB.CBenefactorDB

    On Error GoTo ErrTrap:

    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Call objBenefactor.UpdatePatientCardNumber(lngPatientID, lngPlanID, strCardNumber)
    Set objBenefactor = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Sub

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    Set objBenefactor = Nothing
    
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Public Sub CombinePatients(ByVal lngFromPatID As Long, _
                           ByVal lngToPatID As Long, _
                           ByVal strUserName As String)
    Dim objBenefactor As BenefactorDB.CBenefactorDB

    On Error GoTo ErrTrap:

    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Call objBenefactor.CombinePatients(lngFromPatID, lngToPatID, strUserName)
    Set objBenefactor = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Sub

ErrTrap:
    Set objBenefactor = Nothing
    
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub UpdateCreditCard(ByVal lngPatientID As Long, ByRef varCC As Variant)

    Dim objBenefactor As BenefactorDB.CBenefactorDB

    On Error GoTo ErrTrap:

    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Call objBenefactor.UpdateCC(lngPatientID, varCC)
    Set objBenefactor = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Sub

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    Set objBenefactor = Nothing
    
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
Public Function ValidatePatient(ByVal lngOwnerID As Long, ByVal lngUserName As String, ByVal lngPassword As String) As ADODB.Recordset
    
    Dim objBenefactor As BenefactorDB.CBenefactorDB
    
    On Error GoTo ErrTrap:
    
    
    Set objBenefactor = CreateObject("BenefactorDB.CBenefactorDB")
    Set ValidatePatient = objBenefactor.ValidatePatient(lngOwnerID, lngUserName, lngPassword)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objBenefactor = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
