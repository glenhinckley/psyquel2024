VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApptCategoryBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CApptCategoryBz                                        '
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  MTS business object designed to call methods         '
'              associated with the CApptCategoryDB class.           '
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CApptCategoryBz"

Public Function Fetch(Optional ByVal blnIncludeDisabled As Boolean = False) _
    As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Retrieves records from the tblApptCategory table.    '
'Parameters: blnIncludeDisabled - Optional parameter that identifies'
'              if records flagged as 'Disabled' or 'De-activated'   '
'              are to be included in the record set. The default    '
'              value is False.                                      '
'Returns: Recordset of appointment categories                       '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptCategoryDB
    Dim rstSQL As ADODB.Recordset

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CApptCategoryDB")
    Set rstSQL = objAppt.Fetch(blnIncludeDisabled)

    Set Fetch = rstSQL
    
    Set objAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function Insert(ByVal strDescription As String) As Long
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Inserts a row into the tblApptCategory table.        '
'Parameters: strDescription - The description of the Appt Category  '
'              that will be inserted into the table.                '
'Returns: ID (Primary Key) of the row inserted                      '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptCategoryDB
    Dim lngID As Long
    Dim strErrMessage As String

    On Error GoTo ErrTrap:

    'Verify data before proceeding.
    If Not VerifyData(0, strDescription, strErrMessage) Then
        GoTo ErrTrap
    End If
    
    Set objAppt = CreateObject("ApptDB.CApptCategoryDB")
    lngID = objAppt.Insert(strDescription)

    Insert = lngID
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    'Release resources
    Set objAppt = Nothing
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function


Public Sub Update(ByVal lngID As Long, ByVal strDescription As String)
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Updates a row into the tblApptCategory table.        '
'Parameters:  lngID - ID of the row in the table whose value will be'
'               updated.                                            '
'             strDescription - The appointment Category description '
'                to which the record will be changed.               '
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptCategoryDB
    Dim strErrMessage As String

    On Error GoTo ErrTrap:

    'Verify data before proceeding.
    If Not VerifyData(lngID, strDescription, strErrMessage) Then
        GoTo ErrTrap
    End If
    
    Set objAppt = CreateObject("ApptDB.CApptCategoryDB")
    Call objAppt.Update(lngID, strDescription)

    Set objAppt = Nothing
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Sub


Public Sub Deleted(ByVal blnDeleted As Boolean, ByVal lngID As Long)
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Flags a row in the tblApptCategory table marking the row
'               as deleted or undeleted.                            '
'Parameters: blnDeleted - Boolean value identifying if the record is'
'               to be deleted (True) or undeleted (False).          '
'            lngID - ID of the row in the table whose value will be '
'               updated.                                            '
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptCategoryDB

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CApptCategoryDB")

    Call objAppt.Deleted(blnDeleted, lngID)
    Set objAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Public Function Exists(ByVal strDescription As String) As Boolean
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Determines if an appoinment Category description     '
'               identical to the strDescription parameter already   '
'               exists in the table.                                '
'Parameters: strDescription - Appointment Category name to be checked
'Returns: True if the name exists, false otherwise                  '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CApptCategoryDB
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CApptCategoryDB")
    
    blnExists = objAppt.Exists(strDescription)
    Exists = blnExists

    Set objAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function VerifyData(ByVal lngID As Long, ByVal strDescription As String, _
                            ByRef strErrMessage As String) As Boolean
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Verifies all required data has been provided by the user.
'Parameters:  The values to be checked.                             '
'Returns: Boolean value identifying if all criteria has been satisfied.
'--------------------------------------------------------------------

    If Trim(strDescription) = "" Then
        strErrMessage = "Market Name is required."
        VerifyData = False
        Exit Function
    End If

    'Check for existance only when inserting new data
    If lngID = 0 And Exists(strDescription) Then
        strErrMessage = "Appointment Category '" & strDescription & "' already exists."
        VerifyData = False
        Exit Function
    End If

    'If we get here, all is well...
    VerifyData = True

End Function

