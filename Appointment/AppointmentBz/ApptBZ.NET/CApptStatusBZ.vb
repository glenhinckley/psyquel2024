Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("CApptStatusBZ_NET.CApptStatusBZ")> Public Class CApptStatusBZ
	'--------------------------------------------------------------------
	'Class Name: CApptStatusBz                                          '
	'Date: 08/28/2000                                                   '
	'Author: Rick "Boom Boom" Segura                                    '
	'Description:  MTS business object designed to call methods         '
	'              associated with the CApptStatusDB class.             '
	'--------------------------------------------------------------------
	
	Private Const CLASS_NAME As String = "CApptStatusBz"
	
	Public Function Fetch(Optional ByVal blnIncludeDisabled As Boolean = False) As ADODB.Recordset
		Dim ApptDB As Object
		'--------------------------------------------------------------------
		'Date: 08/28/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Retrieves records from the tblApptStatus table.      '
		'Parameters: blnIncludeDisabled - Optional parameter that identifies'
		'              if records flagged as 'Disabled' or 'De-activated'   '
		'              are to be included in the record set. The default    '
		'              value is False.                                      '
		'Returns: Recordset of appointment statuses                         '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'--------------------------------------------------------------------
		Dim objAppt As ApptDB.CApptStatusDB
		Dim rstSQL As ADODB.Recordset
		
		On Error GoTo ErrTrap
		
		objAppt = CreateObject("ApptDB.CApptStatusDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.Fetch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rstSQL = objAppt.Fetch(blnIncludeDisabled)
		
		Fetch = rstSQL
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		'Signal successful completion
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Function
		
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.
		System.EnterpriseServices.ContextUtil.SetAbort()
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		'UPGRADE_NOTE: Object rstSQL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstSQL = Nothing
		Err.Raise(Err.Number, Err.Source, Err.Description)
		
	End Function
	
	Public Function Insert(ByVal strDescription As String) As Integer
		Dim ApptDB As Object
		'--------------------------------------------------------------------
		'Date: 08/28/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Inserts a row into the tblApptStatus table.          '
		'Parameters: strDescription - The description of the Appt Status    '
		'              that will be inserted into the table.                '
		'Returns: ID (Primary Key) of the row inserted                      '
		'--------------------------------------------------------------------
		Dim objAppt As ApptDB.CApptStatusDB
		Dim lngID As Integer
		Dim strErrMessage As String
		
		On Error GoTo ErrTrap
		
		'Verify data before proceeding.
		If Not VerifyData(0, strDescription, strErrMessage) Then
			GoTo ErrTrap
		End If
		
		objAppt = CreateObject("ApptDB.CApptStatusDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.Insert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngID = objAppt.Insert(strDescription)
		
		Insert = lngID
		
		'Signal successful completion
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		'Release resources
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		Exit Function
		
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.  The
		'condition handles custom business rule checks we may have established.
		System.EnterpriseServices.ContextUtil.SetAbort()
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMessage)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
		
	End Function
	
	
	Public Sub Update(ByVal lngID As Integer, ByVal strDescription As String)
		Dim ApptDB As Object
		'--------------------------------------------------------------------
		'Date: 08/28/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Updates a row into the tblApptStatus table.          '
		'Parameters:  lngID - ID of the row in the table whose value will be'
		'               updated.                                            '
		'             strDescription - The appointment status description   '
		'                to which the record will be changed.               '
		'Returns: Null                                                      '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'                                                                   '
		'--------------------------------------------------------------------
		Dim objAppt As ApptDB.CApptStatusDB
		Dim strErrMessage As String
		
		On Error GoTo ErrTrap
		
		'Verify data before proceeding.
		If Not VerifyData(lngID, strDescription, strErrMessage) Then
			GoTo ErrTrap
		End If
		
		objAppt = CreateObject("ApptDB.CApptStatusDB")
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call objAppt.Update(lngID, strDescription)
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		'Signal successful completion
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Sub
		
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.  The
		'condition handles custom business rule checks we may have established.
		System.EnterpriseServices.ContextUtil.SetAbort()
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		If Err.Number = 0 Then
			Err.Raise(vbObjectError, CLASS_NAME, strErrMessage)
		Else
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
		
	End Sub
	
	
	Public Sub Deleted(ByVal blnDeleted As Boolean, ByVal lngID As Integer)
		Dim ApptDB As Object
		'--------------------------------------------------------------------
		'Date: 08/28/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Flags a row in the tblApptStatus table marking the row
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
		Dim objAppt As ApptDB.CApptStatusDB
		
		On Error GoTo ErrTrap
		
		objAppt = CreateObject("ApptDB.CApptStatusDB")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.Deleted. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call objAppt.Deleted(blnDeleted, lngID)
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		'Signal successful completion
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Sub
		
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.
		System.EnterpriseServices.ContextUtil.SetAbort()
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		Err.Raise(Err.Number, Err.Source, Err.Description)
		
	End Sub
	
	
	Public Function Exists(ByVal strDescription As String) As Boolean
		Dim ApptDB As Object
		'--------------------------------------------------------------------
		'Date: 08/28/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Determines if an appoinment status description       '
		'               identical to the strDescription parameter already   '
		'               exists in the table.                                '
		'Parameters: strDescription - Appointment status name to be checked '
		'Returns: True if the name exists, false otherwise                  '
		'--------------------------------------------------------------------
		'Revision History:                                                  '
		'                                                                   '
		'--------------------------------------------------------------------
		Dim objAppt As ApptDB.CApptStatusDB
		Dim blnExists As Boolean
		
		On Error GoTo ErrTrap
		
		objAppt = CreateObject("ApptDB.CApptStatusDB")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objAppt.Exists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		blnExists = objAppt.Exists(strDescription)
		Exists = blnExists
		
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		
		'Signal successful completion
		System.EnterpriseServices.ContextUtil.SetComplete()
		
		Exit Function
		
ErrTrap: 
		'Signal incompletion and raise the error to the calling environment.
		System.EnterpriseServices.ContextUtil.SetAbort()
		'UPGRADE_NOTE: Object objAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAppt = Nothing
		Err.Raise(Err.Number, Err.Source, Err.Description)
	End Function
	
	Private Function VerifyData(ByVal lngID As Integer, ByVal strDescription As String, ByRef strErrMessage As String) As Boolean
		'--------------------------------------------------------------------
		'Date: 08/28/2000                                                   '
		'Author: Rick "Boom Boom" Segura                                    '
		'Description:  Verifies all required data has been provided by the user.
		'Parameters:  The values to be checked.                             '
		'Returns: Boolean value identifying if all criteria has been satisfied.
		'--------------------------------------------------------------------
		
		If Trim(strDescription) = "" Then
			strErrMessage = "Status description is required."
			VerifyData = False
			Exit Function
		End If
		
		'Check for existance only when inserting new data
		If lngID = 0 And Exists(strDescription) Then
			strErrMessage = "Appointment Status '" & strDescription & "' already exists."
			VerifyData = False
			Exit Function
		End If
		
		'If we get here, all is well...
		VerifyData = True
		
	End Function
End Class