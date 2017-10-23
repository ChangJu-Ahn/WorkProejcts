<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%

	
	Dim lgOpModeCRUD
	On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
    '---------------------------------------Common-----------------------------------------------------------
	 
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	
    Dim iPB6SA15
    Dim iStrData
    Dim exportData2
    Dim exportData
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrCostCenterCd
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
	Dim txtCOST_CENTER_CD
    
    Const C_SHEETMAXROWS_D	= 100
    
    Const C_QueryConut		= 0
    Const C_MaxQueryReCord	= 1
    Const C_COST_CENTER_CD	= 2
   
    Const COST_CD			= 0
	Const COST_NM			= 1
	Const COST_ENG_NM		= 2
	Const COST_TYPE			= 3
	Const DI_FG				= 4
	Const BIZ_AREA_CD		= 5
	Const BIZ_AREA_NM		= 6
	Const BIZ_UNIT_CD		= 7
	Const BIZ_UNIT_NM		= 8
	Const PLANT_CD			= 9
	Const PLANT_NM			= 10
	Const C_ORG_CHANGE_ID	= 11
	Const C_INTERNAL_CD		= 12
	Const C_DEPT_CD			= 13
	Const C_DEPT_NM			= 14
	Const C_CHKFLAG			= 15	

	On Error Resume Next
    Err.Clear
    ReDim exportData(1)
	Const A552_E2_cost_cd	= 0
	Const A552_E2_cost_nm	= 1


	iStrPrevKey				= Trim(Request("lgStrPrevKey"))  
	txtCOST_CENTER_CD		= Trim(Request("txtCOST_CENTER_CD"))  
    iIntMaxRows				= Request("txtMaxRows")
    iIntQueryCount			= Request("lgPageNo")

    If iStrPrevKey = "" Then
		iStrCostCenterCd	= txtCOST_CENTER_CD
	Else
		iStrCostCenterCd	= iStrPrevKey
    End If

    If Len(Trim(iIntQueryCount))  Then                                        '¢Ð : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)
       End If
    Else   
       iIntQueryCount = 0
    End If
    ReDim importArray(2)
    importArray(C_QueryConut)		= iIntQueryCount
    importArray(C_MaxQueryReCord)	= C_SHEETMAXROWS_D
    importArray(C_COST_CENTER_CD)   = iStrCostCenterCd

	Set iPB6SA15 = Server.CreateObject("PB6SA15.cbListCoCenterSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iPB6SA15 = Nothing
       Exit Sub
    End If

	Call iPB6SA15.B_LIST_COST_CENTER_SVR(gStrGloBalCollection, importArray,txtCOST_CENTER_CD, exportData, exportData2)

    If CheckSYSTEMError(Err, True) = True Then
       Set iPB6SA15 = Nothing
       Exit Sub
    End If



    Set iPB6SA15 = Nothing

	iStrData = ""
	iIntLoopCount = 0
	If IsEmpty(exportData2) = False Then
		For iLngRow = 0 To UBound(exportData2, 1)
			iIntLoopCount = iIntLoopCount + 1

			If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then 
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, COST_CD))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, COST_NM))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, COST_ENG_NM))
				iStrData = iStrData & Chr(11) & Trim(exportData2(iLngRow, COST_TYPE))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & Trim(exportData2(iLngRow, DI_FG))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, BIZ_AREA_CD))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, BIZ_AREA_NM))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, BIZ_UNIT_CD))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, BIZ_UNIT_NM))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, PLANT_CD))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, PLANT_NM))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_ORG_CHANGE_ID))
				iStrData = iStrData & Chr(11) & ""				
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_INTERNAL_CD))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_DEPT_CD))
				iStrData = iStrData & Chr(11) & ""	
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_DEPT_NM))
				IF exportData2(iLngRow, C_CHKFLAG) = "Y" Then
					iStrData = iStrData & Chr(11) & 1
				ELSE
					iStrData = iStrData & Chr(11) & 0
				END IF
				iStrData = iStrData & Chr(11) & Cstr(iLngRow + 1) & Chr(11) & Chr(12)
			Else
				iStrPrevKey = exportData2(UBound(exportData2, 1), 0)
				iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next
	End If

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.frm1.txtCost_Center_Nm.value = """ & ConvSPChars(exportData(A552_E2_cost_nm))    & """" & vbCr
	Response.Write "	.frm1.htxtCOST_CENTER_CD.value = """ & ConvSPChars(exportData(A552_E2_cost_cd))    & """" & vbCr
    Response.Write "	.lgPageNo = """ & iIntQueryCount    & """" & vbCr
    Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)    & """" & vbCr
    Response.Write "	.DbQueryOK " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim iPB6SA15 
	Dim iErrorPosition

	On Error Resume Next
    Err.Clear

    Set iPB6SA15 = Server.CreateObject("PB6SA15.cbMngCoCenterSvr")
        
    If CheckSYSTEMError(Err, True) = True Then
       Set iPB6SA15 = Nothing
       Exit Sub
    End If

    Call iPB6SA15.B_MANAGE_COST_CENTER_SVR(gStrGloBalCollection, Request("txtSpread"),iErrorPosition)
		
  	If CheckSYSTEMError2(Err, True,iErrorPosition & "Çà","","","","") = True Then				
       Set iPB6SA15 = Nothing
       Exit Sub
    End If

    Set iPB6SA15 = Nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()	
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	On Error Resume Next
    Err.Clear
End Sub

%>