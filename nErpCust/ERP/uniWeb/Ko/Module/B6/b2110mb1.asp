<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->

<%

	On Error Resume Next
    Err.Clear

	Dim lgOpModeCRUD
    Call HideStatusWnd
    Call LoadBasisGlobalInf()
    '---------------------------------------Common-----------------------------------------------------------
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update             
             Call SubBizSaveMulti()
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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Dim PB6SA00LIST
	Dim iBizUnitCd
	Dim iStrData
	Dim iBizUnitNm
	Dim iBizUnit
	Dim iLngRow,iLngCol
	Dim txtBizUnitCd

	'==============½Å±Ô================
	Dim iIntQueryCount
	Dim iIntLoopCount
	Dim iStrBizUnitCd
	Dim iStrPrevKey
	Dim iIntMaxRows
	'==================================

    Const C_SHEETMAXROWS_D = 100

	On Error Resume Next
    Err.Clear

	iStrPrevKey		= Trim(Request("lgStrPrevKey"))
    iIntMaxRows		= Request("txtMaxRows")

    iIntQueryCount	= Request("lgPageNo")
	txtBizUnitCd	= Request("txtBizUnitCd")

	ReDim iBizUnitNm(1)
	Const 	EA_b_biz_unit_biz_unit_cd		= 0
	Const 	EA_b_biz_unit_biz_unit_nm		= 1

	If iStrPrevKey = "" Then
		iStrBizUnitCd	= txtBizUnitCd
	Else
		iStrBizUnitCd	= iStrPrevKey
    End If


	If Len(Trim(iIntQueryCount))  Then                                        '¢Ð : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)
       End If   
    Else   
       iIntQueryCount = 0
    End If

    Set PB6SA00LIST = Server.CreateObject("PB6SA00.cBListBizUnitSvr")

	If CheckSYSTEMError(Err, True) = True Then
       Set PB6SA00LIST = Nothing
       Exit Sub
    End If

    Call PB6SA00LIST.B_LIST_BIZ_UNIT_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D,iStrBizUnitCd,txtBizUnitCd,iBizUnitNm,iBizUnit)
	
	If CheckSYSTEMError(Err, True) = True Then
       Set PB6SA00LIST = Nothing
       Exit Sub
    End If

    Set PB6SA00LIST = nothing

    iStrData = ""
    iIntLoopCount = 0
	For iLngRow = 0 To UBound(iBizUnit, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			For iLngCol = 0 To UBound(iBizUnit, 2)
			    iStrData = iStrData & Chr(11) & ConvSPChars(iBizUnit(iLngRow, iLngCol))
			Next
				iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
			    iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = iBizUnit(UBound(iBizUnit, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>								 " & vbCr
	Response.Write " With parent											 " & vbCr
    Response.Write "	.ggoSpread.Source		 = .frm1.vspdData			 " & vbCr
    Response.Write "	.ggoSpread.SSShowData	   """ & iStrData		& """" & vbCr
    Response.Write "	.frm1.txtBizUnitNm.value = """ & ConvSPChars(iBizUnitNm(EA_b_biz_unit_biz_unit_nm))		& """" & vbCr
    Response.Write "	.frm1.hBizUnitCd.value	 = """ & ConvSPChars(iBizUnitNm(EA_b_biz_unit_biz_unit_cd))		& """" & vbCr
    Response.Write "	.lgPageNo				 = """ & iIntQueryCount	& """" & vbCr
    Response.Write "	.lgStrPrevKey			 = """ & ConvSPChars(iStrPrevKey)	& """" & vbCr
    Response.Write "	.DbQueryOk	" & vbCr
    Response.Write "End With		" & vbCr
    Response.Write "</Script>		" & vbCr
    
End Sub    	 

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	Dim PB6SA00CUD
	Dim txtSpread
    Dim iErrorPosition

	On Error Resume Next
    Err.Clear

    Set PB6SA00CUD = Server.CreateObject("PB6SA00.cBMngBizUnitSvr")    

    If CheckSYSTEMError(Err, True) = True Then
       Set PB6SA00CUD = Nothing
       Exit Sub
    End If

	Call PB6SA00CUD.B_MANAGE_BIZ_UNIT_SVR(gStrGlobalCollection,request("txtSpread"),iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "Çà","","","","") = True Then
       Set PB6SA00CUD = Nothing
       Exit Sub
    End If

    Set PB6SA00CUD = nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub    

'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
	On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
	On Error Resume Next
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
End Sub

'============================================================================================================
Sub SetErrorStatus()
End Sub
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	On Error Resume Next
    Err.Clear
End Sub

%>

