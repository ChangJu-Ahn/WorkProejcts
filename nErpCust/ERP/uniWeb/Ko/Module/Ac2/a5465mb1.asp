
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5123mb1
'*  4. Program Name         : 회계전표일괄생성 
'*  5. Program Desc         : 각 모쥴에서 생성한 자료를 토대로 일괄적으로 전표처리.
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/09/26 : ..........
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../ag/incAcctMBFunc.asp"  -->
<% 
Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Dim lgStrPrevKeyTempGlDt	
	Dim lgStrPrevKeyBatchNo
    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet

    lgLngMaxRow             = CInt(Request("txtMaxRows"))                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount              = Request("lgMaxCount")                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyTempGlDt    = Trim(Request("lgStrPrevKeyTempGlDt"))
    lgStrPrevKeyBatchNo     = Trim(Request("lgStrPrevKeyBatchNo"))
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    'Call SubOpenDB(lgObjConn)  '11/20정승기                                                      '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
    End Select

    'Call SubCloseDB(lgObjConn)  '11/20정승기                                                     '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Const C_SHEETMAXROWS	= 100
	
	Const A228_I2_gl_dt_from        = 0
    Const A228_I2_gl_dt_to          = 1    
    
    Const A228_I3_gl_dt_previous    = 0
    Const A228_I3_batch_no_previous = 1
    Const A228_I3_gl_input_type     = 2
    Const A228_I3_auto_trans_fg     = 3

    Const A228_I4_from_ref_no    = 0
    Const A228_I4_to_ref_no    = 1
    
    Dim PAGG115_cAListBtchSvr
        
    Dim I1_b_biz_area
    Dim I2_a_batch
    Dim I3_a_batch
    Dim E1_a_batch
    Dim E2_a_batch
    Dim EG1_export_group
	Dim I4_ref_no
	Dim I6_Bp_cd
	Dim I7_trans_type
	
    Dim iStrData			
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iIntMaxRows
    Dim iIntMaxCount
    Dim iIntLoopCount
    
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
  
    
    
    ReDim I2_a_batch(1)
    ReDim I3_a_batch(3)
	ReDim I4_ref_no(1)
                
    I1_b_biz_area							= FilterVar(Trim(Request("txtBizCd")),"","S")
        
    I2_a_batch(A228_I2_gl_dt_from)			= UNIConvDate(Request("txtFromReqDt"))
    I2_a_batch(A228_I2_gl_dt_to)			= UNIConvDate(Request("txtToReqDt"))
    
    I3_a_batch(A228_I3_gl_dt_previous)		= UNIConvDate(Request("lgStrPrevKeyTempGlDt"))
    I3_a_batch(A228_I3_batch_no_previous)	= FilterVar(Trim(Request("lgStrPrevKeyBatchNo")),"","S")
    I4_ref_no(A228_I4_from_ref_no)		= Trim(Request("txtfrRefNo"))
    I4_ref_no(A228_I4_to_ref_no)		= Trim(Request("txttoRefNo"))
    
    I6_Bp_cd		= Trim(Request("txtBpcd"))
    I7_trans_type		= Trim(Request("txtTransType"))
    
    IF UCase(Trim(Request("cboConfFg")))	= "C" Then
		I3_a_batch(A228_I3_auto_trans_fg)	= "Y"
	Else
		I3_a_batch(A228_I3_auto_trans_fg)	= "N"
	End IF
	
     
	Set PAGG115_cAListBtchSvr = Server.CreateObject("PAGG115.cAListBtchSvr")	
	  
	
    If CheckSYSTEMError(Err, True) = True Then
		Call SetErrorStatus()
		Exit Sub
    End If    

	
	Call PAGG115_cAListBtchSvr.A_LIST_BATCH_SVR(gStrGlobalCollection, _
												C_SHEETMAXROWS, _
												I1_b_biz_area, _
												I2_a_batch, _
												I3_a_batch, _
												E1_a_batch, _
												E2_a_batch, _
												EG1_export_group, _
												1, I4_ref_no, _
												I6_Bp_cd, _
												I7_trans_type )
	
'	if err.number <> 0 then
'		Response.Write "xx" & "::" & err.source & "::" & err.description
'		Response.End
'	end if
	
	If CheckSYSTEMError(Err, True) = True Then			
		Set PAGG115_cAListBtchSvr = Nothing
		Call SetErrorStatus
		'Exit Sub
	End If
	
    If lgErrorStatus <> "YES" Then
    
		Set PAGG115_cAListBtchSvr = Nothing
		iStrData = ""
		iIntLoopCount = 0
		IF isempty(EG1_export_group) = FALSE Then
			For iLngRow = 0 To UBound(EG1_export_group, 1) 
				
				iIntLoopCount = iIntLoopCount + 1
			    
			    If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
					
					iStrData = iStrData & Chr(11) & "0"
								
					For iLngCol = 0 To UBound(EG1_export_group, 2)
						IF iLngCol = 0 or iLngCol = 8 Then 
							iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, iLngCol))
						ELSE
							iStrData = iStrData & Chr(11) & EG1_export_group(iLngRow, iLngCol)
						END IF
					Next
                        iStrData = iStrData & Chr(11) & Cstr(iLngRow + 1 + lgLngMaxRow) & Chr(11) & Chr(12)
			    Else
					lgStrPrevKeyTempGlDt = EG1_export_group(UBound(EG1_export_group, 1), 0)
					lgStrPrevKeyBatchNo = EG1_export_group(UBound(EG1_export_group, 1), 1)
					Exit For
					  
				End If
			Next
			
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				lgStrPrevKeyTempGlDt = ""
				lgStrPrevKeyBatchNo = ""
			End If
		END IF
	
	End If
	
		
	
	
		Response.Write " <Script Language=vbscript>			" & vbCr
		Response.Write " With parent						" & vbCr
	
	IF not isEmpty(E1_a_batch) Then
		Response.Write "	.frm1.txtBizCd.value = """ & E1_a_batch(0)			& """" & vbCr 			 
		Response.Write "	.frm1.txtBizNm.value = """ & E1_a_batch(1)			& """" & vbCr			 
	ENd IF
	
	IF not isEmpty(E2_a_batch) Then
		Response.Write "	.frm1.txtGlInputType.value = """ & E2_a_batch(0)	& """" & vbCr 
		Response.Write "	.frm1.txtGlInputTypeNm.value = """ & E2_a_batch(1)	& """" & vbCr 		 
	END IF
		Response.Write "	.ggoSpread.Source = .frm1.vspdData						 " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & ConvSPChars(iStrData)	& """" & vbCr	
		Response.Write "	.lgStrPrevKeyTempGlDt = """ & lgStrPrevKeyTempGlDt  & """" & vbCr
		Response.Write "	.lgStrPrevKeyBatchNo = """ & lgStrPrevKeyBatchNo    & """" & vbCr
		Response.Write "	.frm1.hFromReqDt.value = """ & Trim(Request("FromReqDt"))			& """" & vbCr
		Response.Write "	.frm1.hToReqDt.value = """ & Trim(Request("ToReqDt"))				& """" & vbCr
		Response.Write "	.frm1.hGlInputType.value = """ & Trim(Request("txtGlInputType"))    & """" & vbCr
		Response.Write "	.frm1.hcboConfFg.value = """ & Trim(Request("cboConfFg"))			& """" & vbCr
		Response.Write "	.frm1.hBizCd.value = """ & Trim(Request("txtBizCd"))				& """" & vbCr
	If lgErrorStatus <> "NO" Then
		Response.Write "		.frm1.txtBizCd.value			= """"" & vbCr'
		Response.Write "		.frm1.txtBizNm.value			= """"" & vbCr	
		Response.Write "		.frm1.txtGlInputType.value		= """"" & vbCr
		Response.Write "		.frm1.txtGlInputTypeNm.value	= """"" & vbCr
	End If
		Response.Write " .DbQueryOk   " & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr                                                        '☜: Release RecordSSet
		
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()


	Const A377_I2_from_a_batch_gl_dt			= 0
	Const A377_I2_from_a_batch_gl_input_type	= 1

	Const A377_IG1_a_batch_batch_no				= 0
	Const A377_IG1_a_batch_auto_trans_fg		= 1
 
    Const A228_I4_from_ref_no    = 0
    Const A228_I4_to_ref_no    = 1


	Dim PAGG115_cAMngBtchToGlSvr
	Dim iCommandSent
	Dim I1_b_biz_area
	Dim I2_from_a_batch
	Dim I3_to_a_batch
	Dim IG1_a_batch	
	Dim iErrorPosition
	
	Dim iLngMaxRow
	Dim iLngRow
	Dim iArrTemp
	Dim iArrVal
	Dim iStrStatus
	Dim I4_ref_no
	Dim I6_Bp_cd
	Dim I7_trans_type
	
	ReDim I4_ref_no(1)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
																		'☜: Protect system from crashing

	'If Request("txtMaxRows") = "0" Then
		'Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		'Response.End 
	'End If

	iLngMaxRow = CInt(Request("txtMaxRows"))
    I4_ref_no(A228_I4_from_ref_no)		= Trim(Request("txtfrRefNo"))
    I4_ref_no(A228_I4_to_ref_no)		= Trim(Request("txttoRefNo"))
    
    I6_Bp_cd		= Trim(Request("txtBpcd"))
    I7_trans_type		= Trim(Request("txtTransType"))

	If iLngMaxRow > 0 then
		iArrTemp = Split(Request("txtSpread"), gRowSep)	
		ReDim IG1_a_batch(Ubound(iArrTemp,1) -1, A377_IG1_a_batch_auto_trans_fg)
	
		For iLngRow = 1 To iLngMaxRow    				
			iArrVal = Split(iArrTemp(iLngRow-1), gColSep)		
			iStrStatus = iArrVal(0)														'☜: Row 의 상태 
			Select Case iStrStatus
				Case "U"
					IG1_a_batch(iLngRow - 1, A377_IG1_a_batch_batch_no)		 = FilterVar(Trim(iArrVal(1)),"","S")
					IG1_a_batch(iLngRow - 1, A377_IG1_a_batch_auto_trans_fg) = iArrVal(2)
		    End Select
			    
		Next
	
		Set PAGG115_cAMngBtchToGlSvr = Server.CreateObject("PAGG115.cAMngBtchToGlSvr")
	
		If CheckSYSTEMError(Err, True) = True Then		
			Exit Sub
		End If
	
		Call PAGG115_cAMngBtchToGlSvr.A_MANAGE_BATCH_TO_GL_SVR(gStrGlobalCollection, _
																iCommandSent, _
																I1_b_biz_area, _
																I2_from_a_batch, _
																I3_to_a_batch, _
																IG1_a_batch, _
																iErrorPosition, _
																1, _ 
																I4_ref_no , _
																I6_Bp_cd, _
																I7_trans_type )

'		If CheckSYSTEMError(Err, True) = True Then		
'		   Set PAGG115_cAMngBtchToGlSvr = Nothing
'		   Exit Sub
'		End If
		
		If CheckSYSTEMErrorAcct2(Err, True,iErrorPosition & "","","","","") = True Then
			Set PAGG115_cAMngBtchToGlSvr = Nothing
			Call SetErrorStatus
			Exit Sub
		End If
	    
		Set PAGG115_cAMngBtchToGlSvr  = Nothing
	End If

Response.Write " <Script Language=vbscript>	                        " & vbCr
Response.Write " With parent                                        " & vbCr
Response.Write "	Window.status = ""저장 성공""					" & vbCr 
Response.Write "	.DbSaveOk										" & vbCr 			 
Response.Write " End With											" & vbCr
Response.Write " </Script>											" & vbCr                                              '
    
				

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)


End Sub




'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>


