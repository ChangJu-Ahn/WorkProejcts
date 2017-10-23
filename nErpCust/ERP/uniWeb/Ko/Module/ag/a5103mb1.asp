<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Open A/P Confirm
'*  3. Program ID           : a5103mb1
'*  4. Program Name         : 결의전표승인 
'*  5. Program Desc         : 결의전표에 대하여 승인 또는 승인취소하는 기능 
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +
'*  7. Modified date(First) : 2000/10/14
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Chang Goo,Kang
'* 10. Modifier (Last)      : Ahn Hae Jin
'* 11. Comment              :
'*
'**********************************************************************************************
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../ag/incAcctMBFunc.asp"  -->

<% 
Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")

'    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Dim lgStrPrevKeyTempGlDt	' 이전결의전표일 
	Dim lgStrPrevKeyTempGlNo	' 이전결의전표번호 
    Call HideStatusWnd                                                               '☜: Hide Processing message

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL
    
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                          '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                          '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet

    lgLngMaxRow       = CInt(Request("txtMaxRows"))                                 '☜: Read Operation Mode (CRUD)
    lgMaxCount        = Request("lgMaxCount")										'☜: Fetch count at a time for VspdData
    
    lgStrPrevKeyTempGlDt = Trim(Request("lgStrPrevKeyTempGlDt"))
    lgStrPrevKeyTempGlNo = Trim(Request("lgStrPrevKeyTempGlNo"))
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

   ' Call SubOpenDB(lgObjConn)                                                       '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
    End Select

    'Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Const C_SHEETMAXROWS	= 100
	
    Dim PAGG0015_cAListTmpGlSvr
    
    Dim iStrData    
    Dim iLngRow,iLngCol        
    Dim iIntMaxCount
    Dim iIntLoopCount
        
    Dim pvSheetMaxRowsD
    Dim I1_from_temp_gl_dt
    Dim I2_to_temp_gl_dt
    Dim I3_b_acct_dept
    Dim I4_conf_fg
    Dim I5_gl_input_type
    Dim I6_pvPrevKey
    Dim l7_from_temp_gl_no
    Dim l8_to_temp_gl_no
    Dim l9_from_gl_no
    Dim l10_to_gl_no
    Dim l11_Ref_no
    Dim I12_eWare_flag
    Dim I13_bizarea
    
    Dim E1_b_acct_dept
    Dim E2_bizarea
    Dim EG1_export_grp_temp_gl
    
    Const A051_I3_org_change_id = 0
    Const A051_I3_dept_cd = 1

    Const A051_I6_temp_gl_dt = 0
    Const A051_I6_temp_gl_no = 1
    
    Const A051_E1_dept_cd = 0
    Const A051_E1_dept_nm = 1
    Const A051_E1_gl_input_type = 2
    Const A051_E1_gl_input_type_nm = 3
    
    Const A051_EG1_conf_fg = 0
    Const A051_EG1_E2_temp_gl_dt = 1
    Const A051_EG1_E2_issued_dt = 2
    Const A051_EG1_E2_temp_gl_no = 3
    Const A051_EG1_E1_dept_nm = 4
    Const A051_EG1_E2_dr_amt = 5
    Const A051_EG1_E2_dr_loc_amt = 6
    Const A051_EG1_E3_gl_no = 7
    Const A051_EG1_E2_temp_gl_desc = 8	'2003.02.13 추가 

    ReDim I3_b_acct_dept(1)
    ReDim I6_pvPrevKey(1)
    ReDim I13_bizarea(1)
    ReDim E1_b_acct_dept(3)
    ReDim E2_bizarea(3)
    
	' 권한관리 추가 
	Dim arrAuth(3)

	arrAuth(0) = lgAuthBizAreaCd
	arrAuth(1) = lgInternalCd
	arrAuth(2) = lgSubInternalCd
	arrAuth(3) = lgAuthUsrID
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    pvSheetMaxRowsD 	= C_SHEETMAXROWS    
    I1_from_temp_gl_dt	= UNIConvDate(Request("txtFromReqDt"))
    I2_to_temp_gl_dt	= UNIConvDate(Request("txtToReqDt"))
    I3_b_acct_dept(0)	= Trim(request("hOrgChangeId"))                   'GetGlobalInf("gChangeOrgId")

    I3_b_acct_dept(1)	= Trim(Request("txtDeptCd"))
    I4_conf_fg			= Trim(Request("cboConfFg"))
    I5_gl_input_type	= Trim(Request("txtGlInputType"))
    I6_pvPrevKey(0)		= lgStrPrevKeyTempGlDt
    I6_pvPrevKey(1)		= lgStrPrevKeyTempGlNo  
    l7_from_temp_gl_no	= Request("txtTempGlNoFr")
    l8_to_temp_gl_no	= Request("txtTempGlNoTo")
    l9_from_gl_no	= Request("txtGlNoFr")
    l10_to_gl_no	= Request("txtGlNoTo")
    l11_Ref_no	= Request("txtRefNo")

    
	If gEware = "" Then
		I12_eWare_flag   = "N"
	Else
		I12_eWare_flag   = "Y"	
	End If		
	
	I13_bizarea(0)		= Request("txtBizAreaCd")
	I13_bizarea(1)		= Request("txtBizAreaCd1")
	
	Set PAGG0015_cAListTmpGlSvr = Server.CreateObject("PAGG015.cAListTmpGlSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If    

	Call PAGG0015_cAListTmpGlSvr.A_LIST_TEMP_GL_SVR(gStrGlobalCollection, _
													Clng(C_SHEETMAXROWS), _
													I1_from_temp_gl_dt, _
													I2_to_temp_gl_dt, _
													I3_b_acct_dept, _
													I4_conf_fg, _
													I5_gl_input_type, _
													I6_pvPrevKey, _
													l7_from_temp_gl_no, _
													l8_to_temp_gl_no, _
													l9_from_gl_no, _
													l10_to_gl_no, _
													l11_Ref_no, _
													E1_b_acct_dept, _
													EG1_export_grp_temp_gl, _
													I12_eWare_flag, _
													I13_bizarea, _
													E2_bizarea, _
													arrAuth)

	If CheckSYSTEMError(Err, True) = True Then		
		Set PAGG015_cACnfmTmpGlSvr = Nothing
		Call SetErrorStatus		
		Exit Sub
	End If    
	
	If lgErrorStatus <> "YES" Then	
		Set PAGG0015_cAListTmpGlSvr = Nothing    
		IStrData = ""	
		iIntLoopCount = 0
	
		If isEmpty(EG1_export_grp_temp_gl) = False Then
			For iLngRow = 0 To UBound(EG1_export_grp_temp_gl, 1)
				iIntLoopCount = iIntLoopCount + 1

			    If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
					iStrData = iStrData & Chr(11) & "0"

					For iLngCol = 0 To UBound(EG1_export_grp_temp_gl, 2)
						If iLngCol = A051_EG1_E2_temp_gl_dt  Then '1
							iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_grp_temp_gl(iLngRow, iLngCol))
						Elseif iLngCol = A051_EG1_E2_issued_dt Then	'2
							If EG1_export_grp_temp_gl(iLngRow, iLngCol) <> "" Then
								iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_grp_temp_gl(iLngRow, iLngCol))
							Else
								iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_grp_temp_gl(iLngRow, iLngCol - 1))
							End If
						ElseIf iLngCol = A051_EG1_E1_dept_nm  Then
							iStrData = iStrData & Chr(11) & EG1_export_grp_temp_gl(iLngRow, iLngCol)
							iStrData = iStrData & Chr(11) & ""
						ElseIf iLngCol = A051_EG1_E2_dr_amt  Or iLngCol = A051_EG1_E2_dr_loc_amt Then
							iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_grp_temp_gl(iLngRow, iLngCol), ggAmtOfMoney.DecPoint,0)
						Else
							iStrData = iStrData & Chr(11) & EG1_export_grp_temp_gl(iLngRow, iLngCol)
						End If
					Next

		            iStrData = iStrData & Chr(11) & Cstr(iLngRow + 1 + lgLngMaxRow) & Chr(11) & Chr(12)
			    Else
					lgStrPrevKeyTempGlDt = EG1_export_grp_temp_gl(UBound(EG1_export_grp_temp_gl, 1), 1)
					lgStrPrevKeyTempGlNo = EG1_export_grp_temp_gl(UBound(EG1_export_grp_temp_gl, 1), 3)
					Exit For
				End If
			Next
			
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				lgStrPrevKeyTempGlDt = ""
				lgStrPrevKeyTempGlNo = ""
			End If
		End If
	End If

	Response.Write " <Script Language=vbscript>											 " & vbCr
	Response.Write " With parent														 " & vbCr
	Response.Write "	.ggoSpread.Source    = .frm1.vspdData							 " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & ConvSPChars(iStrData)			& """" & vbCr
	Response.Write "	.lgStrPrevKeyTempGlDt	= """ & lgStrPrevKeyTempGlDt		& """" & vbCr
	Response.Write "	.lgStrPrevKeyTempGlNo	= """ & lgStrPrevKeyTempGlNo		& """" & vbCr
	Response.Write "	.frm1.hFromReqDt.value	= """ & Trim(Request("FromReqDt"))  & """" & vbCr
	Response.Write "	.frm1.hToReqDt.value	= """ & Trim(Request("ToReqDt"))    & """" & vbCr
	Response.Write "	.frm1.hcboConfFg.value	= """ & Trim(Request("cboConfFg"))  & """" & vbCr
	Response.Write "	.frm1.txtDeptCd.value	= """ & Trim(E1_b_acct_dept(0))		& """" & vbCr
	Response.Write "	.frm1.hDeptCd.value	    = """ & Trim(E1_b_acct_dept(0))		& """" & vbCr
	Response.Write "	.frm1.txtDeptNm.value	= """ & Trim(E1_b_acct_dept(1))		& """" & vbCr	
	Response.Write "	.frm1.txtGlInputType.value	= """ & Trim(E1_b_acct_dept(2))		& """" & vbCr	
	Response.Write "	.frm1.txtGlInputTypeNm.value	= """ & Trim(E1_b_acct_dept(3))		& """" & vbCr	
	Response.Write "	.frm1.txtBizAreaCd.value	= """ & Trim(E2_bizarea(0))		& """" & vbCr
	Response.Write "	.frm1.txtBizAreaNm.value	= """ & Trim(E2_bizarea(1))		& """" & vbCr
	Response.Write "	.frm1.txtBizAreaCd1.value	= """ & Trim(E2_bizarea(2))		& """" & vbCr
	Response.Write "	.frm1.txtBizAreaNm1.value	= """ & Trim(E2_bizarea(3))		& """" & vbCr
	Response.Write "	.frm1.htxtBizAreaCd.value	= """ & Trim(E2_bizarea(0))		& """" & vbCr
	Response.Write "	.frm1.htxtBizAreaCd1.value	= """ & Trim(E2_bizarea(2))		& """" & vbCr		
	Response.Write "	If """ & lgErrorStatus & """ = ""NO"" then		" & vbCr
	Response.Write "		.DBQueryOk(""" & 1 & """)								" & vbCr	
	Response.Write "	Else											" & vbCr
	Response.Write "		.frm1.txtDeptCd.value			=		""""" & vbCr
	Response.Write "		.frm1.txtDeptNm.value			=		""""" & vbCr	
	Response.Write "		.frm1.txtGlInputType.value		=		""""" & vbCr
	Response.Write "		.frm1.txtGlInputTypeNm.value	=		""""" & vbCr
	Response.Write "		.frm1.vspdData.Focus						" & vbCr	
	Response.Write "	End If											" & vbCr
	Response.Write " End With											" & vbCr
	Response.Write " </Script>											" & vbCr                                                        '☜: Release RecordSSet		
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim PAGG015_cACnfmTmpGlSvr

	Dim import_grp
	Dim arrVal, arrTemp
	Dim	lGrpCnt	
	Dim LngMaxRow
	Dim LngRowItem
	Dim strStatus
    
	Dim iCommandSent
	Dim I1_from_temp_gl_dt
	Dim I2_to_temp_gl_dt
	Dim I3_b_acct_dept
	Dim I4_gl_input_type
	Dim I5_issued_dt
	Dim l6_from_temp_gl_no
	Dim l7_to_temp_gl_no
	Dim I9_a_data_auth
	Dim IG1_import_grp_temp_gl
    Dim iErrorPosition
    
    Const A051_I9_a_data_auth_data_BizAreaCd = 0
    Const A051_I9_a_data_auth_data_internal_cd = 1
    Const A051_I9_a_data_auth_data_sub_internal_cd = 2
    Const A051_I9_a_data_auth_data_auth_usr_id = 3    
    
    Redim I3_b_acct_dept(1)
    Redim I9_a_data_auth(3)
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    iCommandSent	   = ""  
    I1_from_temp_gl_dt = UNIConvDate(Request("txtFromReqDt"))    
    I2_to_temp_gl_dt   = UNIConvDate(Request("txtToReqDt"))     
    I3_b_acct_dept(0)  = Trim(request("hOrgChangeId"))			'GetGlobalInf("gChangeOrgId")	
    I3_b_acct_dept(1)  = Trim(Request("txtDeptCd"))
    I4_gl_input_type   = Trim(Request("txtGlInputType"))
    I5_issued_dt	   = UNIConvDate(Trim(Request("GIDate")))
    l6_from_temp_gl_no = Request("txtTempGlNoFr")
    l7_to_temp_gl_no   = Request("txtTempGlNoTo")
    
	I9_a_data_auth(A051_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A051_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A051_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A051_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))    
    
    LngMaxRow = CInt(Request("txtMaxRows"))										'☜: 최대 업데이트된 갯수 
	
	If LngMaxRow > 0 Then	
	    Set PAGG015_cACnfmTmpGlSvr = Server.CreateObject("PAGG015.cACnfmTmpGlSvr")
	    
	    If CheckSYSTEMError(Err, True) = True Then					       
	       Call SetErrorStatus
	       Exit Sub
	    End If    

	    '------------------------------------------------------------------------------------------------------
	    'Data manipulate area
	    '-----------------------
	    ' Data 연결 규칙 
	    ' 0: Flag , 1: Row위치, 2~N: 각 데이타 
	    '------------------------------------------------------------------------------------------------------
		
		arrTemp = Split(Request("txtSpread"), gRowSep)	'ITEM SPREAD
		lGrpCnt	= 0		
	    Redim IG1_import_grp_temp_gl(UBound(arrTemp,1) -1,3)
		
		For LngRowItem = 1 To LngMaxRow
			
		    lGrpCnt = lGrpCnt +1
			arrVal = Split(arrTemp(LngRowItem -1), gColSep)
			strStatus = arrVal(0)														'☜: Row 의 상태	
			Select Case strStatus
					Case "U"
					IG1_import_grp_temp_gl(LngRowItem-1,0) = Trim(arrVal(3))
					IG1_import_grp_temp_gl(LngRowItem-1,1) = Trim(arrVal(1))
					IG1_import_grp_temp_gl(LngRowItem-1,2) = UNIConvDate(arrVal(2))					
					IG1_import_grp_temp_gl(LngRowItem-1,3) = Trim(arrVal(4))					
		    End Select
	    Next
	  
		Call PAGG015_cACnfmTmpGlSvr.A_CONFIRM_TEMP_GL_SVR(gStrGlobalCollection, _
														iCommandSent, _
														I1_from_temp_gl_dt, _
														I2_to_temp_gl_dt, _
														I3_b_acct_dept, _
														I4_gl_input_type, _
														I5_issued_dt, _
														l6_from_temp_gl_no, _
														l7_to_temp_gl_no, _
														IG1_import_grp_temp_gl, _
														iErrorPosition, _
														gDsnNo, _
														I9_a_data_auth)		

			
		If CheckSYSTEMError2(Err, True,iErrorPosition & "","","","","") = True Then			
			Set PAGG015_cACnfmTmpGlSvr = Nothing
			Call SetErrorStatus
		End If	
	    
	    Set PAGG015_cACnfmTmpGlSvr = Nothing
    End If
    
    Response.Write " <Script Language=vbscript>	                        " & vbCr	
	Response.Write "	If """ & lgErrorStatus & """ = ""NO"" then		" & vbCr
	Response.Write "		Parent.DbSaveOk()							" & vbCr
	Response.Write "	Else											" & vbCr
	Response.Write "		Parent.Frm1.vspdData.Focus					" & vbCr
	Response.Write "	End If											" & vbCr	
	Response.Write " </Script>											" & vbCr  
End Sub    

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

%>


