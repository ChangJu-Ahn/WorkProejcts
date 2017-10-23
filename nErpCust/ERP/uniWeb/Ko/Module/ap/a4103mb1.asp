<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a4103mb1
'*  4. Program Name         : 채무승인조회 
'*  5. Program Desc         : 채무승인조회 
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/10/13
'*  8. Modified date(Last)  : 2003/04/03
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. 조건부 
'##########################################################################################################
																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
On Error Resume Next														'☜: 
Err.Clear 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'#########################################################################################################
'												2.1 조건 체크 
'##########################################################################################################
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then										'☜: 조회 전용 Biz 이므로 다른값은 그냥 종료함 
	Call ServerMesgBox("700118", vbInformation, I_MKSCRIPT)					'⊙: 조회 전용인데 다른 상태로 요청이 왔을 경우, 필요없으면 빼도 됨, 메세지는 ID값으로 사용해야 함 
	Response.End 
End If

'#########################################################################################################
'												2. 업무 처리 수행부 
'##########################################################################################################

'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
Dim iArrData
Dim iGData
DIm intCount
Dim IntRows
Dim LngMaxRow
Dim iPAPG015
Dim strData
Dim lgCurrency
Dim iIntLoopCount
Dim iIntMaxRows
Dim iIntQueryCount
Dim iStrPrevKey

Dim I1_a_open_AP_conf
Dim I2_a_open_AP_next
DIm E1_a_open_AP_conf
Dim EG1_export_group

Const C_SHEETMAXROWS_D  =100

Const A048_I1_a_open_AP_conf_start_dt = 0
Const A048_I1_a_open_AP_conf_end_dt = 1
Const A048_I1_a_open_AP_conf_fg = 2
Const A048_I1_a_open_AP_deal_bp_cd = 3
Const A048_I1_a_open_AP_dept_cd = 4
Const A048_I1_a_open_AP_org_change_id = 5

'I2_a_open_AP_next
Const A048_I2_a_open_AP_next_query_cnt = 0
Const A048_I2_a_open_AP_next_AP_no = 1
    
'E1_a_open_AP_conf
Const A048_E1_a_open_AP_conf_start_dt = 0
Const A048_E1_a_open_AP_conf_end_dt = 1
Const A048_E1_a_open_AP_conf_bp_cd = 2
Const A048_E1_a_open_AP_conf_bp_nm = 3
Const A048_E1_a_open_AP_conf_dept_cd = 4
Const A048_E1_a_open_AP_conf_dept_nm = 5
Const A048_E1_a_open_AP_conf_org_change_id = 6
Const A048_E1_a_open_AP_conf_fg = 7
Const A048_E1_a_open_AP_AP_loc_amt = 8    
'EG1_export_group
Const A048_EG1_a_open_AP_check_fg = 0
Const A048_EG1_a_open_AP_AP_DT = 1
Const A048_EG1_a_open_AP_Gl_DT = 2
Const A048_EG1_a_open_AP_AP_NO = 3
Const A048_EG1_a_open_AP_BP_NM = 4
Const A048_EG1_a_open_AP_DOC_CUR = 5
Const A048_EG1_a_open_AP_AP_AMT = 6
Const A048_EG1_a_open_AP_AP_LOC_AMT = 7
Const A048_EG1_a_open_AP_dept_cd = 8
Const A048_EG1_a_open_AP_TEMP_GL_NO = 9
Const A048_EG1_a_open_AP_GL_NO = 10
Const A048_EG1_a_open_AP_conf_fg = 11
    
' -- 권한관리추가 
Const A048_I3_a_data_auth_data_BizAreaCd = 0
Const A048_I3_a_data_auth_data_internal_cd = 1
Const A048_I3_a_data_auth_data_sub_internal_cd = 2
Const A048_I3_a_data_auth_data_auth_usr_id = 3

Dim I3_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

Redim I3_a_data_auth(3)
I3_a_data_auth(A048_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
I3_a_data_auth(A048_I3_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
I3_a_data_auth(A048_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
I3_a_data_auth(A048_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

'#########################################################################################################
'												2.2. 요청 변수 처리 
'##########################################################################################################
	LngMaxRow = Request("txtMaxRows")
'#########################################################################################################
'												2.3. 업무 처리 
'##########################################################################################################

	Set iPAPG015 = Server.CreateObject("PAPG015.cALkUpCnfmAPSvr")

	If CheckSYSTEMError(Err, True) = True Then					
		Response.End 
	End If    

	ReDim I1_a_open_AP_conf(A048_I1_a_open_AP_org_change_id)
	REdim I2_a_open_AP_next(A048_I2_a_open_AP_next_AP_no)
	
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
	iStrPrevKey     = Trim(Request("lgStrPrevKey"))
	
	I1_a_open_AP_conf(A048_I1_a_open_AP_conf_start_dt)	= Request("txtFromReqDt")
	I1_a_open_AP_conf(A048_I1_a_open_AP_conf_end_dt)	= Request("txtToReqDt")
	I1_a_open_AP_conf(A048_I1_a_open_AP_conf_fg)		= Trim(Request("cboConfFg"))
	I1_a_open_AP_conf(A048_I1_a_open_AP_deal_bp_cd)		= UCase(Trim(Request("txtBpCd")))
	I1_a_open_AP_conf(A048_I1_a_open_AP_dept_cd)		= UCase(Trim(Request("txtDeptCd")))
	I1_a_open_AP_conf(A048_I1_a_open_AP_org_change_id)	= Trim(request("OrgChangeId"))

	I2_a_open_AP_next(A048_I2_a_open_AP_next_query_cnt) = C_SHEETMAXROWS_D
	
	If iStrPrevKey = "" Then
		I2_a_open_AP_next(A048_I2_a_open_AP_next_AP_no) = ""
	Else
		I2_a_open_AP_next(A048_I2_a_open_AP_next_AP_no) = iStrPrevKey
    End If
		
	Call iPAPG015.A_LOOKUP_CONFIRM_AP_SVR(gStrGloBalCollection ,I1_a_open_AP_conf ,I2_a_open_AP_next, E1_a_open_AP_conf, EG1_export_group, I3_a_data_auth)
		
	If CheckSYSTEMError(Err, True) = True Then					
		Set iPAPG020 = Nothing
		Response.End 
	End If    

	Response.Write "<Script Language=vbscript>  " & vbcr
	Response.Write " With parent.frm1           " & vbcr														'☜: 화면 처리 ASP 를 지칭함 
	Response.Write ".txtDeptCd.Value		= """ & ConvSPChars(E1_a_open_AP_conf(A048_E1_a_open_AP_conf_dept_cd))			& """ " & vbcr
	Response.Write ".txtDeptNm.value		= """ & ConvSPChars(E1_a_open_AP_conf(A048_E1_a_open_AP_conf_dept_nm))			& """ " & vbcr
	Response.Write ".txtBpCd.value			= """ & ConvSPChars(E1_a_open_AP_conf(A048_E1_a_open_AP_conf_bp_cd))			& """ " & vbcr
	Response.Write ".txtBpNM.Value			= """ & ConvSPChars(E1_a_open_AP_conf(A048_E1_a_open_AP_conf_bp_nm))			& """ " & vbcr
	Response.Write ".hOrgChangeId.Value		= """ & ConvSPChars(E1_a_open_AP_conf(A048_E1_a_open_AP_conf_org_change_id))			& """ " & vbcr
	Response.Write ".txtTotAPLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(E1_a_open_AP_conf(A048_E1_a_open_AP_AP_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """ " & vbcr
	Response.Write " End With					" & vbcr		    
	Response.write "</Script>				    " & vbcr  

	iIntLoopCount = 0	

	If IsArray(EG1_export_group) Or IsEmpty(EG1_export_group) = False Then
		strData	 = ""		
		intCount = UBound(EG1_export_group,1)
	  
		For IntRows = 0 To intCount
			iIntLoopCount = iIntLoopCount + 1
		    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then					
				lgCurrency = ConvSPChars(EG1_export_group(IntRows,A048_EG1_a_open_AP_DOC_CUR))

				strData = strData & Chr(11) & "0"
				strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows,A048_EG1_a_open_AP_AP_DT))    
				strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows,A048_EG1_a_open_AP_Gl_DT))
				strData = strData & Chr(11) & UCase(ConvSPChars(EG1_export_group(IntRows,A048_EG1_a_open_AP_AP_NO)))
				strData = strData & Chr(11) & UCase(ConvSPChars(EG1_export_group(IntRows,A048_EG1_a_open_AP_BP_NM)))
				strData = strData & Chr(11) & UCase(ConvSPChars(EG1_export_group(IntRows,A048_EG1_a_open_AP_DOC_CUR)))
				strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A048_EG1_a_open_AP_AP_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
				strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A048_EG1_a_open_AP_AP_LOC_AMT), gCurrency,ggAmtOfMoneyNo, "X" , "X")
				strData = strData & Chr(11) & UCase(ConvSPChars(EG1_export_group(IntRows,A048_EG1_a_open_AP_dept_cd)))
				strData = strData & Chr(11) & UCase(ConvSPChars(EG1_export_group(IntRows,A048_EG1_a_open_AP_TEMP_GL_NO)))
				strData = strData & Chr(11) & UCase(ConvSPChars(EG1_export_group(IntRows,A048_EG1_a_open_AP_GL_NO)))
				strData = strData & Chr(11) & UCase(ConvSPChars(EG1_export_group(IntRows,A048_EG1_a_open_AP_conf_fg)))
			    
				strData = strData & Chr(11) & iIntMaxRows + iIntLoopCount
				strData = strData & Chr(11) & Chr(12)                                    
			Else
				iStrPrevKey = EG1_export_group(UBound(EG1_export_group, 1), A048_EG1_a_open_AP_AP_NO)
				iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If							
		Next
	End If
	
	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = "" 
	End If	

	Response.write "<Script Language=vbscript>									" & vbCr
	Response.write "With parent													" & vbCr												
	Response.Write ".ggoSpread.Source     = .frm1.vspdData						" & vbCr
	Response.Write ".frm1.vspdData.Redraw = False								" & vbCr
	Response.Write ".ggoSpread.SSShowData """ & strData					   & """" & vbCR
	Response.Write  " Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & iIntMaxRows+1  & "," & iIntMaxRows + iIntLoopCount & ",Parent.C_DOC_CUR,Parent.C_AP_AMT,""A"" ,""Q"",""X"",""X"")" & vbCr
	Response.Write ".frm1.vspdData.Redraw = True								" & vbCr
	Response.Write ".lgPageNo             = """ & iIntQueryCount           & """" & vbCr
	Response.Write ".lgStrPrevKey         = """ & ConvSPChars(iStrPrevKey) & """" & vbCr	
	Response.Write ".DbQueryOk													" & vbCr
	Response.write " End With 													" & vbCr													
	Response.Write "</Script>													" & VbCr

%>
