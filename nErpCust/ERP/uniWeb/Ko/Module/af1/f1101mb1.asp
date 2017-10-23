<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f1101mb1
'*  4. Program Name         : 예산통제기간등록 
'*  5. Program Desc         : Register of Control Period
'*  6. Comproxy List        : FB0011, FB0019
'*  7. Modified date(First) : 2000.09.18
'*  8. Modified date(Last)  : 2001.12.04
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Oh, Soo Min
'* 11. Comment              :
'=======================================================================================================

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next														'☜: 
Err.Clear

Call LoadBasisGlobalInf()

Dim lPtxtNoteNo
Dim strMode
Dim txtCtrlYR
Dim cboctrlunit											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 


strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

txtCtrlYR = Request("txtCtrlYR")
cboctrlunit = Request("cboctrlunit")

Call HideStatusWnd

'------ Developer Coding part (End   ) ------------------------------------------------------------------     
Select Case strMode
    Case CStr(UID_M0001)                                                         '☜: Query
         Call SubBizQuery()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSave()
    Case CStr(UID_M0003)                                                         '☜: Save,Update
         Call SubBizDelete()
End Select

'==================================================================================
'	Name : SubBizQuery()
'	Description : 멀티조회 정의 
'==================================================================================
Sub SubBizQuery()

On Error Resume Next
Err.Clear 

Dim objPAFG105Q

Dim I1_f_bdg_perd
Const C_CTRL_YR = 0
Const C_CTRL_UNIT = 1

Dim E1_com_budget_yyyymm
Const E1_txt_1st_from_dt = 0 
Const E1_txt_1st_to_dt = 1
Const E1_txt_2st_from_dt = 2
Const E1_txt_2st_to_dt = 3
Const E1_txt_3st_from_dt = 4
Const E1_txt_3st_to_dt = 5
Const E1_txt_4st_from_dt = 6
Const E1_txt_4st_to_dt = 7



    	
    Set objPAFG105Q = Server.CreateObject("PAFG105.cFLkUpBdgPrdSvr")    

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
    
    ReDim I1_f_bdg_perd(C_CTRL_UNIT)
    
    I1_f_bdg_perd(C_CTRL_YR) = txtCtrlYR
    I1_f_bdg_perd(C_CTRL_UNIT) = cboctrlunit
    
   
    Call objPAFG105Q.F_LOOKUP_BDG_PERD_SVR(gStrGlobalCollection, I1_f_bdg_perd,	E1_com_budget_yyyymm)
				
    If CheckSYSTEMError(Err,True) = True Then
		Set objPAFG105Q = nothing		
		Exit Sub
    End If
    
    Set objPAFG105Q = nothing 
    
    Response.Write " <Script Language=vbscript>											 " & vbCr
	Response.Write " With Parent.frm1													 " & vbCr

		If E1_com_budget_yyyymm(E1_txt_1st_from_dt) <> "" Then
			Response.Write "	.txt1stFrYR.text	  =	""" & UNIMonthClientFormat(E1_com_budget_yyyymm(E1_txt_1st_from_dt))	& """" & vbCr
		End If
		If E1_com_budget_yyyymm(E1_txt_1st_to_dt) <> "" Then
			Response.Write "	.txt1stToYR.text	  =	""" & UNIMonthClientFormat(E1_com_budget_yyyymm(E1_txt_1st_to_dt))		& """" & vbCr
		End If
		
		If E1_com_budget_yyyymm(E1_txt_2st_from_dt) <> "" Then
			Response.Write "	.txt2ndFrYR.text	  =	""" & UNIMonthClientFormat(E1_com_budget_yyyymm(E1_txt_2st_from_dt))	& """" & vbCr
		End If
		If E1_com_budget_yyyymm(E1_txt_2st_to_dt) <> "" Then
			Response.Write "	.txt2ndToYR.text	  =	""" & UNIMonthClientFormat(E1_com_budget_yyyymm(E1_txt_2st_to_dt))		& """" & vbCr
		End If

		If E1_com_budget_yyyymm(E1_txt_3st_from_dt) <> "" Then
			Response.Write "	.txt3rdFrYR.text	  =	""" & UNIMonthClientFormat(E1_com_budget_yyyymm(E1_txt_3st_from_dt))	& """" & vbCr
		End If
		If E1_com_budget_yyyymm(E1_txt_3st_to_dt) <> "" Then
			Response.Write "	.txt3rdToYR.text	  =	""" & UNIMonthClientFormat(E1_com_budget_yyyymm(E1_txt_3st_to_dt))		& """" & vbCr
		End If

		If E1_com_budget_yyyymm(E1_txt_4st_from_dt) <> "" Then
			Response.Write "	.txt4thFrYR.text	  =	""" & UNIMonthClientFormat(E1_com_budget_yyyymm(E1_txt_4st_from_dt))	& """" & vbCr
		End If
		If E1_com_budget_yyyymm(E1_txt_4st_to_dt) <> "" Then
			Response.Write "	.txt4thToYR.text	  =	""" & UNIMonthClientFormat(E1_com_budget_yyyymm(E1_txt_4st_to_dt))		& """" & vbCr
		End If

	if isarray(E1_com_budget_yyyymm) = False Then
		Call DisplayMsgBox("140100", vbInformation, "", "", I_MKSCRIPT)
		Exit Sub	
	End if
	
	Response.Write "End With														" & vbCr
	Response.Write "Parent.DbQueryOk()												" & vbCr
	Response.Write "</Script>														" & vbCr 
End Sub

'==================================================================================
'	Name : SubBizSave()
'	Description : mulity save define
'==================================================================================
Sub SubBizSave()


On Error Resume Next
Err.Clear							'☜: Protect system from crashing

Dim f_bdg_perd
Const C_ctrl_yr    = 0
Const C_ctrl_unit  = 1
Const C_txt1stFrYM = 2
Const C_txt1stToYM = 3
Const C_txt2ndFrYM = 4
Const C_txt2ndToYM = 5
Const C_txt3rdFrYM = 6
Const C_txt3rdToYM = 7
Const C_txt4thFrYM = 8
Const C_txt4thToYM = 9

Dim objPAFG105CU
Dim str1FrYear,str1FrMonth,str1FrDay
Dim str1ToYear,str1ToMonth,str1ToDay
Dim str2FrYear,str2FrMonth,str2FrDay
Dim str2ToYear,str2ToMonth,str2ToDay
Dim str3FrYear,str3FrMonth,str3FrDay
Dim str3ToYear,str3ToMonth,str3ToDay
Dim str4FrYear,str4FrMonth,str4FrDay
Dim str4ToYear,str4ToMonth,str4ToDay
Dim str1FrDt,str1ToDt
Dim str2FrDt,str2ToDt
Dim str3FrDt,str3ToDt
Dim str4FrDt,str4ToDt
Dim lgIntFlgMode


	
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 
	
    Set objPAFG105CU = Server.CreateObject("PAFG105.cFMngBdgPrdSvr")    

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If   
    
    ReDim f_bdg_perd(C_txt4thToYM)
	f_bdg_perd(C_ctrl_yr) = txtCtrlYR
	f_bdg_perd(C_ctrl_unit) = cboctrlunit
	
	Call ExtractDateFrom(Request("txt1stFrYR"),gDateFormatYYYYMM,gComDateType,str1FrYear,str1FrMonth,str1FrDay)	
    str1FrDt = str1FrYear & str1FrMonth
    Call ExtractDateFrom(Request("txt1stToYR"),gDateFormatYYYYMM,gComDateType,str1ToYear,str1ToMonth,str1ToDay)	
    str1ToDt = str1ToYear & str1ToMonth
    Call ExtractDateFrom(Request("txt2ndFrYR"),gDateFormatYYYYMM,gComDateType,str2FrYear,str2FrMonth,str2FrDay)	
    str2FrDt = str2FrYear & str2FrMonth
    Call ExtractDateFrom(Request("txt2ndToYR"),gDateFormatYYYYMM,gComDateType,str2ToYear,str2ToMonth,str2ToDay)	
    str2ToDt = str2ToYear & str2ToMonth
    Call ExtractDateFrom(Request("txt3rdFrYr"),gDateFormatYYYYMM,gComDateType,str3FrYear,str3FrMonth,str3FrDay)	
    str3FrDt = str3FrYear & str3FrMonth
    Call ExtractDateFrom(Request("txt3rdToYR"),gDateFormatYYYYMM,gComDateType,str3ToYear,str3ToMonth,str3ToDay)	
    str3ToDt = str3ToYear & str3ToMonth
    Call ExtractDateFrom(Request("txt4thFrYR"),gDateFormatYYYYMM,gComDateType,str4FrYear,str4FrMonth,str4FrDay)	
    str4FrDt = str4FrYear & str4FrMonth
    Call ExtractDateFrom(Request("txt4thToYR"),gDateFormatYYYYMM,gComDateType,str4ToYear,str4ToMonth,str4ToDay)	
    str4ToDt = str4ToYear & str4ToMonth

	Select Case cboctrlunit    	
		Case "Y"
			f_bdg_perd(C_txt1stFrYM) = str1FrDt
			f_bdg_perd(C_txt1stToYM) = str1ToDt
			f_bdg_perd(C_txt2ndFrYM) = ""
			f_bdg_perd(C_txt2ndToYM) = ""
			f_bdg_perd(C_txt3rdFrYM) = ""
			f_bdg_perd(C_txt3rdToYM) = ""
			f_bdg_perd(C_txt4thFrYM) = ""
			f_bdg_perd(C_txt4thToYM) = ""
		Case "H"
			f_bdg_perd(C_txt1stFrYM) = str1FrDt
			f_bdg_perd(C_txt1stToYM) = str1ToDt
			f_bdg_perd(C_txt2ndFrYM) = str2FrDt
			f_bdg_perd(C_txt2ndToYM) = str2ToDt
			f_bdg_perd(C_txt3rdFrYM) = ""
			f_bdg_perd(C_txt3rdToYM) = ""
			f_bdg_perd(C_txt4thFrYM) = ""
			f_bdg_perd(C_txt4thToYM) = ""
		Case "B"
			f_bdg_perd(C_txt1stFrYM) = str1FrDt
			f_bdg_perd(C_txt1stToYM) = str1ToDt
			f_bdg_perd(C_txt2ndFrYM) = str2FrDt
			f_bdg_perd(C_txt2ndToYM) = str2ToDt
			f_bdg_perd(C_txt3rdFrYM) = str3FrDt
			f_bdg_perd(C_txt3rdToYM) = str3ToDt
			f_bdg_perd(C_txt4thFrYM) = str4FrDt
			f_bdg_perd(C_txt4thToYM) = str4ToDt
	End Select
	
		

	If lgIntFlgMode = OPMD_CMODE Then
		Call objPAFG105CU.F_MANAGE_BDG_PERD_SVR(gStrGlobalCollection, "CREATE",f_bdg_perd)
	Elseif lgIntFlgMode = OPMD_UMODE Then
		Call objPAFG105CU.F_MANAGE_BDG_PERD_SVR(gStrGlobalCollection, "UPDATE",f_bdg_perd)
	End If
	
    If CheckSYSTEMError(Err,True) = True Then
		Set objPAFG105CU = nothing		
		Exit Sub
    End If
    
    Set objPAFG105CU = nothing 

    Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbSaveOk()							" & vbCr
    Response.Write "</Script>									" & vbCr
														
End Sub

'==================================================================================
'	Name : SubBizDelete()
'	Description : mulity save define
'==================================================================================
Sub SubBizDelete()

'On Error Resume Next
Err.Clear							'☜: Protect system from crashing

Dim objPAFG105D
Dim f_bdg_perd
Const C_ctrl_yr_D    = 0
Const C_ctrl_unit_d  = 1
	
	ReDim f_bdg_perd(C_ctrl_unit_d)
	f_bdg_perd(C_ctrl_yr_D) = txtCtrlYR
	f_bdg_perd(C_ctrl_unit_d) = cboctrlunit
	
	Set objPAFG105D = Server.CreateObject("PAFG105.cFMngBdgPrdSvr")    

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If   
    

	Call objPAFG105D.F_MANAGE_BDG_PERD_SVR(gStrGlobalCollection, "DELETE",f_bdg_perd)
	
	 If CheckSYSTEMError(Err,True) = True Then
		Set objPAFG105D = nothing		
		Exit Sub
    End If
    
    Set objPAFG105D = nothing 
    
	Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbDeleteOk()						" & vbCr
    Response.Write "</Script>									" & vbCr
End Sub
%>
