<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1511MB2
'*  4. Program Name         : Quality Configuration
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PD6G020,PD6G010
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
	Const I2_plant_cd = 0
    Const I2_pr_yn_before_receipt = 1
    Const I2_st_yn_after_receipt = 2
    Const I2_modify_yn_after_release = 3
    Const I2_basic_mark_for_insp_dt = 4
    Const I2_basic_mark_for_release_dt = 5
    
	Call LoadBasisGlobalInf
                                         
	On Error Resume Next
	Call HideStatusWnd																'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

	Dim objPD6G010																	
	Dim lgIntFlgMode	
	Dim iCommandSent
	Dim I2_q_configuration
		
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
		
	Set objPD6G010 = Server.CreateObject("PD6G010.cQMaintConfigSvr")

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
	End If

	    
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	Redim I2_q_configuration(5)
	
	I2_q_configuration(I2_plant_cd) = Request("txtPlantCd2")
	I2_q_configuration(I2_pr_yn_before_receipt) = Request("rdoPRYNBeforeReceipt")
	I2_q_configuration(I2_st_yn_after_receipt) = Request("rdoSTYNAftereReceipt") 
	I2_q_configuration(I2_modify_yn_after_release) = Request("rdoModifyYNAfterRelease")
	I2_q_configuration(I2_basic_mark_for_insp_dt) = Request("cboInspDt")
	I2_q_configuration(I2_basic_mark_for_release_dt) = Request("cboReleaseDt")
		
	Call objPD6G010.Q_MAINT_CONFIGURATION_SVR(gStrGlobalCollection, iCommandSent, I2_q_configuration)

	If CheckSYSTEMError(Err,True) = true Then
	   Set objPD6G010 = Nothing
	   Response.End
	End If

	Set objPD6G010 = Nothing
%>
	<Script Language=vbscript>
	With parent																			
		.DbSaveOk
	End With
	</Script>
