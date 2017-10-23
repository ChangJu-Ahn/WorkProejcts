<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1511MB1
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

Const E2_pr_yn_before_receipt = 0
Const E2_st_yn_after_receipt = 1
Const E2_modify_yn_after_release = 2
Const E2_basic_mark_for_insp_dt = 3
Const E2_basic_mark_for_release_dt = 4
Const E2_plant_cd = 5
Const E2_plant_nm = 6
    
    
	Call LoadBasisGlobalInf

	On Error Resume Next

	Call HideStatusWnd																			'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

	Dim objPD6G020																				'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim strFlag
	Dim strPlantCd
	Dim E2_q_configuration
		
	Set objPD6G020 = Server.CreateObject("PD6G020.cQLookUpConfigSvr")    

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
	End If
		    
	strPlantCd = Request("txtPlantCd")
	strFlag = Request("PrevNextFlg")
		        
	Call objPD6G020.Q_LOOK_UP_CONFIGURATION (gStrGlobalCollection, _
											 strFlag, _
											 strPlantCd, _
											 E2_q_configuration)
			
	If CheckSYSTEMError(Err,True) = true Then
	   Set objPD6G020 = Nothing
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write "	Parent.frm1.txtPlantCd1.Focus " & vbCr  	   	  
		Response.Write "</Script>      " & vbCr   
		Response.End
	End If

	Set objPD6G020 = Nothing
	' 다음키와 이전키가 존재하지 않을 경우 Blank로 보내는 로직을 수행함.
%>
	<Script Language=vbscript>
	With parent.frm1		
		.txtPlantCd1.value = "<%=ConvSPChars(E2_q_configuration(E2_plant_cd))%>"							'☆: Plant Code
		.txtPlantNm1.value = "<%=ConvSPChars(E2_q_configuration(E2_plant_nm))%>"							'☆: Plant Name		
		.txtPlantCd2.value = "<%=ConvSPChars(E2_q_configuration(E2_plant_cd))%>"							'☆: Plant Code
		.txtPlantNm2.value = "<%=ConvSPChars(E2_q_configuration(E2_plant_nm))%>"							'☆: Plant Name		
 			
		If UCase("<%=E2_q_configuration(E2_pr_yn_before_receipt)%>") = "Y" Then
			.rdoPRYNBeforeReceipt1.checked = true
		Else
			.rdoPRYNBeforeReceipt2.checked = true
		End If
			
		If UCase("<%=E2_q_configuration(E2_st_yn_after_receipt)%>") = "Y" Then
			.rdoSTYNAftereReceipt1.checked = true
		Else
			.rdoSTYNAftereReceipt2.checked = true
		End If

		If UCase("<%=E2_q_configuration(E2_modify_yn_after_release)%>") = "Y" Then
			.rdoModifyYNAfterRelease1.checked = true
		Else
			.rdoModifyYNAfterRelease2.checked = true
		End If
	
		.cboInspDt.value = "<%= E2_q_configuration(E2_basic_mark_for_insp_dt)%>"
		.cboReleaseDt.value = "<%= E2_q_configuration(E2_basic_mark_for_release_dt)%>"
				
		parent.DbQueryOk																		'☜: 조화가 성공 
	End With
	</Script>
