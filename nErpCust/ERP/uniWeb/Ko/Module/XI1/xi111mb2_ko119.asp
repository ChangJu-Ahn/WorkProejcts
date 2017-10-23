<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : INTERFACE
'*  2. Function Name        : 
'*  3. Program ID           : xi111mb2_ko119.asp	
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2006/04/19
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : HJO
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 
Err.Clear

Dim pPXI1G111_KO119																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_system_config,  iCommandSent
Dim iIntFlgMode

Const I1_system_id = 0
Const I1_system_nm = 1
Const I1_plant_cd = 2
Const I1_usage_flag = 3
Const I1_alias_nm = 4
Const I1_ip_address = 5
Const I1_port_no = 6
Const I1_config_file_nm = 7
Const I1_config_step_nm = 8
Const I1_url = 9
Const I1_e_mail_id = 10
Const I1_login_id = 11
Const I1_login_pwd = 12
Const I1_remark = 13

Redim I1_system_config(I1_remark)
    
iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별

'-----------------------
'Data manipulate area
'-----------------------
I1_system_config(I1_system_id) = ConvSPChars(UCase(Trim(Request("txtSystemId2"))))
I1_system_config(I1_system_nm) = ConvSPChars(Trim(Request("txtSystemIdNm2")))
I1_system_config(I1_plant_cd) = ConvSPChars(UCase(Trim(Request("txtPlantCd"))))
I1_system_config(I1_usage_flag) = ConvSPChars(Trim(Request("txtRdoFlg")))
I1_system_config(I1_alias_nm) = ConvSPChars(Trim(Request("txtAliasNm")))
I1_system_config(I1_ip_address) = ConvSPChars(Trim(Request("txtIPAdd")))
I1_system_config(I1_port_no) = ConvSPChars(Trim(Request("txtPortNo")))
I1_system_config(I1_config_file_nm) = ConvSPChars(Trim(Request("txtConfigFNm")))
I1_system_config(I1_config_step_nm) = ConvSPChars(Trim(Request("txtConfigSNm")))
I1_system_config(I1_url) = ConvSPChars(Trim(Request("txtUrl")))
I1_system_config(I1_e_mail_id) = ConvSPChars(Trim(Request("txtEMail")))
I1_system_config(I1_login_id) = ConvSPChars(Trim(Request("txtLoginId")))
I1_system_config(I1_login_pwd) = ConvSPChars(Trim(Request("txtLoginPwd")))
I1_system_config(I1_remark) =ConvSPChars( Trim(Request("txtRemark")))

If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
ElseIf iIntFlgMode = 1003 Then
	iCommandSent = "DELETE"		
End If

    
'-----------------------
'Com Action Area
'-----------------------
Set pPXI1G111_KO119 = Server.CreateObject("PXI1G111_KO119.cXIInterFaceSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPXI1G111_KO119.XI_MAIN_INTERFACE_SVR(gStrGlobalCollection, iCommandSent, I1_system_config)

If CheckSYSTEMError2(Err, True, "시스템ID", "", "", "", "") = True Then
	Set pPXI1G111_KO119 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPXI1G111_KO119 = Nothing															'☜: Unload Component

'-----------------------
'Result data display area
'----------------------- 
If iIntFlgMode =1003 Then
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbDeleteOk" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End																			'☜: Process End

Else
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End																				'☜: Process End

End If
%>
