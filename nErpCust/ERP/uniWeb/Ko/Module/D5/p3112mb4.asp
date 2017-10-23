<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p3112mb4.asp 
'*  4. Program Name         : Entry Production Configuration
'*  5. Program Desc         :
'*  6. Component List       : PB0C102.cMngPltConfig
'*  7. Modified date(First) : 2000/11/29
'*  8. Modified date(Last)  : 2002/11/11
'*  9. Modifier (First)     : Jung Yu Kyung 
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

Err.Clear
On Error Resume Next
Call HideStatusWnd

Dim pPB0C102                 '☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_plant_cd, I2_p_plant_configuration, iCommandSent
Dim iIntFlgMode

	Const P009_I2_rls_inv_chk_flg = 0
	Const P009_I2_auto_rcpt_flg = 1
	Const P009_I2_prev_opr_chk_flg = 2
	Const P009_I2_prod_etc_mthd = 3
	Const P009_I2_ord_close_mthd = 4
	Const P009_I2_exss_rcpt_flg = 5
	Const P009_I2_prod_mthd = 6
	Const P009_I2_prod_flg = 7
	Const P009_I2_MPS_method = 8
	Const P009_I2_delivery_order_flg = 9
	Const P009_I2_BOM_history_flg = 10
	Const P009_I2_routing_lt_flg = 11
	Const P009_I2_Eng_BOM_flg = 12
	Const P009_I2_Opr_cost_flg = 13
	const P009_I2_backlog_flg = 14
	'Add 2006-07-17
	const P009_I2_prod_child_mthd = 15
	const P009_I2_prod_rsc_mthd = 16
	
ReDim I2_p_plant_configuration(P009_I2_prod_rsc_mthd)

If Request("txtPlantCd1") = "" Then										'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("971012", vbOKOnly, "공장", "", I_MKSCRIPT)	'⊙: 에러메세지는 DB화 한다.		 
	Response.End 
End If

iIntFlgMode = CInt(Request("txtFlgMode"))          '☜: 저장시 Create/Update 판별 

'-----------------------
'Data manipulate area
'-----------------------
I1_plant_cd = Trim(Request("txtPlantCd1"))
 
IF Request("rdoRlsInvChkFlg") = "1" Then
	I2_p_plant_configuration(P009_I2_rls_inv_chk_flg) = "1"
ElseIf Request("rdoRlsInvChkFlg") = "2" Then
	I2_p_plant_configuration(P009_I2_rls_inv_chk_flg) = "2"
Else 
	I2_p_plant_configuration(P009_I2_rls_inv_chk_flg) = "3"
End If
    
IF Request("rdoAutoRcptFlg") = "Y" Then
	I2_p_plant_configuration(P009_I2_auto_rcpt_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_auto_rcpt_flg) = "N"
End If
    
IF Request("rdoPreOprChkFlg") = "Y" Then
	I2_p_plant_configuration(P009_I2_prev_opr_chk_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_prev_opr_chk_flg) = "N"
End If
    
IF Request("rdoProdEtcMthd") = "Y" Then
	I2_p_plant_configuration(P009_I2_prod_etc_mthd) = "Y"
Else
	I2_p_plant_configuration(P009_I2_prod_etc_mthd) = "N"
End If
    
IF Request("rdoProdFlg") = "Y" Then
	I2_p_plant_configuration(P009_I2_prod_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_prod_flg) = "N"
End If
    
IF Request("rdoOrdCloseMthd") = "Y" Then
	I2_p_plant_configuration(P009_I2_ord_close_mthd) = "Y"
Else
	I2_p_plant_configuration(P009_I2_ord_close_mthd) = "N"
End If
    
IF Request("rdoExssRcptFlg") = "Y" Then
	I2_p_plant_configuration(P009_I2_exss_rcpt_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_exss_rcpt_flg) = "N"
End If
    
Select Case Request("rdoProdMthd") 
	Case "1" 
		I2_p_plant_configuration(P009_I2_prod_mthd) = "1"
	Case "2"
		I2_p_plant_configuration(P009_I2_prod_mthd) = "2"
	Case "3"
		I2_p_plant_configuration(P009_I2_prod_mthd) = "3"
	Case "4"
		I2_p_plant_configuration(P009_I2_prod_mthd) = "4"
	Case "5"
		I2_p_plant_configuration(P009_I2_prod_mthd) = "5"
End Select

'2003-03-18 추가/변경 START
Select Case Request("rdoMPSMETHOD") 
	Case "1" 
		I2_p_plant_configuration(P009_I2_MPS_method) = "N"
	Case "2"
		I2_p_plant_configuration(P009_I2_MPS_method) = "S"
	Case "3"
		I2_p_plant_configuration(P009_I2_MPS_method) = "C"
End Select

IF Request("rdoDELIVERYORDERFLG") = "Y" Then
	I2_p_plant_configuration(P009_I2_delivery_order_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_delivery_order_flg) = "N"
End If

IF Request("rdoBOMHISTORYFLG") = "Y" Then
	I2_p_plant_configuration(P009_I2_BOM_history_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_BOM_history_flg) = "N"
End If

IF Request("rdoRoutingLTFlg") = "Y" Then
	I2_p_plant_configuration(P009_I2_routing_lt_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_routing_lt_flg) = "N"
End If

'2003-03-18 추가/변경 END

'2005-03-07 Add start
IF Request("rdoENGBOMFLG") = "Y" Then
	I2_p_plant_configuration(P009_I2_Eng_BOM_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_Eng_BOM_flg) = "N"
End If
'2005-03-07 Add end

'2005-09-17 Add start
IF Request("rdoOprCostFlg") = "Y" Then
	I2_p_plant_configuration(P009_I2_Opr_cost_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_Opr_cost_flg) = "N"
End If
'2005-09-17 Add end

'2006-04-17 Add start
IF Request("rdoBacklogFlg") = "Y" Then
	I2_p_plant_configuration(P009_I2_backlog_flg) = "Y"
Else
	I2_p_plant_configuration(P009_I2_backlog_flg) = "N"
End If
'2006-04-17 Add end

'2006-07-17 Add START
Select Case Request("rdoProdChildMthd") 
	Case "1" 
		I2_p_plant_configuration(P009_I2_prod_child_mthd) = "1"
	Case "2"
		I2_p_plant_configuration(P009_I2_prod_child_mthd) = "2"
	Case "3"
		I2_p_plant_configuration(P009_I2_prod_child_mthd) = "3"
	Case "4"
		I2_p_plant_configuration(P009_I2_prod_child_mthd) = "4"
	Case "5"
		I2_p_plant_configuration(P009_I2_prod_child_mthd) = "5"
End Select

Select Case Request("rdoProdRscMthd") 
	Case "1" 
		I2_p_plant_configuration(P009_I2_prod_rsc_mthd) = "1"
	Case "2"
		I2_p_plant_configuration(P009_I2_prod_rsc_mthd) = "2"
	Case "3"
		I2_p_plant_configuration(P009_I2_prod_rsc_mthd) = "3"
	Case "4"
		I2_p_plant_configuration(P009_I2_prod_rsc_mthd) = "4"
	Case "5"
		I2_p_plant_configuration(P009_I2_prod_rsc_mthd) = "5"
End Select

'2006-07-17 Add END


If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If
    
'-----------------------
'Com Action Area
'-----------------------
Set pPB0C102 = Server.CreateObject("PB0C102.cPMngPltConfig")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB0C102.P_MANAGE_PLANT_CONFIGURE(gStrGlobalCollection, iCommandSent, I1_plant_cd, I2_p_plant_configuration)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB0C102 = Nothing               '☜: Unload Component
	Response.End
End If

Set pPB0C102 = Nothing               '☜: Unload Component

'-----------------------
'Result data display area
'----------------------- 
Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End                    '☜: Process End
%>
