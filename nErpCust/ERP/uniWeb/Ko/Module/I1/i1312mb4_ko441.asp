<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Im
'*  2. Function Name        : Move Type
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +i14111 LookUp
'		          +i14119 Insert
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2003/05/13
'*  9. Modifier (First)     : Mr Kim Nam hoon
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%												

On Error Resume Next							

Dim iPI0C010									

Dim I1_i_movetype_configuration
	Const C_I1_mov_type = 0
	Const C_I1_trns_type = 1
Redim I1_i_movetype_configuration(C_I1_trns_type)


Dim E1_i_movetype_configuration
	Const C_E1_stck_type_ctrl_flag = 2
	Const C_E1_stck_type_flag_origin = 3
	Const C_E1_gui_control_flag = 9
	Const C_E1_gui_control_flag2 = 10
	Const C_E1_gui_control_flag3 = 11
	Const C_E1_gui_control_flag4 = 12

Dim E2_move1_b_minor

	Call LoadBasisGlobalInf()
	    
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

	Err.Clear                                                         

	If Request("txtMovType") = "" Then								
		Response.End 
	End If
	'-----------------------
	'Data manipulate  area
	'-----------------------
	I1_i_movetype_configuration(C_I1_mov_type) = Trim(Request("txtMovType"))

	Set iPI0C010 = Server.CreateObject("PI0C010.cICheck")

	If CheckSYSTEMError(Err,True) = True Then
		Set iPI0C010 = Nothing
		Response.End
	End If

	Call iPI0C010.I_LOOK_UP_MOVETYPE_CONF(gStrGlobalCollection, _
										I1_i_movetype_configuration, _
										E1_i_movetype_configuration, _
										E2_move1_b_minor)
	
	If CheckSYSTEMError(Err,True) = True Then
		Set iPI0C010 = Nothing
		Response.Write "<Script Language=vbscript> " & vbCrlf
		Response.Write "	Call parent.ClickTab1 " & vbCrlf 
		Response.Write "	parent.frm1.txtMovType.focus " & vbCrlf 
		Response.Write "</Script>" & vbCrlf
		Response.End																	
	End If

    Response.Write "<Script Language=vbscript> " & vbCrlf
	Response.Write "With parent" & vbCrlf
	if E1_i_movetype_configuration(C_E1_stck_type_ctrl_flag) = "A" then
		Response.Write "	.frm1.vspdData.Col = 18" & vbCrlf
		Response.Write "	.frm1.vspdData.ColHidden = True " & vbCrlf
		Select Case E1_i_movetype_configuration(C_E1_stck_type_flag_origin)
	       Case "G"
				Response.Write "	.frm1.vspdData.Text = ""양품"" " & vbCrlf
    	   Case "B"
  				Response.Write "	.frm1.vspdData.Text = ""불량품"" " & vbCrlf
	       Case "Q"
				Response.Write "	.frm1.vspdData.Text = ""검사품"" " & vbCrlf
	       Case "T"
				Response.Write "	.frm1.vspdData.Text = ""이동품"" " & vbCrlf
		End Select
	End if
	
	if  E1_i_movetype_configuration(C_E1_gui_control_flag) <> "Y" then
		Response.Write "	.txtPlantCd2Title.style.display = ""none"" " & vbCrlf
		Response.Write "	.frm1.txtPlantCd2.style.display = ""none"" " & vbCrlf
		Response.Write "	.frm1.txtPlantNm2.style.display = ""none"" " & vbCrlf
		Response.Write "	.frm1.btnPlantCd.style.display  = ""none"" " & vbCrlf
		Response.Write "	.frm1.txtPlantCd2.value 	   = .frm1.txtPlantCd1.value" & vbCrlf
		Response.Write "	.frm1.txtPlantCd2.tag = ""25XXXU"" " & vbCrlf
		Response.Write "	.ggoOper.SetReqAttr .frm1.txtPlantCd2, ""Q"" " & vbCrlf
		Response.Write "	.txtCostCd2Title.style.display = ""none"" " & vbCrlf
		Response.Write "	.frm1.txtCostCd2.value 	   = .frm1.txtCostCd1.value " & vbCrlf
		Response.Write "	.ggoOper.SetReqAttr .frm1.txtCostCd2, ""Q"" " & vbCrlf
			
	else
		Response.Write "	.txtPlantCd2Title.style.display = """" " & vbCrlf
		Response.Write "	.frm1.txtPlantCd2.style.display = """" " & vbCrlf
		Response.Write "	.frm1.txtPlantNm2.style.display = """" " & vbCrlf
		Response.Write "	.frm1.btnPlantCd.style.display  = """" " & vbCrlf
		Response.Write "	.frm1.txtPlantCd2.tag = ""23XXXU"" " & vbCrlf  
		Response.Write "	.ggoOper.SetReqAttr .frm1.txtPlantCd2, ""N"" " & vbCrlf
		Response.Write "	.txtCostCd2Title.style.display = """" " & vbCrlf
		Response.Write "	.ggoOper.SetReqAttr .frm1.txtCostCd2, ""D"" " & vbCrlf
	end if
		
	Response.Write "	.frm1.hGuiControlFlag2.value = """ & E1_i_movetype_configuration(C_E1_gui_control_flag2) & " """ & vbCrlf
	if   E1_i_movetype_configuration(C_E1_gui_control_flag2)  <> "Y" then
		Response.Write " 	.txtSLCd2Title.style.display = ""none"" " & vbCrlf
		Response.Write "	.frm1.txtSLCd2.tag = ""25XXXU"" " & vbCrlf 
		Response.Write "	.ggoOper.SetReqAttr .frm1.txtSLCd2, ""Q"" " & vbCrlf
	else
		Response.Write "	.txtSLCd2Title.style.display = """" " & vbCrlf
		Response.Write "	.frm1.txtSLCd2.tag = ""23XXXU"" " & vbCrlf  
		Response.Write "	.ggoOper.SetReqAttr .frm1.txtSLCd2, ""N"" " & vbCrlf
	end if  
		    
		Response.Write ".frm1.hGuiControlFlag3.value = 	""" & E1_i_movetype_configuration(C_E1_gui_control_flag3) & """" & vbcrlf
	if  E1_i_movetype_configuration(C_E1_gui_control_flag3) = "Y" then
		Response.Write "if .lgIntFlgMode = .Parent.OPMD_CMODE then " & vbcrlf
		Response.Write "	.ggoSpread.SSSetRequired 	.C_TrnsItemCd, -1, -1 " & vbcrlf
		Response.Write "	else" & vbCrlf
		Response.Write "	.ggoSpread.SSSetProtected 	.C_TrnsItemCd, -1, -1 " & vbcrlf
		Response.Write " End if " & vbcrlf	
	End if    			 
		
	if  E1_i_movetype_configuration(C_E1_gui_control_flag4) <> "Y" then
		Response.Write "	.txtTrackingNoTitle.style.display = ""none"" " & vbCrlf
		Response.Write "	.frm1.txtTrackingNo.style.display =""none"" " & vbCrlf
		Response.Write "	.frm1.txtTrackingNo.value 	 = """" " & vbCrlf
		Response.Write "	.frm1.btnTrackingNo.style.display  = ""none"" " & vbCrlf
		Response.Write "	.frm1.txtTrackingNo.tag = ""25XXXU"" " & vbCrlf 
		Response.Write "	.ggoOper.SetReqAttr .frm1.txtTrackingNo, ""Q"" " & vbCrlf
	else
		Response.Write "	.txtTrackingNoTitle.style.display = """" " & vbCrlf
		Response.Write "	.frm1.txtTrackingNo.style.display = """" " & vbCrlf
		Response.Write "	.frm1.btnTrackingNo.style.display = """" " & vbCrlf
		Response.Write "	.frm1.txtTrackingNo.tag = ""23XXXU"" " & vbCrlf  
		Response.Write "	.ggoOper.SetReqAttr .frm1.txtTrackingNo, ""N"" " & vbCrlf
	End if    			 
	
	Response.Write "	.gMovTypeFlag = ""Y"" " & vbCrlf
	Response.Write "	.MovTypeDbQueryOk" & vbCrlf													
	Response.Write "End With" & vbCrlf
    Response.Write "</Script>" & vbCrlf
	 
	Set iPI0C010 = Nothing									
	Response.End											
%>
