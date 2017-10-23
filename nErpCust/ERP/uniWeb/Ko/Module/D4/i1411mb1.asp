<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Move Type
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +i14111 LookUp
'		          +i14119 Insert
'*  7. Modified date(First) : 2000/03/23
'*  8. Modified date(Last)  : 2000/03/23
'*  9. Modifier (First)     : Mr  Koh
'* 10. Modifier (Last)      : Mr  Koh
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change"
'*                            this mark(¢Á) Means that "may  change"
'*                            this mark(¡Ù) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf()
												
Call HideStatusWnd 
On Error Resume Next							

Dim pI14119										

Dim strMode										

Dim I015_i_movetype_configuration
	Const I015_I1_mov_type		= 0
	Const I015_I1_trns_type		= 1
ReDim I015_i_movetype_configuration(I015_I1_trns_type)
Dim I2_PrevNextFlag

Dim E015_i_movetype_configuration
    Const I015_E1_mov_type					= 0
	Const I015_E1_debit_credit_flag			= 1
	Const I015_E1_stck_type_ctrl_flag		= 2
	Const I015_E1_stck_type_flag_origin		= 3
	Const I015_E1_stck_type_flag_dest		= 4
	Const I015_E1_price_ctrl_flag			= 5
	Const I015_E1_trns_type					= 6
	Const I015_E1_revrse_mov_type			= 7
	Const I015_E1_post_ctrl_flag			= 8
	Const I015_E1_gui_control_flag			= 9
	Const I015_E1_gui_control_flag2			= 10
	Const I015_E1_gui_control_flag3			= 11
	Const I015_E1_gui_control_flag4			= 12
	Const I015_E1_matl_cost_dist_indctr		= 13
Dim E015_move1_b_minor
	Const I015_E1_minor_nm = 0


strMode = Request("txtMode")	

Err.Clear                                  


	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	I015_i_movetype_configuration(I015_I1_mov_type) = Request("txtMovType1")
	I2_PrevNextFlag	= Request("PrevNextFlg")

	Set pI14119 = Server.CreateObject("PI0C010.cICheck")    	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End										
	End If    
	
    Call pI14119.I_LOOK_UP_MOVETYPE_CONF(gStrGlobalCollection, _
									I015_i_movetype_configuration, _
									E015_i_movetype_configuration, _
									E015_move1_b_minor, _
									I2_PrevNextFlag)
                                                            
	If CheckSYSTEMError(Err, True) = True Then
		Set pI14119 = Nothing                                       
		Response.End
	End If

	Set pI14119 = Nothing   

	'-----------------------
	'Result data display area
	'----------------------- 
%>

<Script Language=vbscript>
	With parent.frm1

		.txtMovType1.value		= "<%=ConvSPChars(E015_i_movetype_configuration(I015_E1_mov_type))%>"
		.txtMovTypeNm1.Value	= "<%=ConvSPChars(E015_move1_b_minor(I015_E1_minor_nm))%>"
		.txtMovType2.value		= "<%=ConvSPChars(E015_i_movetype_configuration(I015_E1_mov_type))%>"
		.txtMovTypeNm2.Value	= "<%=ConvSPChars(E015_move1_b_minor(I015_E1_minor_nm))%>"

		If "<%=E015_i_movetype_configuration(I015_E1_debit_credit_flag)%>" = "D" then
			.optDebitCreditFlag(0).Checked = True
		Else
			.optDebitCreditFlag(1).Checked = True
		End If
			
		.cboStckTypeFlagOrigin.Value	= "<%=E015_i_movetype_configuration(I015_E1_stck_type_flag_origin)%>"
		.cboStckTypeFlagDest.Value		= "<%=E015_i_movetype_configuration(I015_E1_stck_type_flag_dest)%>" 
			
		If "<%=E015_i_movetype_configuration(I015_E1_price_ctrl_flag)%>" = "Y" then
			.optPriceCtrlFlag(0).Checked = True 
		Else
			.optPriceCtrlFlag(1).Checked = True 
		End if
							
		.cboTrnsType.Value	    = "<%=E015_i_movetype_configuration(I015_E1_trns_type)%>" 		
			
		If "<%=E015_i_movetype_configuration(I015_E1_post_ctrl_flag)%>" = "Y" then
			.optPostCtrlFlag(0).Checked = True
		Else
			.optPostCtrlFlag(1).Checked = True
		End If
				
		If "<%=E015_i_movetype_configuration(I015_E1_gui_control_flag)%>" = "Y" then
			.optPlantMovFlag(0).Checked = True 
		Else
			.optPlantMovFlag(1).Checked = True 
		End If
			
		If "<%=E015_i_movetype_configuration(I015_E1_gui_control_flag2)%>" = "Y" then
			.optSLMovFlag(0).Checked = True 
		Else
			.optSLMovFlag(1).Checked = True 
		End If
				
		If "<%=E015_i_movetype_configuration(I015_E1_gui_control_flag3)%>" = "Y" then
			.optItemMovFlag(0).Checked = True 
		Else
			.optItemMovFlag(1).Checked = True
		End If
				
		If "<%=E015_i_movetype_configuration(I015_E1_gui_control_flag4)%>"= "Y" then 
			.optTrackingNoMovFlag(0).Checked = True
		Else
			.optTrackingNoMovFlag(1).Checked = True 
		End If
			
		parent.DbQueryOk
	
	End With

</Script>