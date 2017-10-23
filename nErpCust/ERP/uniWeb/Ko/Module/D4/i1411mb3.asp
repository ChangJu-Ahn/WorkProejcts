<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Im
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

Dim pPD4G010								

Dim lgIntFlgMode

Dim I009_i_movetype_configuration
	Const I009_I1_mov_type					= 0    
	Const I009_I1_debit_credit_flag			= 1
	Const I009_I1_stck_type_ctrl_flag		= 2
	Const I009_I1_stck_type_flag_origin		= 3
	Const I009_I1_stck_type_flag_dest		= 4
	Const I009_I1_price_ctrl_flag			= 5
	Const I009_I1_trns_type					= 6
	Const I009_I1_post_ctrl_flag			= 7
	Const I009_I1_gui_control_flag			= 8
	Const I009_I1_gui_control_flag2			= 9
	Const I009_I1_gui_control_flag3			= 10
	Const I009_I1_gui_control_flag4			= 11
	Const I009_I1_matl_cost_dist_indctr		= 12
ReDim I009_i_movetype_configuration(I009_I1_matl_cost_dist_indctr)	
Dim I009_ief_supplied_select_char


	Err.Clear                                        
	
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	I009_ief_supplied_select_char = "D"	
	
	I009_i_movetype_configuration(I009_I1_mov_type) = Request("txtMovType1")		


	Set pPD4G010 = Server.CreateObject("PD4G010.cIMaintMoveTypeSvr")    	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End										
	End If    
	
    Call pPD4G010.I_MAINT_MOVE_TYPE_SVR(gStrGlobalCollection, _
										I009_i_movetype_configuration, _
										I009_ief_supplied_select_char)
                                                
	If CheckSYSTEMError(Err, True) = True Then
		Set pPD4G010 = Nothing                            
		Response.End
	End If

	Set pPD4G010 = Nothing                                
	
	'-----------------------
	'Result data display area
	'----------------------- 
%>
<Script Language=vbscript>
	Call parent.DbDeleteOk()
</Script> 
	        