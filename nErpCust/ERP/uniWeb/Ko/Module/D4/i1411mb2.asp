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

Dim pPD4G010					

Dim lgIntFlgMode

Dim i_movetype_configuration
	Const I1_mov_type					= 0    
	Const I1_debit_credit_flag			= 1
	Const I1_stck_type_ctrl_flag		= 2
	Const I1_stck_type_flag_origin		= 3
	Const I1_stck_type_flag_dest		= 4
	Const I1_price_ctrl_flag			= 5
	Const I1_trns_type					= 6
	Const I1_post_ctrl_flag				= 7
	Const I1_gui_control_flag			= 8
	Const I1_gui_control_flag2			= 9
	Const I1_gui_control_flag3			= 10
	Const I1_gui_control_flag4			= 11
	Const I1_matl_cost_dist_indctr		= 12
ReDim i_movetype_configuration(I1_matl_cost_dist_indctr)	
Dim ief_supplied_select_char


	Err.Clear																
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))								
	
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	i_movetype_configuration(I1_mov_type)				= Request("txtMovType2")				
	i_movetype_configuration(I1_debit_credit_flag)		= Request("optDebitCreditFlag")	       
	i_movetype_configuration(I1_stck_type_ctrl_flag)	= "A"      
	i_movetype_configuration(I1_stck_type_flag_origin)	= Request("cboStckTypeFlagOrigin")
	i_movetype_configuration(I1_stck_type_flag_dest)	= Request("cboStckTypeFlagDest")    
	i_movetype_configuration(I1_price_ctrl_flag)		= Request("optPriceCtrlFlag")           
	i_movetype_configuration(I1_trns_type)				= Request("cboTrnsType")                      
	i_movetype_configuration(I1_post_ctrl_flag)			= Request("optPostCtrlFlag")	
	i_movetype_configuration(I1_gui_control_flag)		= Request("optPlantMovFlag")			
	i_movetype_configuration(I1_gui_control_flag2)		= Request("optSLMovFlag")			
	i_movetype_configuration(I1_gui_control_flag3)		= Request("optItemMovFlag")			
	i_movetype_configuration(I1_gui_control_flag4)		= Request("optTrackingNoMovFlag")
	i_movetype_configuration(I1_matl_cost_dist_indctr)	= Request("cboCostFlag")
	
	If lgIntFlgMode = OPMD_CMODE Then
		ief_supplied_select_char = "C"							
	End If
	
	Set pPD4G010 = Server.CreateObject("PD4G010.cIMaintMoveTypeSvr")    	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End											
	End If    
	
    Call pPD4G010.I_MAINT_MOVE_TYPE_SVR(gStrGlobalCollection, _
										i_movetype_configuration, _
										ief_supplied_select_char)
                                                
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
	With parent
		.DbSaveOk
	End With
</Script>
	
