<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Comon Popup Business Part													*
'*  3. Program ID           : i1411bp1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 이동유형팝업																*
'*  7. Modified date(First) : 2000/02/29																*
'*  8. Modified date(Last)  : 2000/02/29																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : An Chang Hwan																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :																			*
'*                            2000/02/29 : Coding Start													*
'********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf()
																		
On Error Resume Next
Call HideStatusWnd 
Dim objPopUp
Dim strMode																

Dim lgStrPrevKey	
Dim strData
Dim LngMaxRow		
Dim LngRow
Dim PvArr

Const C_SHEETMAXROWS_D = 100

    '-----------------------
    'IMPORTS View
    '-----------------------
    Dim I1_i_movetype_configuration
		Const I013_I1_mov_type	= 0
		Const I013_I1_trns_type = 1    
    ReDim I1_i_movetype_configuration(I013_I1_trns_type)
	'-----------------------
	'EXPORTS View
	'-----------------------
    Dim EG1_export_group
		Const I013_EG1_E1_mov_type					= 0
		Const I013_EG1_E1_debit_credit_flag			= 1
		Const I013_EG1_E1_stck_type_ctrl_flag		= 2
		Const I013_EG1_E1_stck_type_flag_origin		= 3
		Const I013_EG1_E1_stck_type_flag_dest		= 4
		Const I013_EG1_E1_price_ctrl_flag			= 5
		Const I013_EG1_E1_trns_type					= 6
		Const I013_EG1_E1_post_ctrl_flag			= 7
		Const I013_EG1_E1_gui_control_flag			= 8
		Const I013_EG1_E1_gui_control_flag2			= 9
		Const I013_EG1_E1_gui_control_flag3			= 10
		Const I013_EG1_E1_gui_control_flag4			= 11
		Const I013_EG1_E1_matl_cost_dist_indctr		= 12
		Const I013_EG1_E2_minor_nm					= 13
		Const I013_EG1_E3_minor_nm					= 14
    Dim E1_i_movetype_configuration_mov_type


	lgStrPrevKey = Request("lgStrPrevKey")
	
	I1_i_movetype_configuration(I013_I1_mov_type)  = Request("txtMovType")
	I1_i_movetype_configuration(I013_I1_trns_type) = Request("txtTrnsType")
	
	If Trim(lgStrPrevKey) <> "" Then I1_i_movetype_configuration(I013_I1_mov_type)  = lgStrPrevKey
	

	Set objPopUp = Server.CreateObject("PD4G020.cIListMoveTypeSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End											
	End If    
	
	Call objPopUp.I_LIST_MOVE_TYPE_CONF(gStrGlobalCollection, C_SHEETMAXROWS_D, _
								I1_i_movetype_configuration, _
								EG1_export_group, _
								E1_i_movetype_configuration_mov_type)

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
    	Set objPopUp = Nothing												
		Response.End														
	End If

	Set objPopUp = Nothing
	
	if isEmpty(EG1_export_group) then
		Response.End													
	end if
	
	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	ReDim PvArr(ubound(EG1_export_group,1))
	
	For LngRow = 0 To ubound(EG1_export_group,1)
		
		strData = Chr(11) & ConvSPChars(EG1_export_group(LngRow, I013_EG1_E1_mov_type)) & _
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow, I013_EG1_E2_minor_nm))

		If EG1_export_group(LngRow, I013_EG1_E1_debit_credit_flag) = "D" then
		strData = strData & Chr(11) & "증가"
		Else
		strData = strData & Chr(11) & "감소"
		End if
		      
		strData = strData &		Chr(11) & EG1_export_group(LngRow, I013_EG1_E1_price_ctrl_flag)			& _
								Chr(11) & EG1_export_group(LngRow, I013_EG1_E1_post_ctrl_flag)			& _
								Chr(11) & EG1_export_group(LngRow, I013_EG1_E1_matl_cost_dist_indctr)	& _
								Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
		
		PvArr(LngRow) = strData
	Next
    
	strData = Join(PvArr, "")
	
	If E1_i_movetype_configuration_mov_type = EG1_export_group(LngRow, I013_EG1_E1_mov_type) Then 
		lgStrPrevKey = ""
	Else
		lgStrPrevKey = E1_i_movetype_configuration_mov_type
	End If
%>

<Script Language=vbscript>
 With Parent

   .ggoSpread.SSShowData "<%=strData%>"
   .vspdData.focus 
	
   .lgStrPrevKey  = "<%=ConvSPChars(lgStrPrevKey)%>"

	if .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) and .lgStrPrevKey <> "" Then  
		.lgQueryFlag = "0"
		.DbQuery
	Else
		.lgQueryFlag = "1"
		.DbQueryOk
	End If
       
End With 
</Script>