<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : i2231mb1.asp
'*  4. Program Name         : Plant query
'*  5. Program Desc         :
'*  6. Comproxy List        : +B25019LookUpPlant
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2000/03/27
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Im Hyun Soo
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change"
'*                            this mark(¢Á) Means that "may  change"
'*                            this mark(¡Ù) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%													
Call LoadBasisGlobalInf()
Call HideStatusWnd									

' E3_b_plant
Const B153_E3_plant_cd = 0
Const B153_E3_plant_nm = 1
Const B153_E3_cur_cd = 2
Const B153_E3_plan_hrzn = 3
Const B153_E3_llc_given_dt = 4
Const B153_E3_bom_last_updt_dt = 5
Const B153_E3_mps_firm_dt = 6
Const B153_E3_dtf_for_mps = 7
Const B153_E3_ptf_for_mps = 8
Const B153_E3_ptf_for_mrp = 9
Const B153_E3_inv_cls_dt = 10
Const B153_E3_inv_open_dt = 11
Const B153_E3_valid_from_dt = 12
Const B153_E3_valid_to_dt = 13


On Error Resume Next														
Err.Clear                                           

Dim iPB6S101										

Dim I1_select_char
Dim I2_plant_cd
Dim E1_b_biz_area
Dim E2_b_currency
Dim E3_b_plant
Dim E4_p_mfg_calendar_type
Dim prStatusCodePrevNext
		
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I2_plant_cd = Request("txtPlantCd")
    	
    Set iPB6S101 = Server.CreateObject("PB6S101.cBLkUpPlt")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
    
    '-----------------------
    'Com action area
    '-----------------------
    Call iPB6S101.B_LOOK_UP_PLANT_SVR (gStrGlobalCollection, I1_select_char, I2_plant_cd, E1_b_biz_area, _
						E2_b_currency, E3_b_plant, E4_p_mfg_calendar_type, prStatusCodePrevNext)			

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	
	If CheckSYSTEMError(Err,True) = True Then
		Set iPB6S101 = Nothing								
		Response.End										
	End If
	
	Set iPB6S101 = Nothing								
	
	'-----------------------
	'Result data display area
	'----------------------- 
	
Response.Write "<Script Language=vbscript>	" & vbcr
Response.Write "With parent.frm1			" & vbcr

Response.Write "	.txtInvClsDt.text =	""" & UniMonthClientFormat(E3_b_plant(B153_E3_inv_cls_dt)) & """ " & vbcr
Response.Write "	parent.DbQueryOk	"	& vbcr
Response.Write "End With				"	& vbcr
Response.Write "	</Script>			"	& vbcr	

%>
