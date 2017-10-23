<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : B2801mb1.asp
'*  4. Program Name         : 창고정보등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B28019Lookup
'
'*  7. Modified date(First) : 2000/04/25
'*  8. Modified date(Last)  : 2000/04/25
'*  9. Modifier (First)     : Kweon Soon Ho
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%
Call LoadBasisGlobalInf()
								

Call HideStatusWnd				

On Error Resume Next			
    Err.Clear                   

Dim PB6C020						


Dim strMode     



'-----------------------
'IMPORTS View
'-----------------------
Dim I1_b_plant_plant_cd
Dim I2_b_storage_location_sl_cd

'-----------------------
'EXPORTS View
'-----------------------
Dim E1_b_plant
	Const I005_E1_plant_cd		= 0
	Const I005_E1_plant_nm		= 1    
Dim E2_b_biz_partner
	Const I005_E2_bp_nm			= 0
	Const I005_E2_bp_cd			= 1
Dim E3_b_storage_location
	Const I005_E3_sl_cd			= 0
	Const I005_E3_sl_type		= 1
	Const I005_E3_sl_nm			= 2
	Const I005_E3_inv_mgr		= 3
	Const I005_E3_sl_group_cd	= 4
	Const I005_E3_ext_sl_type	= 5
	Const I005_E3_mrp_used_flg	= 6
	Const I005_E3_tax_class		= 7

	
	strMode = Request("txtMode")

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I1_b_plant_plant_cd = Request("txtPlantCd")
	I2_b_storage_location_sl_cd = Request("txtSLCd")

   

    Set PB6C020 = Server.CreateObject("PB6C020.cBLookUpStorageLoc")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If


    '-----------------------
    'Com action area
    '-----------------------
	Call PB6C020.B_LOOK_UP_STORAGE_LOCATION(gStrGlobalCollection, _
											I1_b_plant_plant_cd, _
											I2_b_storage_location_sl_cd, _
											E1_b_plant, _
											E2_b_biz_partner, _
											E3_b_storage_location)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set PB6C020 = Nothing
		Response.End
	End If
	

	Set PB6C020 = Nothing	
	
	'-----------------------
	'Result data display area
	'----------------------- 
%>
<Script Language=vbscript>
	With parent
		.frm1.txtSLCd1.value		= "<%=ConvSPChars(E3_b_storage_location(I005_E3_sl_cd))%>"
		.frm1.txtSLNm1.value		= "<%=ConvSPChars(E3_b_storage_location(I005_E3_sl_nm))%>"
		.frm1.cboSLType.value		= "<%=UCase(E3_b_storage_location(I005_E3_sl_type))%>"	
		.frm1.cboExtSLType.value	= "<%=UCase(E3_b_storage_location(I005_E3_ext_sl_type))%>"
		.frm1.cboSLGroup.value		= "<%=UCase(E3_b_storage_location(I005_E3_sl_group_cd))%>"
		.frm1.cboInvMgr.value		= "<%=UCase(E3_b_storage_location(I005_E3_inv_mgr))%>"	
		.frm1.cboTaxClass.value		= "<%=UCase(E3_b_storage_location(I005_E3_tax_class))%>"
				
		if "<%=E3_b_storage_location(I005_E3_mrp_used_flg)%>" = "Y" then
			.frm1.optMrpUsedFlg(0).Checked = True			
		else
			.frm1.optMrpUsedFlg(1).Checked = True
		end if
						
		.frm1.txtBPCd.value			= "<%=ConvSPChars(E2_b_biz_partner(I005_E2_bp_cd))%>"
		.frm1.txtBPNm.value			= "<%=ConvSPChars(E2_b_biz_partner(I005_E2_bp_nm))%>"
		.frm1.txthPlantCd.value		= "<%=ConvSPChars(E1_b_plant(I005_E1_plant_cd))%>"	

		.lgNextNo = ""		
		.lgPrevNo = ""		
		
		.DbQueryOk			
	End With
</Script>
