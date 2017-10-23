<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : b2801mb2.asp	
'*  4. Program Name         : Storage Location Entry
'*  5. Program Desc         :
'*  6. Comproxy List        : +B28011ManageStorageLocation

'*  7. Modified date(First) : 2000/04/25
'*  8. Modified date(Last)  : 2000/04/25
'*  9. Modifier (First)     : Kweon Soon Ho
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<%												
Call LoadBasisGlobalInf()

On Error Resume Next									
Err.Clear

Call HideStatusWnd										

Dim pPB6G010											
'Dim strCode											
Dim lgIntFlgMode
Dim iCommandSent

Dim I1_b_plant_plant_cd
Dim I2_b_biz_partner_bp_cd
Dim I3_b_storage_location
	Const I3_sl_cd			= 0
	Const I3_sl_type		= 1
	Const I3_sl_nm			= 2
	Const I3_inv_mgr		= 3
	Const I3_sl_group_cd	= 4
	Const I3_ext_sl_type	= 5
	Const I3_mrp_used_flg	= 6
	Const I3_tax_class		= 7
ReDim I3_b_storage_location(I3_tax_class)


    If Request("txtSLCd1") = "" Then							
		Call DisplayMsgBox("169902", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	End If
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))
	
	Set pPB6G010 = Server.CreateObject("PB6G010.cBManageStorageLoc")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

    '-----------------------
    'Data manipulate area
    '-----------------------
	I1_b_plant_plant_cd 					= UCase(Trim(Request("txtPlantCd")))
	I3_b_storage_location(I3_sl_cd) 		= UCase(Trim(Request("txtSLCd1")))
	I3_b_storage_location(I3_sl_nm) 		= Request("txtSLNm1")
	I3_b_storage_location(I3_sl_type)		= UCase(Trim(Request("cboSLType")))
	I2_b_biz_partner_bp_cd 						= UCase(Trim(Request("txtBPCd"))) 
	I3_b_storage_location(I3_inv_mgr)   	= UCase(Trim(Request("cboInvMgr")))
	I3_b_storage_location(I3_sl_group_cd)	= UCase(Trim(Request("cboSLGroup")))
	I3_b_storage_location(I3_tax_class)		= UCase(Trim(Request("cboTaxClass")))
	
	If	UCase(Request("cboSLType")) = "E" Then	
		I3_b_storage_location(I3_ext_sl_type) = UCase(Trim(Request("cboExtSLType")))
	End If

	I3_b_storage_location(I3_mrp_used_flg) = UCase(Trim(Request("optMrpUsedFlg")))
    
    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "create"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "update"
    End If
    
   
    
    '-----------------------
	'Com Action Area
	'-----------------------
	Call pPB6G010.B_MANAGE_STORAGE_LOCATION(gStrGlobalCollection, iCommandSent, _
											I1_b_plant_plant_cd, _
											I2_b_biz_partner_bp_cd, _
											I3_b_storage_location)
	If CheckSYSTEMError(Err, True) = True Then
		Set pPB6G010 = Nothing	
		Response.End
	End If

    Set pPB6G010 = Nothing		
	


	'-----------------------
	'Result data display area
	'----------------------- 

%>
<Script Language=vbscript>
	With parent																			
		.DbSaveOk
	End With
</Script>
