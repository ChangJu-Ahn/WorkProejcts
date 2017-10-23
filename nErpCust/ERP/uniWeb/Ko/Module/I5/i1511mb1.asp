<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI Storage Location
'*  3. Program ID           : i1511mb1.asp
'*  4. Program Name         : VMI 창고정보등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : PI5G010
'
'*  7. Modified date(First) : 2003/01/02
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : 
'* 11. Comment              : VB Conversion
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%	
Call LoadBasisGlobalInf()
Call HideStatusWnd														

On Error Resume Next										
Err.Clear                                                   

Dim PI5G010													
'-----------------------
'IMPORTS View
'-----------------------
Dim I1_b_plant_plant_cd
Dim I2_i_vmi_storage_location_sl_cd
'-----------------------
'EXPORTS View
'-----------------------
Dim E1_i_vmi_storage_location
	Const I501_E1_sl_cd			= 0
	Const I501_E1_sl_nm			= 1
	Const I501_E1_sl_group_cd	= 2
	Const I501_E1_inv_mgr		= 3
REDim E1_i_vmi_storage_location(I501_E1_sl_group_cd)

    If Request("txtPlantCd") = "" or Request("txtSLCd") = "" Then								
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)         
		Response.End 
	End If

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I1_b_plant_plant_cd				= Request("txtPlantCd")
	I2_i_vmi_storage_location_sl_cd = Request("txtSLCd")


    Set PI5G010 = Server.CreateObject("PI5G010.cIVMILookUpStgLoc")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

    '-----------------------
    'Com action area
    '-----------------------
	Call PI5G010.I_LOOK_UP_VMI_STORAGE_LOCATION(gStrGlobalCollection, _
											I1_b_plant_plant_cd, _
											I2_i_vmi_storage_location_sl_cd, _
											E1_i_vmi_storage_location)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set PI5G010 = Nothing
		Response.End
	End If
	
	Set PI5G010 = Nothing	
	
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write "With Parent" & vbcr
	
	Response.Write "    .frm1.txtSLCd1.value  = """ & ConvSPChars(E1_i_vmi_storage_location(I501_E1_sl_cd)) & """" & vbcr
	Response.Write "    .frm1.txtSLNm1.value = """ & ConvSPChars(E1_i_vmi_storage_location(I501_E1_sl_nm))    & """" & vbcr
	Response.Write "    .frm1.cboSLGroup.value = """ & UCase(E1_i_vmi_storage_location(I501_E1_sl_group_cd))           & """" & vbcr
	Response.Write "    .frm1.cboInvMgr.value = """ & UCase(E1_i_vmi_storage_location(I501_E1_inv_mgr))       & """" & vbcr
	Response.Write "    .frm1.txthPlantCd.value = """ & ConvSPChars(I1_b_plant_plant_cd)       & """" & vbcr

	Response.Write "    .DbQueryOk " & vbcr

	Response.Write "End With"       & vbcr
	Response.Write "</Script>"      & vbcr

	Response.End 

%>
