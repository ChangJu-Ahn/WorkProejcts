<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI STORAGE LOCATION ENTRY
'*  3. Program ID           : i1511mb2.asp	
'*  4. Program Name         : VMI Storage Location Entry
'*  5. Program Desc         :
'*  6. Comproxy List        : PI5G020

'*  7. Modified date(First) : 2003/01/03
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              : VB Conversion
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<%													
Call LoadBasisGlobalInf()

On Error Resume Next												
Err.Clear

Call HideStatusWnd														

Dim PI5G020															

Dim I1_b_plant_plant_cd
Dim I2_i_vmi_storage_location 
Dim iCommandSent

Const I500_I2_sl_cd			= 0
Const I500_I2_sl_nm			= 1
Const I500_I2_sl_group_cd	= 2
Const I500_I2_inv_mgr		= 3

redim I2_i_vmi_storage_location(I500_I2_inv_mgr)
	
    If Request("txtSLCd1") = "" Then											
		Call DisplayMsgBox("169902", vbOKOnly, "", "", I_MKSCRIPT)  
		Response.End 
	End If
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))								
	
	Set PI5G020 = Server.CreateObject("PI5G020.cIVMIManageStgLoc")
										
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
    '-----------------------
    'Data manipulate area
    '-----------------------
	I1_b_plant_plant_cd 							= UCase(Trim(Request("txtPlantCd")))
	I2_i_vmi_storage_location(I500_I2_sl_cd) 		= UCase(Trim(Request("txtSLCd1")))
	I2_i_vmi_storage_location(I500_I2_sl_nm) 		= Request("txtSLNm1")
	I2_i_vmi_storage_location(I500_I2_sl_group_cd)	= UCase(Trim(Request("cboSLGroup")))
	I2_i_vmi_storage_location(I500_I2_inv_mgr)   	= UCase(Trim(Request("cboInvMgr")))
	
    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    Else
		If Request("txtCommand") = "DELETE" Then
			iCommandSent = "DELETE"
		Else
			iCommandSent = "UPDATE"
		End If
	End If

    '-----------------------
	'Com Action Area
	'-----------------------
	Call PI5G020.I_MANAGE_VMI_STORAGE_LOCATION(gStrGlobalCollection, iCommandSent, _
											I1_b_plant_plant_cd, _
											I2_i_vmi_storage_location)

	If CheckSYSTEMError(Err, True) = True Then
		Set PI5G020 = Nothing														
		Response.End
	End If

    Set PI5G020 = Nothing																

%>
<Script Language=vbscript>
	With parent																			
		If "<%=iCommandSent%>" = "DELETE" Then
			.DbDeleteOk
		Else 
			.DbSaveOk
		End if 
	End With
</Script>
<%					
	Response.End 
%>

