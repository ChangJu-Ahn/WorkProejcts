<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : Vendor Managed Inventory
'*  3. Program ID           : i1531mb4.asp
'*  4. Program Name         : Receipt Purchase Order for VMI
'*  5. Program Desc         : Receipt Purchase Order for VMI
'*  6. Modified date(First) : 2003-01-23
'*  7. Modified date(Last)  : 2003-01-23
'*  8. Modifier (First)     : Ahn, Jung Je
'*  9. Modifier (Last)      : Ahn, Jung Je
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(¢Ð) means that "Do not change"
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%	
	Call LoadBasisGlobalInf()								
    
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd 

	Dim pPI5S230											
	Dim txtSpread
	
	Dim I1_b_plant_plant_cd
    Dim I2_i_vmi_goods_mvmt_hdr
		Const I513_I2_trns_type = 0
		Const I513_I2_document_dt = 1
	ReDim I2_i_vmi_goods_mvmt_hdr(I513_I2_document_dt)
    Dim I3_b_pur_GRP_pur_GRP
    Dim I4_m_mvmt_type_io_type_cd
    
    Dim E1_m_pur_goods_mvmt_rcpt_no
    Dim prErrorPosition

	
	I1_b_plant_plant_cd	= Request("txthPlantCd")
	I2_i_vmi_goods_mvmt_hdr(I513_I2_trns_type) = "FR"
	I2_i_vmi_goods_mvmt_hdr(I513_I2_document_dt) = UNIConvDate(Request("txtGRDt")) 
	I3_b_pur_GRP_pur_GRP = Request("txtGroupCd")
	I4_m_mvmt_type_io_type_cd = Request("cboMvmtType")
	txtSpread = Request("txtSpread")
	
	  
	Set pPI5S230 = Server.CreateObject("PI5S230.cIVMIManagePurItem")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
	
	Call pPI5S230.I_VMI_MANAGE_PUR_ITEM(gStrGlobalCollection, _
										I1_b_plant_plant_cd, _
										I2_i_vmi_goods_mvmt_hdr, _
										I3_b_pur_GRP_pur_GRP, _
										I4_m_mvmt_type_io_type_cd, _
										txtSpread, _
										E1_m_pur_goods_mvmt_rcpt_no, _
										prErrorPosition)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPI5S230 = Nothing														
		If prErrorPosition <> "" Then
			Response.Write "<Script Language=VBScript>" & vbCrLF
			Response.Write "Call parent.SheetFocus(" & prErrorPosition & ")" & vbCrLF
			Response.Write "</Script>" & vbCrLF
			Response.End
		End If
		Response.End
	End If
			
	Set pPI5S230 = Nothing													

Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.frm1.txtMvmtNo.value = """ & E1_m_pur_goods_mvmt_rcpt_no & """" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>