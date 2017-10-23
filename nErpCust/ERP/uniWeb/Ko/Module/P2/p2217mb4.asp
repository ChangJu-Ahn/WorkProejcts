<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2217mb4.asp
'*  4. Program Name         : Plant query
'*  5. Program Desc         :
'*  6. Comproxy List        : PB6S101.cBLkUpPlt
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)		: Lee Hyun Jae
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change"
'**********************************************************************************************-->
<% Call LoadBasisGlobalInf
   Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

Call HideStatusWnd

On Error Resume Next

Dim pPB6S101

Dim I2_plant_cd
Dim E1_b_biz_area, E3_b_plant, E4_p_mfg_calendar_type, iStatusCodeOfPrevNext

Const P058_E1_biz_area_cd = 0
Const P058_E1_biz_area_nm = 1
Const P058_E1_biz_area_eng_nm = 2


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

Const P058_E4_cal_type = 0
Const P058_E4_cal_type_nm = 1


I2_plant_cd = Request("txtPlantCd")
    
Set pPB6S101 = Server.CreateObject("PB6S101.cBLkUpPlt")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB6S101.B_LOOK_UP_PLANT_SVR(gStrGlobalCollection, "", I2_plant_cd, E1_b_biz_area, , _
                     E3_b_plant, E4_p_mfg_calendar_type, iStatusCodeOfPrevNext)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB6S101 = Nothing
	Response.End
End If

Set pPB6S101 = Nothing

If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
	Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
End If
	

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "With parent.frm1" & vbCrLf
Response.Write "	.txtPlantNm.value = """ & ConvSPChars(E3_b_plant(B153_E3_plant_nm)) & """" & vbCrLf
Response.Write "	.txtPH.text = """ & UniDateClientFormat(UniDateAdd("d", E3_b_plant(B153_E3_plan_hrzn), GetSvrDate, gServerDateFormat)) & """" & vbCrLf
Response.Write "	.txtDTF.text = """ & UniDateClientFormat(UniDateAdd("d", E3_b_plant(B153_E3_dtf_for_mps), GetSvrDate, gServerDateFormat)) & """" & vbCrLf
Response.Write "	.txtPTF.text = """ & UniDateClientFormat(UniDateAdd("d", E3_b_plant(B153_E3_ptf_for_mps), GetSvrDate, gServerDateFormat)) & """" & vbCrLf
Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
%>
