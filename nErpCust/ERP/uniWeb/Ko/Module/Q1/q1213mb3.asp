<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality
'*  2. Function Name        : 
'*  3. Program ID           : q1213mb3.asp
'*  4. Program Name         : 품목별 검사기준의 검사방식 LOOK UP
'*  5. Program Desc         : 
'*  6. Comproxy List        : PQBG300.cQLookUpInspMethodSvr
'                             
'*  7. Modified date(First) : 2003/06/17
'*  8. Modified date(Last)  : 2003/06/17
'*  9. Modifier (First)     : Jaewoo Koh
'* 10. Modifier (Last)      : Jaewoo Koh
'* 11. Comment              :
'**********************************************************************************************
On Error Resume Next

Call HideStatusWnd
	
'Export Views
'B_Plant
Const Q300_E1_plant_cd = 0
Const Q300_E1_plant_nm = 1
    
'B_Item
Const Q300_E2_item_cd = 0
Const Q300_E2_item_nm = 1
    
'Q_Inspection_Item
Const Q300_E3_inspection_item_cd = 0
Const Q300_E3_inspection_item_nm = 1
    
'B_Minor (For Insp method)
Const Q300_E4_minor_insp_method_cd = 0
Const Q300_E4_minor_insp_method_nm = 1
    
	
Dim objPQBG300

Dim E1_b_plant
Dim E2_b_item
Dim E3_q_inspection_item
Dim E4_b_minor_insp_method
	
Set objPQBG300 = Server.CreateObject("PQBG300.cQLookUpInspMethodSvr")    

If CheckSYSTEMError(Err,True) = True Then
   Response.End
End if
	    
Call objPQBG300.Q_LOOK_UP_INSP_METHOD_SVR(gStrGlobalCollection, _
											Request("txtPlantCd"), _
											Request("txtInspClassCd"), _
											Request("txtItemCd"), _
											Request("txtInspItemCd"), _
											E1_b_plant, _
											E2_b_item, _
											E3_q_inspection_item, _
											E4_b_minor_insp_method)
	
If CheckSYSTEMError(Err,True) = true Then
   Set objPQBG300 = Nothing
   Response.End
End if
Set objPQBG300 = Nothing
%>
<Script Language=vbscript>
With parent.frm1		
	.txtPlantNm.value = "<%=ConvSPChars(Trim(E1_b_plant(Q300_E1_plant_nm)))%>"
	.txtItemNm.value = "<%=ConvSPChars(Trim(E2_b_item(Q300_E2_item_nm)))%>"
	.txtInspItemNm.value = "<%=ConvSPChars(Trim(E3_q_inspection_item(Q300_E3_inspection_item_nm)))%>"
	.txtInspMthdCd.value = "<%=ConvSPChars(Trim(E4_b_minor_insp_method(Q300_E4_minor_insp_method_cd)))%>"
	.txtInspMthdNm.value = "<%=ConvSPChars(Trim(E4_b_minor_insp_method(Q300_E4_minor_insp_method_nm)))%>"
	
End With
	
Call parent.LookUpInspMethodOk()	
</Script>
