<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","PB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1211PB1
'*  4. Program Name         : 품목별 검사항목 팝업 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG020
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd 

Dim LngRow
Dim LngMaxRow
Dim intGroupCount 
Dim PQBG120

Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strRoutNo
Dim strOprNo
Dim strInspMthdCd
Dim lgStrPrevKey
Dim i
Dim strData
Dim strInspItemCd
Dim StrNextKey
	
strPlantCd = Request("txtPlantCd")
strItemCd = Request("txtItemCd")
strRoutNo = Request("txtRoutNo")
strOprNo = Request("txtOprNo")
strInspMthdCd = Request("txtInspMthdCd")
strInspItemCd = Request("txtInspItemCd")
lgStrPrevKey = Request("lgStrPrevKey")
LngMaxRow = Request("txtMaxRows")
		
Redim strInspClassCd(1)
Const Q039_I2_insp_class_cd = 0
Const Q039_I2_insp_method_cd = 1   


Dim E1_q_inspection_item
	    
Dim E3_b_plant
Redim E3_b_plant(1)
Const Q039_E1_plant_cd = 0
Const Q039_E1_plant_nm = 1    

Dim E2_b_item
Redim E2_b_item(1)
Const Q039_E2_item_cd = 0
Const Q039_E2_item_nm = 1

Dim E4_p_routing_header
Dim E5_p_routing_detail

Dim EG1_group_export
Dim EG2_group_export
Dim EG3_group_export
Dim EG4_group_export
Dim EG5_group_export
Dim EG6_group_export
Dim EG7_group_export
Dim EG8_group_export   
    
Const C_SHEETMAXROWS_D = 100	
		
strInspClassCd(Q039_I2_insp_class_cd) = Request("txtInspClassCd")
strInspClassCd(Q039_I2_insp_method_cd) = strInspMthdCd

If lgStrPrevKey <> "" Then
	strInspItemCd = lgStrPrevKey
End If

Set PQBG120 = Server.CreateObject("PQBG120.cQListInspStdItemSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

CALL PQBG120.Q_LIST_INSP_STD_BY_ITEM_ON_POPUP_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D,strInspItemCd ,strItemCd, _
													strInspClassCd, strPlantCd, strRoutno, strOprNo, E1_q_inspection_item, _
													E2_b_item, E3_b_plant, E4_p_routing_header, E5_p_routing_detail, _
													EG1_group_export, EG2_group_export, EG3_group_export, _
													EG4_group_export, EG5_group_export, EG6_group_export, _
													EG7_group_export, EG8_group_export)
						
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Set PQBG120 = Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQBG120 = Nothing

Dim TmpBuffer
Dim iTotalStr

For i = 0 To UBound(EG8_group_export, 1)
	If i < C_SHEETMAXROWS_D Then
		ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		
		strData = strData & Chr(11) & Trim(ConvSPChars(EG8_group_export(i, 11)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG7_group_export(i, 0)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG7_group_export(i, 1)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG8_group_export(i, 1)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG3_group_export(i, 0)))
		strData = strData & Chr(11) & LngMaxRow + i + 1	
  		strData = strData & Chr(11) & Chr(12)	
		TmpBuffer(LngRow) = strData
    ELSE 
		StrNextKey = EG7_group_export(i, 0)
    End If
Next 

iTotalStr = Join(TmpBuffer, "")
%>		    
<Script Language="vbscript">     
With parent    
	.ggoSpread.Source = .vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
	.vspdData.focus
	
	.txtPlantNm.Value = "<%=ConvSPChars(E3_b_plant(1))%>"
	.txtItemNm.Value = "<%=ConvSPChars(E2_b_item(1))%>"
	.txtRoutNoDesc.Value = "<%=ConvSPChars(E4_p_routing_header(1))%>"
	
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	
	If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		<% ' Request값을 hidden Varialbe로 넘겨줌 %>
		.hInspItemCd = "<%=ConvSPChars(strInspItemCd)%>"
		.DbQueryOk  
	End If
End with
</Script>
<%
Set pq12118 = Nothing
%>
