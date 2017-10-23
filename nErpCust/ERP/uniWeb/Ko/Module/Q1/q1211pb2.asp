<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% 
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1211PB2
'*  4. Program Name         : 공장별 품목 팝업 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG001
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
Call loadInfTB19029B("I", "*", "NOCOOKIE","PB")
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
																			
Dim StrNextKey1																'Code로 조회시 
Dim lgStrPrevKey1															'코드 이전 값 
Dim lgStrPrevKey2															'코드 이전 값 
Dim StrNextKey2
Dim StrConFlg	
Dim LngMaxRow																' 현재 그리드의 최대Row
	
Dim PQBG001
Dim lgstr
Dim StrData
Dim i
		
Const C_SHEETMAXROWS_D = 100
		
Dim I3_b_item
	Const Q505_I3_item_cd = 0
	Const Q505_I3_item_nm = 1		
ReDim I3_b_item(Q505_I3_item_nm)

Dim I5_search_char
	Const Q505_I5_select_char = 0
	Const Q505_I5_main_class_cd = 1		
ReDim I5_search_char(Q505_I5_main_class_cd)


Dim E2_q_wks_interface_flag

Dim EG1_group_export
	Const Q505_EG1_E1_item_cd = 0
	Const Q505_EG1_E1_item_nm = 1
	Const Q505_EG1_E1_spec = 2
	
Dim strlgInspClassCd
Dim strPlantCd

	strlgInspClassCd = Request("cboInspClassCd")
	strPlantCd = Request("PlantCd")
	StrConFlg = Request("lgConFlg")	
	LngMaxRow = Request("txtMaxRows")
	I5_search_char(Q505_I5_main_class_cd) = Request("lgInspClassCd")
		
	I3_b_item(Q505_I3_item_cd) = Request("txtItemCd")
	I3_b_item(Q505_I3_item_nm) = Request("txtItemNm")

	I5_search_char(Q505_I5_select_char) = Request("search_char")		'검색조건 추가 
	
	Set PQBG001 = Server.CreateObject("PQBG001.cQListItemByPlantSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
		
	Call PQBG001.Q_LIST_ITEM_BY_PLANT_SVR(gStrGlobalCollection, _
	  	  								  C_SHEETMAXROWS_D, _
				 						  strlgInspClassCd, _
										  strPlantCd, _
										  I3_b_item, _
										  StrConFlg, _
										  I5_search_char, _
										  EG1_group_export, _
										  E2_q_wks_interface_flag)

	If CheckSYSTEMError(Err,True) = True Then
		Set PQBG001 = Nothing
		Response.End
	End If		

	StrConFlg = E2_q_wks_interface_flag 

	Dim TmpBuffer
	Dim iTotalStr
	ReDim TmpBuffer(UBound(EG1_group_export, 1))

	For i = 0 To UBound(EG1_group_export, 1)
	    If i < C_SHEETMAXROWS_D Then
			
			TmpBuffer(i) = Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q505_EG1_E1_item_cd))) & _
						   Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q505_EG1_E1_item_nm))) & _
						   Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q505_EG1_E1_spec))) & _
						   Chr(11) & LngMaxRow + i	+ 1 & _
						   Chr(11) & Chr(12)	
	    Else
			If StrConFlg = "C" Then
				StrNextKey1 = EG1_group_export(i, Q505_EG1_E1_item_cd)
			Else
				StrNextKey1 = EG1_group_export(i, Q505_EG1_E1_item_cd)
				StrNextKey2 = EG1_group_export(i, Q505_EG1_E1_item_nm)			
			End If
	    End If
	Next  

	iTotalStr = Join(TmpBuffer, "")	

	Set PQBG001 = nothing	

%>		    
<Script Language="vbscript">   
With parent
	.ggoSpread.Source = .vspdData
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"	
	
	.lgStrPrevKey1 = "<%=StrNextKey1%>"
	.lgStrPrevKey2 = "<%=StrNextKey2%>"
	.lgConFlg = "<%=StrConFlg%>"	
	
	If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.hItemCd.value = "<%=ConvSPChars(Request("txtItemCd"))%>"	
		.hItemNm.value = "<%=ConvSPChars(Request("txtItemNm"))%>"
		.DbQueryOk
	End If	
	.vspdData.focus
End With
</Script>