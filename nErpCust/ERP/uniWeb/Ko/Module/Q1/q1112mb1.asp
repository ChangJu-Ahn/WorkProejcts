<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1112MB1
'*  4. Program Name         : 검사항목등록 
'*  5. Program Desc         : 검사항목등록 
'*  6. Component List       : PQBG040
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
Call LoadBasisGlobalInf

On Error Resume Next													
Call HideStatusWnd 

Dim PQBG040	
													
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim strInspItemCd	  
Dim EG1_export_group
Dim strData

Const C_SHEETMAXROWS_D = 100

	lgStrPrevKey = Request("lgStrPrevKey")
	LngMaxRow = Request("txtMaxRows")
	strInspitemCd = Request("txtInspItemCd")

	If lgStrPrevKey <> "" Then
		strInspitemCd = lgStrPrevKey
	End If


	Set PQBG040 = Server.CreateObject ("PQBG040.cQListInspItemSvr")


	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	Call PQBG040.Q_LIST_INSP_ITEM_SVR (gStrGlobalCollection, _
									   C_SHEETMAXROWS_D, _
									   strInspitemCd, _
									   EG1_export_group)
			
	If CheckSYSTEMError(Err,True) = True Then
		Set PQBG040 = Nothing
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

	Dim TmpBuffer
	Dim iTotalStr
	ReDim TmpBuffer(UBound(EG1_export_group, 1))

	For LngRow = 0 To UBound(EG1_export_group, 1)
	    If LngRow < C_SHEETMAXROWS_D Then    
			
			strData = Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, 0))) & _
					  Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, 1))) & _
					  Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, 2))) & _
					  Chr(11) & "" & _
					  Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, 3))) & _
					  Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, 5))) & _
					  Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, 4))) & _
					  Chr(11) & LngMaxRow + LngRow + 1 & _
					  Chr(11) & Chr(12)
			TmpBuffer(LngRow) = strData
	    Else
			StrNextKey = EG1_export_group(LngRow,0)
	    End If
	Next

	iTotalStr = Join(TmpBuffer, "")
%>

<Script Language=vbscript>
With Parent		
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"		
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else		
		' Request값을 hidden input으로 넘겨줌 
		.frm1.hInspItemCd.value = "<%=ConvSPChars(strInspItemCd)%>"
		.DbQueryOk
    End If		
	
	If UCase(Trim(.frm1.txtInspItemCd.Value)) = "<%=ConvSPChars(EG1_export_group(0,0))%>" Then
		.frm1.txtInspItemCd.Value = "<%=ConvSPChars(EG1_export_group(0,0))%>"
		.frm1.txtInspItemNm.Value = "<%=ConvSPChars(EG1_export_group(0,1))%>"    
	Else
		.frm1.txtInspItemNm.Value = ""
	End If
End with		
</Script>	
<%
Set PQBG040 = Nothing
%>