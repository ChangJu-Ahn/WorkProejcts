<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1113MB1
'*  4. Program Name         : 불량률단위등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG050
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/10
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
Call LoadBasisGlobalInf
		
On Error Resume Next										
Call HideStatusWnd 

Dim PQBG060													'☆ : 조회용 ComProxy Dll 사용 변수 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim strDefectRatioUnitCd
Dim strData
Dim EG1_group_export
	
Const C_SHEETMAXROWS_D = 100

	lgStrPrevKey = Request("lgStrPrevKey")
	LngMaxRow = Request("txtMaxRows")
	strDefectRatioUnitCd = Request("txtDefectRatioUnitCd")

	If lgStrPrevKey <> "" Then
		strDefectRatioUnitCd = lgStrPrevKey
	End If

	Set PQBG060 = Server.CreateObject("PQBG060.cQListDfctRatUnitSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	Call PQBG060.Q_LIST_DFCT_RATIO_UNIT_SVR (gStrGlobalCollection, _
											 C_SHEETMAXROWS_D, _
											 strDefectRatioUnitCd, _
											 EG1_group_export)

	If CheckSYSTEMError(Err,True) = True Then
		Set PQBG060 = Nothing
		Response.End
	End If

	Dim TmpBuffer
	Dim iTotalStr
	ReDim TmpBuffer(UBound(EG1_group_export, 1))

	For LngRow = 0 To UBound(EG1_group_export, 1)
	    
	    If LngRow < C_SHEETMAXROWS_D Then
			strData = Chr(11) & Trim(ConvSPChars(EG1_group_export(LngRow, 0))) & _
					  Chr(11) & Trim(ConvSPChars(EG1_group_export(LngRow, 1))) & _
					  Chr(11) & Trim(ConvSPChars(EG1_group_export(LngRow, 2))) & _
					  Chr(11) & LngMaxRow + LngRow + 1 & _
					  Chr(11) & Chr(12)
			TmpBuffer(LngRow) = strData
	    Else
			StrNextKey = EG1_group_export(LngRow, 0)
	    End If
	Next

	iTotalStr = Join(TmpBuffer, "")  

	Set PQBG060 = Nothing
%>
<Script Language=vbscript>
With Parent
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"		

	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.hDefectRatioUnitCd.value = "<%=ConvSPChars(strDefectRatioUnitCd)%>"
		.DbQueryOk
	End If	
	
	If Trim(.frm1.txtDefectRatioUnitCd.Value) = Trim("<%=ConvSPChars(EG1_group_export(0,0))%>") Then
		.frm1.txtDefectRatioUnitCd.Value = "<%=ConvSPChars(EG1_group_export(0,0))%>"
		.frm1.txtDefectRatioUnitNm.Value = "<%=ConvSPChars(EG1_group_export(0,1))%>"		
	Else
		.frm1.txtDefectRatioUnitNm.Value = ""	
	End If
End with			
</Script>	
