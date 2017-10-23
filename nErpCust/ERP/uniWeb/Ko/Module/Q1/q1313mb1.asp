<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1312MB1
'*  4. Program Name         : 부적합처리정보등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG270
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

Dim PQBG270													'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strData
	
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount          
Dim strDispositionCd   
Dim strInspClassCd
	
Dim EG1_group_export

Const C_SHEETMAXROWS_D = 100

Const Q092_EG1_E1_disposition_cd = 0
Const Q092_EG1_E1_disposition_nm = 1
Const Q092_EG1_E1_stock_type_cd = 2
Const Q092_EG1_E1_insp_class_cd = 3
Const Q092_EG1_E1_stock_type_nm = 4
Const Q092_EG1_E1_insp_class_nm = 5

lgStrPrevKey = Request("lgStrPrevKey")
LngMaxRow = Request("txtMaxRows")
strDispositionCd = Request("txtDispositionCd")
strInspClassCd = Request("txtInspClassCd")

Set PQBG270 = Server.CreateObject("PQBG270.cQListDispositionSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

if lgStrPrevKey <> "" Then
	strDispositionCd = lgStrPrevKey
end if

Call PQBG270.Q_LIST_DISPOSITION_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D,  strDispositionCd, strInspClassCd, EG1_group_export)

If CheckSYSTEMError(Err,True) = True Then
	Set PQBG270 = Nothing
	Response.End
End If

Dim TmpBuffer
Dim iTotalStr
	
Dim i
For i = 0 To UBound(EG1_group_export, 1)
    If i < C_SHEETMAXROWS_D Then		
		ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q092_EG1_E1_disposition_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q092_EG1_E1_disposition_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q092_EG1_E1_stock_type_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q092_EG1_E1_insp_class_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q092_EG1_E1_stock_type_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q092_EG1_E1_insp_class_cd)))
		strData = strData & Chr(11) & LngMaxRow + i + 1
		strData = strData & Chr(11) & Chr(12)
		TmpBuffer(i) = strData
    Else
		StrNextKey = EG1_group_export(i, Q092_EG1_E1_disposition_cd)
    End If
Next  

iTotalStr = Join(TmpBuffer, "")
Set PQBG270 = Nothing
%>
<Script Language=vbscript>
With Parent
   
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
		
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.hDispositionCd.Value = "<%=ConvSPChars(strNextKey)%>"
		.frm1.hInspClassCd.Value = "<%=ConvSPChars(strInspClassCd)%>"
		.DbQueryOk
    End If	

    If Trim(UCase(.frm1.txtDispositionCd.Value)) = "<%=Trim(UCase(ConvSPChars(EG1_group_export(0,Q092_EG1_E1_disposition_cd))))%>" Then
		.frm1.txtDispositionCd.Value = "<%=Trim(ConvSPChars(EG1_group_export(0, Q092_EG1_E1_disposition_cd)))%>"
		.frm1.txtDispositionNm.Value = "<%=Trim(ConvSPChars(EG1_group_export(0, Q092_EG1_E1_disposition_nm)))%>"
	End If        			
End with
</Script>	
