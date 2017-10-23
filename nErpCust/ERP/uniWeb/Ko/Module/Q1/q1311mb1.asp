<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1311MB1
'*  4. Program Name         : 불량유형 정보등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG230
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

Dim PQBG230													'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount          
Dim strPlantCd
Dim strData


Dim strInspClassCd
ReDim strInspClassCd(1)
Const Q081_I1_insp_class_cd = 0
Const Q081_I1_defect_type_cd = 1
	
Dim E1_b_plant
Dim E2_q_defect_type
Dim EG1_group_export
	
'Dim E1_b_plant
				   
lgStrPrevKey = Request("lgStrPrevKey")
LngMaxRow = Request("txtMaxRows")
strPlantCd = Request("txtplantCd")

strInspClassCd(Q081_I1_insp_class_cd)	= Request("cboInspClassCd")

Const C_SHEETMAXROWS_D = 100

if lgStrPrevKey <> "" Then
	strInspClassCd(Q081_I1_defect_type_cd) = lgStrPrevKey
end if

Set PQBG230 = Server.CreateObject ("PQBG230.cQListDefTypeSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

CALL PQBG230.Q_LIST_DEFECT_TYPE_SVR (gStrGlobalCollection, C_SHEETMAXROWS_D, strInspClassCd, strPlantCd, _
									E1_b_plant, E2_q_defect_type,  EG1_group_export )
									
If CheckSYSTEMError(Err,True) = True Then
	Set PQBG230 = Nothing	
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQBG230 = Nothing

If IsEmpty(EG1_group_export) = true then
	Set PQBG230 = Nothing
	Response.End
End If

Dim TmpBuffer
Dim iTotalStr
	
Dim i
For i = 0 To UBound(EG1_group_export, 1)
	If i < C_SHEETMAXROWS_D Then
		ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
			
			strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, 0)))
			strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, 1)))
			strData = strData & Chr(11) & LngMaxRow + i + 1
       		strData = strData & Chr(11) & Chr(12)
		TmpBuffer(i) = strData
	ELSE 
		StrNextKey = EG1_group_export(i, 0)
    End If
Next

iTotalStr = Join(TmpBuffer, "")
%>
<Script Language=vbscript>
	
With Parent
	.frm1.txtPlantNm.Value = "<%=ConvSPChars(E1_b_plant(1))%>"
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		<% ' Request값을 hidden input으로 넘겨줌 %>
		.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		.frm1.hInspClassCd.value = "<%=ConvSPChars(strInspClassCd((Q081_I1_insp_class_cd)))%>"
		.DbQueryOk
    End If		
End with
</Script>	
<%
Set PQBG230 = Nothing
%>