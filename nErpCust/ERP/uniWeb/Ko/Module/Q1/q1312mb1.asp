<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1312MB1
'*  4. Program Name         : �ҷ����� ������� 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG250
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next												
Call HideStatusWnd 

Dim PQBG250													'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim strMode
Dim strData
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim intGroupCount          
Dim strPlantCd
Dim strInspClassCd
Dim E1_q_defect_cause
Dim E2_b_plant
Dim EG1_group_export
	
ReDim strInspClassCd(1)
Const Q086_I1_insp_class_cd = 0
Const Q086_I1_defect_cause_cd = 1

lgStrPrevKey = Request("lgStrPrevKey")
LngMaxRow = Request("txtMaxRows")
strPlantCd = Request("txtPlantCd")

strInspClassCd(Q086_I1_insp_class_cd) = Request("cboInspClassCd")

Const C_SHEETMAXROWS_D = 100

if lgStrPrevKey <> "" Then
	strInspClassCd(Q086_I1_defect_cause_cd) = lgStrPrevKey
end if


Set PQBG250 = Server.CreateObject("PQBG250.cQListDefCauseSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

CALL PQBG250.Q_LIST_DEFECT_CAUSE_SVR (gStrGlobalCollection, C_SHEETMAXROWS_D, strInspClassCd, _
									strPlantCd, E1_q_defect_cause, E2_b_plant, EG1_group_export)

If CheckSYSTEMError(Err,True) = True Then
	Set PQBG250 = Nothing	
	Response.End														'��: �����Ͻ� ���� ó���� ������ 
End If

Set PQBG250 = Nothing

If IsEmpty(EG1_group_export) = true then
	Set PQBG250 = Nothing
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
	.frm1.txtPlantNm.Value = "<%=ConvSPChars(E2_b_plant(1))%>"
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		<% ' Request���� hidden input���� �Ѱ��� %>
		.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		.frm1.hInspClassCd.value = "<%=ConvSPChars(strInspClassCd(Q086_I1_insp_class_cd))%>"
		.DbQueryOk
    End If		
End with
</Script>	
<%
Set pq13128 = Nothing
%>
