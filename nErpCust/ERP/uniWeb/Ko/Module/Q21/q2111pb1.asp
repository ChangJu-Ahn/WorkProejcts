<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","PB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1211PA1
'*  4. Program Name         : 
'*  5. Program Desc         : 검사현황 팝업 
'*  6. Component List       : 
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
													 '☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngRow
Dim LngMaxRow
Dim intGroupCount 
Dim StrNextKey
Dim PQIG050
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strInspReqNo
Dim strBpCd
Dim strCustCd
Dim strWcCd
Dim strFrInspDt
Dim strToInspDt
Dim strStatusFlag
Dim strDecision
Dim lgStrPrevKey
DIM I1_q_inspection_result
Dim i
Dim StrData

Dim E1_q_inspection_request
Dim E2_q_inspection_result

Dim EG2_group_export
Dim EG3_group_export
Dim EG4_group_export
Dim EG5_group_export
Dim EG6_group_export
Dim EG7_group_export
Dim EG8_group_export
Dim EG9_group_export
Dim EG10_group_export
		
REDIM I1_q_inspection_result(6)	
	
Const Q220_I1_insp_class_cd = 0
Const Q220_I1_decision = 1
Const Q220_I1_bp_cd = 2
Const Q220_I1_wc_cd = 3
Const Q220_I1_item_cd = 4
Const Q220_I1_plant_cd = 5
Const Q220_I1_status_flag = 6		
		
CONST C_SHEETMAXROWS_D = 100
		
strPlantCd			= Request("txtPlantCd")
strItemCd			= Request("txtItemCd")
strInspClassCd		= Request("txtInspClassCd")
strInspReqNo		= Request("txtInspReqNo")
strBpCd				= Request("txtBpCd")
strCustCd			= Request("txtCustCd")
strWcCd				= Request("txtWcCd")
strFrInspDt			= Request("txtFrInspDt")
strToInspDt			= Request("txtToInspDt")
strStatusFlag		= Request("txtStatusFlag")
strDecision			= Request("txtDecision")
lgStrPrevKey		= Request("lgStrPrevKey")
LngMaxRow			= Request("txtMaxRows")


Set PQIG050 = Server.CreateObject("PQIG050.cQListInspResultSvr")


If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Select Case strInspClassCd
	Case "R"
		I1_q_inspection_result(Q220_I1_bp_cd) = strBpCd
	Case "P"
		I1_q_inspection_result(Q220_I1_wc_cd) = strWcCd
	Case "F"
	
	Case "S"
		strCustCd = strCustCd
End Select 


if strFrInspDt <> "" then
	strFrInspDt = UNIConvDate(strFrInspDt)
End If
if strToInspDt <> "" then
	strToInspDt = UNIConvDate(strToInspDt)
End If
If strStatusFlag <> "" Then
	I1_q_inspection_result(Q220_I1_status_flag) = strStatusFlag
End If
If strDecision <> "" Then
	I1_q_inspection_result(Q220_I1_decision) = strDecision
End If

If lgStrPrevKey = "" then
	strInspReqNo = strInspReqNo
Else
	strInspReqNo = lgStrPrevKey
End If

I1_q_inspection_result(Q220_I1_insp_class_cd) = strInspClassCd
I1_q_inspection_result(Q220_I1_plant_cd) = strPlantCd
I1_q_inspection_result(Q220_I1_item_cd) = strItemCd

 			
CALL PQIG050.Q_LIST_INSP_RESULT_SVR (gStrGlobalCollection, C_SHEETMAXROWS_D, I1_q_inspection_result, _
									strFrInspDt, strToInspDt, strInspReqNo, lgStrPrevKey,  , EG2_group_export , _
									EG3_group_export, EG4_group_export , EG5_group_export , _
									EG6_group_export, EG7_group_export , EG8_group_export , _
									EG9_group_export, EG10_group_export, E1_q_inspection_request , E2_q_inspection_result )

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG050 = Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQIG050 = Nothing

If IsEmpty(EG2_group_export) = true then
	Set PQIG050 = Nothing
	Response.End
End If

For i = 0 To UBound(EG2_group_export, 1)
	If i < C_SHEETMAXROWS_D Then		
		strData = strData & Chr(11) & Trim(ConvSPChars(E1_q_inspection_request(i, 0)))
		strData = strData & Chr(11) & Trim(ConvSPChars(E2_q_inspection_result(i, 0)))			
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 6)))	
		strData = strData & Chr(11) & Trim(ConvSPChars(EG4_group_export(i, 0)))

		Select Case UCase(I1_q_inspection_result(Q220_I1_insp_class_cd))
			CASE "R"							
				strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 4)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG6_group_export(i, 0)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 5)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG5_group_export(i, 0)))	
			CASE "P"	
				strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 4)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG6_group_export(i, 0)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 4)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG6_group_export(i, 0)))
			CASE "F"
				strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 4)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG6_group_export(i, 0)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 4)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG6_group_export(i, 0)))
			CASE "S"
				strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 4)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG6_group_export(i, 0)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 5)))
				strData = strData & Chr(11) & Trim(ConvSPChars(EG5_group_export(i, 0)))	
		END SELECT

		strData = strData & Chr(11) & Trim(ConvSPChars(EG9_group_export(i, 0)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 1)))	
		strData = strData & Chr(11) & Trim(ConvSPChars(EG7_group_export(i, 0)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 0)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 8)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 9)))
		strData = strData & Chr(11) & UniNumClientFormat(EG2_group_export(i, 10), ggQty.DecPoint ,0)	'Lot Size
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 12)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 13)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 14)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG8_group_export(i, 0)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 11)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG10_group_export(i, 0)))
		strData = strData & Chr(11) & LngMaxRow + i	+ 1
		strData = strData & Chr(11) & Chr(12)
	ELSE	
		StrNextKey = E1_q_inspection_request(i, 0)
	End If
Next
%>		    
<Script Language="vbscript">   
With parent
	.ggoSpread.Source = .vspdData 
	.ggoSpread.SSShowDataByClip "<%=strData%>"
	.vspdData.focus
		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> "" Then
		 .DbQuery
	Else
		<% ' Request값을 hidden Varialbe로 넘겨줌 %>
		.hItemCd = "<%=ConvSPChars(strItemCd)%>"
		.hInspReqNo = "<%=ConvSPChars(strInspReqNo)%>"
		.hBpCd = "<%=ConvSPChars(strBpCd)%>"
		.hCustCd = "<%=ConvSPChars(strCustCd)%>"
		.hWcCd = "<%=ConvSPChars(strWcCd)%>"
		.hFrInspDt = "<%=strFrInspDt%>"
		.hToInspDt = "<%=strToInspDt%>"
		.hStatusFlag = "<%=ConvSPChars(strStatusFlag)%>"
		.hDecision = "<%=ConvSPChars(strDecision)%>"
		.DbQueryOk
	End If		
End with
</Script>
<%
Set PQIG050 = Nothing
%>
