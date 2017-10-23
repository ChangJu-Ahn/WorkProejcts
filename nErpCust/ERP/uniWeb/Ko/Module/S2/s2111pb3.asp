<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : s2111pb3.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         :																			*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : Oh, Sang Eun																*
'* 10. Modifier (Last)      : sonbumyeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/29 : Coding Start												*
'********************************************************************************************************

Response.Expires = -1															'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True															'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call HideStatusWnd

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Const lsPLANNUM  = "PLANNUM"	'계획차수 

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 

Select Case strMode
Case CStr(lsPLANNUM)			'계획차수 조회 

    Dim pS21113
    Dim LngRow
    Dim GroupCount
    
	Set pS21113 = Server.CreateObject("S21113.S21113ListPlanSeqSvr")
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
    If Err.Number <> 0 Then
		Set pS21113 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	pS21113.ImpSItemGroupSalesPlanSalesOrg = Trim(Request("txtConSalesOrg"))
	pS21113.ImpSItemGroupSalesPlanSpYear = Trim(Request("txtConSpYear"))
	pS21113.ImpSItemGroupSalesPlanPlanFlag = Trim(Request("txtConPlanTypeCd"))
	pS21113.ImpSItemGroupSalesPlanExportFlag = Trim(Request("txtConDealTypeCd"))
	pS21113.ImpSItemGroupSalesPlanCur = Trim(Request("txtConCurr"))
	pS21113.ImpIefSuppliedSelectChar = Trim(Request("txtSelectChr"))

	pS21113.ServerLocation = ggServerIP
	pS21113.ComCfg = gConnectionString

	pS21113.Execute

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set pS21113 = Nothing
		Call ServerMesgBox(Err.Description, vbCritical, I_MKSCRIPT)              
		Response.End 
    End If

    '-----------------------
    'Com action result check area(DB,internal)
    '-----------------------
    If Not (pS21113.OperationStatusMessage = MSG_OK_STR) Then
		Call DisplayMsgBox(pS21113.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		Set pS21113 = Nothing
		Response.End 
    End If
    
	GroupCount = pS21113.ExpGrpCount

%>
<Script Language=vbscript>

    Dim LngMaxRow       
    Dim strData

	With parent

		LngMaxRow = .vspdData.MaxRows											'Save previous Maxrow

<%
		For LngRow = 1 To GroupCount
%>

			<% '계획차수 %>
			strData = strData & Chr(11) & "<%=pS21113.ExpItemSItemGroupSalesPlanPlanSeq(LngRow)%>"
			strData = strData & Chr(11) & Chr(12)
<%
		Next	
%>
		.ggoSpread.Source = .vspdData 
		.ggoSpread.SSShowDataByClip strData

		If Trim(.txtPlanSeq.value) <> "" Then
			.vspdData.Row = .txtPlanSeq.value
		Else
			.vspdData.Row = <%=GroupCount%>
		End If

		.vspdData.Col = .C_PlanSeq
		.vspdData.Action = 0
		.vspdData.Focus

	End With
</Script>
<%

	Set pS21113 = Nothing
	Response.End																'☜: Process End

End Select
%>
