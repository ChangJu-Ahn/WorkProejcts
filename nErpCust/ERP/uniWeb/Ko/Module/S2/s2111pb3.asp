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
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/29 : Coding Start												*
'********************************************************************************************************

Response.Expires = -1															'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True															'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Call HideStatusWnd

Dim strMode																		'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Const lsPLANNUM  = "PLANNUM"	'��ȹ���� 

strMode = Request("txtMode")													'�� : ���� ���¸� ���� 

Select Case strMode
Case CStr(lsPLANNUM)			'��ȹ���� ��ȸ 

    Dim pS21113
    Dim LngRow
    Dim GroupCount
    
	Set pS21113 = Server.CreateObject("S21113.S21113ListPlanSeqSvr")
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
    If Err.Number <> 0 Then
		Set pS21113 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
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

			<% '��ȹ���� %>
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
	Response.End																'��: Process End

End Select
%>
