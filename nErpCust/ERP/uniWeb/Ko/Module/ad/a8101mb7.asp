<%
'=======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : a53017
'*  4. Program Name         :
'*  5. Program Desc         : ������ �����׸� ����Ÿ ��ȸ 
'*  6. Modified date(First) : 2000/10/9
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 

Dim pA53017											'��ȸ�� ComProxy Dll ��� ���� 

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          
Dim strItemSeq
Dim AcctNm

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

On Error Resume Next

'Call SvrMsgBox("Condition ->" & Request("txtAcctCd") & " : " & Request("txtItemSeq") , vbInformation, I_MKSCRIPT)

Select Case strMode

	Case CStr(UID_M0001)								'��: ���� ��ȸ/Prev/Next ��û�� ���� 

	lgStrPrevKey = Request("lgStrPrevKey")
	strItemSeq   = Request("txtItemSeq")
	
    Set pA53017 = Server.CreateObject("A53017.ALookupTempGlDtlSvr")
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pA53017 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
    pA53017.InTempGlATempGlTempGlNo    = Request("txtTempGlNo")
    pA53017.InTemSeqATempGlItemItemSeq = Request("txtItemSeq")
    pA53017.ServerLocation             = ggServerIP

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
	pA53017.ComCfg = gConnectionString
    pA53017.Execute

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		Set pA53017 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	'-----------------------------------------
	'Com action result check area(DB,internal)
	'-----------------------------------------
	If Not (pA53017.OperationStatusMessage = MSG_OK_STR) Then
		Call DisplayMsgBox(pA53017.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		Set pA53017 = Nothing												'��: ComProxy Unload
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If    
    
	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
   	GroupCount = pA53017.OutGrpTempGlDtlCount

	' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
	' ����/���� �� ���, ���ƿ� �°� ó���� 
'	If pA53017.ExportPIndReqIndReqmtNo(GroupCount) = pA53017.ExportNextPMPSRequirementIndReqmtNo Then
'		StrNextKey = ""
'	Else
'		StrNextKey = pA53017.ExportNextPMPSRequirementIndReqmtNo
'    End If
%>

<Script Language=vbscript>
    Dim lngMaxRows       
    Dim strData
    Dim lRows
    Dim tmpDrCrFg	
	
	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		
	lngMaxRows = .frm1.vspdData3.MaxRows
	.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=GroupCount%>)
<%      
	For LngRow = 1 To GroupCount
%>
<%'@ExportView - ����� %>
        strData = strData & Chr(11) & "<%=pA53017.OutGrpATempGlDtlDtlSeq(LngRow)%>"				'1   
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemCtrlCd(LngRow))%>"              '2 
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemCtrlNm(LngRow))%>"              '3
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpATempGlDtlCtrlVal(LngRow))%>"                                                        '4  
        strData = strData & Chr(11) & ""        						'5
        
		If "<%=ConvSPChars(pA53017.OutGrpACtrlItemTblId(LngRow))%>" = "" And "<%=ConvSPChars(pA53017.OutGrpACtrlItemColmDataType(LngRow))%>" = "D" then
	        strData = strData & Chr(11) & "(Format : YYYY-MM-DD)"                               '6
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpEabACtrlValRtnCtrlValC(LngRow))%>" '6
		End If

        strData = strData & Chr(11) & "<%=strItemSeq%>"											'7	
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemTblId(LngRow))%>" 				'8
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataColmId(LngRow))%>"			'9
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataColmNm(LngRow))%>"			'10
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemColmDataType(LngRow))%>"        '11
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataLen(LngRow))%>"        		'12
        strData = strData & Chr(11) & "<%=ConvSPChars(pA53017.OutGrpAAssignAcctInputFg(LngRow))%>"			'13
        strData = strData & Chr(11) & <%=LngRow%>												'14
        strData = strData & Chr(11) & Chr(12)

		.frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
        .frm1.vspdData3.Col = 1
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 2
        .frm1.vspdData3.Text = "<%=pA53017.OutGrpATempGlDtlDtlSeq(LngRow)%>"
        .frm1.vspdData3.Col = 3
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemCtrlCd(LngRow))%>"
        .frm1.vspdData3.Col = 4
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemCtrlNm(LngRow))%>"
        .frm1.vspdData3.Col = 5
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpATempGlDtlCtrlVal(LngRow))%>"
        .frm1.vspdData3.Col = 6 
        .frm1.vspdData3.Text =  ""
        .frm1.vspdData3.Col = 7
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpEabACtrlValRtnCtrlValC(LngRow))%>"
        .frm1.vspdData3.Col = 8
        .frm1.vspdData3.Text = "<%=strItemSeq%>"
        .frm1.vspdData3.Col = 9
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemTblId(LngRow))%>"
	
        If "<%=ConvSPChars(pA53017.OutGrpACtrlItemTblId(LngRow))%>" = "" And "<%=ConvSPChars(pA53017.OutGrpACtrlItemColmDataType(LngRow))%>" = "D" then
			.frm1.vspdData3.Col = 7
        	.frm1.vspdData3.Text = "(Format : YYYY-MM-DD)"
		End If

        .frm1.vspdData3.Col = 10
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataColmId(LngRow))%>"
        .frm1.vspdData3.Col = 11
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemDataColmNm(LngRow))%>"
        .frm1.vspdData3.Col = 12
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpACtrlItemColmDataType(LngRow))%>"
        .frm1.vspdData3.Col = 13
        .frm1.vspdData3.Text = "<%=pA53017.OutGrpACtrlItemDataLen(LngRow)%>"
        .frm1.vspdData3.Col = 14
        .frm1.vspdData3.Text = "<%=ConvSPChars(pA53017.OutGrpAAssignAcctInputFg(LngRow))%>"
<%      
    Next
%>    
    .frm1.vspdData2.MaxRows = 0
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowData strData



	.DbQueryOk2
		
	End With
</Script>	
<% 
    Set pA53017 = Nothing

End Select
%>
</Script>
