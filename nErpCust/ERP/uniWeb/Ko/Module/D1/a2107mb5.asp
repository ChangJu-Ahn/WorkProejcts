<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_ACCT_TRANS_TYPE
'*  3. Program ID        : a2107mb
'*  4. Program �̸�      : �а����� ��� 
'*  5. Program ����      : �а����� ��� ���� ���� ��ȸ 
'*  6. Comproxy ����Ʈ   : a2107ma
'*  7. ���� �ۼ������   : 2000/10/02
'*  8. ���� ���������   : 2000/10/02
'*  9. ���� �ۼ���       : ����ȯ 
'* 10. ���� �ۼ���       : Cho Ig Sung
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'*                         -2000/10/02 : ..........
'**********************************************************************************************

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call HideStatusWnd

'On Error Resume Next
Response.Write "mb5"
Response.End

Dim pAb0019											'��ȸ�� ComProxy Dll ��� ���� 

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'Dim StrNextKeyCtr5		' ���� �� 
Dim StrNextKeyThree_CtrlCd
'Dim lgStrPrevKeyCtr5	' ���� �� 
Dim lgStrPrevKeyThree_CtrlCd

'Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          
'Dim strItemSeq
Dim AcctNm

'@Var_Declare

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

'On Error Resume Next

Select Case strMode

	Case CStr(UID_M0001)								'��: ���� ��ȸ/Prev/Next ��û�� ���� 

		lgStrPrevKeyThree_CtrlCd = Request("lgStrPrevKeyThree_CtrlCd")
'		lgStrPrevKeyCtr5 = Request("lgStrPrevKeyCtr5")

	    Set pAb0019 = Server.CreateObject("Ab0019.ALookupAcctSvr")
	    '-----------------------------------------
	    'Com action result check area(OS,internal)
	    '-----------------------------------------
	    If Err.Number <> 0 Then
			Set pAb0019 = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

	    '-----------------------------------------
	    'Data manipulate  area(import view match)
	    '-----------------------------------------
	    pAb0019.ImportAAcctAcctCd = Trim(Request("txtAcctCd"))   
		
	    pAb0019.CommandSent = "lookupac"
	    
	    pAb0019.ServerLocation = ggServerIP

	    '-----------------------------------------
	    'Com Action Area
	    '-----------------------------------------
	    pAb0019.Comcfg = gConnectionString
	    pAb0019.Execute
	    
	    AcctNm = pAb0019.ExportAAcctAcctNm

	    '-----------------------------------------
	    'Com action result check area(OS,internal)
	    '-----------------------------------------
	    If Err.Number <> 0 Then
			Set pAb0019 = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

		'-----------------------------------------
		'Com action result check area(DB,internal)
		'-----------------------------------------
		If Not (pAb0019.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(pAb0019.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)

			Set pAb0019 = Nothing												'��: ComProxy Unload
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If    
	    
		'LngMaxRow = Request("txtMaxRows5")										'Save previous Maxrow                                                
	   	GroupCount = pAb0019.ExportGroupCount

		' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
		' ����/���� �� ���, ���ƿ� �°� ó���� 
		'If pAb0019.ExportPIndReqIndReqmtNo(GroupCount) = pAb0019.ExportNextPMPSRequirementIndReqmtNo Then
		'	StrNextKeyCtr5 = ""
		'Else
		'	StrNextKeyCtr5 = pAb0019.ExportNextPMPSRequirementIndReqmtNo
		'End If
%>

<Script Language=vbscript>
		Dim lngMaxRows       
		Dim strData
		Dim lRows
		Dim tmpDrCrFg	
		Dim CtrlCtrlCnt
		
		With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		
			lngMaxRows = .frm1.vspdData3.MaxRows
			.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=GroupCount%>)
			CtrlCtrlCnt = 0
<%      
			For LngRow = 1 To GroupCount
%>
				CtrlCtrlCnt = CtrlCtrlCnt + 1
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlCd(LngRow))%>"			' �����׸��ڵ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlNm(LngRow))%>"			' �����׸�� 
				strData = strData & Chr(11) & .frm1.txtTransType.value								' �ŷ����� 
				strData = strData & Chr(11) & .frm1.txtJnlCd.value									' �ŷ��׸� 
				strData = strData & Chr(11) & "<%=Request("txtFormSeq")%>"
				strData = strData & Chr(11) & CtrlCtrlCnt
				strData = strData & Chr(11) & .frm1.txtDrCrFgCd.value									' ���뱸�� 
				strData = strData & Chr(11) & .frm1.txtAcctCD.value									' �������� 
'				strData = strData & Chr(11) & .frm1.txtBizAreaCd.value									' �������� 
				strData = strData & Chr(11) & ""													' ���̺��ID
				strData = strData & Chr(11) & ""													' �÷���ID
				strData = strData & Chr(11) & ""													' �ڷ������ڵ� 
				strData = strData & Chr(11) & ""													' �ڷ������� 
				strData = strData & Chr(11) & ""													' Key�÷�ID1
				strData = strData & Chr(11) & ""													' �ڷ������ڵ�1
				strData = strData & Chr(11) & ""													' �ڷ�������1
				strData = strData & Chr(11) & ""													' Key�÷�ID2
				strData = strData & Chr(11) & ""													' �ڷ������ڵ�2
				strData = strData & Chr(11) & ""													' �ڷ�������2
				strData = strData & Chr(11) & ""													' Key�÷�ID3
				strData = strData & Chr(11) & ""													' �ڷ������ڵ�3
				strData = strData & Chr(11) & ""													' �ڷ�������3
				strData = strData & Chr(11) & ""													' Key�÷�ID4
				strData = strData & Chr(11) & ""													' �ڷ������ڵ�4
				strData = strData & Chr(11) & ""													' �ڷ�������4
				strData = strData & Chr(11) & ""													' Key�÷�ID5
				strData = strData & Chr(11) & ""													' �ڷ������ڵ�5
				strData = strData & Chr(11) & ""													' �ڷ�������5
				strData = strData & Chr(11) & <%=LngRow%>
				strData = strData & Chr(11) & Chr(12)

				.ggoSpread.Source = .frm1.vspdData2
				
				.frm1.vspdData3.Row = lngMaxRows + Clng(<%=LngRow%>)
				.frm1.vspdData3.Col = 0:	.frm1.vspdData3.Text = .ggoSpread.InsertFlag
				.frm1.vspdData3.Col = 1:	.frm1.vspdData3.Text = "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlCd(LngRow))%>"
				.frm1.vspdData3.Col = 2:	.frm1.vspdData3.Text = "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlNm(LngRow))%>"
				.frm1.vspdData3.Col = 3:	.frm1.vspdData3.Text = .frm1.txtTransType.value
				.frm1.vspdData3.Col = 4:	.frm1.vspdData3.Text = .frm1.txtJnlCd.value
				.frm1.vspdData3.Col = 5:	.frm1.vspdData3.Text = "<%=Request("txtFormSeq")%>"
				.frm1.vspdData3.Col = 6:	.frm1.vspdData3.Text = CtrlCtrlCnt
				.frm1.vspdData3.Col = 7:	.frm1.vspdData3.Text = .frm1.txtDrCrFgCd.value
				.frm1.vspdData3.Col = 8:	.frm1.vspdData3.Text = .frm1.txtAcctCD.value
'				.frm1.vspdData3.Col = 7:	.frm1.vspdData3.Text = .frm1.txtBizAreaCd.value
				.frm1.vspdData3.Col = 9:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 10:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 11:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 12:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 13:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 14:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 15:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 16:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 17:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 18:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 19:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 20:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 21:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 22:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 23:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 24:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 25:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 26:	.frm1.vspdData3.Text = ""
				.frm1.vspdData3.Col = 27:	.frm1.vspdData3.Text = ""
<%
		    Next
%>
		    .frm1.vspdData2.MaxRows = 0
			.ggoSpread.Source = .frm1.vspdData2
			.ggoSpread.SSShowData strData

			For lRows = 1 To .frm1.vspdData2.MaxRows
			    .frm1.vspdData2.Row = lRows
				.frm1.vspdData2.Col = 0
			    .frm1.vspdData2.Text = .ggoSpread.InsertFlag
			Next
				
'		.frm1.vspdData.Row = .frm1.vspdData.ActiveRow
'		.frm1.vspdData.Col = 8  '�����ڵ�� 
'		.frm1.vspdData.Text = "<%=ConvSPChars(AcctNm)%>"

			.DbQuery_ThreeOk
			
		End With
</Script>	
<% 
	    Set pAb0019 = Nothing
End Select
%>
</Script>
