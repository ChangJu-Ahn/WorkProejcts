
<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : 
'*  3. Program ID        : a6104mb1
'*  4. Program �̸�      : �ΰ���������ȸ 
'*  5. Program ����      : �ΰ���������ȸ 
'*  6. Comproxy ����Ʈ   : a6104mb1
'*  7. ���� �ۼ������   : 2000/04/22
'*  8. ���� ���������   : 2000/04/22
'*  9. ���� �ۼ���       : ����ȯ 
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'*
'**********************************************************************************************

Response.Expires = -1		'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True		'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next			' ��: 

Dim Ag0104						' ��ȸ�� ComProxy Dll ��� ���� 

Dim strMode						'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey			' ���� �� 
Dim lgStrPrevKey		' ���� �� 
Dim LngMaxRow			' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          

Dim StrErr
Dim StrDebug

strMode = Request("txtMode")	'�� : ���� ���¸� ���� 

GetGlobalVar

Select Case strMode

	Case CStr(UID_M0001)			'��: ���� ��ȸ/Prev/Next ��û�� ���� 

		lgStrPrevKey = Request("lgStrPrevKey")
	
		Set Ag0104 = Server.CreateObject("Ag0104.AQueryVatListSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set Ag0104 = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
			Call HideStatusWnd
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		Ag0104.ImportStartAVatIssuedDt = UNIConvDate(Request("txtIssueDT1"))
		Ag0104.ImportEndAVatIssuedDt   = UNIConvDate(Request("txtIssueDT2"))
		Ag0104.ImportAVatIoFg          = Request("cboIOFlag")
		Ag0104.ImportAVatVatType       = Request("cboVatType")
		Ag0104.ImportBBizPartnerBpCd   = UCase(Trim(Request("txtBPCd")))
		Ag0104.ImportBBizAreaBizAreaCd = UCase(Trim(Request("txtBizAreaCd")))
	    Ag0104.ImportIefSuppliedCount  = lgStrPrevKey
		Ag0104.CommandSent             = "QUERY"
		Ag0104.ServerLocation          = ggServerIP

		'-----------------------
		'Com Action Area
		'-----------------------
		Ag0104.ComCfg = gConnectionString
		Ag0104.Execute
		
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set Ag0104 = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
			Call HideStatusWnd
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (Ag0104.OperationStatusMessage = MSG_OK_STR) Then
		    Select Case Ag0104.OperationStatusMessage
		        Case MSG_DEADLOCK_STR
		            Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
		        Case MSG_DBERROR_STR
		            Call DisplayMsgBox2(Ag0104.ExportErrEabSqlCodeSqlcode, _
										Ag0104.ExportErrEabSqlCodeSeverity, _
										Ag0104.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
		        Case Else
		            Call DisplayMsgBox(Ag0104.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		    End Select
			Call HideStatusWnd
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

		GroupCount = Ag0104.ExportGroupCount

		' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
		' ����/���� �� ���, ���ƿ� �°� ó���� 
		If lgStrPrevKey >= Ag0104.ExportNextIefSuppliedCount Then
			StrNextKey = 0
		Else
			StrNextKey = Ag0104.ExportNextIefSuppliedCount
		End If

%>
<Script Language=vbscript>
		Dim LngMaxRow       
	    Dim LngRow
		Dim strData
		
	
		With parent									'��: ȭ�� ó�� ASP �� ��Ī�� 

			LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                

<%
			For LngRow = 1 To GroupCount
%>
				
				' UNIDateClientFormat(pDate)
				' UNINumClientFormat(pNum, iDecPoint, pDefault)
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(Ag0104.ExportAVatIssuedDt(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Ag0104.ExportAVatIoFg(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Ag0104.ExportAVatIoFg(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Ag0104.ExportBBizPartnerBpCd(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Ag0104.ExportBBizPartnerBpNm(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Ag0104.ExportAVatOwnRgstNo(LngRow))%>"
				strData = strData & Chr(11) & "<%=UNINumClientFormat(Ag0104.ExportAVatNetLocAmt(LngRow), ggAmtOfMoney.DecPoint, 0)%>"
				strData = strData & Chr(11) & "<%=UNINumClientFormat(Ag0104.ExportAVatVatLocAmt(LngRow), ggAmtOfMoney.DecPoint, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Ag0104.ExportAVatVatType(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Ag0104.ExportAVatVatType(LngRow))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
				strData = strData & Chr(11) & Chr(12)
<%      
		    Next
%>    
			.frm1.txtBizAreaCd.value = UCase(Trim("<%=ConvSPChars(Ag0104.ExportBBizAreaBizAreaCd)%>"))
			.frm1.txtBizAreaNm.value = "<%=ConvSPChars(Ag0104.ExportBBizAreaBizAreaNm)%>"
			.frm1.txtBpCd.value = UCase(Trim("<%=ConvSPChars(Ag0104.ExpBBizPartnerBpCd)%>"))
			.frm1.txtBpNm.value = "<%=ConvSPChars(Ag0104.ExpBBizPartnerBpNm)%>"

			.ggoSpread.Source = .frm1.vspdData 
			.ggoSpread.SSShowData strData

			' ���� 
			.frm1.txtCntSumO.value = "<%=UNINumClientFormat(Ag0104.ExportSumOutIefSuppliedCount, ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtAmtSumO.value = "<%=UNINumClientFormat(Ag0104.ExportSumOutAVatNetLocAmt,    ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtVatSumO.value = "<%=UNINumClientFormat(Ag0104.ExportSumOutAVatVatLocAmt,    ggAmtOfMoney.DecPoint, 0)%>"
			' ���� 
			.frm1.txtCntSumI.value = "<%=UNINumClientFormat(Ag0104.ExportSumInIefSuppliedCount, ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtAmtSumI.value = "<%=UNINumClientFormat(Ag0104.ExportSumInAVatNetLocAmt,    ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtVatSumI.value = "<%=UNINumClientFormat(Ag0104.ExportSumInAVatVatLocAmt,    ggAmtOfMoney.DecPoint, 0)%>"

			.lgStrPrevKey = "<%=StrNextKey%>"
			
			<% ' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ %>
			If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> "0" Then
				.DbQuery
			Else
				.frm1.hBPCd.value      = UCase(Trim("<%=Request("txtBPCd")%>"))
				.frm1.hIssueDT1.value  = "<%=Request("txtIssueDT1")%>"
				.frm1.hIssueDT2.value  = "<%=Request("txtIssueDT2")%>"
				.frm1.hIOFlag.value    = "<%=ConvSPChars(Request("cboIOFlag"))%>"
				.frm1.hVatType.value   = "<%=ConvSPChars(Request("cboVatType"))%>"
				.frm1.hBizAreaCd.value = UCase(Trim("<%=ConvSPChars(Request("txtBizAreaCd"))%>"))
				.DbQueryOk
			End If

		End With
</Script>	
<%
	    Set Ag0104 = Nothing

		Call HideStatusWnd

End Select

%>
