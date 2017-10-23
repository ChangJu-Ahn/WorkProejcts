<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_RECEIPT
'*  3. Program ID        : f5501mb1
'*  4. Program �̸�      : ī������ ��� 
'*  5. Program ����      : ī������ ��� ���� ���� ��ȸ 
'*  6. Comproxy ����Ʈ   : f5501mb1
'*  7. ���� �ۼ������   : 2002/03/06
'*  8. ���� ���������   : 
'*  9. ���� �ۼ���       : CHO IG SUNG
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'**********************************************************************************************



Call HideStatusWnd

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next			' ��: 


Call LoadBasisGlobalInf()    
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")    

Dim Fn0028						' ��ȸ�� ComProxy Dll ��� ���� 

Dim strMode						'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          

Dim StrErr

strMode = Request("txtMode")	'�� : ���� ���¸� ���� 

GetGlobalVar

Select Case strMode

	Case CStr(UID_M0001)			'��: ���� ��ȸ/Prev/Next ��û�� ���� 

		lgStrPrevKey = Request("lgStrPrevKey")
	
		Set Fn0028 = Server.CreateObject("Fn0021.FListNoteDtlSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set Fn0028 = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
			Call HideStatusWnd
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		Fn0028.ImportFNoteNoteNo  = Trim(Request("txtNoteNo"))
		Fn0028.ImportFNoteItemSeq = lgStrPrevKey
		Fn0028.CommandSent		  = "QUERY"
		Fn0028.ServerLocation     = ggServerIP

		'-----------------------
		'Com Action Area
		'-----------------------
		Fn0028.ComCfg = gConnectionString
		'Fn0028.ComCfg = "TCP LETITBE 2050"
		Fn0028.Execute

		If Fn0028.ExportGroupCount = 0 Then
			.DbQueryOk
			Set Fn0028 = Nothing
			Call HideStatusWnd
			Response.End																		'��: Process End
		
		End If
		
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set Fn0028 = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
			Call HideStatusWnd
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (Fn0028.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(Fn0028.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
			Set Fn0028 = Nothing												'��: ComProxy Unload
			Call HideStatusWnd
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If    

		GroupCount = Fn0028.ExportGroupCount

		' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
		' ����/���� �� ���, ���ƿ� �°� ó���� 
		If Fn0028.ExportFNoteItemSeq(GroupCount) = Fn0028.ExportNextFNoteItemSeq Then
			StrNextKey = 0
		Else
			StrNextKey = Fn0028.ExportNextFNoteItemSeq
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
				strData = strData & Chr(11) & "<%=ConvSPChars(Fn0028.ExportFNoteItemNoteSts(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Fn0028.ExportFNoteItemNoteSts(LngRow))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(Fn0028.ExportFNoteItemStsDt(LngRow))%>"
				strData = strData & Chr(11) & "<%=UNINumClientFormat(Fn0028.ExportFNoteItemAmt(LngRow), ggAmtOfMoney.DecPoint, 0)%>"
				strData = strData & Chr(11) & "<%=UNINumClientFormat(Fn0028.ExportFNoteItemDcRate(LngRow), ggExchRate.DecPoint, 0)%>"
				strData = strData & Chr(11) & "<%=UNINumClientFormat(Fn0028.ExportFNoteItemIntAmt(LngRow), ggAmtOfMoney.DecPoint, 0)%>"
				strData = strData & Chr(11) & "<%=UNINumClientFormat(Fn0028.ExportFNoteItemChargeAmt(LngRow), ggAmtOfMoney.DecPoint, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Fn0028.ExportBBankBankCd(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Fn0028.ExportBBankBankNm(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Fn0028.ExportBBizPartnerBpCd(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Fn0028.ExportBBizPartnerBpNm(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Fn0028.ExportFNoteItemGlNo(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(Fn0028.ExportFNoteItemTempGlNo(LngRow))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
				strData = strData & Chr(11) & Chr(12)
<%      
		    Next
%>    
			.ggoSpread.Source = .frm1.vspdData 
			.ggoSpread.SSShowData strData

			.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
   
			<% ' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ %>
			If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> 0 Then
				.DbQuery2
			Else
				.DbQueryOk2
			End If

		End With
</Script>	
<%
	    Set Fn0028 = Nothing
		Call HideStatusWnd
		Response.End																		'��: Process End
    
End Select

%>
