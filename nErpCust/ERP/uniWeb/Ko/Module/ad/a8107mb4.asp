<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : ���������� 
'*  3. Program ID        : a8107mb4
'*  4. Program �̸�      : ��������ǥ���� 
'*  5. Program ����      : ������ǥ List
'*  6. Comproxy ����Ʈ   : 
'*  7. ���� �ۼ������   : 2001/01/19
'*  8. ���� ���������   : 
'*  9. ���� �ۼ���       : hersheys
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'*                         -2000/10/02 : ..........
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

Dim Ab010m4						' �Է�/������ ComProxy Dll ��� ���� 
Dim A53018BR						' ��ȸ�� ComProxy Dll ��� ���� 

Dim strMode						'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim LngMaxRow			' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          

Dim StrErr

Dim StrDebug

strMode = Request("txtMode")	'�� : ���� ���¸� ���� 

GetGlobalVar

Select Case strMode

	Case CStr(UID_M0001)			'��: ���� ��ȸ/Prev/Next ��û�� ���� 

		Set A53018BR = Server.CreateObject("A53018BR.A53018BrListTempGlBrchSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set A53018BR = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
			Call HideStatusWnd
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		A53018BR.ImportBBizAreaBizAreaCd     = Trim(Request("txtBizAreaCd"))
		A53018BR.ImportConfFgATempGlHqBrchNo =  Trim(Request("txtHqBrchNo"))
		A53018BR.ImportConfFgATempGlTempGlNo =  Trim(Request("txtTempGLNo"))		
		A53018BR.ImportBAcctDeptOrgChangeId  = gChangeOrgID
		A53018BR.ServerLocation		         = ggServerIP

		'-----------------------
		'Com Action Area
		'-----------------------
		A53018BR.Comcfg = gConnectionString
		A53018BR.Execute		

		If A53018BR.ExportGrpTempGlCount = 0 Then
%>
	<Script Language=vbscript>
'			Parent.frm1.txtCtrlCnt.value = 0
'			Parent.frm1.hCtrlCnt.value = 0
	</Script>
<%
			Set A53018BR = Nothing
			Call HideStatusWnd
			Response.End																		'��: Process End
		End If

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set A53018BR = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
			Call HideStatusWnd
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		'If Not (A53018BR.OperationStatusMessage = MSG_OK_STR) Then
		'	Call DisplayMsgBox(A53018BR.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		'	Set A53018BR = Nothing												'��: ComProxy Unload
		'	Call HideStatusWnd
		'	Response.End														'��: �����Ͻ� ���� ó���� ������ 
		'End If    
		'If Not (A53018BR.OperationStatusMessage = MSG_OK_STR) Then
		'    Select Case A53018BR.OperationStatusMessage
		'        Case MSG_DEADLOCK_STR
		'            Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
		'        Case MSG_DBERROR_STR
		'            Call DisplayMsgBox2(A53018BR.ExportErrEabSqlCodeSqlcode, _
		'								A53018BR.ExportErrEabSqlCodeSeverity, _
		'								A53018BR.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
		'        Case Else
		'            Call DisplayMsgBox(A53018BR.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		'    End Select
		'	Call HideStatusWnd
		'	Response.End														'��: �����Ͻ� ���� ó���� ������ 
		'End If

		GroupCount = A53018BR.ExportGrpTempGlCount

		' ���� �κ�: Next Key���� ���� ����Ÿ(�׷���)�� ������ ���� ������ ���� ����Ÿ�� �����Ƿ� Ű ������ ������ ���� �ʱ�ȭ�� 
		' ����/���� �� ���, ���ƿ� �°� ó���� 
		'If A53018BR.ExportACtrlItemCtrlCd(GroupCount) = A53018BR.ExportNextACtrlItemCtrlCd Then
		'	StrNextKeyCtrl4 = ""
		'Else
		'	StrNextKeyCtrl4 = A53018BR.ExportNextACtrlItemCtrlCd
		'End If
	
%>
<Script Language=vbscript>
		Dim LngMaxRow       
	    Dim LngRow
		Dim strData

		With parent									'��: ȭ�� ó�� ASP �� ��Ī�� 

			LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow                                                
<%
			For LngRow = 1 To GroupCount
%>
				strData = strData & Chr(11) & "<%=A53018BR.ExportItemATempGlConfFg(LngRow)%>"
				strData = strData & Chr(11) & " " 'Conf_Nm
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(A53018BR.ExportItemATempGlTempGlDt(LngRow))%>"
				strData = strData & Chr(11) & "<%=A53018BR.ExportItemATempGlTempGlNo(LngRow)%>"
				strData = strData & Chr(11) & "<%=A53018BR.ExportItemBAcctDeptDeptNm(LngRow)%>"
				strData = strData & Chr(11) & " " 'Currency
				strData = strData & Chr(11) & "<%=A53018BR.ExportItemATempGlDrAmt(LngRow)%>"
				strData = strData & Chr(11) & "<%=A53018BR.ExportItemATempGlDrLocAmt(LngRow)%>"
				If "<%=A53018BR.ExportItemAGlGlDt(LngRow)%>" < "1900-01-01" Then
					strData = strData & Chr(11) & ""
				Else
					strData = strData & Chr(11) & "<%=UNIDateClientFormat(A53018BR.ExportItemAGlGlDt(LngRow))%>"
				End If
				strData = strData & Chr(11) & "<%=A53018BR.ExportItemAGlGlNo(LngRow)%>"
				strData = strData & Chr(11) & "<%=A53018BR.ExportItemATempGlGlInputType(LngRow)%>"
				strData = strData & Chr(11) & " " 'InputTypeNm
				strData = strData & Chr(11) & "<%=A53018BR.ExportItemATempGlHqBrchNo(LngRow)%>"
				strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
				strData = strData & Chr(11) & Chr(12)
<%
		    Next
%>
			.ggoSpread.Source = .frm1.vspdData2
			.ggoSpread.SSShowData strData

'			.lgStrPrevKeyCtrl4 = "<%=StrNextKeyCtrl4%>"
   
'			<% ' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ %>
'			If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKeyCtrl4 <> "" Then
'				.DbQuery2
'			Else
'				.frm1.hFormCnt.value = CtrlCtrlCnt
'				.frm1.hTransType.value = "<%=Request("txtTransType")%>"
'				.frm1.hJnlCd.value = "<%=Request("txtJnlCd")%>"
'				.frm1.hDrCrFgCd.value = "<%=Request("txtDrCrFgCd")%>"
'				.frm1.hAcctCd.value = "<%=Request("txtAcctCd")%>"
'				.frm1.hBizAreaCd.value = "<%=Request("txtBizAreaCd")%>"
'				.frm1.hCtrlCD.value = "<%=Request("txtCtrlCD")%>"
				.DbQueryOk2
'			End If

		End With
</Script>	
<%
	    Set A53018BR = Nothing
		Call HideStatusWnd
		Response.End																		'��: Process End
    
	Case CStr(UID_M0002)					'��: ���� ��û�� ���� 
										
	    Err.Clear							'��: Protect system from crashing

		If Request("txtMaxRows4") = "" Then
%>
	<Script Language=vbscript>
			Call DisplayMsgBox("700100", "X", "X", "X")
	</Script>
<%
			'Call ServerMesgBox("txtMaxRows ���ǰ��� ����ֽ��ϴ�!",vbInformation, I_MKSCRIPT)              
			Call HideStatusWnd
			Response.End 
		End If

		LngMaxRow = Request("txtMaxRows4")		'��: �ִ� ������Ʈ�� ���� 
		
	    Set Ab010m4 = Server.CreateObject("Ab0101.AManageAcctSpclJnlCtrlSvr")

	    '-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If Err.Number <> 0 Then
			Set Ab010m4 = Nothing					'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)	'��:
			Call HideStatusWnd
			Response.End							'��: �����Ͻ� ���� ó���� ������ 
		End If

		Dim arrVal, arrTemp							'��: Spread Sheet �� ���� ���� Array ���� 
		Dim strStatus								'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
		Dim	lGrpCnt									'��: Group Count
		Dim strCode									'��: Lookup �� ���� ���� 
		
		arrTemp = Split(Request("txtSpread4"), gRowSep)	' Spread Sheet ������ ��� �ִ� Element�� 
		
	    Ab010m4.ServerLocation = ggServerIP
	    
	    lGrpCnt = 0
	    
	    For LngRow = 1 To LngMaxRow
	    
			lGrpCnt = lGrpCnt +1					'��: Group Count
			arrVal = Split(arrTemp(LngRow-1), gColSep)
			strStatus = arrVal(0)					'��: Row �� ���� 

			'Response.Write arrTemp(LngRow-1) & "<br>"
			
			Select Case strStatus

	            Case "C"							'��: Create
	                Ab010m4.ImportIefSuppliedSelectChar(lGrpCnt) = "C"
	                Ab010m4.ImportAAcctTransTypeTransType(lGrpCnt) = arrVal(2)
	                Ab010m4.ImportAJnlItemJnlCd(lGrpCnt) =  Trim(arrVal(3))
	                Ab010m4.ImportAJnlFormDrCrFg(lGrpCnt) = arrVal(4)
	                Ab010m4.ImportAAcctAcctCd(lGrpCnt) =  Trim(arrVal(5))
	                Ab010m4.ImportBBizAreaBizAreaCd(lGrpCnt) =  Trim(arrVal(6))
	                Ab010m4.ImportACtrlItemCtrlCd(lGrpCnt) =  Trim(arrVal(7))
	                Ab010m4.ImportASpclJnlCtrlAssnTblId(lGrpCnt) = arrVal(8)
	                Ab010m4.ImportASpclJnlCtrlAssnDataColmId(lGrpCnt) = arrVal(9)
	                Ab010m4.ImportASpclJnlCtrlAssnDataType(lGrpCnt) = arrVal(10)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyColmId1(lGrpCnt) = arrVal(11)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyDataType1(lGrpCnt) = arrVal(12)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyColmId2(lGrpCnt) = arrVal(13)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyDataType2(lGrpCnt) = arrVal(14)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyColmId3(lGrpCnt) = arrVal(15)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyDataType3(lGrpCnt) = arrVal(16)
	                Ab010m4.ImportASpclJnlCtrlAssnInsrtUserId(lGrpCnt) = gUsrID
				Case "U"
	                Ab010m4.ImportIefSuppliedSelectChar(lGrpCnt) = "U"
	                Ab010m4.ImportAAcctTransTypeTransType(lGrpCnt) = arrVal(2)
	                Ab010m4.ImportAJnlItemJnlCd(lGrpCnt) =  Trim(arrVal(3))
	                Ab010m4.ImportAJnlFormDrCrFg(lGrpCnt) = arrVal(4)
	                Ab010m4.ImportAAcctAcctCd(lGrpCnt) =  Trim(arrVal(5))
	                Ab010m4.ImportBBizAreaBizAreaCd(lGrpCnt) =  Trim(arrVal(6))
	                Ab010m4.ImportACtrlItemCtrlCd(lGrpCnt) =  Trim(arrVal(7))
	                Ab010m4.ImportASpclJnlCtrlAssnTblId(lGrpCnt) = arrVal(8)
	                Ab010m4.ImportASpclJnlCtrlAssnDataColmId(lGrpCnt) = arrVal(9)
	                Ab010m4.ImportASpclJnlCtrlAssnDataType(lGrpCnt) = arrVal(10)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyColmId1(lGrpCnt) = arrVal(11)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyDataType1(lGrpCnt) = arrVal(12)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyColmId2(lGrpCnt) = arrVal(13)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyDataType2(lGrpCnt) = arrVal(14)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyColmId3(lGrpCnt) = arrVal(15)
	                Ab010m4.ImportASpclJnlCtrlAssnKeyDataType3(lGrpCnt) = arrVal(16)
	                Ab010m4.ImportASpclJnlCtrlAssnUpdtUserId(lGrpCnt) = gUsrID
	            Case "D"
	                Ab010m4.ImportIefSuppliedSelectChar(lGrpCnt) = "D"
	                Ab010m4.ImportAAcctTransTypeTransType(lGrpCnt) = arrVal(2)
	                Ab010m4.ImportAJnlItemJnlCd(lGrpCnt) =  Trim(arrVal(3))
	                Ab010m4.ImportAJnlFormDrCrFg(lGrpCnt) = arrVal(4)
	                Ab010m4.ImportAAcctAcctCd(lGrpCnt) =  Trim(arrVal(5))
	                Ab010m4.ImportBBizAreaBizAreaCd(lGrpCnt) =  Trim(arrVal(6))
	                Ab010m4.ImportACtrlItemCtrlCd(lGrpCnt) =  Trim(arrVal(7))
	        End Select
			
			If lGrpCnt > 50 Or Cint(LngRow) = Cint(LngMaxRow) Then		' 50���� Group����, ������ �϶� 
	            Ab010m4.Comcfg = gConnectionString
	            Ab010m4.Execute
	                   
	            '-----------------------
	            'Com action result check area(OS,internal)
	            '-----------------------
	            If Err.Number <> 0 Then
					Set Ab010m4 = Nothing
					Call ServerMesgBox(Err.Description, vbCritical, I_MKSCRIPT)
					Call HideStatusWnd
					Response.End 
	            End If

	            '-----------------------
	            'Com action result check area(DB,internal)
	            '-----------------------
				'If Not (Ab010m4.OperationStatusMessage = MSG_OK_STR) Then
				'	Call DisplayMsgBox(Ab010m4.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
				'	Set Ab010m4 = Nothing												'��: ComProxy Unload
				'	Call HideStatusWnd
				'	Response.End														'��: �����Ͻ� ���� ó���� ������ 
				'End If    
				If Not (Ab010m4.OperationStatusMessage = MSG_OK_STR) Then
				    Select Case Ab010m4.OperationStatusMessage
				        Case MSG_DEADLOCK_STR
				            Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
				        Case MSG_DBERROR_STR
				            Call DisplayMsgBox2(Ab010m4.ExportErrEabSqlCodeSqlcode, _
												Ab010m4.ExportErrEabSqlCodeSeverity, _
												Ab010m4.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
				        Case Else
				            Call DisplayMsgBox(Ab010m4.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
				    End Select
					Call HideStatusWnd
					Response.End														'��: �����Ͻ� ���� ó���� ������ 
				End If

	            Ab010m4.Clear
	            lGrpCnt = 0
	            
	            'ggoSpread.SSDeleteFlag lStartRow, lEndRow	'���ϴ°��� �𸣰��� 
			End If
	        
	    Next

%>
	<Script Language=vbscript>
		Parent.DbSaveOk("<%=Request("txtTransType")%>")				'��: ȭ�� ó�� ASP �� ��Ī�� 
	</Script>
<%					
	    Set Ab010m4 = Nothing       '��: Unload Comproxy
		Call HideStatusWnd
		Response.End																		'��: Process End

End Select

%>
