<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : 본지점관리 
'*  3. Program ID        : a8107mb4
'*  4. Program 이름      : 본지점전표승인 
'*  5. Program 설명      : 지점전표 List
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2001/01/19
'*  8. 최종 수정년월일   : 
'*  9. 최초 작성자       : hersheys
'* 10. 최종 작성자       : 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'*                         -2000/10/02 : ..........
'**********************************************************************************************

Response.Expires = -1		'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True		'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next			' ☜: 

Dim Ab010m4						' 입력/수정용 ComProxy Dll 사용 변수 
Dim A53018BR						' 조회용 ComProxy Dll 사용 변수 

Dim strMode						'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim LngMaxRow			' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

Dim StrErr

Dim StrDebug

strMode = Request("txtMode")	'☜ : 현재 상태를 받음 

GetGlobalVar

Select Case strMode

	Case CStr(UID_M0001)			'☜: 현재 조회/Prev/Next 요청을 받음 

		Set A53018BR = Server.CreateObject("A53018BR.A53018BrListTempGlBrchSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set A53018BR = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
			Call HideStatusWnd
			Response.End														'☜: 비지니스 로직 처리를 종료함 
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
			Response.End																		'☜: Process End
		End If

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set A53018BR = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
			Call HideStatusWnd
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		'If Not (A53018BR.OperationStatusMessage = MSG_OK_STR) Then
		'	Call DisplayMsgBox(A53018BR.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		'	Set A53018BR = Nothing												'☜: ComProxy Unload
		'	Call HideStatusWnd
		'	Response.End														'☜: 비지니스 로직 처리를 종료함 
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
		'	Response.End														'☜: 비지니스 로직 처리를 종료함 
		'End If

		GroupCount = A53018BR.ExportGrpTempGlCount

		' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
		' 문자/숫자 일 경우, 문맥에 맞게 처리함 
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

		With parent									'☜: 화면 처리 ASP 를 지칭함 

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
   
'			<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
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
		Response.End																		'☜: Process End
    
	Case CStr(UID_M0002)					'☜: 저장 요청을 받음 
										
	    Err.Clear							'☜: Protect system from crashing

		If Request("txtMaxRows4") = "" Then
%>
	<Script Language=vbscript>
			Call DisplayMsgBox("700100", "X", "X", "X")
	</Script>
<%
			'Call ServerMesgBox("txtMaxRows 조건값이 비어있습니다!",vbInformation, I_MKSCRIPT)              
			Call HideStatusWnd
			Response.End 
		End If

		LngMaxRow = Request("txtMaxRows4")		'☜: 최대 업데이트된 갯수 
		
	    Set Ab010m4 = Server.CreateObject("Ab0101.AManageAcctSpclJnlCtrlSvr")

	    '-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If Err.Number <> 0 Then
			Set Ab010m4 = Nothing					'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)	'⊙:
			Call HideStatusWnd
			Response.End							'☜: 비지니스 로직 처리를 종료함 
		End If

		Dim arrVal, arrTemp							'☜: Spread Sheet 의 값을 받을 Array 변수 
		Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
		Dim	lGrpCnt									'☜: Group Count
		Dim strCode									'⊙: Lookup 용 리턴 변수 
		
		arrTemp = Split(Request("txtSpread4"), gRowSep)	' Spread Sheet 내용을 담고 있는 Element명 
		
	    Ab010m4.ServerLocation = ggServerIP
	    
	    lGrpCnt = 0
	    
	    For LngRow = 1 To LngMaxRow
	    
			lGrpCnt = lGrpCnt +1					'☜: Group Count
			arrVal = Split(arrTemp(LngRow-1), gColSep)
			strStatus = arrVal(0)					'☜: Row 의 상태 

			'Response.Write arrTemp(LngRow-1) & "<br>"
			
			Select Case strStatus

	            Case "C"							'☜: Create
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
			
			If lGrpCnt > 50 Or Cint(LngRow) = Cint(LngMaxRow) Then		' 50개를 Group으로, 나머지 일때 
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
				'	Set Ab010m4 = Nothing												'☜: ComProxy Unload
				'	Call HideStatusWnd
				'	Response.End														'☜: 비지니스 로직 처리를 종료함 
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
					Response.End														'☜: 비지니스 로직 처리를 종료함 
				End If

	            Ab010m4.Clear
	            lGrpCnt = 0
	            
	            'ggoSpread.SSDeleteFlag lStartRow, lEndRow	'뭐하는건지 모르겠음 
			End If
	        
	    Next

%>
	<Script Language=vbscript>
		Parent.DbSaveOk("<%=Request("txtTransType")%>")				'☜: 화면 처리 ASP 를 지칭함 
	</Script>
<%					
	    Set Ab010m4 = Nothing       '☜: Unload Comproxy
		Call HideStatusWnd
		Response.End																		'☜: Process End

End Select

%>
