
<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : 
'*  3. Program ID        : a6104mb1
'*  4. Program 이름      : 부가세내역조회 
'*  5. Program 설명      : 부가세내역조회 
'*  6. Comproxy 리스트   : a6104mb1
'*  7. 최초 작성년월일   : 2000/04/22
'*  8. 최종 수정년월일   : 2000/04/22
'*  9. 최초 작성자       : 김종환 
'* 10. 최종 작성자       : 
'* 11. 전체 comment      :
'*
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

Dim Ag0104						' 조회용 ComProxy Dll 사용 변수 

Dim strMode						'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey			' 다음 값 
Dim lgStrPrevKey		' 이전 값 
Dim LngMaxRow			' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

Dim StrErr
Dim StrDebug

strMode = Request("txtMode")	'☜ : 현재 상태를 받음 

GetGlobalVar

Select Case strMode

	Case CStr(UID_M0001)			'☜: 현재 조회/Prev/Next 요청을 받음 

		lgStrPrevKey = Request("lgStrPrevKey")
	
		Set Ag0104 = Server.CreateObject("Ag0104.AQueryVatListSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set Ag0104 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
			Call HideStatusWnd
			Response.End														'☜: 비지니스 로직 처리를 종료함 
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
			Set Ag0104 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
			Call HideStatusWnd
			Response.End														'☜: 비지니스 로직 처리를 종료함 
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
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If

		GroupCount = Ag0104.ExportGroupCount

		' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
		' 문자/숫자 일 경우, 문맥에 맞게 처리함 
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
		
	
		With parent									'☜: 화면 처리 ASP 를 지칭함 

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

			' 매출 
			.frm1.txtCntSumO.value = "<%=UNINumClientFormat(Ag0104.ExportSumOutIefSuppliedCount, ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtAmtSumO.value = "<%=UNINumClientFormat(Ag0104.ExportSumOutAVatNetLocAmt,    ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtVatSumO.value = "<%=UNINumClientFormat(Ag0104.ExportSumOutAVatVatLocAmt,    ggAmtOfMoney.DecPoint, 0)%>"
			' 매입 
			.frm1.txtCntSumI.value = "<%=UNINumClientFormat(Ag0104.ExportSumInIefSuppliedCount, ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtAmtSumI.value = "<%=UNINumClientFormat(Ag0104.ExportSumInAVatNetLocAmt,    ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtVatSumI.value = "<%=UNINumClientFormat(Ag0104.ExportSumInAVatVatLocAmt,    ggAmtOfMoney.DecPoint, 0)%>"

			.lgStrPrevKey = "<%=StrNextKey%>"
			
			<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
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
