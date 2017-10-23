<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_RECEIPT
'*  3. Program ID        : f5501mb1
'*  4. Program 이름      : 카드정보 등록 
'*  5. Program 설명      : 카드정보 등록 수정 삭제 조회 
'*  6. Comproxy 리스트   : f5501mb1
'*  7. 최초 작성년월일   : 2002/03/06
'*  8. 최종 수정년월일   : 
'*  9. 최초 작성자       : CHO IG SUNG
'* 10. 최종 작성자       : 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************



Call HideStatusWnd

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next			' ☜: 


Call LoadBasisGlobalInf()    
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")    

Dim Fn0028						' 조회용 ComProxy Dll 사용 변수 

Dim strMode						'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

Dim StrErr

strMode = Request("txtMode")	'☜ : 현재 상태를 받음 

GetGlobalVar

Select Case strMode

	Case CStr(UID_M0001)			'☜: 현재 조회/Prev/Next 요청을 받음 

		lgStrPrevKey = Request("lgStrPrevKey")
	
		Set Fn0028 = Server.CreateObject("Fn0021.FListNoteDtlSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set Fn0028 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
			Call HideStatusWnd
			Response.End														'☜: 비지니스 로직 처리를 종료함 
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
			Response.End																		'☜: Process End
		
		End If
		
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set Fn0028 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
			Call HideStatusWnd
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (Fn0028.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(Fn0028.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
			Set Fn0028 = Nothing												'☜: ComProxy Unload
			Call HideStatusWnd
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If    

		GroupCount = Fn0028.ExportGroupCount

		' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
		' 문자/숫자 일 경우, 문맥에 맞게 처리함 
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
	
		With parent									'☜: 화면 처리 ASP 를 지칭함 

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
   
			<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
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
		Response.End																		'☜: Process End
    
End Select

%>
