<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_ACCT_TRANS_TYPE
'*  3. Program ID        : a2107mb
'*  4. Program 이름      : 분개형태 등록 
'*  5. Program 설명      : 분개형태 등록 수정 삭제 조회 
'*  6. Comproxy 리스트   : a2107ma
'*  7. 최초 작성년월일   : 2000/10/02
'*  8. 최종 수정년월일   : 2000/10/02
'*  9. 최초 작성자       : 김종환 
'* 10. 최종 작성자       : Cho Ig Sung
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'*                         -2000/10/02 : ..........
'**********************************************************************************************

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

'On Error Resume Next
Response.Write "mb5"
Response.End

Dim pAb0019											'조회용 ComProxy Dll 사용 변수 

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'Dim StrNextKeyCtr5		' 다음 값 
Dim StrNextKeyThree_CtrlCd
'Dim lgStrPrevKeyCtr5	' 이전 값 
Dim lgStrPrevKeyThree_CtrlCd

'Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          
'Dim strItemSeq
Dim AcctNm

'@Var_Declare

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 

'On Error Resume Next

Select Case strMode

	Case CStr(UID_M0001)								'☜: 현재 조회/Prev/Next 요청을 받음 

		lgStrPrevKeyThree_CtrlCd = Request("lgStrPrevKeyThree_CtrlCd")
'		lgStrPrevKeyCtr5 = Request("lgStrPrevKeyCtr5")

	    Set pAb0019 = Server.CreateObject("Ab0019.ALookupAcctSvr")
	    '-----------------------------------------
	    'Com action result check area(OS,internal)
	    '-----------------------------------------
	    If Err.Number <> 0 Then
			Set pAb0019 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
			Response.End														'☜: 비지니스 로직 처리를 종료함 
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
			Set pAb0019 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.Number & Err.description, vbCritical, I_MKSCRIPT)
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If

		'-----------------------------------------
		'Com action result check area(DB,internal)
		'-----------------------------------------
		If Not (pAb0019.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(pAb0019.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)

			Set pAb0019 = Nothing												'☜: ComProxy Unload
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If    
	    
		'LngMaxRow = Request("txtMaxRows5")										'Save previous Maxrow                                                
	   	GroupCount = pAb0019.ExportGroupCount

		' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
		' 문자/숫자 일 경우, 문맥에 맞게 처리함 
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
		
		With parent																	'☜: 화면 처리 ASP 를 지칭함 
		
			lngMaxRows = .frm1.vspdData3.MaxRows
			.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=GroupCount%>)
			CtrlCtrlCnt = 0
<%      
			For LngRow = 1 To GroupCount
%>
				CtrlCtrlCnt = CtrlCtrlCnt + 1
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlCd(LngRow))%>"			' 관리항목코드 
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0019.ExportACtrlItemCtrlNm(LngRow))%>"			' 관리항목명 
				strData = strData & Chr(11) & .frm1.txtTransType.value								' 거래유형 
				strData = strData & Chr(11) & .frm1.txtJnlCd.value									' 거래항목 
				strData = strData & Chr(11) & "<%=Request("txtFormSeq")%>"
				strData = strData & Chr(11) & CtrlCtrlCnt
				strData = strData & Chr(11) & .frm1.txtDrCrFgCd.value									' 차대구분 
				strData = strData & Chr(11) & .frm1.txtAcctCD.value									' 계정과목 
'				strData = strData & Chr(11) & .frm1.txtBizAreaCd.value									' 계정과목 
				strData = strData & Chr(11) & ""													' 테이블명ID
				strData = strData & Chr(11) & ""													' 컬럼명ID
				strData = strData & Chr(11) & ""													' 자료유형코드 
				strData = strData & Chr(11) & ""													' 자료유형명 
				strData = strData & Chr(11) & ""													' Key컬럼ID1
				strData = strData & Chr(11) & ""													' 자료유형코드1
				strData = strData & Chr(11) & ""													' 자료유형명1
				strData = strData & Chr(11) & ""													' Key컬럼ID2
				strData = strData & Chr(11) & ""													' 자료유형코드2
				strData = strData & Chr(11) & ""													' 자료유형명2
				strData = strData & Chr(11) & ""													' Key컬럼ID3
				strData = strData & Chr(11) & ""													' 자료유형코드3
				strData = strData & Chr(11) & ""													' 자료유형명3
				strData = strData & Chr(11) & ""													' Key컬럼ID4
				strData = strData & Chr(11) & ""													' 자료유형코드4
				strData = strData & Chr(11) & ""													' 자료유형명4
				strData = strData & Chr(11) & ""													' Key컬럼ID5
				strData = strData & Chr(11) & ""													' 자료유형코드5
				strData = strData & Chr(11) & ""													' 자료유형명5
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
'		.frm1.vspdData.Col = 8  '계정코드명 
'		.frm1.vspdData.Text = "<%=ConvSPChars(AcctNm)%>"

			.DbQuery_ThreeOk
			
		End With
</Script>	
<% 
	    Set pAb0019 = Nothing
End Select
%>
</Script>
