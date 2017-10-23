<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_ACCT_TRANS_TYPE
'*  3. Program ID        : a2107mb
'*  4. Program 이름      : 분개형태 등록 
'*  5. Program 설명      : 분개형태 등록 수정 삭제 조회 
'*  6. Comproxy 리스트   : a2107mb
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

Response.Expires = -1		'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True		'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

'On Error Resume Next
Response.Write "mb2"
Response.End

Dim pAb0121						' 입력/수정용 ComProxy Dll 사용 변수 
Dim pAb0128						' 조회용 ComProxy Dll 사용 변수 

Dim strMode						'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey2			' 다음 값 
Dim StrNextKeyTwo_CtrlCd
Dim StrNextKeyAcct2
Dim StrNextKeyDrCrFg2
Dim StrNextKeyJnl2
Dim StrNextKeyCtrl2
'Dim lgStrPrevKey2		' 이전 값 
Dim lgStrPrevKeyTwo_CtrlCd
'Dim lgStrPrevKeyAcct2
'Dim lgStrPrevKeyDrCrFg2
'Dim lgStrPrevKeyJnl2
'Dim lgStrPrevKeyCtrl2
Dim LngMaxRow			' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

Dim StrErr

Dim StrDebug

strMode = Request("txtMode")	'☜ : 현재 상태를 받음 

Select Case strMode

	Case CStr(UID_M0001)			'☜: 현재 조회/Prev/Next 요청을 받음 
Response.End 
		lgStrPrevKeyTwo_CtrlCd = Request("lgStrPrevKeyTwo_CtrlCd")
'		lgStrPrevKeyAcct2 = Request("lgStrPrevKeyAcct2")
'		lgStrPrevKeyDrCrFg2 = Request("lgStrPrevKeyDrCrFg2")
'		lgStrPrevKeyJnl2 = Request("lgStrPrevKeyJnl2")
'		lgStrPrevKeyCtrl2 = Request("lgStrPrevKeyCtrl2")
	
		Set pAb0128 = Server.CreateObject("Ab0128.AListJnlCtrlAssnSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set pAb0128 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		pAb0128.ImportAAcctTransTypeTransType	= Request("txtTransType")
		pAb0128.ImportAJnlItemJnlCd				= Trim(Request("txtJnlCd"))
		pAb0128.ImportAJnlFormSeq				= Request("txtFormSeq")
		pAb0128.ImportAJnlFormDrCrFg			= Request("txtDrCrFgCd")
		pAb0128.ImportAAcctAcctCd				= Trim(Request("txtAcctCd"))
		pAb0128.ImportACtrlItemCtrlCd		= lgStrPrevKeyTwo_CtrlCd
'		pAb0128.ImportACtrlItemCtrlCd = Trim(Request("txtCtrlCD"))
'		pAb0128.CommandSent			= "QUERY"
		pAb0128.ServerLocation		= ggServerIP

		'-----------------------
		'Com Action Area
		'-----------------------
		pAb0128.Comcfg = gConnectionString
		pAb0128.Execute		

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------

		If Err.Number <> 0 Then
			Set pAb0128 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'⊙:
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (pAb0128.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(pAb0128.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)

			Set pAb0128 = Nothing												'☜: ComProxy Unload
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If    

		If pAb0128.ExportGroupCount = 0 Then
%>
	<Script Language=vbscript>
			Parent.frm1.txtCtrlCnt.value = 0
			Parent.frm1.hCtrlCnt.value = 0
	</Script>
<%
			Set pAb0128 = Nothing

			Response.End																		'☜: Process End
		End If

		GroupCount = pAb0128.ExportGroupCount

		' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
		' 문자/숫자 일 경우, 문맥에 맞게 처리함 
		If pAb0128.ExportItemACtrlItemCtrlCd(GroupCount) = pAb0128.ExportNextACtrlItemCtrlCd Then
			StrNextKeyTwo_CtrlCd = ""
		Else
			StrNextKeyTwo_CtrlCd = pAb0128.ExportNextACtrlItemCtrlCd
		End If
	
%>
<Script Language=vbscript>
		Dim LngMaxRow       
	    Dim LngRow
		Dim strData
		Dim CtrlCtrlCnt

		With parent									'☜: 화면 처리 ASP 를 지칭함 

			LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                
			CtrlCtrlCnt = 0

<%
			For LngRow = 1 To GroupCount
%>
				CtrlCtrlCnt = CtrlCtrlCnt + 1
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemACtrlItemCtrlCd(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemACtrlItemCtrlNm(LngRow))%>"
				strData = strData & Chr(11) & "<%=Request("txtTransType")%>"
				strData = strData & Chr(11) & "<%=Request("txtJnlCd")%>"
				strData = strData & Chr(11) & "<%=Request("txtFormSeq")%>"
				strData = strData & Chr(11) & CtrlCtrlCnt
				strData = strData & Chr(11) & "<%=Request("txtDrCrFgCd")%>"
				strData = strData & Chr(11) & "<%=Request("txtAcctCd")%>"
'				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemAJnlCtrlAssnTblId(LngRow))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemAJnlCtrlAssnDataColmId(LngRow))%>"
				strData = strData & Chr(11) & "<%=pAb0128.ExportItemAJnlCtrlAssnDataType(LngRow)%>"			' Code
				strData = strData & Chr(11) & ""														' Name
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemAJnlCtrlAssnKeyColmId1(LngRow))%>"
				strData = strData & Chr(11) & "<%=pAb0128.ExportItemAJnlCtrlAssnKeyDataType1(LngRow)%>"		' Code
				strData = strData & Chr(11) & ""														' Name
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemAJnlCtrlAssnKeyColmId2(LngRow))%>"
				strData = strData & Chr(11) & "<%=pAb0128.ExportItemAJnlCtrlAssnKeyDataType2(LngRow)%>"		' Code
				strData = strData & Chr(11) & ""														' Name
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemAJnlCtrlAssnKeyColmId3(LngRow))%>"
				strData = strData & Chr(11) & "<%=pAb0128.ExportItemAJnlCtrlAssnKeyDataType3(LngRow)%>"		' Code
				strData = strData & Chr(11) & ""														' Name
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemAJnlCtrlAssnKeyColmId4(LngRow))%>"
				strData = strData & Chr(11) & "<%=pAb0128.ExportItemAJnlCtrlAssnKeyDataType4(LngRow)%>"		' Code
				strData = strData & Chr(11) & ""														' Name
				strData = strData & Chr(11) & "<%=ConvSPChars(pAb0128.ExportItemAJnlCtrlAssnKeyColmId5(LngRow))%>"
				strData = strData & Chr(11) & "<%=pAb0128.ExportItemAJnlCtrlAssnKeyDataType5(LngRow)%>"		' Code
				strData = strData & Chr(11) & ""														' Name
				strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
				strData = strData & Chr(11) & Chr(12)
<%      
		    Next
%>    
			.ggoSpread.Source = .frm1.vspdData3
			.ggoSpread.SSShowData strData

			.lgStrPrevKeyTwo_CtrlCd = "<%=StrNextKeyTwo_CtrlCd%>"
   
			<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
			If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKeyTwo_CtrlCd <> "" Then
				.DbQuery_Two
			Else
				.frm1.hFormCnt.value = CtrlCtrlCnt
				.frm1.hTransType.value = "<%=Request("txtTransType")%>"
				.frm1.hJnlCd.value = "<%=Request("txtJnlCd")%>"
				.frm1.hDrCrFgCd.value = "<%=Request("txtDrCrFgCd")%>"
				.frm1.hAcctCd.value = "<%=Request("txtAcctCd")%>"
				.frm1.hCtrlCD.value = "<%=Request("txtCtrlCD")%>"
				.DbQuery_TwoOk
			End If

		End With
</Script>	
<%
	    Set pAb0128 = Nothing

		Response.End																		'☜: Process End
    
'***************************************************************************************************    
'                                              SAVE
'***************************************************************************************************
	Case CStr(UID_M0002)					'☜: 저장 요청을 받음 


	    Err.Clear							'☜: Protect system from crashing
		Response.Write  "mb2"
		Response.End 
		LngMaxRow = Request("txtMaxRows_Two")		'☜: 최대 업데이트된 갯수 
		Response.Write  Request("txtSpread2")
		Response.Write  "<br>"
		Response.Write  LngMaxRow
		Response.Write  "<br>"
		Response.Write  Request("txtTransType")
		Response.Write  "<br>"
		Response.Write  Request("txtTransType")
		Response.Write  "<br>"
		
		
		
	    Set pAb0121 = Server.CreateObject("Ab0121.AManageJnlCtrlAssnSvr")

	    '-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If Err.Number <> 0 Then
			Set pAb0121 = Nothing					'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)	'⊙:
			Response.End							'☜: 비지니스 로직 처리를 종료함 
		End If

		Dim arrVal, arrTemp							'☜: Spread Sheet 의 값을 받을 Array 변수 
		Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
		Dim	lGrpCnt									'☜: Group Count
		Dim strCode									'⊙: Lookup 용 리턴 변수 
		
		arrTemp = Split(Request("txtSpread2"), gRowSep)	' Spread Sheet 내용을 담고 있는 Element명 
		
	    pAb0121.ServerLocation = ggServerIP
	    
	    lGrpCnt = 0
	    
	    For LngRow = 1 To LngMaxRow
	    
			lGrpCnt = lGrpCnt +1					'☜: Group Count
			arrVal = Split(arrTemp(LngRow-1), gColSep)
			strStatus = arrVal(0)					'☜: Row 의 상태 

			'Response.Write arrTemp(LngRow-1) & "<br>"
			
			Select Case strStatus

	            Case "C"							'☜: Create

	                pAb0121.ImportItemIefSuppliedSelectChar(lGrpCnt) = "C"
	                pAb0121.ImportItemAAcctTransTypeTransType(lGrpCnt) = arrVal(2)
	                pAb0121.ImportItemAJnlItemJnlCd(lGrpCnt) = Trim(arrVal(3))
	                pAb0121.ImportItemAJnlFormSeq(lGrpCnt) = arrVal(4)
	                pAb0121.ImportItemAJnlFormDrCrFg(lGrpCnt) = arrVal(5)
	                pAb0121.ImportItemAAcctAcctCd(lGrpCnt) = Trim(arrVal(6))
	                pAb0121.ImportItemACtrlItemCtrlCd(lGrpCnt) = Trim(arrVal(7))
	                pAb0121.ImportItemAJnlCtrlAssnTblId(lGrpCnt) = arrVal(8)
	                pAb0121.ImportItemAJnlCtrlAssnDataColmId(lGrpCnt) = arrVal(9)
	                pAb0121.ImportItemAJnlCtrlAssnDataType(lGrpCnt) = arrVal(10)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId1(lGrpCnt) = arrVal(11)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType1(lGrpCnt) = arrVal(12)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId2(lGrpCnt) = arrVal(13)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType2(lGrpCnt) = arrVal(14)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId3(lGrpCnt) = arrVal(15)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType3(lGrpCnt) = arrVal(16)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId4(lGrpCnt) = arrVal(17)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType4(lGrpCnt) = arrVal(18)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId5(lGrpCnt) = arrVal(19)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType5(lGrpCnt) = arrVal(20)
	                pAb0121.ImportItemAJnlCtrlAssnInsrtUserId(lGrpCnt) = gUsrID
	                pAb0121.ImportItemAJnlCtrlAssnUpdtUserId(lGrpCnt) = gUsrID
				Case "U"

	                pAb0121.ImportItemIefSuppliedSelectChar(lGrpCnt) = "U"
	                pAb0121.ImportItemAAcctTransTypeTransType(lGrpCnt) = arrVal(2)
	                pAb0121.ImportItemAJnlItemJnlCd(lGrpCnt) = Trim(arrVal(3))
	                pAb0121.ImportItemAJnlFormSeq(lGrpCnt) = arrVal(4)
	                pAb0121.ImportItemAJnlFormDrCrFg(lGrpCnt) = arrVal(5)
	                pAb0121.ImportItemAAcctAcctCd(lGrpCnt) = Trim(arrVal(6))
	                pAb0121.ImportItemACtrlItemCtrlCd(lGrpCnt) = Trim(arrVal(7))
	                pAb0121.ImportItemAJnlCtrlAssnTblId(lGrpCnt) = arrVal(8)
	                pAb0121.ImportItemAJnlCtrlAssnDataColmId(lGrpCnt) = arrVal(9)
	                pAb0121.ImportItemAJnlCtrlAssnDataType(lGrpCnt) = arrVal(10)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId1(lGrpCnt) = arrVal(11)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType1(lGrpCnt) = arrVal(12)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId2(lGrpCnt) = arrVal(13)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType2(lGrpCnt) = arrVal(14)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId3(lGrpCnt) = arrVal(15)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType3(lGrpCnt) = arrVal(16)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId4(lGrpCnt) = arrVal(17)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType4(lGrpCnt) = arrVal(18)
	                pAb0121.ImportItemAJnlCtrlAssnKeyColmId5(lGrpCnt) = arrVal(19)
	                pAb0121.ImportItemAJnlCtrlAssnKeyDataType5(lGrpCnt) = arrVal(20)
	                pAb0121.ImportItemAJnlCtrlAssnUpdtUserId(lGrpCnt) = gUsrID
	            Case "D"
         
	                pAb0121.ImportItemIefSuppliedSelectChar(lGrpCnt) = "D"
	                pAb0121.ImportItemAAcctTransTypeTransType(lGrpCnt) = arrVal(2)
	                pAb0121.ImportItemAJnlItemJnlCd(lGrpCnt) = Trim(arrVal(3))
	                pAb0121.ImportItemAJnlFormSeq(lGrpCnt) = arrVal(4)
	                pAb0121.ImportItemAJnlFormDrCrFg(lGrpCnt) = arrVal(5)
	                pAb0121.ImportItemAAcctAcctCd(lGrpCnt) = Trim(arrVal(6))
	                pAb0121.ImportItemACtrlItemCtrlCd(lGrpCnt) = Trim(arrVal(7))
	        End Select
			

	        
	    Next

%>
	<Script Language=vbscript>
		Parent.DbSave_OneOk("<%=Request("txtTransType")%>")				'☜: 화면 처리 ASP 를 지칭함 
	</Script>
<%					
	    'Set pAb0121 = Nothing

		Response.End																		'☜: Process End

End Select

%>
