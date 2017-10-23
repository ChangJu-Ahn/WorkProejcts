<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"		-->
<!-- #Include file="../../inc/adovbs.inc"				-->
<!-- #Include file="../../inc/lgSvrVariables.inc"	-->
<!-- #Include file="../../inc/incServeradodb.asp"	-->
<!-- #Include file="../../inc/incSvrDate.inc"		-->
<!-- #Include file="../../inc/incSvrNumber.inc"		-->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Call HideStatusWnd																		'☜: Hide Processing message
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q","M", "NOCOOKIE", "MB")

	Dim userDN, userInfo, smartBillID, smartBillPW

	Dim gtaxBillNos
	Dim i, nMaxRows
	
	Dim nSucess, nFail

	nSucess = 0
	nFail = 0
	
	lgErrorStatus = "NO"
	lgErrorPos    = ""																		'☜: Set to space

	Call CreateSBObject()
	Call SubOpenDB(lgObjConn)
	Call ReadSmartBillUserInfo()

	For i = 0 To nMaxRows
		Call SendDTInfomation(gtaxBillNos(i), i)
	Next

	Set lgObjRs = Nothing
	Call SubCloseDB(lgObjConn)

'===========================================================================================================
Sub CreateSBObject()
%>
<Script Language=vbscript>
	Dim objSBI, objDTI
	' 공통변수
	Dim smartBillID, smartBillPW
	Dim nMaxRows, returnCode, strErrorMsg, dtiMessage, dtiResult, returnCode1
	Dim whereFlag, ChgRegion, ChgRemark1, ChgRemark2, ChgRemark3
	
	'거래별 변수(배열로 처리해야 하는 부분)
	Dim itemInfo1, itemInfo2, itemCount
	Dim generalInfo, supplierInfo, buyerInfo, brokerInfo, settlementInfos
	Dim mainConvID, generalInfo_T
	Dim SendingResult, strErrDesc, taxBillNos
	On Error Resume Next																		'☜: Protect system from crashing
	Err.Clear																					'☜: Clear Error status

	Set objSBI = CreateObject("SBIHandler.SBIInterface")
	If err.number <> 0 Then
		MsgBox "스마트빌에서 제공한 서버 컴포넌트(SBIHandler)가 설치가 되어 있지 않습니다. 확인하시고 다시 실행하십시오."
	End If

	'Set objDTI = Server.CreateObject("FSSmartBillDTI.DTIServerInterface")
	Set objDTI = CreateObject("FSSmartBillDTI.DTIInterface")
	If err.number <> 0 Then
		Set objSBI = Nothing
		MsgBox "스마트빌에서 제공한 서버 컴포넌트(FSSmartBillDTI)가 설치가 되어 있지 않습니다. 확인하시고 다시 실행하십시오."
	End If
	On Error Goto 0
	Err.Clear
	
<%
	gtaxBillNos = Split(Request("txtSpread"), gRowSep)
	nMaxRows = UBound(gtaxBillNos) - 1
%>
	nMaxRows = <%=nMaxRows%>
	
	Redim whereFlag(<%=nMaxRows%>)
	Redim ChgRegion(<%=nMaxRows%>)
	Redim ChgRemark1(<%=nMaxRows%>)
	Redim ChgRemark2(<%=nMaxRows%>)
	Redim ChgRemark3(<%=nMaxRows%>)
	
	Redim SendingResult(<%=nMaxRows%>)
	Redim strErrDesc(<%=nMaxRows%>)
	Redim taxBillNos(<%=nMaxRows%>)
	Redim generalInfo(<%=nMaxRows%>)
	Redim supplierInfo(<%=nMaxRows%>)
	Redim buyerInfo(<%=nMaxRows%>)
	Redim brokerInfo(<%=nMaxRows%>)
	Redim settlementInfo(<%=nMaxRows%>)
	Redim itemInfo1(<%=nMaxRows%>)
	Redim itemInfo2(<%=nMaxRows%>)
	Redim mainConvID(<%=nMaxRows%>)
	Redim generalInfo_T(<%=nMaxRows%>)
	Redim itemCount(<%=nMaxRows%>)
<%
End Sub

'===========================================================================================================
Sub ReadSmartBillUserInfo()
	On Error Resume Next																		'☜: Protect system from crashing
	Err.Clear																					'☜: Clear Error status

	lgStrSQL = "" & _
"SELECT dt_id, dt_pw " & vbCrLf & _
"  FROM dt_user_info " & vbCrLf & _
" WHERE user_id = " & FilterVar(gUsrId, "''", "S")

	If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		Call DisplayMsgBox("205921", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
	Else	%>
		smartBillID = "<%=ConvSPChars(lgObjRs("dt_id")) %>"
		smartBillPW = "<%=ConvSPChars(lgObjRs("dt_pw")) %>"
<%	End If
End Sub

'===========================================================================================================
Function SendDTInfomation(taxBillNo, nRow)
	On Error Resume Next																		'☜: Protect system from crashing
	Err.Clear																					'☜: Clear Error status

	Dim generalInfo, supplierInfo, buyerInfo, brokerInfo, settlementInfo, sendData
	Dim loopLength, j, k, IntRetCD, net_loc_amt, vat_amt_loc
	Dim itemInfo1, itemInfo2, addItemInfo
	Dim isSucess, ErrDesc, mainConvID, generalInfo_T
	Dim arrColums
	
	isSucess = "N"
	ErrDesc = ""
	arrColums = Split(taxBillNo, gColSep)
	
	lgStrSQL = "EXEC dbo.usp_dt_send_tax_smartbill_1 " & FilterVar(arrColums(0), "", "S") & ", " & _
																		 FilterVar(arrColums(1), "", "S") & ", " & _
																		 FilterVar(gUsrId, "", "S") & ", " & _
																		 FilterVar(arrColums(3), "", "S") & ", " & _
																		 FilterVar(arrColums(2), "", "S") & ", " & _
																		 FilterVar(arrColums(4), "", "S") & ", " & _
																		 FilterVar(arrColums(5), "", "S")

	If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found.
		lgErrorStatus  = "YES"
		Exit Function
	End If

	generalInfo = lgObjRs("generalInfo")
   generalInfo_T = lgObjRs("generalInfo_T")
   mainConvID = lgObjRs("mainConvId")

	supplierInfo = lgObjRs("supplierInfo")
	buyerInfo = lgObjRs("buyerInfo")
	brokerInfo = lgObjRs("brokerInfo")
	settlementInfo = lgObjRs("settlementInfo")

	net_loc_amt = lgObjRs("net_loc_amt")
	vat_amt_loc = lgObjRs("vat_amt_loc")
	Set lgObjRs = Nothing
	
	lgStrSQL = "EXEC dbo.usp_dt_send_tax_smartbill_2 " & FilterVar(arrColums(0), "", "S") & ", " & _
																		 FilterVar(arrColums(1), "", "S") & ", " & _
																		 FilterVar(gUsrId, "", "S")

	Dim Rs
	If FncOpenRs("R", lgObjConn, Rs, lgStrSQL, "X", "X") = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found.
		lgErrorStatus  = "YES"
		Exit Function
	End If
	
	' 품목이 4건이 넘을 경우
	j = 0
	Do While Not Rs.EOF
		If j = 0 Then
			addItemInfo = Split(Rs("itemInfo"), ";#;")
		End If

		itemInfo1 = itemInfo1 & Rs("itemInfo")
		itemInfo2 = itemInfo2 & Rs("itemInfo_T")
		
		Rs.MoveNext
		j = j + 1
	Loop

	If j >= 4 Then
		addItemInfo(0) = addItemInfo(0) & ";#;"
		addItemInfo(1) = addItemInfo(1) & " 외 " & (j - 1) & "건;#;"
		addItemInfo(2) = ";#;"
		addItemInfo(3) = ";#;"
		addItemInfo(4) = ";#;"
		addItemInfo(5) = "0;#;"
		addItemInfo(6) = "0;#;"
		addItemInfo(7) = CStr(net_loc_amt) & ";#;"
		addItemInfo(8) = CStr(vat_amt_loc) & ";#;"

		itemInfo1 = Join(addItemInfo, "")
	End If
%>	
	'===============================================================================================================
   ' 여기부터 거래명세서 및 전송 부분.
	whereFlag(<%=nRow%>) = "<%=arrColums(0)%>"
	taxBillNos(<%=nRow%>) = "<%=arrColums(1)%>" 
	ChgRegion(<%=nRow%>) = "<%=arrColums(2)%>"
	ChgRemark1(<%=nRow%>) = "<%=arrColums(3)%>"
	ChgRemark2(<%=nRow%>) = "<%=arrColums(4)%>"
	ChgRemark3(<%=nRow%>) = "<%=arrColums(5)%>"

	generalInfo(<%=nRow%>) = "<%=generalInfo %>"
	generalInfo_T(<%=nRow%>) = "<%=generalInfo_T %>"
	mainConvID(<%=nRow%>) = "<%=mainConvID %>"

	supplierInfo(<%=nRow%>) = "<%=supplierInfo%>"
	buyerInfo(<%=nRow%>) = "<%=buyerInfo %>"
	brokerInfo(<%=nRow%>) = "<%=brokerInfo %>"
	settlementInfo(<%=nRow%>) = "<%=settlementInfo %>"

	itemInfo1(<%=nRow%>) = "<%=ConvSPChars(itemInfo1) %>"
	itemInfo2(<%=nRow%>) = "<%=ConvSPChars(itemInfo2) %>"
	itemCount(<%=nRow%>) = <%=j %>
<%
End Function	%>

	Dim strFilePath
	Dim strFileName
	Dim i
	
	For i = 0 To nMaxRows
		returnCode = objDTI.makeDTTFrameWorkForHubBulkWithoutPKIV3(generalInfo_T(i), supplierInfo(i), buyerInfo(i), brokerInfo(i), settlementInfo(i), 4, itemInfo2(i), itemCount(i), "C:\SBCSolution\ARISSUETXT\")
		dtiResult = Split(returnCode, ";#;")

		If dtiResult(0) <> "0" Then
			SendingResult(i) = "N"
			strErrDesc(i) = "거래명세서 작성 오류 입니다."
		Else	'사용자 화면 없이 주어진 디렉토리에 저장 한다. 단 존재하는 디렉토리어야 한다.
			strFilePath = objDTI.getFilePath()
			strFileName = objDTI.getFileName()

			returnCode = objSBI.processServiceForServerV3(mainConvID(i), "15003", strFilePath, strFileName, 0, smartBillID, smartBillPW, "", "DETAILARISSUE")
			dtiResult = Split(returnCode, ";#;")

         ' 스마트빌로 전송 처리 성공 ==> 전자세금계산서 전송 시작
			If( dtiResult(0) = "30000") Then	' 세금 계산서 전송
                If itemCount(i) > 4 Then 
                    itemCount(i) = 1
                End If
                
				' 전송부문
				returnCode = objDTI.makeDTIFrameWorkForHUBBulkV3(generalInfo(i), supplierInfo(i), buyerInfo(i), brokerInfo(i), settlementInfo(i), 4, itemInfo1(i), itemCount(i), parent.frm1.txtuserInfo.value, "")
				dtiResult = Split(returnCode, ";#;")

				If (dtiResult(0) <> "0") Then						' 세금계산서 XML파일을 생성하지 못했을 시
					SendingResult(i) = "N"
					strErrDesc(i) = dtiResult(0) & " - 전자 세금계산서 원본파일 생성 오류입니다."
				Else
					strFilePath = objDTI.getFilePath()
					strFileName = objDTI.getFileName()

					returnCode = objSBI.dtiServiceWithoutGUIV3(mainConvID(i), "15000", strFilePath, strFileName, 0, smartBillID, smartBillPW, parent.frm1.txtuserDN.value, dtiResult(2), smartBillID, "", "", "ARISSUE")

					If returnCode = "30000" Then
						SendingResult(i) = "Y"
						strErrDesc(i) = ""
					Else
						SendingResult(i) = "N"
						strErrDesc(i) = returnCode & "-" & objSBI.getErrorMsg()
					End If
				End If
			Else
				SendingResult(i) = "N"
				strErrDesc(i) = dtiResult(0) & "-" & objSBI.getErrorMsg()
			End If
		End If
	Next

	Set objSBI = Nothing
	Set objDTI = Nothing

	' 결과를 저장하기 위해 문자열을 조합한다.
	Dim strSendData
	With parent.parent
	For i = 0 To nMaxRows
		strSendData = strSendData & whereFlag(i) & .gColSep & _
											 taxBillNos(i) & .gColSep & _
											 mainConvID(i) & .gColSep & _
											 SendingResult(i) & .gColSep & _
											 strErrDesc(i) & .gColSep & _
						  					 ChgRegion(i) & .gColSep & _
											 ChgRemark1(i) & .gColSep & _
											 ChgRemark2(i) & .gColSep & _
											 ChgRemark3(i) & .gRowSep
	Next
	End With
	parent.frm1.txtSpread.value = strSendData
	parent.SaveResult()
</Script>