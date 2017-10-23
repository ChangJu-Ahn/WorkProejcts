<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3221mb7.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Import Local L/C 등록 Save Transaction 처리용 ASP							*
'*  7. Modified date(First) : 2000/05/02																*
'*  8. Modified date(Last)  : 2000/05/02																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/27 : Coding Start												*
'********************************************************************************************************

Response.Expires = -1															'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True															'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call HideStatusWnd

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")	

Select Case strMode
	Case CStr(UID_M0002)														'☜: 현재 Save 요청을 받음 
		Dim M32211																' Master L/C Amend Header Save용 Object
'		Dim B1H019																' Transport Lookup용 Object
'		Dim strTransport
		
		Err.Clear																'☜: Protect system from crashing

		lgIntFlgMode = CInt(Request("txtFlgMode"))								'☜: 저장시 Create/Update 판별 
	
		'⊙: 각 화면당 Relation이 되어 있지 않는 Field들에 대해서는 Lookup을 행한다.

		'⊙: Lookup Pad 동작후 정상적인 데이타 이면, 저장 로직 시작 
		Set M32211 = Server.CreateObject("M32211.M32211MaintLcAmendHdrSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M32211 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		M32211.ImportMLcAmendHdrRemark = Trim(Request("txtRemark"))
		M32211.ImportMLcAmendHdrCurrency = UCase(Trim(Request("txtCurrency")))

		If Len(Trim(Request("txtBeDocAmt"))) Then
			M32211.ImportMLcAmendHdrBeDocAmt = UNIConvNum(Request("txtBeDocAmt"),0)
		End If

		If Len(Trim(Request("txtOpenDt"))) Then
			strConvDt = UNIConvDate(Request("txtOpenDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtOpenDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Response.End													'☜: Process End
			Else
				M32211.ImportMLcAmendHdrOpenDt = strConvDt
			End If
		End If

'		If Len(Trim(Request("txtOpenDt"))) Then
'			M32211.ImportMLcAmendHdrOpenDt = UNIConvDate(Request("txtOpenDt"))
'		End If
		
		If Len(Trim(Request("txtAmendReqDt"))) Then
			M32211.ImportMLcAmendHdrAmendReqDt = UNIConvDate(Request("txtAmendReqDt"))
		End If		

		M32211.ImportMLcAmendHdrOpenBank = UCase(Trim(Request("txtOpenBank")))
		M32211.ImportMLcAmendHdrAdvNo = UCase(Trim(Request("txtAdvNo")))
		M32211.ImportMLcAmendHdrBeneficiary = UCase(Trim(Request("txtBeneficiary")))
		M32211.ImportMLcAmendHdrLcAmdNo = UCase(Trim(Request("txtHLCAmdNo")))
		M32211.ImportMLcAmendHdrLcDocNo = UCase(Trim(Request("txtLCDocNo")))
		M32211.ImportMLcAmendHdrLcNo = UCase(Trim(Request("txtLCNo")))
		M32211.ImportMLcAmendHdrPurGrp = UCase(Trim(Request("txtPurGrp")))
		M32211.ImportMLcAmendHdrPurOrg = UCase(Trim(Request("txtPurOrg")))
		
		If Len(Trim(Request("txtLCAmendSeq"))) Then
			M32211.ImportMLcAmendHdrLcAmendSeq = UNIConvNum(Request("txtLCAmendSeq"),0)
		End If

		M32211.ImportMLcAmendHdrApplicant = UCase(Trim(Request("txtApplicant")))

		If Len(Trim(Request("txtAmendDt"))) Then
			strConvDt = UNIConvDate(Request("txtAmendDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtAmendDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Response.End													'☜: Process End
			Else
				M32211.ImportMLcAmendHdrAmendDt = strConvDt
			End If
		End If

'		If Len(Trim(Request("txtAmendDt"))) Then
'			M32211.ImportMLcAmendHdrAmendDt = UNIConvDate(Request("txtAmendDt"))
'		End If

		If Not ISEMPTY(Request("chkAtDocAmt")) Then
			If Request("rdoAtDocAmt") = "I" Then
				If Len(Trim(Request("txtAmendAmt"))) Then
					M32211.ImportMLcAmendHdrIncAmt = UNIConvNum(Request("txtAmendAmt"),0)
					M32211.ImportMLcAmendHdrAtDocAmt = CDbl(UNIConvNum(Request("txtBeDocAmt"),0)) + CDbl(UNIConvNum(Request("txtAmendAmt"),0))
				End If
			ElseIf Request("rdoAtDocAmt") = "D" Then
				If Len(Trim(Request("txtAmendAmt"))) Then
					M32211.ImportMLcAmendHdrDecAmt = UNIConvNum(Request("txtAmendAmt"),0)
					M32211.ImportMLcAmendHdrAtDocAmt = CDbl(UNIConvNum(Request("txtBeDocAmt"),0)) - CDbl(UNIConvNum(Request("txtAmendAmt"),0))
				End If
			End If

		Else
			If Len(Trim(Request("txtBeDocAmt"))) Then
				M32211.ImportMLcAmendHdrAtDocAmt = UNIConvNum(Request("txtBeDocAmt"),0)
			End If
			
			If Len(Trim(Request("txtBeLocAmt"))) Then
				M32211.ImportMLcAmendHdrAtLocAmt = UNIConvNum(Request("txtBeLocAmt"),0)
			End If

			M32211.ImportMLcAmendHdrIncAmt = UNIConvNum(Request("txtIncAmt"),0)
			M32211.ImportMLcAmendHdrDecAmt = UNIConvNum(Request("txtDecAmt"),0)
		End If		
		
'		If Not ISEMPTY(Request("chkAtDocAmt")) Then
'			If Request("rdoAtDocAmt") = "I" Then
'				If Len(Trim(Request("txtAtDocAmt"))) Then
'					M32211.ImportMLcAmendHdrIncAmt = UNIConvNum(Request("txtAtDocAmt"),0)
'					M32211.ImportMLcAmendHdrAtDocAmt = CDbl(UNIConvNum(Request("txtBeDocAmt"),0)) + CDbl(UNIConvNum(Request("txtAtDocAmt"),0))
'				End If
'			ElseIf Request("rdoAtDocAmt") = "D" Then
'				If Len(Trim(Request("txtAtDocAmt"))) Then
'					M32211.ImportMLcAmendHdrDecAmt = UNIConvNum(Request("txtAtDocAmt"),0)
'					M32211.ImportMLcAmendHdrAtDocAmt = CDbl(UNIConvNum(Request("txtBeDocAmt"),0)) - CDbl(UNIConvNum(Request("txtAtDocAmt"),0))
'				End If
'			End If
'		Else
'			If Len(Trim(Request("txtAmendAmt"))) Then
'				M32211.ImportMLcAmendHdrAtDocAmt = UNIConvNum(Request("txtAmendAmt"),0)
'			End If
'		End If

		If Not ISEMPTY(Request("chkAtExpiryDt")) Then
			If Len(Trim(Request("txtAtExpiryDt"))) Then
				strConvDt = UNIConvDate(Request("txtAtExpiryDt"))

				If strConvDt = "" Then
					Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
					Call LoadTab("parent.frm1.txtAtExpiryDt", 1, I_MKSCRIPT)

					Set M32211 = Nothing											'☜: ComProxy UnLoad

					Response.End													'☜: Process End
				Else
					M32211.ImportMLcAmendHdrAtExpiryDt = strConvDt
				End If
			End If
		Else
			If Len(Trim(Request("txtHExpiryDt"))) Then
				strConvDt = UNIConvDate(Request("txtHExpiryDt"))

				If strConvDt = "" Then
					Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
					'Call LoadTab("parent.frm1.txtBeExpiryDt", 1, I_MKSCRIPT)

					Set M32211 = Nothing											'☜: ComProxy UnLoad

					Response.End													'☜: Process End
				Else
					M32211.ImportMLcAmendHdrAtExpiryDt = strConvDt
				End If
			Else
				If Len(Trim(Request("txtBeExpiryDt"))) Then
					strConvDt = UNIConvDate(Request("txtBeExpiryDt"))

					If strConvDt = "" Then
						Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
						Call LoadTab("parent.frm1.txtBeExpiryDt", 1, I_MKSCRIPT)

						Set M32211 = Nothing											'☜: ComProxy UnLoad

						Response.End													'☜: Process End
					Else
						M32211.ImportMLcAmendHdrAtExpiryDt = strConvDt
					End If
				End If
			End If
		End If

'		If Not ISEMPTY(Request("chkAtExpiryDt")) Then
'			If Len(Trim(Request("txtAtExpiryDt"))) Then
'				M32211.ImportMLcAmendHdrAtExpiryDt = UNIConvDate(Request("txtAtExpiryDt"))
'			End If
'		Else
'			If Len(Trim(Request("txtBeExpiryDt"))) Then
'				M32211.ImportMLcAmendHdrAtExpiryDt = UNIConvDate(Request("txtBeExpiryDt"))
'			End If
'		End If

		If Len(Trim(Request("txtBeExpiryDt"))) Then
			strConvDt = UNIConvDate(Request("txtBeExpiryDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtBeExpiryDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Response.End													'☜: Process End
			Else
				M32211.ImportMLcAmendHdrBeExpiryDt = strConvDt
			End If
		End If

'		If Len(Trim(Request("txtBeExpiryDt"))) Then
'			M32211.ImportMLcAmendHdrBeExpiryDt = UNIConvDate(Request("txtBeExpiryDt"))
'		End If

		If Not ISEMPTY(Request("chkAtLatestShipDt")) Then
			If Len(Trim(Request("txtAtLatestShipDt"))) Then
				strConvDt = UNIConvDate(Request("txtAtLatestShipDt"))

				If strConvDt = "" Then
					Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
					Call LoadTab("parent.frm1.txtAtLatestShipDt", 1, I_MKSCRIPT)

					Set M32211 = Nothing											'☜: ComProxy UnLoad

					Response.End													'☜: Process End
				Else
					M32211.ImportMLcAmendHdrAtLatestShipDt = strConvDt
				End If
			End If
		Else
			If Len(Trim(Request("txtHLatestShipDt"))) Then
				strConvDt = UNIConvDate(Request("txtHLatestShipDt"))

				If strConvDt = "" Then
					Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
					'Call LoadTab("parent.frm1.txtBeExpiryDt", 1, I_MKSCRIPT)

					Set M32211 = Nothing											'☜: ComProxy UnLoad

					Response.End													'☜: Process End
				Else
					M32211.ImportMLcAmendHdrAtLatestShipDt = strConvDt
				End If
			Else
				If Len(Trim(Request("txtBeLatestShipDt"))) Then
					strConvDt = UNIConvDate(Request("txtBeLatestShipDt"))

					If strConvDt = "" Then
						Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
						Call LoadTab("parent.frm1.txtBeLatestShipDt", 1, I_MKSCRIPT)

						Set M32211 = Nothing											'☜: ComProxy UnLoad

						Response.End													'☜: Process End
					Else
						M32211.ImportMLcAmendHdrAtLatestShipDt = strConvDt
					End If
				End If
			End If
		End If

'		If Not ISEMPTY(Request("chkAtLatestShipDt")) Then
'			If Len(Trim(Request("txtAtLatestShipDt"))) Then
'				M32211.ImportMLcAmendHdrAtLatestShipDt = UNIConvDate(Request("txtAtLatestShipDt"))
'			End If
'		Else
'			If Len(Trim(Request("txtBeLatestShipDt"))) Then
'				M32211.ImportMLcAmendHdrAtLatestShipDt = UNIConvDate(Request("txtBeLatestShipDt"))
'			End If
'		End If

		If Len(Trim(Request("txtBeLatestShipDt"))) Then
			strConvDt = UNIConvDate(Request("txtBeLatestShipDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtBeLatestShipDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Response.End													'☜: Process End
			Else
				M32211.ImportMLcAmendHdrBeLatestShipDt = strConvDt
			End If
		End If

'		If Len(Trim(Request("txtBeLatestShipDt"))) Then
'			M32211.ImportMLcAmendHdrBeLatestShipDt = UNIConvDate(Request("txtBeLatestShipDt"))
'		End If

		M32211.ImportMLcAmendHdrAtTranshipmentAsString = "N"
		M32211.ImportMLcAmendHdrBeTranshipmentAsString = "N"

		If Not ISEMPTY(Request("chkAtPartialShip")) Then
			M32211.ImportMLcAmendHdrAtPartialShipAsString = Request("rdoAtPartialShip")
		Else
			If Len(Trim(Request("txtHPartialShip"))) Then
				M32211.ImportMLcAmendHdrAtPartialShipAsString = Request("txtHPartialShip")
			Else
				M32211.ImportMLcAmendHdrAtPartialShipAsString = Request("txtBePartialShip")
			End If
		End If

'		If Not ISEMPTY(Request("chkAtPartialShip")) Then
'			M32211.ImportMLcAmendHdrAtPartialShipAsString = Request("rdoAtPartialShip")
'		Else
'			M32211.ImportMLcAmendHdrAtPartialShipAsString = Request("txtBePartialShip")
'		End If

		M32211.ImportMLcAmendHdrBePartialShipAsString = Request("txtBePartialShip")

		M32211.ImportMLcAmendHdrAtTransferAsString = "N"
		M32211.ImportMLcAmendHdrBeTransferAsString = "N"
		
'		M32211.ImportMLcAmendHdrBeTransport = ""
'		M32211.ImportMLcAmendHdrAtTransport = ""

		M32211.ImportMLcAmendHdrLcKindAsString = "L"
		M32211.ImportMLcAmendHdrInsrtUserId = UCase(Trim(Request("txtInsrtUserId")))
		M32211.ImportMLcAmendHdrUpdtUserId = UCase(Trim(Request("txtUpdtUserId")))

		If lgIntFlgMode = OPMD_CMODE Then
			M32211.CommandSent = "CREATE"
		ElseIf lgIntFlgMode = OPMD_UMODE Then
			M32211.CommandSent = "UPDATE"
		End If

		M32211.ServerLocation = ggServerIP

		'-----------------------
		'Com action area
		'-----------------------
		M32211.ComCfg = gConnectionString
		M32211.Execute

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set M32211 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (M32211.OperationStatusMessage = MSG_OK_STR) Then
			Select Case M32211.OperationStatusMessage
				Case MSG_DEADLOCK_STR
					Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
				Case MSG_DBERROR_STR
					Call DisplayMsgBox2(M32211.ExportErrEabSqlCodeSqlcode, _
							    M32211.ExportErrEabSqlCodeSeverity, _
							    M32211.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
				Case Else
					Call DisplayMsgBox(M32211.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
			End Select

			Set M32211 = Nothing
			Response.End 
		End If

		'-----------------------
		'Result data display area
		'-----------------------
%>
<Script Language=VBScript>
	With parent
		.frm1.txtLCAmdNo.value = "<%=ConvSPChars(M32211.ExportMLcAmendHdrLcAmdNo)%>"
		.DbSaveOk
	End With
</Script>
<%
		Set M32211 = Nothing														'☜: Unload Comproxy

		Response.End																'☜: Process End

	Case Else
		Response.End
End Select
%>
