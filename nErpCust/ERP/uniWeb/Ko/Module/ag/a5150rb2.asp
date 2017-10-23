<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->


<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다
On Error Resume Next
Err.Clear

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1	, rs2, rs3, rs4	       '☜ : DBAgent Parameter 선언
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iPrevEndRow
Dim iEndRow

Dim lgObjComm  
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

Dim strOpenType
Dim strFrOpenDt
Dim strToOpenDt
Dim strDocCur
Dim strOrgChangeId
Dim strDeptCd
Dim strBpCd
Dim strBizCd
Dim strBpCd2
Dim strProject
Dim strAllcAmt
Dim strRefNo
Dim strAcctCd
Dim strGlNo
Dim strMgntCd1
Dim strMgntCd2
Dim strCardCoCd
Dim strCardNo
Dim strFrCardUserId
Dim strToCardUserId
Dim strChkLocalCur
Dim strCond
Dim strFrDueDt
Dim strToDueDt
Dim strParentGlNo

Dim PAmt
Dim strMsgCd
Dim strMsg1
Dim skip_rs3,skip_rs4,no_mgnt1,no_mgnt2

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","A","NOCOOKIE","RB")
    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = UNICInt(Request("lgMaxCount") ,0)                          '☜ : 한번에 가져올수 있는 데이타 건수
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    iPrevEndRow	   = 0
    iEndRow        = 0

    Call SubOpenDB(lgObjConn)                                               '☜: Make a DB Connection
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    Call SubCloseDB(lgObjConn)                                              '☜: Close DB Connection    

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr,iRowStrtmp
    Dim iCtrl_cd,iCtrl_val
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    'If strAllcAmt <> 0 Then
    if 1=1 then
		Do While Not (Rs0.EOF Or Rs0.BOF)
			PAmt = strAllcAmt 
			strAllcAmt = strAllcAmt - UNIConvNum(Rs0(21) ,0)

		    iRowStr = ""
		   
		    
			For ColCnt = 0 To UBound(lgSelectListDT) - 1
			
				If ColCnt = 21  Then '반제할금액 셋팅 
					'If strAllcAmt > 0 Then 
						iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt)) ' 잔액을 그대로 보여줌.
					'Else
					'	iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),PAmt)
					'End If
			
				elseIf ColCnt = 26  Then '

					
					iCtrl_cd = Trim(rs0(ColCnt-2))
					iCtrl_val = Trim(rs0(ColCnt-1))

					If iCtrl_val <> ""  Then 
					    Call SubCreateCommandObject(lgObjComm)
						With lgObjComm
						    .CommandText = "USP_A_MGNT_NAME"
						    .CommandType = adCmdStoredProc
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_CD"  ,adVarWChar,adParamInput,Len(iCtrl_cd),iCtrl_cd)
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_VAL" ,adVarWChar,adParamInput,Len(iCtrl_val),iCtrl_val)
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@OUT_DATA_COLM_NM",adVarWChar,adParamOutput ,128)
							    
						    lgObjComm.Execute ,, adExecuteNoRecords
						End With
							
						If  Err.number = 0 Then
							iRowStrtmp =  lgObjComm.Parameters("@OUT_DATA_COLM_NM").Value
						End If
						Call SubCloseCommandObject(lgObjComm)
						
					Else
					
						iRowStrtmp =  ""
						
					End If

				iRowStr = iRowStr & Chr(11) & iRowStrtmp
			
				elseIf ColCnt = 29  Then '

					iCtrl_cd = Trim(rs0(ColCnt-2))
					iCtrl_val = Trim(rs0(ColCnt-1))

					If iCtrl_val <> ""  Then 
					Call SubCreateCommandObject(lgObjComm)
						With lgObjComm
						    .CommandText = "USP_A_MGNT_NAME"
						    .CommandType = adCmdStoredProc
		
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_CD"  ,adVarWChar,adParamInput,Len(iCtrl_cd),iCtrl_cd)
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_VAL" ,adVarWChar,adParamInput,Len(iCtrl_val),iCtrl_val)
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@OUT_DATA_COLM_NM",adVarWChar,adParamOutput ,128)
							    
						    lgObjComm.Execute ,, adExecuteNoRecords
						End With
							
						If  Err.number = 0 Then
							iRowStrtmp =  lgObjComm.Parameters("@OUT_DATA_COLM_NM").Value
						End If
						Call SubCloseCommandObject(lgObjComm)
						
					Else
						iRowStrtmp =  ""
					End If
					
					
				iRowStr = iRowStr & Chr(11) & iRowStrtmp


				Else					
					iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
				End If	
				
				
				
			Next

			lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
			
			iEndRow = iLoopCount
			
			iLoopCount = iLoopCount + 1

			'If strAllcAmt <= 0 Then 
				'Exit Do
			'End If	
			        
			rs0.MoveNext
		Loop
	Else
		If CDbl(lgPageNo) > 0 Then
			iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)
			rs0.Move = iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		End If

		iLoopCount = -1
    
		Do While Not (rs0.EOF Or rs0.BOF)
		    iLoopCount =  iLoopCount + 1
		    iRowStr = ""
			For ColCnt = 0 To UBound(lgSelectListDT) - 1 
		        iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			Next

		    If  iLoopCount < lgMaxCount Then
		        lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
		    Else
		        lgPageNo = lgPageNo + 1
		        Exit Do
		    End If
		    rs0.MoveNext
		Loop
	End if

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If 
    	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim temp_nm1,temp_nm2
	Dim stbl_id,scol_id,stbl_id2,scol_id2

    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(4,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수

	Select Case UCase(strOpenType)
		Case "AR"
			UNISqlId(0) = "A5150RA201"
			UNISqlId(1) = "ADEPTNM"
			UNISqlId(2) = "ABPNM"
			UNISqlId(3) = "ABIZNM"
			UNISqlId(4) = "ABPNM"									

			UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
			UNIValue(1,1) = UCase(" " & FilterVar(strOrgChangeId, "''", "S") & " " )
			UNIValue(2,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & " " )
			UNIValue(3,0) = UCase(" " & FilterVar(strBizCd, "''", "S") & " " )
			UNIValue(4,0) = UCase(" " & FilterVar(strBpCd2, "''", "S") & " " )			
		Case "AP"
			UNISqlId(0) = "A5150RA202"
			UNISqlId(1) = "ADEPTNM"
			UNISqlId(2) = "ABPNM"
			UNISqlId(3) = "ABIZNM"
			UNISqlId(4) = "ABPNM"
			
			UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
			UNIValue(1,1) = UCase(" " & FilterVar(strOrgChangeId, "''", "S") & " " )
			UNIValue(2,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & " " )
			UNIValue(3,0) = UCase(" " & FilterVar(strBizCd, "''", "S") & " " )
			UNIValue(4,0) = UCase(" " & FilterVar(strBpCd2, "''", "S") & " " )						
		Case "PP"
			UNISqlId(0) = "A5150RA203"		
			UNISqlId(1) = "ADEPTNM"
			UNISqlId(2) = "ABPNM"
			UNISqlId(3) = "ABIZNM"

			UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
			UNIValue(1,1) = UCase(" " & FilterVar(strOrgChangeId, "''", "S") & " " )
			UNIValue(2,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & " " )
			UNIValue(3,0) = UCase(" " & FilterVar(strBizCd, "''", "S") & " " )
		Case "PR"
			UNISqlId(0) = "A5150RA204"		
			UNISqlId(1) = "ADEPTNM"
			UNISqlId(2) = "ABPNM"
			UNISqlId(3) = "ABIZNM"
			
			UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
			UNIValue(1,1) = UCase(" " & FilterVar(strOrgChangeId, "''", "S") & " " )
			UNIValue(2,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & " " )
			UNIValue(3,0) = UCase(" " & FilterVar(strBizCd, "''", "S") & " " )
		Case "SS"
			UNISqlId(0) = "A5150RA205"		
			UNISqlId(1) = "ADEPTNM"
			UNISqlId(2) = "ABPNM"
			UNISqlId(3) = "ABIZNM"
			
			UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
			UNIValue(1,1) = UCase(" " & FilterVar(strOrgChangeId, "''", "S") & " " )
			UNIValue(2,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & " " )
			UNIValue(3,0) = UCase(" " & FilterVar(strBizCd, "''", "S") & " " )			
		Case "U6" '미결(신용카드)
			UNISqlId(0) = "A5150RA207"		
			UNISqlId(1) = "ADEPTNM"
			UNISqlId(2) = "CommonQry"
			UNISqlId(3) = "CommonQry"
			UNISqlId(4) = "CommonQry"			
			
			UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
			UNIValue(1,1) = UCase(" " & FilterVar(strOrgChangeId, "''", "S") & " " )			
			UNIValue(2,0) = " select card_co_cd,card_co_nm from b_card_co where card_co_cd = " & FilterVar(strCardCoCd,"''","S")
			UNIValue(3,0) = " select emp_no,name from Haa010t where emp_no = " & FilterVar(strFrCardUserId,"''","S")
			UNIValue(4,0) = " select emp_no,name from Haa010t where emp_no = " & FilterVar(strToCardUserId,"''","S")
		Case "U9" '미결(기타)
			UNISqlId(0) = "A5150RA206"		
			UNISqlId(1) = "ADEPTNM"
			UNISqlId(2) = "AACCTNM"
			UNISqlId(3) = "CommonQry"
			UNISqlId(4) = "CommonQry"

			UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
			UNIValue(1,1) = UCase(" " & FilterVar(strOrgChangeId, "''", "S") & " " )			
			UNIValue(2,0) = UCase(" " & FilterVar(strAcctCd, "''", "S") & " " )

			If strMgntCd1 <> "" Then
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " select b.tbl_id,b.key_colm_id1,b.data_colm_nm from a_acct a , a_ctrl_item b "
				lgStrSQL = lgStrSQL & " where a.mgnt_cd1 = b.ctrl_cd and a.acct_cd =  " & FilterVar(strAcctCd , "''", "S") & ""
				        
				If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
					no_mgnt1 = "TRUE"
					UNIValue(3,0) = " "
				Else	
					If lgStrflag = 1 Then
						stbl_id = Trim(lgObjRs("tbl_id"))
						scol_id = Trim(lgObjRs("key_colm_id1"))
						temp_nm1 = Trim(lgObjRs("data_colm_nm"))

						UNIValue(3,0) = " select distinct(a.mgnt_val1),b."&temp_nm1 & " from a_open_acct a , "&stbl_id & " b"
						UNIValue(3,0) = UNIValue(3,0) & " where a.acct_cd =  " & FilterVar(strAcctCd , "''", "S") & ""
						UNIValue(3,0) = UNIValue(3,0) & "  and  a.mgnt_val1 = b."&scol_id		
						UNIValue(3,0) = UNIValue(3,0) & "  and  a.mgnt_val1 =  " & FilterVar(strMgntCd1 , "''", "S") & ""
					Else
						skip_rs3 = "TRUE"
						UNIValue(3,0) = " "
					End If
				End If

				Call SubCloseRs(lgObjRs)															'☜: Release RecordSSet	
			Else		
				skip_rs3 = "TRUE"
				UNIValue(3,0) = " "
			End If

			If strMgntCd2 <> "" Then
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " select b.tbl_id,b.key_colm_id1,b.data_colm_nm from a_acct a , a_ctrl_item b "
				lgStrSQL = lgStrSQL & " where a.mgnt_cd1 = b.ctrl_cd and a.acct_cd =  " & FilterVar(strAcctCd , "''", "S") & ""
				        
				If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
					no_mgnt1 = "TRUE"
					UNIValue(4,0) = " "
				Else	
					If lgStrflag = 1 then
						stbl_id = Trim(lgObjRs("tbl_id"))
						scol_id = Trim(lgObjRs("key_colm_id1"))
						temp_nm1 = Trim(lgObjRs("data_colm_nm"))

						UNIValue(4,0) = " select distinct(a.mgnt_val1),b."&temp_nm1 & " from a_open_acct a , "&stbl_id & " b"
						UNIValue(4,0) = UNIValue(3,0) & " where a.acct_cd =  " & FilterVar(strAcctCd , "''", "S") & ""
						UNIValue(4,0) = UNIValue(3,0) & "  and  a.mgnt_val1 = b."&scol_id		
						UNIValue(4,0) = UNIValue(3,0) & "  and  a.mgnt_val1 =  " & FilterVar(strMgntCd2 , "''", "S") & ""
					Else
						skip_rs4 = "TRUE"
						UNIValue(4,0) = " "
					End if
				End If

				Call SubCloseRs(lgObjRs)															'☜: Release RecordSSet	
			Else		
				skip_rs4 = "TRUE"
				UNIValue(4,0) = " "
			End If								
	End Select 	

    UNIValue(0,0) = lgSelectList                                          '☜: Select list
	UNIValue(0,1) = strCond

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언

    Set lgADF = Server.CreateObject("prjPublic.cCtlTake")
    
    Select Case UCase(strOpenType)
		Case "AR","AP","U9","U6"
			lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1 , rs2, rs3 , rs4)
		Case "PP","PR","SS"
			lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1 , rs2, rs3 )
	End Select	
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strDeptCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDeptCd_alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.txtDeptNm.value = ""
		End With
		</Script>
<%
    Else
%>
		<Script Language=vbScript>
		With parent
			.txtDeptCd.value = "<%=Trim(ConvSPChars(rs1(0)))%>"
			.txtDeptNm.value = "<%=Trim(ConvSPChars(rs1(1)))%>"
		End With
		</Script>
<%
    End If

	Set rs1 = Nothing 


	Select Case UCase(strOpenType)
		Case "AR","AP","PR","PP","SS"
		    If (rs2.EOF And rs2.BOF) Then
				If strMsgCd = "" And strBpCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBpCd_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtBpNm.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.txtBpCd.value = "<%=Trim(ConvSPChars(rs2(0)))%>"
					.txtBpNm.value = "<%=Trim(ConvSPChars(rs2(1)))%>"
				End With
				</Script>
		<%
		    End If
		Case "U6"
		    If (rs2.EOF And rs2.BOF) Then
				If strMsgCd = "" And strCardCoCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtCardCoCd_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtCardCoNm.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.txtCardCoCd.value = "<%=Trim(ConvSPChars(rs2(0)))%>"
					.txtCardCoNm.value = "<%=Trim(ConvSPChars(rs2(1)))%>"
				End With
				</Script>
		<%
		    End If
		Case "U9"
		    If (rs2.EOF And rs2.BOF) Then
				If strMsgCd = "" And strAcctCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtAcctCd_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtAcctNm.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
'					.txtAcctCd.value = "<%=Trim(ConvSPChars(rs2(0)))%>"
					.txtAcctNm.value = "<%=Trim(ConvSPChars(rs2(0)))%>"
				End With
				</Script>
		<%
		    End If
	End Select	

	Set rs2 = Nothing 		

	Select Case UCase(strOpenType)
		Case "AR","AP","PP","PR","SS"
		    If (rs3.EOF And rs3.BOF) Then
				If strMsgCd = "" And strBizCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBizCd_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtBizNm.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.txtBizCd.value = "<%=Trim(ConvSPChars(rs3(0)))%>"
					.txtBizNm.value = "<%=Trim(ConvSPChars(rs3(1)))%>"
				End With
				</Script>
		<%
		    End If
		Case "U6"			    
		    If (rs3.EOF And rs3.BOF) Then
				If strMsgCd = "" And strFrCardUserId <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtFrCardUserId_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtFrCardUserNm.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.txtFrCardUserId.value = "<%=Trim(ConvSPChars(rs3(0)))%>"
					.txtFrCardUserNm.value = "<%=Trim(ConvSPChars(rs3(1)))%>"
				End With
				</Script>
		<%
		    End If		    
		Case "U9"
			If UCase(Trim(skip_rs3)) = "TRUE" Then
			
			Else
			    If (rs3.EOF And rs3.BOF) Then
					If strMsgCd = "" And strMgntCd1 <> "" Then
						strMsgCd = "970000"		'Not Found
						strMsg1 = Request("txtMgntCd1_alt")
					End If
			%>
					<Script Language=vbScript>
					With parent
						.txtMgntCd1Nm.value = ""
					End With
					</Script>
			<%
			    Else
			%>
					<Script Language=vbScript>
					With parent
						.txtMgntCd1.value   = "<%=Trim(ConvSPChars(rs3(0)))%>"
						.txtMgntCd1Nm.value = "<%=Trim(ConvSPChars(rs3(1)))%>"
					End With
					</Script>
			<%
			    End If		
			End If	    
		Case Else
	End Select	

	Set rs3 = Nothing 
	
	Select Case UCase(strOpenType)
		Case "AR","AP"
		    If (rs4.EOF And rs4.BOF) Then
				If strMsgCd = "" And strBpCd2 <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBpCd2_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtBpNm2.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.txtBpCd2.value = "<%=Trim(ConvSPChars(rs4(0)))%>"
					.txtBpNm2.value = "<%=Trim(ConvSPChars(rs4(1)))%>"
				End With
				</Script>
		<%
		    End If
		Case "U6"			    
		    If (rs4.EOF And rs4.BOF) Then
				If strMsgCd = "" And strToCardUserId <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtToCardUserId_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtToCardUserNm.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.txtToCardUserId.value = "<%=Trim(ConvSPChars(rs4(0)))%>"
					.txtToCardUserNm.value = "<%=Trim(ConvSPChars(rs4(1)))%>"
				End With
				</Script>
		<%
		    End If		    
		Case "U9"
			If UCase(Trim(skip_rs4)) = "TRUE" Then
			
			Else		
			    If (rs4.EOF And rs4.BOF) Then
					If strMsgCd = "" And strMgntCd2 <> "" Then
						strMsgCd = "970000"		'Not Found
						strMsg1 = Request("txtMgntCd2_alt")
					End If
			%>
					<Script Language=vbScript>
					With parent
						.txtMgntCd2Nm.value = ""
					End With
					</Script>
			<%
			    Else
			%>
					<Script Language=vbScript>
					With parent
						.txtMgntCd2.value   = "<%=Trim(ConvSPChars(rs4(0)))%>"
						.txtMgntCd2Nm.value = "<%=Trim(ConvSPChars(rs4(1)))%>"
					End With
					</Script>
			<%
			    End If
			End If
		Case Else
		
	End Select		

	Set rs4 = Nothing 

	If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
		Response.End													'☜: 비지니스 로직 처리를 종료함
    End If

    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

	strOpenType     = Trim(Request("txtOpenType"))
	strFrOpenDt		= Trim(Request("txtFrOpenDt"))
	strToOpenDt		= Trim(Request("txtToOpenDt"))
	strDocCur		= Trim(Request("txtDocCur"))
	strOrgChangeId	= Trim(Request("txtOrgChangeId"))
	strDeptCd		= Trim(Request("txtDeptCd"))
	strBpCd			= Trim(Request("txtBpCd"))
	strBizCd		= Trim(Request("txtBizCd"))
	strBpCd2		= Trim(Request("txtBpCd2"))
	strProject		= Trim(Request("txtProject"))
	strAllcAmt		= Trim(Request("txtAllcAmt"))
	strRefNo		= Trim(Request("txtRefNo"))
	strAcctCd		= Trim(Request("txtAcctCd"))
	strGlNo			= Trim(Request("txtGlNo"))
	strMgntCd1		= Trim(Request("txtMgntCd1"))
	strMgntCd2		= Trim(Request("txtMgntCd2"))
	strCardCoCd		= Trim(Request("txtCardCoCd"))
	strCardNo		= Trim(Request("txtCardNo"))
	strFrCardUserId = Trim(Request("txtFrCardUserId"))
	strToCardUserId = Trim(Request("txtToCardUserId"))
	strParentGlNo   = Trim(Request("txtParentGLNo"))
	
	strFrDueDt		= Trim(Request("txtFrDueDt"))
	strToDueDt  	= Trim(Request("txtToDueDt"))
	
	if strToDueDt="" then strToDueDt="2999-12-31"
	
	strChkLocalCur  = Trim(Request("chkLocalCur"))

	' 권한관리 추가
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

	Select Case UCase(strOpenType)
		Case "AR"
			strCond = strCond & " AND A.ar_dt >= " & FilterVar(strFrOpenDt , "''", "S") & ""
			strCond = strCond & " AND A.ar_dt <= " & FilterVar(strToOpenDt , "''", "S") & ""
			If strChkLocalCur = "Y" Then 
				strCond = strCond & " AND A.doc_cur = " & FilterVar(gCurrency,"''","S") & ""
			Else
				strCond = strCond & " AND A.doc_cur = " & FilterVar(strDocCur,"''","S") & ""
			End If

			If strBpCd <> "" Then
				strCond = strCond & " AND A.pay_bp_cd = " & FilterVar(strBpCd , "''", "S") & ""
			End If			

			If strDeptCd <> "" Then
				strCond = strCond & " AND A.dept_cd = " & FilterVar(strDeptCd , "''", "S") & ""
				strCond = strCond & " AND A.org_change_id = " & FilterVar(strOrgChangeId , "''", "S") & ""
			End If

			If strBizCd <> "" Then
				strCond = strCond & " AND A.biz_area_cd = " & FilterVar(strBizCd , "''", "S") & ""
			End If

			If strBpCd2 <> "" Then
				strCond = strCond & " AND A.deal_bp_cd = " & FilterVar(strBpCd2 , "''", "S") & ""
			End If			

			If strProject <> "" Then
				strCond = strCond & " AND A.project_no like " & FilterVar("%" & strProject , "''", "S") & ""
			End If

			If strRefNo <> "" Then
				strCond = strCond & " AND A.ref_no like " & FilterVar("%" & strRefNo , "''", "S") & ""
			End If
			
			If strParentGlNo <> "" Then
				strCond = strCond & " AND A.gl_no <> " & FilterVar(strParentGlNo , "''", "S") & ""
			End If

				strCond = strCond & " AND A.ar_due_dt between " & FilterVar(strFrDueDt , "''", "S") & " and " & FilterVar(strToDueDt , "''", "S") & ""
				
			' 권한관리 추가
			If lgAuthBizAreaCd <> "" Then
				strCond = strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S") & ""
			End If

			If lgInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S") & ""
			End If

			If lgSubInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
			End If

			If lgAuthUsrID <> "" Then
				strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S") & ""
			End If
		Case "AP"
			strCond = strCond & " AND A.ap_dt >= " & FilterVar(strFrOpenDt , "''", "S") & ""
			strCond = strCond & " AND A.ap_dt <= " & FilterVar(strToOpenDt , "''", "S") & ""

			If strChkLocalCur = "Y" Then 
				strCond = strCond & " AND A.doc_cur = " & FilterVar(gCurrency,"''","S") & ""
			Else
				strCond = strCond & " AND A.doc_cur = " & FilterVar(strDocCur,"''","S") & ""
			End If

			If strBpCd <> "" Then
				strCond = strCond & " AND A.pay_bp_cd = " & FilterVar(strBpCd , "''", "S") & ""
			End If			

			If strDeptCd <> "" Then
				strCond = strCond & " AND A.dept_cd = " & FilterVar(strDeptCd , "''", "S") & ""
				strCond = strCond & " AND A.org_change_id = " & FilterVar(strOrgChangeId , "''", "S") & ""
			End If

			If strBizCd <> "" Then
				strCond = strCond & " AND A.biz_area_cd =" & FilterVar(strBizCd , "''", "S") & ""
			End If

			If strBpCd2 <> "" Then
				strCond = strCond & " AND A.deal_bp_cd = " & FilterVar(strBpCd2 , "''", "S") & ""
			End If			

			If strRefNo <> "" Then
				strCond = strCond & " AND A.ref_no like " & FilterVar("%" & strRefNo , "''", "S") & ""
			End If
			
			If strParentGlNo <> "" Then
				strCond = strCond & " AND A.gl_no <> " & FilterVar(strParentGlNo , "''", "S") & ""
			End If			
			strCond = strCond & " AND A.ap_due_dt between " & FilterVar(strFrDueDt , "''", "S") & " and " & FilterVar(strToDueDt , "''", "S") & ""
			
			' 권한관리 추가
			If lgAuthBizAreaCd <> "" Then
				strCond = strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S") & ""
			End If

			If lgInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S") & ""
			End If

			If lgSubInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S") & ""
			End If

			If lgAuthUsrID <> "" Then
				strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S") & ""
			End If			
		Case "PR"
			strCond = strCond & " AND A.prrcpt_dt >= " & FilterVar(strFrOpenDt , "''", "S") & ""
			strCond = strCond & " AND A.prrcpt_dt <= " & FilterVar(strToOpenDt , "''", "S") & ""

			If strChkLocalCur = "Y" Then 
				strCond = strCond & " AND A.doc_cur = " & FilterVar(gCurrency,"''","S") & ""
			Else
				strCond = strCond & " AND A.doc_cur = " & FilterVar(strDocCur,"''","S") & ""
			End If			
			
			If strBpCd <> "" Then
				strCond = strCond & " AND A.bp_cd = " & FilterVar(strBpCd , "''", "S") & ""
			End If						

			If strDeptCd <> "" Then
				strCond = strCond & " AND A.dept_cd = " & FilterVar(strDeptCd , "''", "S") & ""
				strCond = strCond & " AND A.org_change_id = " & FilterVar(strOrgChangeId , "''", "S") & ""
			End If

			If strBizCd <> "" Then
				strCond = strCond & " AND A.biz_area_cd =" & FilterVar(strBizCd , "''", "S") & ""
			End If

			If strProject <> "" Then
				strCond = strCond & " AND A.project_no like " & FilterVar("%" & strProject , "''", "S") & ""
			End If

			If strRefNo <> "" Then
				strCond = strCond & " AND A.ref_no like " & FilterVar("%" & strRefNo , "''", "S") & ""
			End If
			
			If strParentGlNo <> "" Then
				strCond = strCond & " AND A.gl_no <> " & FilterVar(strParentGlNo , "''", "S") & ""
			End If			
			
			' 권한관리 추가
			If lgAuthBizAreaCd <> "" Then
				strCond = strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S") & ""
			End If

			If lgInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S") & ""
			End If

			If lgSubInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S") & ""
			End If

			If lgAuthUsrID <> "" Then
				strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S") & ""
			End If			
		Case "PP"
			strCond = strCond & " AND A.prpaym_dt >= " & FilterVar(strFrOpenDt , "''", "S") & ""
			strCond = strCond & " AND A.prpaym_dt <= " & FilterVar(strToOpenDt , "''", "S") & ""
			
			If strChkLocalCur = "Y" Then 
				strCond = strCond & " AND A.doc_cur = " & FilterVar(gCurrency,"''","S") & ""
			Else
				strCond = strCond & " AND A.doc_cur = " & FilterVar(strDocCur,"''","S") & ""
			End If

			If strBpCd <> "" Then
				strCond = strCond & " AND A.bp_cd = " & FilterVar(strBpCd , "''", "S") & ""
			End If

			If strDeptCd <> "" Then
				strCond = strCond & " AND A.dept_cd = " & FilterVar(strDeptCd , "''", "S") & ""
				strCond = strCond & " AND A.org_change_id = " & FilterVar(strOrgChangeId , "''", "S") & ""
			End If

			If strBizCd <> "" Then
				strCond = strCond & " AND A.biz_area_cd =" & FilterVar(strBizCd , "''", "S") & ""
			End If

			If strRefNo <> "" Then
				strCond = strCond & " AND A.ref_no like " & FilterVar("%" & strRefNo , "''", "S") & ""
			End If
			
			If strParentGlNo <> "" Then
				strCond = strCond & " AND A.gl_no <> " & FilterVar(strParentGlNo , "''", "S") & ""
			End If			
			
			' 권한관리 추가
			If lgAuthBizAreaCd <> "" Then
				strCond = strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S") & ""
			End If

			If lgInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S") & ""
			End If

			If lgSubInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S") & ""
			End If

			If lgAuthUsrID <> "" Then
				strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S") & ""
			End If			
		Case "SS"
			strCond = strCond & " AND A.rcpt_dt >= " & FilterVar(strFrOpenDt , "''", "S") & ""
			strCond = strCond & " AND A.rcpt_dt <= " & FilterVar(strToOpenDt , "''", "S") & ""
			
			If strChkLocalCur = "Y" Then 
				strCond = strCond & " AND A.doc_cur = " & FilterVar(gCurrency,"''","S") & ""
			Else
				strCond = strCond & " AND A.doc_cur = " & FilterVar(strDocCur,"''","S") & ""
			End If			

			If strBpCd <> "" Then
				strCond = strCond & " AND A.bp_cd = " & FilterVar(strBpCd , "''", "S") & ""
			End If

			If strDeptCd <> "" Then
				strCond = strCond & " AND A.dept_cd = " & FilterVar(strDeptCd , "''", "S") & ""
				strCond = strCond & " AND A.org_change_id = " & FilterVar(strOrgChangeId , "''", "S") & ""
			End If

			If strBizCd <> "" Then
				strCond = strCond & " AND A.biz_area_cd like %" & FilterVar(strBizCd , "''", "S") & ""
			End If

			If strProject <> "" Then
				strCond = strCond & " AND A.project_no like  " & FilterVar("%" & strProject , "''", "S") & ""
			End If

			If strRefNo <> "" Then
				strCond = strCond & " AND A.ref_no like " & FilterVar("%" & strRefNo , "''", "S") & ""
			End If

			If strParentGlNo <> "" Then
				strCond = strCond & " AND A.gl_no <> " & FilterVar(strParentGlNo , "''", "S") & ""
			End If
			
			' 권한관리 추가
			If lgAuthBizAreaCd <> "" Then
				strCond = strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S") & ""
			End If

			If lgInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S") & ""
			End If

			If lgSubInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S") & ""
			End If

			If lgAuthUsrID <> "" Then
				strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S") & ""
			End If
		Case "U9"
			strCond = strCond & " AND A.gl_dt >= " & FilterVar(strFrOpenDt , "''", "S") & ""
			strCond = strCond & " AND A.gl_dt <= " & FilterVar(strToOpenDt , "''", "S") & ""
			
			If strChkLocalCur = "Y" Then 
				strCond = strCond & " AND A.doc_cur = " & FilterVar(gCurrency,"''","S") & ""
			Else
				strCond = strCond & " AND A.doc_cur = " & FilterVar(strDocCur,"''","S") & ""
			End If			

			If strDeptCd <> "" Then
				strCond = strCond & " AND f.dept_cd = " & FilterVar(strDeptCd , "''", "S") & ""
				strCond = strCond & " AND f.org_change_id = " & FilterVar(strOrgChangeId , "''", "S") & ""
			End If
					
			If strAcctCd <>  "" Then
			   strCond = strCond & "  And a.acct_cd = " & FilterVar(strAcctCd , "''", "S") & ""  
			End If 
	
			If strGlNo <> "" Then
				strCond = strCond & " And a.gl_no LIKE " & FilterVar(strGlNo & "%", "''", "S") & ""  
			End If
	
			If strMgntCd1 <> "" Then
				strCond = strCond & " And a.mgnt_val1 LIKE " & FilterVar(strMgntCd1 & "%", "''", "S") & ""  
			End If
	
			If strMgntCd2 <> "" Then
				strCond = strCond & " And a.mgnt_val2 LIKE " & FilterVar(strMgntCd2 & "%", "''", "S") & ""  
			End If

			If strParentGlNo <> "" Then
				strCond = strCond & " AND A.gl_no <> " & FilterVar(strParentGlNo , "''", "S") & ""
			End If
			strCond = strCond & " AND A.due_dt between " & FilterVar(strFrDueDt , "''", "S") & " and " & FilterVar(strToDueDt , "''", "S") & ""
			
			' 권한관리 추가
			If lgAuthBizAreaCd <> "" Then
				strCond = strCond & " AND f.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S") & ""
			End If

			If lgInternalCd <> "" Then
				strCond = strCond & " AND f.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S") & ""
			End If

			If lgSubInternalCd <> "" Then
				strCond = strCond & " AND f.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S") & ""
			End If

			If lgAuthUsrID <> "" Then
				strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S") & ""
			End If					
		Case "U6"
			strCond = strCond & " AND A.gl_dt >= " & FilterVar(strFrOpenDt , "''", "S") & ""
			strCond = strCond & " AND A.gl_dt <= " & FilterVar(strToOpenDt , "''", "S") & ""
			
			If strChkLocalCur = "Y" Then 
				strCond = strCond & " AND A.doc_cur = " & FilterVar(gCurrency,"''","S") & ""
			Else
				strCond = strCond & " AND A.doc_cur = " & FilterVar(strDocCur,"''","S") & ""
			End If			

			If strDeptCd <> "" Then
				strCond = strCond & " AND f.dept_cd = " & FilterVar(strDeptCd , "''", "S") & ""
				strCond = strCond & " AND f.org_change_id = " & FilterVar(strOrgChangeId , "''", "S") & ""
			End If		
		
			If strCardCoCd <> "" Then	
				strCond = strCond & " AND c.card_co_cd  =  " & FilterVar(strCardCoCd , "''", "S") & ""
			End If
				
			If strCardNo <> "" Then		
				strCond = strCond & " AND c.credit_no =  " & FilterVar(strCardNo , "''", "S") & ""
			End If	
			
			If strFrCardUserId <> "" Then		
				strCond = strCond & " AND c.bp_cd  >=  " & FilterVar(strFrCardUserId , "''", "S") & ""
			End If				
			
			If strToCardUserId <> "" Then		
				strCond = strCond & " AND c.bp_cd  <=  " & FilterVar(strToCardUserId , "''", "S") & ""
			End If							

			If strParentGlNo <> "" Then
				strCond = strCond & " AND A.gl_no <> " & FilterVar(strParentGlNo , "''", "S") & ""
			End If
			strCond = strCond & " AND A.due_dt between " & FilterVar(strFrDueDt , "''", "S") & " and " & FilterVar(strToDueDt , "''", "S") & ""
			
			' 권한관리 추가
			If lgAuthBizAreaCd <> "" Then
				strCond = strCond & " AND f.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S") & ""
			End If

			If lgInternalCd <> "" Then
				strCond = strCond & " AND f.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S") & ""
			End If

			If lgSubInternalCd <> "" Then
				strCond = strCond & " AND f.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S") & ""
			End If

			If lgAuthUsrID <> "" Then
				strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S") & ""
			End If					
	End Select
End Sub





%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			Parent.htxtOpenType.value = Parent.cboOpenType.value
			Parent.htxtFrOpenDt.value = Parent.txtFrOpenDt.Text
			Parent.htxtToOpenDt.value = Parent.txtToOpenDt.Text
			Parent.htxtDocCur.value   = Parent.txtDocCur.value
'			Parent.hOrgChangeId.value = Parent.hOrgChangeId.value
			Parent.htxtDeptCd.value   = Parent.txtDeptCd.value
			Parent.htxtBpCd.value     = Parent.txtBpCd.value
			Parent.htxtBizCd.value    = Parent.txtBizCd.value
			Parent.htxtBpCd2.value    = Parent.txtBpCd2.value
			Parent.htxtProject.value  = Parent.txtProject.value
			Parent.htxtAllcAmt.value  = Parent.txtAllcAmt.Text
			Parent.htxtRefNo.value    = Parent.txtRefNo.value
			Parent.htxtAcctCd.value   = Parent.txtAcctCd.value
			Parent.htxtGlNo.value     = Parent.txtGlNo.value
			Parent.htxtMgntCd1.value  = Parent.txtMgntCd1.value
			Parent.htxtMgntCd2.value  = Parent.txtMgntCd2.value
			Parent.htxtCardCoCd.value = Parent.txtCardCoCd.value
			Parent.htxtCardNo.value   = Parent.txtCardNo.value
			Parent.hChkLocalCur.value = Parent.ChkLocalCur.value
       End If

       'Show multi spreadsheet data from this line
       
		Parent.ggoSpread.Source		= Parent.vspdData
		Parent.vspdData.Redraw = False
		Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data

		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",10),"A", "I" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",11),"A", "I" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",12),"A", "I" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",22),"A", "I" ,"X","X")
		Parent.vspdData.Redraw = True
		Parent.lgPageNo				=  "<%=lgPageNo%>"               '☜ : Next next data tag
			       
		Parent.DbQueryOk
    End If   

</Script>	

<%
%>

