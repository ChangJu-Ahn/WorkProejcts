<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3102mb1
'*  4. Program Name         : 예적금조회 
'*  5. Program Desc         : Query of Deposit
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.11
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2003/06/12 Oh, Soo Min (MA의 C_SHEETMAXROWS_D 변수 삭제)
'=======================================================================================================


%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6                            '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgStrPrevKey
Dim lgTailList                                                              '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strBankCd, strBankAcctNo, strDateFr, strDateTo, strDocCur
Dim PreAmt, PreLocAmt, RcptAmt, RcptLocAmt, PaymAmt, PaymLocAmt
Dim strMsgCd, strMsg1, strMsg2
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1
Dim strWhere

Dim  iLoopCount
Dim  LngMaxRow

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q","F","NOCOOKIE","QB")
    Call HideStatusWnd 

    lgPageNo			= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    lgStrPrevKey		= Request("lgStrPrevKey")
    lgMaxCount			= 100
    lgSelectList		= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList			= Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist			= "No"
	LngMaxRow			= CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    	
    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(lgMaxCount) * CDbl(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
     
    'rs0에 대한 결과 
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
'			Call ServerMesgBox(lgstrData, vbInformation, I_MKSCRIPT)
            
        Else
			Call ServerMesgBox("lgPageNo : " & lgPageNo, vbInformation, I_MKSCRIPT)
        
            lgPageNo = lgPageNo + 1
            lgStrPrevKey = rs0(8)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
        lgStrPrevKey = ""
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
                    '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(6)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(6,6)

    UNISqlId(0) = "F3102MA101KO441"	'예적금조회 
    UNISqlId(1) = "F3102MA102KO441"	'은행명 
    UNISqlId(2) = "F3102MA103KO441"	'이월금액 
    UNISqlId(3) = "F3102MA104KO441"	'입출합계 
    UNISqlId(4) = "F3102MA105KO441"	'계좌번호 
	UNISqlId(5) = "A_GETBIZ"
    UNISqlId(6) = "A_GETBIZ"
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,1) = FilterVar(strBankCd, "''", "S") 
	UNIValue(0,2) = FilterVar(strBankAcctNo, "''", "S")
	UNIValue(0,3) = FilterVar(strDateFr, "''", "S")
	UNIValue(0,4) = FilterVar(strDateTo, "''", "S")
	UNIValue(0,5) = strWhere

	UNIValue(1,0) = FilterVar(strBankCd, "''", "S")

	UNIValue(2,0) = FilterVar(strBankCd, "''", "S")
	UNIValue(2,1) = FilterVar(strBankAcctNo, "''", "S")
	UNIValue(2,2) = FilterVar(strDateFr, "''", "S") 
	UNIValue(2,3) = strWhere

	UNIValue(3,0) = FilterVar(strBankCd, "''", "S") 
	UNIValue(3,1) = FilterVar(strBankAcctNo, "''", "S") 
	UNIValue(3,2) = FilterVar(strDateFr, "''", "S") 
	UNIValue(3,3) = FilterVar(strDateTo, "''", "S") 
	UNIValue(3,4) = strWhere
	
	UNIValue(4,0) = FilterVar(strBankCd, "''", "S") 
	UNIValue(4,1) = FilterVar(strBankAcctNo, "''", "S") 
	
	UNIValue(5,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(6,0)  = FilterVar(strBizAreaCd1, "''", "S")
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6 )
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strBankCd <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBankCd_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
				.txtBankCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
				.txtBankNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
			End With
		</Script>
<%	
	End If
	
	rs1.Close
	Set rs1 = Nothing
	
	If (rs4.EOF And rs4.BOF) Then
		If strMsgCd = "" And strBankAcctNo <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBankAcctNo_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
				.txtBankAcctNo.value = "<%=ConvSPChars(Trim(rs4(0)))%>"
			End With
		</Script>
<%	
	End If
	
	rs4.Close
	Set rs4 = Nothing
	
	If (rs5.EOF And rs5.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd.value = "<%=Trim(rs5(0))%>"
		.frm1.txtBizAreaNm.value = "<%=Trim(rs5(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs5.Close
	Set rs5 = Nothing
	
	
	If (rs6.EOF And rs6.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT1")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd1.value = "<%=Trim(rs6(0))%>"
		.frm1.txtBizAreaNm1.value = "<%=Trim(rs6(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs6.Close
	Set rs6 = Nothing
	
	
	
	PreAmt     = 0
	PreLocAmt  = 0
	RcptAmt    = 0
	RcptLocAmt = 0
	PaymAmt    = 0
	PaymLocAmt = 0
	
	If Not(rs2.EOF And rs2.BOF) Then
		If IsNull(rs2(0)) = False Then PreAmt    = rs2(0)
		If IsNull(rs2(1)) = False Then PreLocAmt = rs2(1)
	End If
	
	rs2.Close
	Set rs2 = Nothing
	
	If Not(rs3.EOF And rs3.BOF) Then
		If IsNull(rs3(0)) = False Then RcptAmt    = rs3(0)
		If IsNull(rs3(1)) = False Then RcptLocAmt = rs3(1)
		If IsNull(rs3(2)) = False Then PaymAmt    = rs3(2)
		If IsNull(rs3(3)) = False Then PaymLocAmt = rs3(3)
	End If

	rs3.Close
	Set rs3 = Nothing

    If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
'		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If

	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strBankCd		= UCase(Trim(Request("txtBankCd")))
    strBankAcctNo	= UCase(Trim(Request("txtBankAcctNo")))
    strDateFr		= UniConvDate(Request("txtDateFr"))
    strDateTo		= UniConvDate(Request("txtDateTo"))

    If Trim(Request("txtDocCur")) = "" Then
		strDocCur = ""
	Else
		strDocCur = " AND B.DOC_CUR = " & Filtervar(UCase(Trim(Request("txtDocCur"))), "''", "S")
	End If
	
	strWhere	= strDocCur

	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))            '사업장From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))            '사업장To
	
	if strBizAreaCd <> "" then
		strWhere = strWhere & " AND B.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere = strWhere & " AND B.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	end if


	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND B.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	' 권한관리 추가 
	strWhere	= strWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL


	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
With parent
	
	If "<%=lgDataExist%>" = "Yes" Then
	   If "<%=strDocCur%>" <> "" Then
		.frm1.txtPreAmt.Text     = "<%=UNINumClientFormat(PreAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtRcptAmt.Text    = "<%=UNINumClientFormat(RcptAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtPaymAmt.Text    = "<%=UNINumClientFormat(PaymAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtBalAmt.Text     = "<%=UNINumClientFormat(Cdbl(PreAmt) + Cdbl(RcptAmt) - Cdbl(PaymAmt), ggAmtOfMoney.DecPoint, 0)%>"	
	   Else 
	    .frm1.txtPreAmt.Text     = "<%=UNINumClientFormat(PreAmt, 2, 0)%>"
		.frm1.txtRcptAmt.Text    = "<%=UNINumClientFormat(RcptAmt, 2, 0)%>"
		.frm1.txtPaymAmt.Text    = "<%=UNINumClientFormat(PaymAmt, 2, 0)%>"
		.frm1.txtBalAmt.Text     = "<%=UNINumClientFormat(Cdbl(PreAmt) + Cdbl(RcptAmt) - Cdbl(PaymAmt), 2, 0)%>"	
	   End If	  	

		.frm1.txtPreLocAmt.Text  = "<%=UNINumClientFormat(PreLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtRcptLocAmt.Text = "<%=UNINumClientFormat(RcptLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtPaymLocAmt.Text = "<%=UNINumClientFormat(PaymLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtBalLocAmt.Text  = "<%=UNINumClientFormat(Cdbl(PreLocAmt) + Cdbl(RcptLocAmt) - Cdbl(PaymLocAmt), ggAmtOfMoney.DecPoint, 0)%>"

        .ggoSpread.Source    = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		.lgStrPrevKey        = "<%=ConvSPChars(lgStrPrevKey)%>"                       '☜: set next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",4),parent.GetKeyPos("A",5),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",4),parent.GetKeyPos("A",7),   "A" ,"I","X","X")
	'	Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1 , -1 ,parent.GetKeyPos("A",4),parent.GetKeyPos("A",5),   "A" ,"Q","X","X")
	'	Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1 , -1, parent.GetKeyPos("A",4),parent.GetKeyPos("A",7),   "A" ,"Q","X","X")
		.frm1.vspdData.Redraw = True
		.DbQueryOk()
	End If
End with
	
</Script>	


