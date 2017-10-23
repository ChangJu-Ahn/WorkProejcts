<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3104mb2
'*  4. Program Name         : 예적금입출내역조회 
'*  5. Program Desc         : Query of Deposit Income/Outgo
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.12
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  사업장코드, 은행코드 오류 Check
'=======================================================================================================


%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strBankCd																'⊙ : 은행 
Dim strBankAcctNo															'⊙ : 계좌번호 
Dim strDateFr, strDateTo													'⊙ : 입출일자 
Dim PreAmt, PreLocAmt, RcptAmt, RcptLocAmt, PaymAmt, PaymLocAmt, BalAmt, BalLocAmt
Dim	strWhere

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
    Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")
    Call HideStatusWnd 

    lgPageNo			= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    lgStrPrevKey		= Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount			= 100
    lgSelectList		= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList			= Request("lgTailList")                                 '☜ : Orderby value
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
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            lgStrPrevKey = rs0(5)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                   '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
        lgStrPrevKey = ""
    End If
  	
'	rs0.Close
'    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	Redim UNIValue(2,5)

	UNISqlId(0) = "F3104MA102"
	UNISqlId(1) = "F3104MA103"	'이월금액 
	UNISqlId(2) = "F3104MA103"	'입출합계 

	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	UNIValue(0,0) = lgSelectList                                          '☜: Select list
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,1) = FilterVar(strBankCd , "''", "S") 
	UNIValue(0,2) = FilterVar(strBankAcctNo , "''", "S") 
	UNIValue(0,3) = " and A.trans_dt >= " & FilterVar(strDateFr , "''", "S") & " and A.trans_dt <= " & FilterVar(strDateTo , "''", "S") & " "
	UNIValue(0,3) = UNIValue(0,3) & strWhere
		    
	UNIValue(1,0) = FilterVar(strBankCd , "''", "S") 
	UNIValue(1,1) = FilterVar(strBankAcctNo , "''", "S") 
	UNIValue(1,2) = " and A.trans_dt < " & FilterVar(strDateFr , "''", "S") & " "
	UNIValue(1,2) = UNIValue(1,2) & strWhere

	UNIValue(2,0) = FilterVar(strBankCd , "''", "S") 
	UNIValue(2,1) = FilterVar(strBankAcctNo , "''", "S") 
	UNIValue(2,2) = " and A.trans_dt >= " & FilterVar(strDateFr , "''", "S") & " and A.trans_dt <= " & FilterVar(strDateTo , "''", "S") & " "
	UNIValue(2,2) = UNIValue(2,2) & strWhere

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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
	
	PreAmt     = 0
	PreLocAmt  = 0
	RcptAmt    = 0
	RcptLocAmt = 0
	PaymAmt    = 0
	PaymLocAmt = 0
    
	If Not(rs1.EOF And rs1.BOF) Then
        If Not IsNull(rs1(0)) Then RcptAmt    = rs1(0)
        If Not IsNull(rs1(1)) Then PaymAmt    = rs1(1)
        If Not IsNull(rs1(2)) Then RcptLocAmt = rs1(2)
        If Not IsNull(rs1(3)) Then PaymLocAmt = rs1(3)
	End If
	
	PreAmt    = CCur(RcptAmt)    - CCur(PaymAmt)
	PreLocAmt = CCur(RcptLocAmt) - CCur(PaymLocAmt)
	RcptAmt    = 0
	RcptLocAmt = 0
	PaymAmt    = 0
	PaymLocAmt = 0
	
	rs1.Close
	Set rs1 = Nothing
	
	If Not(rs2.EOF And rs2.BOF) Then
        If Not IsNull(rs2(0)) Then RcptAmt    = rs2(0)
        If Not IsNull(rs2(1)) Then PaymAmt    = rs2(1)
        If Not IsNull(rs2(2)) Then RcptLocAmt = rs2(2)
        If Not IsNull(rs2(3)) Then PaymLocAmt = rs2(3)
	End If
	
	BalAmt     = CCur(PreAmt)    + CCur(RcptAmt)    - CCur(PaymAmt)
	BalLocAmt  = CCur(PreLocAmt) + CCur(RcptLocAmt) - CCur(PaymLocAmt)

	rs2.Close
	Set rs2 = Nothing

    If rs0.EOF And rs0.BOF Then
'		Call DisplayMsgBox("140600", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set lgADF = Nothing                                             '☜: ActiveX Data Factory Object Nothing
		Exit Sub
'		Response.End													'☜: 비지니스 로직 처리를 종료함 
		
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     strBankCd			= Request("txtBankCd")
     strBankAcctNo		= Request("txtBankAcctNo")
     strDateFr			= UniConvDate(Request("txtDateFr"))
     strDateTo			= UniConvDate(Request("txtDateTo"))

	strWhere = ""

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	' 권한관리 추가 
	strWhere	= strWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL



    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>


    With parent
         .ggoSpread.Source    = .frm1.vspdData2 

		.frm1.vspdData2.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
		.lgStrPrevKey_B       = "<%=ConvSPChars(lgStrPrevKey)%>"     '☜ :  set next data tag
		.lgPageNo_B           =  "<%=lgPageNo%>"                     '☜ : Next next data tag
		Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData2,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.Frm1.txtDocCur1.Value ,parent.GetKeyPos("B",1),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData2,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.Frm1.txtDocCur1.Value ,parent.GetKeyPos("B",2),   "A" ,"I","X","X")
		.frm1.vspdData2.Redraw = True

	End with
    With parent.frm1
		.txtPreAmt.Text     = "<%=UNINumClientFormat(PreAmt    , ggAmtOfMoney.DecPoint, 0)%>"
		.txtPreLocAmt.Text  = "<%=UNINumClientFormat(PreLocAmt , ggAmtOfMoney.DecPoint, 0)%>"
		.txtRcptAmt.Text    = "<%=UNINumClientFormat(RcptAmt   , ggAmtOfMoney.DecPoint, 0)%>"
		.txtRcptLocAmt.Text = "<%=UNINumClientFormat(RcptLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.txtPaymAmt.Text    = "<%=UNINumClientFormat(PaymAmt   , ggAmtOfMoney.DecPoint, 0)%>"
		.txtPaymLocAmt.Text = "<%=UNINumClientFormat(PaymLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.txtBalAmt.Text     = "<%=UNINumClientFormat(BalAmt    , ggAmtOfMoney.DecPoint, 0)%>"
		.txtBalLocAmt.Text  = "<%=UNINumClientFormat(BalLocAmt , ggAmtOfMoney.DecPoint, 0)%>"
	End With
	Call parent.DbQueryOk2
</Script>
