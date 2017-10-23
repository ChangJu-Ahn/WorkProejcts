<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")    

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2        '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim txtCostCd																'⊙ : 코스트센터코드 
Dim txtCondCostCd															'⊙ : 조건부 코스트센터코드 
Dim txtYYYYMM																'⊙ : 년월 
Dim txtAcctCd
Dim txtDiFlag																'⊙ : 계정코드 
Dim txtTotWorkinAmt															'⊙ : 제공배부합계액 
Dim txtTotItemAmt															'⊙ : 근거금액(수량)



'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgPageNo       = Trim(Request("lgPageNo"))                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = Trim(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    
    
    txtCostCd = Trim(Request("txtCostCd"))
    txtYYYYMM = Trim(Request("txtYYYYMM"))
    txtAcctCd = Trim(Request("txtAcctCd"))
    txtDiFlag = Trim(Request("txtDiFlag"))
    
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr

   	Const C_SHEETMAXROWS_D  = 100 
   	
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
	
    lgstrData = ""
    lgDataExist    = "Yes"

	
    If UniConvNumStringToDouble(lgPageNo,0) > 0 Then
       rs0.Move     = UniConvNumStringToDouble(lgMaxCount,0) * UniConvNumStringToDouble(lgPageNo,0)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
	
    Do while Not (rs0.EOF Or rs0.BOF)
        
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
		
        If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = Cstr(UniConvNumStringToDouble(lgPageNo,0) + 1)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
    
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim strWhere
	
    Redim UNIValue(2,2)

    UNISqlId(0) = "C3606MA102"
    UNISqlId(1) = "C3606MA104"
    UNISqlId(2) = "C3606MA105"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strWhere = ""
    strWhere = strWhere & " AND f.yyyymm = " & FilterVar(txtYYYYMM ,"''"	,"S")
    
    IF txtCostCd <> "" Then
	    strWhere = strWhere & " AND b.cost_cd = " & FilterVar(txtCostCd ,"''"	,"S")
	END If
	
	IF txtAcctCd <> "" Then
	    strWhere = strWhere & " AND a.acct_cd = " & FilterVar(txtAcctCd ,"''"	,"S")
	END IF
    
    IF txtDiFlag <> "" Then
		strWhere = strWhere & " AND f.di_flag = " & FilterVar(txtDiFlag ,"''"	,"S")
	END If

    
    UNIValue(0,1) = strWhere
    UNIValue(1,0) = strWhere
    UNIValue(2,0) = strWhere
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("233600", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()

    End If

		'rs1'⊙ : 제공배부합계액 
		If Not (rs1.EOF OR rs1.BOF) Then
			txtTotWorkinAmt = Trim(rs1(0))
		Else
			txtTotWorkinAmt = ""
		End IF
		rs1.Close
		Set rs1 = Nothing

		'rs2'⊙ : 근거금액(수량)
		If Not (rs2.EOF OR rs2.BOF) Then
			txtTotItemAmt = Trim(rs2(0))
		Else
			txtTotItemAmt = ""
		End IF
		rs2.Close
		Set rs2 = Nothing
		
	
End Sub

%>

<Script Language=vbscript>
With Parent

	.frm1.txtTotWorkinAmt.text = "<%=UNINumClientFormat(txtTotWorkinAmt,ggAmtOfMoney.Decpoint,0)%>"
	.frm1.txtTotItemAmt.text = "<%=UNINumClientFormat(txtTotItemAmt,ggAmtOfMoney.Decpoint,0)%>"

    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
       End If


		.ggoSpread.Source  = .frm1.vspdData2
		.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
		.lgPageNo_B      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		.DbQueryOk("2")
	End If

	
End With
</Script>
