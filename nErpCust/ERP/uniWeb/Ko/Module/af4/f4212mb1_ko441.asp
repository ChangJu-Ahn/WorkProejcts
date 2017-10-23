
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4212mb1
'*  4. Program Name         : 차입금현황조회 
'*  5. Program Desc         : Query of Loan State
'*  6. Comproxy List        : DB AGENT
'*  7. Modified date(First) : 2002.04.17
'*  8. Modified date(Last)  : 2003.05.19
'*  9. Modifier (First)     : Park, Joon Won
'* 10. Modifier (Last)      : Ahn do hyun
'* 11. Comment              :
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("Q", "A","NOCOOKIE","QB")
%>	                                                                      
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 


On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5     '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Const C_SHEETMAXROWS_D = 100

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strLoanFg, cboConfFg, cboApSts, txtDocCur
Dim txtLoanPlcFg, txtLoanPlcCd, txtLoanPlcNm
Dim txtLoanType, txtLoanTypeNm
Dim strLoanDtFr
Dim strLoanDtTo
'Dim strIntDtFr
'Dim strIntDtTo
Dim strPaymDtFr
Dim strPaymDtTo
Dim strWhere1                                                               '⊙ : Where 조건 
Dim strMsgCd, strMsg1, strMsg2
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1
'DIM COST_NM
DIM LoanSum, LoanLocSum                                                                 '⊙ : 차입금액합 
DIM IntSum, IntLocSum                                                                 '⊙ : 지급이자액합 
DIM RdpSum, RdpLocSum                                                                  '⊙ : 상환금액합 
DIM BalSum, BalLocSum                                                                  '⊙ : 차입금잔액합 

Dim  iLoopCount
Dim  LngMaxRow

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL
 
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
 
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

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
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
      
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	
    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(5,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "F4212MA1"
    UNISqlId(1) = "F4212MA2"					'SUM
    If txtLoanPlcFg = "BK" Then
		UNISQLID(2) = "ABANKNM"
	Else
		UNISQLID(2) = "ABPNM"
	End If
	
	UNISQLID(3) = "AMINORNM"
	UNISqlId(4) = "A_GETBIZ"
    UNISqlId(5) = "A_GETBIZ"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	
	UNIValue(0,1)  = UCase(Trim(strWhere1))
	UNIValue(1,0)  = UCase(Trim(strWhere1))	'rs1에 대한 Value값 setting(총계)

	If txtLoanPlcFg = "BK" Then
		UNIValue(2,0) = FilterVar(txtLoanPlcCd , "''", "S")
	Else
		UNIValue(2,0) = FilterVar(txtLoanPlcCd , "''", "S")
	End If

	UNIValue(3,0) = FilterVar("F1000",  ""  ,"S")
	UNIValue(3,1) = FilterVar(txtLoanType ,  ""  ,"S")
	UNIValue(4,0) = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(5,0) = FilterVar(strBizAreaCd1, "''", "S")
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF
                                                                 '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
	iStr = Split(lgstrRetMsg,gColSep)
   
    If iStr(0) <> "0" Then
        Call ServerMesgBox(iStr(1), vbInformation, I_MKSCRIPT)
    End If 
    
    
     'rs2에 대한 결과 
    If txtLoanPlcCd <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
			txtLoanPlcNm = Trim(rs2(1))
		Else
			txtLoanPlcNm = ""
			If txtLoanPlcFg = "BK" Then
				Call DisplayMsgBox("800123", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
			Else
				Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
			End If
	        rs2.Close
		    Set rs2 = Nothing
			Exit sub
		End IF
	rs2.Close
	Set rs2 = Nothing
	End If
	
	'rs3
	If txtLoanType <> "" Then
	    If Not (rs3.EOF OR rs3.BOF) Then
			txtLoanTypeNm = Trim(rs3(1))
		Else
			txtLoanTypeNm = ""
			Call DisplayMsgBox("140936", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs3.Close
		    Set rs3 = Nothing 
			Exit sub
		End IF
		rs3.Close
		Set rs3 = Nothing
	End If
    
    'rs4
    If Not( rs4.EOF OR rs4.BOF) Then
   		strBizAreaCd = Trim(rs4(0))
		strBizAreaNm = Trim(rs4(1))
	Else
		strBizAreaCd = ""
		strBizAreaNm = ""
		
    End IF
    
    rs4.Close
    Set rs4 = Nothing
    
    ' rs5
    If Not( rs5.EOF OR rs5.BOF) Then
   		strBizAreaCd1 = Trim(rs5(0))
		strBizAreaNm1 = Trim(rs5(1))
	Else
		strBizAreaCd1 = ""
		strBizAreaNm1 = ""
		
    End IF
    
    rs5.Close
    Set rs5 = Nothing
    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("140900", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	%>
		<Script Language=vbscript>
			parent.frm1.txtLoanDtFr.focus
		</Script>
	<%
        rs0.Close
        Set rs0 = Nothing 
        Exit Sub
    Else    

        Call  MakeSpreadSheetData()
		'rs1에 대한 결과 
		IF NOT (rs1.EOF or rs1.BOF) then	'0,1,2,3 :		'4,5,6,7 : local amount
			LoanSum = rs1(0)                                    			' SUM(A.LOAN_AMT)
		    IntSum  = rs1(1)                                                 ' SUM(A.INT_PAY_AMT)
		    RdpSum  = rs1(2)                                                 ' SUM(A.RDP_AMT)
		    BalSum  = rs1(3) 
			LoanLocSum = rs1(4)                                    			' SUM(A.LOAN_AMT)
		    IntLocSum  = rs1(5)                                                 ' SUM(A.INT_PAY_AMT)
		    RdpLocSum  = rs1(6)                                                 ' SUM(A.RDP_AMT)
		    BalLocSum  = rs1(7) 
		                                           ' SUM(A.LOAN_BAL_AMT)	
		ELSE
		    LoanSum = ""
		    LoanLocSum = ""
			IntSum = ""
			IntLocSum = ""
			RdpSum = ""
			RdpLocSum = ""
			BalSum = ""
			BalLocSum = ""
		End if
		rs1.Close
		Set rs1 = Nothing 

    End If
  
End Sub


'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strLoanFg		= UCase(Trim(Request("cboLoanFg")))
	cboConfFg		= Trim(Request("cboConfFg"))
	cboApSts		= Trim(Request("cboApSts"))
	txtDocCur		= UCase(Trim(Request("txtDocCur")))
    txtLoanPlcFg    = UCase(Trim(Request("txtLoanPlcFg")))
    txtLoanPlcCd    = UCase(Trim(Request("txtLoanPlcCd")))
	txtLoanType		= UCase(Trim(Request("txtLoanType")))	
	strLoanDtFr		= UNIConvDate(Trim(Request("txtLoanDtFr")))
	strLoanDtTo		= UNIConvDate(Trim(Request("txtLoanDtTo")))
	strPaymDtFr		= Request("txtPaymDtFr")
	strPaymDtTo	    = Request("txtPaymDtTo")
	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))            '사업장From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))            '사업장To

    strWhere1 = ""
	strWhere1 = strWhere1 & " LN.loan_dt between  " & FilterVar( strLoanDtFr, "''", "S") & " and  " & FilterVar(strLoanDtTo, "''", "S") & " "

	If strPaymDtFr <> "" Then strWhere1 = strWhere1 & " and LN.due_dt >=  " & FilterVar(UNIConvDate(strPaymDtFr), "''", "S") & " "
	If strPaymDtTo <> "" Then strWhere1 = strWhere1 & " and LN.due_dt <=  " & FilterVar(UNIConvDate(strPaymDtTo), "''", "S") & " "
	If txtDocCur   <> "" Then strWhere1 = strWhere1 & " and LN.Doc_Cur   = " & FilterVar(txtDocCur ,"''"	,"S")
	If strLoanFg   <> "" Then strWhere1 = strWhere1 & " and LN.loan_fg   = " & FilterVar(strLoanFg ,"''"	,"S")
	If txtLoanType <> "" Then strWhere1 = strWhere1 & " and LN.loan_type = " & FilterVar(txtLoanType ,"''"	,"S")
	If txtLoanPlcFg <> "" Then strWhere1 = strWhere1 & " and LN.Loan_Plc_Type   =  " & FilterVar(txtLoanPlcFg , "''", "S") & " " 
	If txtLoanPlcCd <> "" Then
		If txtLoanPlcFg = "BK" Then
			strWhere1 = strWhere1 & " and LN.Loan_bank_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")
		Else
			strWhere1 = strWhere1 & " and LN.bp_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")
		End If
	End If
	If cboConfFg	= "C" Then	strWhere1 = strWhere1 & " and LN.conf_fg IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " ) " 
	If cboConfFg	= "U" Then	strWhere1 = strWhere1 & " and LN.conf_fg   =  " & FilterVar(cboConfFg , "''", "S") & " " 
	If cboApSts		<> "" Then	strWhere1 = strWhere1 & " and LN.rdp_cls_fg   =  " & FilterVar(cboApSts , "''", "S") & " " 
	
	if strBizAreaCd <> "" then
		strWhere1 = strWhere1 & " AND LN.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere1 = strWhere1 & " AND LN.BIZ_AREA_CD >= " & FilterVar("", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere1 = strWhere1 & " AND LN.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere1 = strWhere1 & " AND LN.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND LN.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND LN.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND LN.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND LN.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	strWhere1	= strWhere1	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

End Sub

%>

<Script Language=vbscript>
With parent
	If "<%=lgDataExist%>" = "Yes" Then
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.hConfFg.value		= "<%=ConvSPChars(cboConfFg)%>" 
			.frm1.hApSts.value		= "<%=ConvSPChars(cboApSts)%>" 
			.Frm1.hLoanFg.Value		= "<%=ConvSPChars(strLoanFg)%>"                  'For Next Search
			.Frm1.hLoanPlcFg.Value	= "<%=ConvSPChars(txtLoanPlcFg)%>"                  'For Next Search
			.Frm1.hLoanPlcCd.Value	= "<%=ConvSPChars(txtLoanPlcCd)%>"                  'For Next Search
			.Frm1.hDocCur.Value		= "<%=ConvSPChars(txtDocCur)%>"                  'For Next Search
			.Frm1.hLoanType.Value	= "<%=ConvSPChars(txtLoanType)%>"                  'For Next Search
			.Frm1.hLoanDtFr.Value	= "<%=strLoanDtFr%>"                  'For Next Search
			.Frm1.hLoanDtTo.Value	= "<%=strLoanDtTo%>"                  'For Next Search
			.Frm1.hPaymDtFr.Value	= "<%=strPaymDtFr%>"                  'For Next Search
			.Frm1.hPaymDtTo.Value	= "<%=strPaymDtTo%>"                  'For Next Search
			.Frm1.htxtBizAreaCd.value = Trim(.Frm1.txtBizAreaCd.value)
			.Frm1.htxtBizAreaCd1.value = Trim(.Frm1.txtBizAreaCd1.value)
		End If
		'If "<%=txtDocCur%>"   <> "" Then
		'	.frm1.txtLoan.Text ="<%=UNINumClientFormat(LoanSum,ggAmtOfMoney.DecPoint, 0)%>"
		'	.frm1.txtInt.Text ="<%=UNINumClientFormat(IntSum,ggAmtOfMoney.DecPoint, 0)%>"
		'	.frm1.txtRdp.Text ="<%=UNINumClientFormat(RdpSum,ggAmtOfMoney.DecPoint, 0)%>"
		'	.frm1.txtBal.Text ="<%=UNINumClientFormat(BalSum,ggAmtOfMoney.DecPoint, 0)%>"
		'Else
			.frm1.txtLoan.Text = "<%=UNINumClientFormat(LoanSum,2, 0)%>"
			.frm1.txtInt.Text = "<%=UNINumClientFormat(IntSum,2, 0)%>"
			.frm1.txtRdp.Text = "<%=UNINumClientFormat(RdpSum,2, 0)%>"
			.frm1.txtBal.Text = "<%=UNINumClientFormat(BalSum,2, 0)%>"
		'End If
		.frm1.txtLoanLoc.Text = "<%=UNINumClientFormat(LoanLocSum,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtIntLoc.Text = "<%=UNINumClientFormat(IntLocSum,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtRdpLoc.Text = "<%=UNINumClientFormat(RdpLocSum,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtBalLoc.Text = "<%=UNINumClientFormat(BalLocSum,ggAmtOfMoney.DecPoint, 0)%>"

		.ggoSpread.Source    = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",2),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",3),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",5),   "A" ,"I","X","X")
		.frm1.vspdData.Redraw = True
	End If  
	.DbQueryOk()
	.frm1.txtLoanPlcNm.value = "<%=ConvSPChars(txtLoanPlcNm)%>"
	.frm1.txtLoanTypeNm.value = "<%=ConvSPChars(txtLoanTypeNm)%>"
	
	.frm1.txtBizAreaCd.value="<%=strBizAreaCd%>"
	.frm1.txtBizAreaNm.value="<%=strBizAreaNm%>"
	.frm1.txtBizAreaCd1.value="<%=strBizAreaCd1%>"
	.frm1.txtBizAreaNm1.value="<%=strBizAreaNm1%>"

End with
</Script>	

