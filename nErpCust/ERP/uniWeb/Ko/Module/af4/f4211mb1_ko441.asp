<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%'======================================================================================================
'*  1. Module Name          : Accounting - Treasury
'*  2. Function Name        : Loan
'*  3. Program ID           : f4205rb1
'*  4. Program Name         : 차입금번호팝업 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001.02.19
'*  7. Modified date(Last)  : 2003.04.29
'*  8. Modifier (First)     : Song, Mun Gil
'*  9. Modifier (Last)      : Oh, Soo Min
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

%>

<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("Q", "A","NOCOOKIE","QB")
%>

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5     '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Const C_SHEETMAXROWS_D = 100

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim	txtLoanType, txtLoanTypeNm
Dim cboConfFg, cboApSts
Dim cboLoanFg, txtDocCur, txtLoanPlcFg, txtLoanPlcCd, txtLoanPlcNm
Dim	txtLoan_From_DT
Dim	txtLOAN_To_Dt
Dim	txtBase_Dt
Dim	txtDue_From_Dt,txtDue_To_Dt
Dim txtLoanTotAmt, txtPrPayTotAmt, txtBalTotAmt, txtIntPayTotAmt
Dim txtLoanTotLocAmt, txtPrPayTotLocAmt, txtBalTotLocAmt, txtIntPayTotLocAmt
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1

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

	cboConfFg		= Trim(Request("cboConfFg"))
	cboApSts		= Trim(Request("cboApSts"))
	cboLoanFg		= Trim(Request("cboLoanFg"))
	txtDocCur		= Trim(Request("txtDocCur"))
	txtLoanPlcFg	= Trim(Request("txtLoanPlcFg"))
	txtLoanPlcCd	= Trim(Request("txtLoanPlcCd"))
	txtLoanType		= Trim(Request("txtLoanType"))
	txtLoan_From_DT = Request("txtLoan_From_DT")
	txtLOAN_To_Dt   = Request("txtLOAN_To_Dt")
	txtBase_Dt		= Request("txtBase_Dt")
	txtDue_From_Dt	= Request("txtDue_From_Dt")
	txtDue_To_Dt	= Request("txtDUE_To_DT")
	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))            '사업장From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))            '사업장To

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    
    Dim  iRowStr

    lgstrData = ""

    lgDataExist    = "Yes"

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
	Dim strWhere
	Dim strConfFg

    Redim UNIValue(5,11)

    UNISQLID(0) = "F4211MA101"
	If txtLoanPlcFg = "BK" Then
		UNISQLID(1) = "ABANKNM"
	Else
		UNISQLID(1) = "ABPNM"
	End If
	UNISQLID(2) = "AMINORNM"
    UNISQLID(3) = "F4211MA102"
    UNISqlId(4) = "A_GETBIZ"
    UNISqlId(5) = "A_GETBIZ"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    strWhere = ""
    strConfFg = ""

    If txtDue_From_Dt	<> "" Then strWhere = strWhere & " AND a.due_dt >= " & FilterVar(UNIConvDate(txtDue_From_Dt),null	,"S")
	If txtDue_To_Dt		<> "" Then strWhere = strWhere & " AND a.due_dt <= " & FilterVar(UNIConvDate(txtDue_To_Dt),null	,"S")
	If txtLoanType		<> "" Then strWhere = strWhere & " AND a.loan_type = " & FilterVar(txtLoanType ,"''"	,"S")
	If cboConfFg		= "C" Then strConfFg = strConfFg & " and a.conf_fg IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " ) " 
	If cboConfFg		= "U" Then strConfFg = strConfFg & " and a.conf_fg   =  " & FilterVar(cboConfFg , "''", "S") & " " 
'	If cboApSts			<> "" Then strWhere = strWhere & " AND a.rdp_cls_fg = " & FilterVar(cboApSts ,"''"	,"S")
	If cboApSts			= "Y" Then strWhere = strWhere & " AND (HI.CHG_DT IS NOT NULL OR ISNULL(ITEM.PR_PAY_AMT,0) + ISNULL(LN1.BAS_RDP_AMT,0) >= ISNULL(A.LOAN_AMT, 0)) "
	If cboApSts			= "N" Then strWhere = strWhere & " AND HI.CHG_DT IS NULL AND ISNULL(ITEM.PR_PAY_AMT,0) + ISNULL(LN1.BAS_RDP_AMT,0) < ISNULL(A.LOAN_AMT, 0) "
	If cboLoanFg		<> "" Then strWhere = strWhere & " AND a.loan_fg = " & FilterVar(cboLoanFg ,"''"	,"S")
	If txtDocCur		<> "" Then strWhere = strWhere & " AND a.doc_cur = " & FilterVar(txtDocCur ,"''"	,"S")
	If txtLoanPlcFg		<> "" Then strWhere = strWhere & " AND a.loan_plc_type = " & FilterVar(txtLoanPlcFg ,"''"	,"S")
	If txtLoanPlcCd		<> "" Then
		If txtLoanPlcFg = "BK" Then
			strWhere = strWhere & " and a.Loan_bank_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")
		Else
			strWhere = strWhere & " and a.bp_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")
		End If
	End If
	
	if strBizAreaCd		<> "" then
		strWhere = strWhere & " AND a.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere = strWhere & " AND a.BIZ_AREA_CD >= " & FilterVar("0", "''", "S") & " "
	end if
	
	if strBizAreaCd1	<> "" then
		strWhere = strWhere & " AND a.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere = strWhere & " AND a.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if

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

	strWhere	= strWhere	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	
    UNIValue(0,1) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
'    UNIValue(0,2) = strConfFg 
    UNIValue(0,3) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(0,4) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(0,5) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(0,6) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(0,7) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(0,8) = FilterVar(UNIConvDate(txtLoan_From_DT), "''", "S") 
    UNIValue(0,9) = FilterVar(UNIConvDate(txtLoan_To_DT), "''", "S") 
    UNIValue(0,10) = strWhere & strConfFg

	If txtLoanPlcFg = "BK" Then
		UNIValue(1,0) = FilterVar(txtLoanPlcCd ,""       ,"S")
	Else
		UNIValue(1,0) = FilterVar(txtLoanPlcCd ,"''"       ,"S")
	End If
    UNIValue(2,0) = FilterVar("F1000" , "''", "S") 
    UNIValue(2,1) = FilterVar(txtLoanType ,""	,"S")
    UNIValue(3,0) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
'    UNIValue(3,1) = strConfFg
    UNIValue(3,2) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(3,3) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(3,4) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(3,5) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(3,6) = FilterVar(UNIConvDate(txtBase_Dt), "''", "S") 
    UNIValue(3,7) = FilterVar(UNIConvDate(txtLoan_From_DT), "''", "S") 
    UNIValue(3,8) = FilterVar(UNIConvDate(txtLoan_To_DT), "''", "S") 
    UNIValue(3,9) = strWhere & strConfFg
    
    UNIValue(4,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(5,0)  = FilterVar(strBizAreaCd1, "''", "S")
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    'rs1
	If txtLoanPlcCd <> "" Then
	    If Not (rs1.EOF OR rs1.BOF) Then
			txtLoanPlcNm = Trim(rs1(1))
		Else
			txtLoanPlcNm = ""
			If txtLoanPlcFg = "BK" Then
				Call DisplayMsgBox("800123", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
			Else
				Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
			End If			
	        rs1.Close
		    Set rs1 = Nothing 
			Exit sub
		End IF
		rs1.Close
		Set rs1 = Nothing
	End If

    'rs2
	If txtLoanType <> "" Then
	    If Not (rs2.EOF OR rs2.BOF) Then
			txtLoanTypeNm = Trim(rs2(1))
		Else
			txtLoanTypeNm = ""
			Call DisplayMsgBox("140936", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs2.Close
		    Set rs2 = Nothing 
			Exit sub
		End IF
		rs2.Close
		Set rs2 = Nothing
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
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()

	    'rs3
		If Not (rs3.EOF OR rs3.BOF) Then
			'LocalAmt(b,d,f,h)
			txtLoanTotAmt = Trim(rs3("a"))
			txtLoanTotLocAmt = Trim(rs3("b"))
			txtPrPayTotAmt = Trim(rs3("c"))
			txtPrPayTotLocAmt = Trim(rs3("d"))
			txtIntPayTotAmt = Trim(rs3("e"))
			txtIntPayTotLocAmt = Trim(rs3("f"))
			txtBalTotAmt = Trim(rs3("g"))
			txtBalTotLocAmt = Trim(rs3("h"))
		Else
			txtLoanTotAmt = ""
			txtLoanTotLocAmt = ""
			txtPrPayTotAmt = ""
			txtPrPayTotLocAmt = ""
			txtIntPayTotAmt = ""
			txtIntPayTotLocAmt = ""
			txtBalTotAmt = ""
			txtBalTotLocAmt = ""
		End IF
		rs3.Close
		Set rs3 = Nothing
    End If
    
End Sub
%>
<Script Language=vbscript>
With Parent
	If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.hConfFg.value		= "<%=ConvSPChars(cboConfFg)%>" 
			.frm1.hApSts.value		= "<%=ConvSPChars(cboApSts)%>" 
			.frm1.hLoanFg.value		= "<%=ConvSPChars(cboLoanFg)%>" 
			.frm1.hDocCur.value		= "<%=ConvSPChars(txtDocCur)%>" 
			.frm1.hLoanPlcFg.value	= "<%=ConvSPChars(txtLoanPlcFg)%>" 
			.frm1.hLoanPlcCd.value	= "<%=ConvSPChars(txtLoanPlcCd)%>" 
			.Frm1.hLoanType.Value		= "<%=ConvSPChars(txtLoanType)%>"			'For Next Search
			.Frm1.hLoanFromDt.Value	= "<%=txtLoan_From_DT%>"                  'For Next Search
			.Frm1.hLoanToDt.Value		= "<%=txtLOAN_To_Dt%>"                  'For Next Search
			.Frm1.hDueFromDt.Value	= "<%=txtDue_From_Dt%>"                  'For Next Search
			.Frm1.hDueToDt.Value		= "<%=txtDue_To_Dt%>"                  'For Next Search
			.Frm1.hBaseDt.Value		= "<%=txtBase_Dt%>"                  'For Next Search
			.Frm1.htxtBizAreaCd.value = Trim(.Frm1.txtBizAreaCd.value)
			.Frm1.htxtBizAreaCd1.value = Trim(.Frm1.txtBizAreaCd1.value)
		End If
'		If "<%=txtDocCur%>"   <> "" Then
'			.frm1.txtLoanTotAmt.Text	= "<%=UNINumClientFormat(txtLoanTotAmt,ggAmtOfMoney.DecPoint, 0)%>"			'rs3 값 
'			.frm1.txtPrPayTotAmt.Text	= "<%=UNINumClientFormat(txtPrPayTotAmt,ggAmtOfMoney.DecPoint, 0)%>"
'			.frm1.txtIntPayTotAmt.Text	= "<%=UNINumClientFormat(txtIntPayTotAmt,ggAmtOfMoney.DecPoint, 0)%>"
'			.frm1.txtBalTotAmt.Text		= "<%=UNINumClientFormat(txtBalTotAmt,ggAmtOfMoney.DecPoint, 0)%>"
'		Else
			.frm1.txtLoanTotAmt.Text	= "<%=UNINumClientFormat(txtLoanTotAmt,2, 0)%>"			'rs3 값 
			.frm1.txtPrPayTotAmt.Text	= "<%=UNINumClientFormat(txtPrPayTotAmt,2, 0)%>"
			.frm1.txtIntPayTotAmt.Text	= "<%=UNINumClientFormat(txtIntPayTotAmt,2, 0)%>"
			.frm1.txtBalTotAmt.Text		= "<%=UNINumClientFormat(txtBalTotAmt,2, 0)%>"
'		End If
		.frm1.txtLoanTotLocAmt.Text	= "<%=UNINumClientFormat(txtLoanTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"			'rs3 값 
		.frm1.txtPrPayTotLocAmt.Text= "<%=UNINumClientFormat(txtPrPayTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtIntPayTotLocAmt.Text = "<%=UNINumClientFormat(txtIntPayTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtBalTotLocAmt.Text	= "<%=UNINumClientFormat(txtBalTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"

		.ggoSpread.Source  = .frm1.vspdData
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
	.frm1.txtLoanPlcNm.value = "<%=ConvSPChars(txtLoanPlcNm)%>"			'rs1 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 
	.frm1.txtLoanTypeNm.value = "<%=ConvSPChars(txtLoanTypeNm)%>"			'rs1 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 
	.frm1.txtBizAreaCd.value="<%=strBizAreaCd%>"
	.frm1.txtBizAreaNm.value="<%=strBizAreaNm%>"
	.frm1.txtBizAreaCd1.value="<%=strBizAreaCd1%>"
	.frm1.txtBizAreaNm1.value="<%=strBizAreaNm1%>"
End With
</Script>

