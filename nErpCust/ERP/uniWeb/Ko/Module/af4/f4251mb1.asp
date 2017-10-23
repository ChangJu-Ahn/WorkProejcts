<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4251mb1
'*  4. Program Name         : 차입금상환내역조회 
'*  5. Program Desc         : Query of Loan Repay
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2002.04.12
'*  8. Modified date(Last)  : 2003.05.05
'*  9. Modifier (First)     : Hwang Eun Hee
'* 10. Modifier (Last)      : Ahn do hyun
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  차입금번호 오류 Check
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4          '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Const C_SHEETMAXROWS_D = 30

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim cboConfFg, cboApSts
Dim cboLoanFg, txtDocCur, txtLoanPlcFg, txtLoanPlcCd
Dim txtLoanType, txtLoanTypeNm, txtLoanPlcNm
Dim strLoanDtFr, strLoanDtTo
Dim strIntDtFr, strIntDtTo
Dim strPaymDtFr, strPaymDtTo
Dim strWhere1, strWhere2													'⊙ : Where 조건 
Dim strMsgCd, strMsg1, strMsg2
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
	txtDocCur		= UCase(Trim(Request("txtDocCur")))
	txtLoanPlcFg	= Trim(Request("txtLoanPlcFg"))
	txtLoanPlcCd	= UCase(Trim(Request("txtLoanPlcCd")))
	txtLoanType		= UCase(Trim(Request("txtLoanType")))	
	strPaymDtFr		= Request("txtPaymDtFr")
	strPaymDtTo		= Request("txtPaymDtTo")
	strIntDtFr		= Request("txtIntDtFr")
	strIntDtTo		= Request("txtIntDtTo")
	strLoanDtFr		= Request("txtLoanDtFr")
	strLoanDtTo		= Request("txtLoanDtTo")
	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))            '사업장From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))            '사업장To

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

    lgstrData = ""

    lgDataExist    = "Yes"

    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
  	
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(4,11)

    UNISqlId(0) = "F4251MB01"
	If txtLoanPlcFg = "BK" Then
		UNISQLID(1) = "ABANKNM"
	Else
		UNISQLID(1) = "ABPNM"
	End If
	UNISQLID(2) = "AMINORNM"
	UNISqlId(3) = "A_GETBIZ"
    UNISqlId(4) = "A_GETBIZ"
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = FilterVar(UniConvDate(strIntDtFr),""       ,"S") 
    UNIValue(0,2) = FilterVar(UniConvDate(strIntDtTo),""       ,"S") 
    UNIValue(0,3) = FilterVar(UniConvDate(strIntDtFr),""       ,"S") 
    UNIValue(0,4) = FilterVar(UniConvDate(strIntDtTo),""       ,"S") 
    UNIValue(0,5) = FilterVar(UniConvDate(strIntDtFr),""       ,"S") 
    UNIValue(0,6) = FilterVar(UniConvDate(strIntDtTo),""       ,"S") 
    UNIValue(0,7) = FilterVar(UniConvDate(strIntDtFr),""       ,"S") 
    UNIValue(0,8) = FilterVar(UniConvDate(strIntDtTo),""       ,"S") 
    UNIValue(0,9) = UCase(Trim(strWhere1))
    UNIValue(0,10) = UCase(Trim(strWhere2))

	If txtLoanPlcFg = "BK" Then
		UNIValue(1,0) = FilterVar(txtLoanPlcCd ,""       ,"S")
	Else
		UNIValue(1,0) = FilterVar(txtLoanPlcCd ,"''"       ,"S")
	End If

	UNIValue(2,0) = FilterVar("F1000" ,""       ,"S")
	UNIValue(2,1) = FilterVar(txtLoanType ,""       ,"S")
	UNIValue(3,0) = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(4,0) = FilterVar(strBizAreaCd1, "''", "S")

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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    Set lgADF = Nothing  

    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then		
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
  
	If txtLoanPlcCd <> "" Then
		If Not (rs1.EOF OR rs1.BOF) Then
			txtLoanPlcNm = Trim(rs1(1))
		Else
			txtLoanPlcNm = ""			
			Call DisplayMsgBox("140900", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs1.Close
		    Set rs1 = Nothing
			Exit sub
		End IF
	rs1.Close
	Set rs1 = Nothing
	End If

	If txtLoanType <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
			txtLoanTypeNm = Trim(rs2(1))
		Else
			txtLoanTypeNm = ""			
			Call DisplayMsgBox("140900", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs2.Close
		    Set rs2 = Nothing
			Exit sub
		End IF
		rs2.Close
		Set rs2 = Nothing
	End If
	
	'rs3
    If Not( rs3.EOF OR rs3.BOF) Then
   		strBizAreaCd = Trim(rs3(0))
		strBizAreaNm = Trim(rs3(1))
	Else
		strBizAreaCd = ""
		strBizAreaNm = ""
		
    End IF
    
    rs3.Close
    Set rs3 = Nothing
    
    ' rs4
    If Not( rs4.EOF OR rs4.BOF) Then
   		strBizAreaCd1 = Trim(rs4(0))
		strBizAreaNm1 = Trim(rs4(1))
	Else
		strBizAreaCd1 = ""
		strBizAreaNm1 = ""
		
    End IF
    
    rs4.Close
    Set rs4 = Nothing
    
    If rs0.EOF And rs0.BOF Then
	'	If strMsgCd = "" Then strMsgCd = "900014"
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		%><Script Language=vbscript>parent.frm1.txtPaymDtFr.focus</Script><%
		Exit sub
	Else
		Call MakeSpreadSheetData()
	End If

	rs0.Close
	Set rs0 = Nothing 
                                                  '☜: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
	End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Dim strAnd
	strWhere1=""
	strWhere1 = strWhere1 & " Where pay_dt >= " & FilterVar(UniConvDate(strIntDtFr),null	,"S")
	strWhere1 = strWhere1 & " AND pay_dt <= " & FilterVar(UniConvDate(strIntDtTo),null	,"S")
	If cboConfFg	= "C" Then strWhere1 = strWhere1 & " and conf_fg IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " ) " 
	If cboConfFg	= "U" Then	strWhere1 = strWhere1 & " and conf_fg   =  " & FilterVar(cboConfFg , "''", "S") & " " 
	
	strWhere2 = ""
	strAnd = " WHERE"
	If txtLoanPlcFg <> "" Then	strWhere2 = strWhere2 & strAnd & " A.loan_plc_type = " & FilterVar(txtLoanPlcFg ,"''"	,"S")		: strAnd = " AND"
	If strLoanDtFr <> "" Then	strWhere2 = strWhere2 & strAnd & " A.loan_dt >= " & FilterVar(UniConvDate(strLoanDtFr),null,"S")	: strAnd = " AND"
	If strLoanDtTo <> "" Then	strWhere2 = strWhere2 & strAnd & " A.loan_dt <= " & FilterVar(UniConvDate(strLoanDtTo),null,"S")	: strAnd = " AND"
	If strPaymDtFr <> "" Then	strWhere2 = strWhere2 & strAnd & " A.due_dt >= " & FilterVar(UniConvDate(strPaymDtFr),null	,"S")	: strAnd = " AND"
	If strPaymDtTo <> "" Then	strWhere2 = strWhere2 & strAnd & " A.due_dt <= " & FilterVar(UniConvDate(strPaymDtTo),null	,"S")	: strAnd = " AND"
	If cboApSts <> "" Then		strWhere2 = strWhere2 & strAnd & " A.rdp_cls_fg = " & FilterVar(cboApSts ,"''"	,"S")
	If cboLoanFg <> "" Then		strWhere2 = strWhere2 & strAnd & " A.loan_fg = " & FilterVar(cboLoanFg ,"''"	,"S")
	If txtDocCur <> "" Then		strWhere2 = strWhere2 & strAnd & " A.doc_cur = " & FilterVar(txtDocCur ,"''"	,"S")	: strAnd = " AND"
	If txtLoanPlcCd <> "" Then
		If txtLoanPlcFg = "BK" Then
			strWhere2 = strWhere2 & strAnd & " A.Loan_bank_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")	: strAnd = " AND"
		Else
			strWhere2 = strWhere2 & strAnd & " A.bp_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")	: strAnd = " AND"
		End If
	End If
	If txtLoanType <> "" Then strWhere2 = strWhere2 & strAnd & " A.loan_type = " & FilterVar(txtLoanType ,"''"	,"S")
	
	if strBizAreaCd <> "" then
		strWhere2 = strWhere2 & " AND A.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere2 = strWhere2 & " AND A.BIZ_AREA_CD >= " & FilterVar("", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere2 = strWhere2 & " AND A.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere2 = strWhere2 & " AND A.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
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

	strWhere2	= strWhere2	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	
     '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------ 
End Sub

%>

<Script Language=vbscript>
With Parent
	If "<%=lgDataExist%>" = "Yes" Then

		If "<%=lgPageNo%>" = "1" Then									' "1" means that this query is first and next data exists
			.frm1.hApSts.value			= "<%=ConvSPChars(cboApSts)%>" 
			.frm1.hConfFg.value			= "<%=ConvSPChars(cboConfFg)%>" 
			.frm1.hLoanFg.value			= "<%=ConvSPChars(cboLoanFg)%>" 
			.frm1.hDocCur.value		="<%=ConvSPChars(txtDocCur)%>"
			.frm1.hLoanPlcFg.value	="<%=ConvSPChars(txtLoanPlcFg)%>"
			.frm1.hLoanPlcCd.value	="<%=ConvSPChars(txtLoanPlcCd)%>"
			.frm1.hLoanType.value	="<%=ConvSPChars(txtLoanType)%>"
			.frm1.HLoanDtFr.value		="<%=strLoanDtFr%>" 
			.frm1.HLoanDtTo.value		="<%=strLoanDtTo%>" 
			.frm1.HPaymDtFr.value		="<%=strPaymDtFr%>" 
			.frm1.HPaymDtTo.value		="<%=strPaymDtTo%>" 
			.frm1.HIntDtFr.value		="<%=strIntDtFr%>" 
			.frm1.HIntDtTo.value		="<%=strIntDtTo%>"
			.Frm1.htxtBizAreaCd.value = Trim(.Frm1.txtBizAreaCd.value)
			.Frm1.htxtBizAreaCd1.value = Trim(.Frm1.txtBizAreaCd1.value)
	    End If
		.ggoSpread.Source    = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
		.lgPageNo_A      =  "<%=lgPageNo%>"               '☜ : Next next data tag

		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("A",2),.GetKeyPos("A",3),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("A",2),.GetKeyPos("A",4),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("A",2),.GetKeyPos("A",5),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("A",2),.GetKeyPos("A",6),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("A",2),.GetKeyPos("A",7),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("A",2),.GetKeyPos("A",8),   "A" ,"I","X","X")
		.DbQueryOk("1")
		.frm1.vspdData.Redraw = True
	End If  
	.frm1.txtLoanPlcNm.value = "<%=ConvSPChars(txtLoanPlcNm)%>"
	.frm1.txtLoanTypeNm.value = "<%=ConvSPChars(txtLoanTypeNm)%>"
	.frm1.txtBizAreaCd.value="<%=strBizAreaCd%>"
	.frm1.txtBizAreaNm.value="<%=strBizAreaNm%>"
	.frm1.txtBizAreaCd1.value="<%=strBizAreaCd1%>"
	.frm1.txtBizAreaNm1.value="<%=strBizAreaNm1%>"
End with
</Script>	


    
       
