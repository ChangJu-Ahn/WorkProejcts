<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3                           '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Const C_SHEETMAXROWS_D = 30

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strDB
Dim strCond
Dim strLoanFrDt, strLoanToDt, txtDueFrDt, txtDueToDt, txtDocCur, txtLoanFg, txtLoanType, txtLoanNo
Dim txtLoanPlcFg, txtLoanPlcCd, txtLoanPlcNm, txtLoanNm, txtLoanTypeNm
Dim strPgmId
Dim strMsgCd, strMsg1, strMsg2

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
	Call LoadInfTB19029B("I","*","NOCOOKIE","QB")

    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

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
                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(3,3)

    UNISqlId(0) = "F4232RA1"
    
	If txtLoanPlcFg = "BK" Then
		UNISQLID(1) = "ABANKNM"
	Else
		UNISQLID(1) = "ABPNM"
	End If
    
    UNISqlId(2) = "commonqry"
    UNISqlId(3) = "AMINORNM"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strDB
    UNIValue(0,2) = strCond
    
	If txtLoanPlcFg = "BK" Then
		UNIValue(1,0) = FilterVar(txtLoanPlcCd , "''", "S")
	Else
		UNIValue(1,0) = FilterVar(txtLoanPlcCd , "''", "S")
	End If

	UNIValue(2,0) = "select loan_nm from f_ln_info Where loan_no=" & FilterVar(txtLoanNo , "''", "S")

    UNIValue(3,0) = FilterVar("F1000" , "''", "S")
    UNIValue(3,1) = FilterVar(txtLoanType , "''", "S")

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

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
	iStr = Split(lgstrRetMsg,gColSep)

   
    If iStr(0) <> "0" Then
        Call ServerMesgBox(iStr(1), vbInformation, I_MKSCRIPT)
    End If 
 
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
		End If
		rs1.Close
		Set rs1 = Nothing
	End If

	'rs2
	If txtLoanNo <> "" Then
	    If Not (rs2.EOF OR rs2.BOF) Then
			txtLoanNm = Trim(rs2(0))
		Else
			txtLoanNm = ""
			Call DisplayMsgBox("140900", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
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


    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
	%>
		<Script Language=vbscript>
			parent.frm1.txtLoanFrDt.focus
		</Script>
	<%
        rs0.Close
        Set rs0 = Nothing 
        Exit Sub
    Else    
        Call MakeSpreadSheetData()
    End If

	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strLoanFrDt	= UniConvDate(Request("txtLoanFrDt"))
    strLoanToDt	= UniConvDate(Request("txtLoanToDt"))
    txtDueFrDt	= Request("txtDueFrDt")
    txtDueToDt	= Request("txtDueToDt")
    txtDocCur	= UCase(Request("txtDocCur"))
    txtLoanFg	= UCase(Request("txtLoanFg"))
    txtLoanType	= UCase(Request("txtLoanType"))
    txtLoanNo	= UCase(Request("txtLoanNo"))
    txtLoanPlcFg	= UCase(Request("txtLoanPlcFg"))
    txtLoanPlcCd	= UCase(Request("txtLoanPlcCd"))
    strPgmId	    = Request("txtPgmId")
    
    	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

	Select Case strPgmId
		Case "F4232MA1"														'유동성전환 
			strCond = strCond & " AND A.CONF_FG	IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " )"  
			strCond = strCond & " AND A.RDP_CLS_FG	= " & FilterVar("N", "''", "S") & " "  
			strCond = strCond & " AND (A.loan_fg=" & FilterVar("LL", "''", "S") & "  or A.loan_fg = " & FilterVar("LN", "''", "S") & " ) "  
			strDB = ""
		Case "F4233MA1"														'유동성전환취소 
			strCond = strCond & " AND D.Chg_Type = " & FilterVar("FC", "''", "S") & " "
			strCond = strCond & " AND A.loan_no = D.loan_no"
			strDB = ",f_ln_his D"
		Case "F4220BA1"														'차입금상환전개 
	'		strCond = strCond & " AND A.CONF_FG	IN ('C','E')"  
		Case "F4223MA1"														'차입금계획변경 
	'		strCond = strCond & " AND A.CONF_FG	IN ('C','E')"
		Case "F4231MA1"														'이자율변경등록 
	'		strCond = strCond & " AND A.CONF_FG	IN ('C','E')"  
			strCond = strCond & " and A.Int_Votl = " & FilterVar("F", "''", "S") & "  "
	End Select

	If strLoanFrDt <> "" Then strCond = strCond & " and A.loan_dt >=  " & FilterVar(strLoanFrDt , "''", "S") & " "
	If strLoanToDt <> "" Then strCond = strCond & " and A.loan_dt <=  " & FilterVar(strLoanToDt , "''", "S") & " "
	If txtDueFrDt <> "" Then strCond = strCond & " and A.Due_Dt >=  " & FilterVar(txtDueFrDt , "''", "S") & " "
	If txtDueToDt <> "" Then strCond = strCond & " and A.Due_Dt <=  " & FilterVar(txtDueToDt , "''", "S") & " "
	If txtDocCur   <> "" Then strCond = strCond & " and A.Doc_Cur   = " & Filtervar(txtDocCur ,"''"	,"S")
	If txtLoanFg   <> "" Then strCond = strCond & " and A.Loan_Fg   =  " & FilterVar(txtLoanFg , "''", "S") & " "
	If txtLoanType   <> "" Then strCond = strCond & " and A.Loan_Type   = " & Filtervar(txtLoanType ,"''"	,"S")
	If txtLoanNo   <> "" Then strCond = strCond & " and A.Loan_No   = " & Filtervar(txtLoanNo ,"''"	,"S")
	If txtLoanPlcFg   <> "" Then strCond = strCond & " and A.loan_plc_type   =  " & FilterVar(txtLoanPlcFg , "''", "S") & " "
	If txtLoanPlcCd <> "" Then
		If txtLoanPlcFg = "BK" Then
			strCond = strCond & " and A.Loan_bank_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")
		Else
			strCond = strCond & " and A.bp_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")
		End If
	End If
	
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
	
	strCond = strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    With parent
		If "<%=lgDataExist%>" = "Yes" Then
			If "<%=lgPageNo%>" = "1" Then
				.frm1.hLoanFrDt.value = "<%=strLoanFrDt%>"
				.frm1.hLoanToDt.value   = "<%=strLoanToDt%>"
				.frm1.hDueFrDt.value   = "<%=txtDueFrDt%>"
				.frm1.hDueToDt.value   = "<%=txtDueToDt%>"
				.frm1.hDocCur.value   = "<%=ConvSPChars(txtDocCur)%>"
				.frm1.hLoanFg.value   = "<%=ConvSPChars(txtLoanFg)%>"
				.frm1.hLoanType.value   = "<%=ConvSPChars(txtLoanType)%>"
				.frm1.hLoanNo.value   = "<%=ConvSPChars(txtLoanNo)%>"
				.frm1.hLoanPlcFg.value = "<%=ConvSPChars(txtLoanPlcFg)%>"
				.frm1.hLoanPlcCd.value = "<%=ConvSPChars(txtLoanPlcCd)%>"
			End if
			.ggoSpread.Source  = .frm1.vspdData
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
			Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",3),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
			.frm1.vspdData.Redraw = True

		End If	
		.frm1.txtLoanTypeNm.value = "<%=ConvSPChars(txtLoanTypeNm)%>"
		.frm1.txtLoanNm.value = "<%=ConvSPChars(txtLoanNm)%>"
		.frm1.txtLoanPlcNm.value = "<%=ConvSPChars(txtLoanPlcNm)%>"
		.DbQueryOk()
	End with
</Script>	

