<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 

Call LoadBasisGlobalInf()

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2					'�� : DBAgent Parameter ���� 
Dim lgstrData																'�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Const C_SHEETMAXROWS_D = 30

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
dim txtLoanFromDt
dim txtLoanToDt
dim txtDocCur, txtLoanfg, txtLoanType, txtLoanPlcfg, txtLoanPlcCd
Dim txtLoanNo, txtLoanTypeNm, txtLoanPlcNm
Dim arrLoanNo, arrPayPlanDt

Dim  iLoopCount
Dim  LngMaxRow

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- ������ coding part(��������,End)----------------------------------------------------------

	Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1


	txtLoanFromDt = Request("hLoanFromDt")
	txtLoanToDt = Request("hLoanToDt")
	txtDocCur = UCase(Trim(Request("hDocCur")))
	txtLoanfg = Request("hLoanfg")
	txtLoanType = UCase(Trim(Request("hLoanType")))
	txtLoanPlcfg = Trim(Request("hLoanPlcfg"))
	txtLoanPlcCd = UCase(Trim(Request("hLoanPlcCd")))
	txtLoanNo = UCase(Trim(Request("hLoanNo")))

'	txtLoanFromDt = Request("txtLoanFromDt")
'	txtLoanToDt = Request("txtLoanToDt")
'	txtDocCur = UCase(Trim(Request("txtDocCur")))
'	txtLoanfg = Request("txtLoanfg")
'	txtLoanType = UCase(Trim(Request("txtLoanType")))
'	txtLoanPlcfg = Trim(Request("txtLoanPlcfg"))
'	txtLoanPlcCd = UCase(Trim(Request("txtLoanPlcCd")))
'	txtLoanNo = UCase(Trim(Request("txtLoanNo")))

	arrLoanNo = split(Request("hParentLoanNo"), chr(11))
	arrPayPlanDt = split(Request("hParentPayPlanDt"), chr(11))

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))	

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
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
    If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	Dim strWhere, strGroup
	Dim arrCnt
	strWhere = ""
    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "F4250RA101"
	If txtLoanPlcFg = "BK" Then
		UNISQLID(1) = "ABANKNM"
	Else
		UNISQLID(1) = "ABPNM"
	End If
	UNISQLID(2) = "AMINORNM"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	txtDocCur = Request("txtDocCur")
	txtLoanfg = Request("txtLoanfg")
	txtLoanType = Request("txtLoanType")
	txtLoanPlcfg = Trim(Request("txtLoanPlcfg"))
	txtLoanPlcCd = Trim(Request("txtLoanPlcCd"))
	txtLoanNo = Trim(Request("txtLoanNo"))


	strWhere = strWhere & " AND A.CONF_FG IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " )"
	strWhere = strWhere & " AND b.pay_plan_dt >=" & FilterVar(txtLoanFromDt ,"''"     ,"S")
	strWhere = strWhere & " AND b.pay_plan_dt <=" & FilterVar(txtLoanToDt ,"''"       ,"S")
    If txtDocCur <> "" Then		strWhere = strWhere & " AND a.doc_cur = " & FilterVar(txtDocCur ,"''"       ,"S")
	Select Case txtLoanfg
		Case "SL"
			strWhere = strWhere & " AND a.loan_fg IN (" & FilterVar("SL", "''", "S") & " ," & FilterVar("SN", "''", "S") & " ) "
		Case "LL"
			strWhere = strWhere & " AND a.loan_fg IN (" & FilterVar("LL", "''", "S") & " ," & FilterVar("LN", "''", "S") & " ) "
		Case "SLLL"
			strWhere = strWhere & " AND a.loan_fg IN (" & FilterVar("SL", "''", "S") & " ," & FilterVar("SN", "''", "S") & " ," & FilterVar("LL", "''", "S") & " ," & FilterVar("LN", "''", "S") & " ) "
	End Select		
    
    If txtLoanType <> "" Then	strWhere = strWhere & " AND a.loan_type = " & FilterVar(txtLoanType ,"''"       ,"S")
    
    If txtLoanPlcfg <> "" Then	strWhere = strWhere & " AND a.loan_plc_type = " & FilterVar(txtLoanPlcfg ,"''"       ,"S")
	
	If txtLoanPlcCd <> "" Then
		if txtLoanPlcFg = "BK" Then
			strWhere = strWhere & " AND a.loan_bank_cd = " & FilterVar(txtLoanPlcCd ,"''"       ,"S")
		Else
			strWhere = strWhere & " AND a.bp_cd = " & FilterVar(txtLoanPlcCd ,"''"       ,"S")
		End If
	End If
	
	If txtLoanNo <> "" Then		strWhere = strWhere & " AND a.loan_no = " & FilterVar(txtLoanNo ,"''"       ,"S")
	
   	For arrCnt = 0 to ubound(arrLoanNo) - 1
		If Trim(arrLoanNo(arrCnt)) <> "" Then
			strWhere = strWhere & " AND (b.loan_no not in ("	& FilterVar(Trim(arrLoanNo(arrCnt))		, "''", "S") & ")"
			strWhere = strWhere & " or b.pay_plan_dt not in ("	& FilterVar(Trim(arrPayPlanDt(arrCnt))	, "''", "S") & "))"
		End If
	Next
	
	' ���Ѱ��� �߰� 
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

	' ���Ѱ��� �߰� 
	strWhere = strWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL



	UNIValue(0,1)  = strWhere

	If txtLoanPlcFg = "BK" Then
		UNIValue(1,0) = FilterVar(txtLoanPlcCd , "''", "S")
	Else
		UNIValue(1,0) = FilterVar(txtLoanPlcCd , "''", "S")
	End If
	
    UNIValue(2,0) = FilterVar("F1000", "''", "S")
    UNIValue(2,1) = FilterVar(txtLoanType , "''", "S")

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing

	iStr = Split(lgstrRetMsg,gColSep)
   
    If iStr(0) <> "0" Then
        Call ServerMesgBox(iStr(1), vbInformation, I_MKSCRIPT)
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
			txtLoanTypeNm = Trim(rs2("minor_nm"))
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


    If  rs0.EOF Or rs0.BOF Then
		Call DisplayMsgBox("140900", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else
        Call MakeSpreadSheetData()
    End If
    
End Sub
%>

<Script Language=vbscript>

With Parent

	If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists			
		
			.Frm1.hLoanFromDt.value		= "<%=txtLoanFromDt%>"
			.Frm1.hLoanToDt.value		= "<%=txtLoanToDt%>"
			.Frm1.hDocCur.value			= "<%=ConvSPChars(txtDocCur)%>"
			.Frm1.hLoanfg.value			= "<%=ConvSPChars(txtLoanfg)%>"
			.Frm1.hLoanType.value		= "<%=ConvSPChars(txtLoanType)%>"
			.Frm1.hLoanPlcfg.value		= "<%=ConvSPChars(txtLoanPlcfg)%>"
			.Frm1.hLoanPlcCd.value		= "<%=ConvSPChars(txtLoanPlcCd)%>"
			.Frm1.hLoanNo.value			= "<%=ConvSPChars(txtLoanNo)%>"
			
		End If
       
       'Show multi spreadsheet data from this line
		.ggoSpread.Source  = Parent.frm1.vspdData
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '�� : Display data
		.FRM1.lgPageNo.VALUE      =  "<%=lgPageNo%>"               '�� : Next next data tag
		.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",8),parent.GetKeyPos("A",9),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",8),parent.GetKeyPos("A",11),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",8),parent.GetKeyPos("A",13),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",8),parent.GetKeyPos("A",19),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",8),parent.GetKeyPos("A",21),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",8),parent.GetKeyPos("A",23),   "A" ,"I","X","X")
		.frm1.vspdData.Redraw = True
    End If

	.DbQueryOk()
	.frm1.txtLoanPlcNm.value = "<%=ConvSPChars(txtLoanPlcNm)%>"			'rs1 �� �ޱ� �˾����� ���ϰ� �׳� �Է������� ���־��ֱ� 
	.frm1.txtLoanTypeNm.value = "<%=ConvSPChars(txtLoanTypeNm)%>"			'rs2 �� �ޱ� �˾����� ���ϰ� �׳� �Է������� ���־��ֱ� 
	 
End With

</Script>
