<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 

Call LoadBasisGlobalInf()

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1						'�� : DBAgent Parameter ���� 
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
dim txtPayFromDt
dim txtPayToDt
dim txtPayNo
Dim txtBankPayCd
Dim txtBankPayNm
Dim txthPayNo

Dim  iLoopCount
Dim  LngMaxRow

'--------------- ������ coding part(��������,End)----------------------------------------------------------

	Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

	txtPayFromDt = Request("txtPayFromDt")
	txtPayToDt = Request("txtPayToDt")
	txtPayNo = Trim(Request("txtPayNo"))
	txtBankPayCd = Trim(Request("txtBankPayCd"))

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
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	Dim strWhere, strGroup
	strWhere = ""
    Redim UNIValue(1,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "F4255RA201"
    UNISQLID(1) = "commonqry"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    

    strWhere = strWhere & " AND A.PAY_DT >=" & FilterVar(txtPayFromDt ,"''"     ,"S")
    strWhere = strWhere & " AND A.PAY_DT <=" & FilterVar(txtPayToDt ,"''"       ,"S")
    strWhere = strWhere & " AND A.PAY_COND = " & FilterVar("M", "''", "S") & "  "
    If txtPayNo <> "" Then
	    strWhere = strWhere & " AND A.PAY_NO >= " & FilterVar(txtPayNo ,"''"       ,"S")
	End If
    If txtBankPayCd <> "" Then
	    strWhere = strWhere & " AND A.BANK_CD = " & FilterVar(txtBankPayCd ,"''"       ,"S")
	End If
	strWhere = strWhere & " and A.ST_ADV_INT_FG <> " & FilterVar("IA", "''", "S") & "  "
	UNIValue(0,1)  = strWhere
	UNIValue(1,0) = "select bank_nm from b_bank Where bank_cd=" & FilterVar(txtBankPayCd ,"''"       ,"S")
    
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	'rs1
	If txtBankPayCd <> "" Then
		If Not (rs1.EOF OR rs1.BOF) Then
			txtBankPayNm = Trim(rs1("Bank_Nm"))
		Else
			txtBankPayNm = ""
			Call DisplayMsgBox("800123", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs1.Close
		    Set rs1 = Nothing
			Exit sub
		End IF
		rs1.Close
		Set rs1 = Nothing
	End If

    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
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
			.Frm1.hPayFromDt.value	= "<%=txtPayFromDt%>"
			.Frm1.hPayToDt.value	= "<%=txtPayToDt%>"
			.Frm1.hBankPayCd.value	= "<%=ConvSPChars(txtBankPayCd)%>"
			.Frm1.hPayNo.value		= "<%=ConvSPChars(txthPayNo)%>"
		End If
       
       'Show multi spreadsheet data from this line
		.ggoSpread.Source  = Parent.frm1.vspdData
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '�� : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",3),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",3),parent.GetKeyPos("A",5),   "A" ,"I","X","X")
		.frm1.vspdData.Redraw = True
    End If

	.DbQueryOk()
	.frm1.txtBankPayNm.value = "<%=ConvSPChars(txtBankPayNm)%>"			'rs1 �� �ޱ� �˾����� ���ϰ� �׳� �Է������� ���־��ֱ� 
	 
End With

</Script>
