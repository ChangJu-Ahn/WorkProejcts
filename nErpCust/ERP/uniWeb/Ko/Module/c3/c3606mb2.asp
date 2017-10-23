<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")    

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2        '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim txtCostCd																'�� : �ڽ�Ʈ�����ڵ� 
Dim txtCondCostCd															'�� : ���Ǻ� �ڽ�Ʈ�����ڵ� 
Dim txtYYYYMM																'�� : ��� 
Dim txtAcctCd
Dim txtDiFlag																'�� : �����ڵ� 
Dim txtTotWorkinAmt															'�� : ��������հ�� 
Dim txtTotItemAmt															'�� : �ٰűݾ�(����)



'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgPageNo       = Trim(Request("lgPageNo"))                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = Trim(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
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
   	
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time
	
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

    If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then                                            '��: Check if next data exists
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
	Dim strWhere
	
    Redim UNIValue(2,2)

    UNISqlId(0) = "C3606MA102"
    UNISqlId(1) = "C3606MA104"
    UNISqlId(2) = "C3606MA105"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
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
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
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

		'rs1'�� : ��������հ�� 
		If Not (rs1.EOF OR rs1.BOF) Then
			txtTotWorkinAmt = Trim(rs1(0))
		Else
			txtTotWorkinAmt = ""
		End IF
		rs1.Close
		Set rs1 = Nothing

		'rs2'�� : �ٰűݾ�(����)
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
		.ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
		.lgPageNo_B      =  "<%=lgPageNo%>"               '�� : Next next data tag
		.DbQueryOk("2")
	End If

	
End With
</Script>
