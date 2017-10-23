<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3104mb2
'*  4. Program Name         : ���������⳻����ȸ 
'*  5. Program Desc         : Query of Deposit Income/Outgo
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.12
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  ������ڵ�, �����ڵ� ���� Check
'=======================================================================================================


%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strBankCd																'�� : ���� 
Dim strBankAcctNo															'�� : ���¹�ȣ 
Dim strDateFr, strDateTo													'�� : �������� 
Dim PreAmt, PreLocAmt, RcptAmt, RcptLocAmt, PaymAmt, PaymLocAmt, BalAmt, BalLocAmt
Dim	strWhere

Dim  iLoopCount
Dim  LngMaxRow

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")
    Call HideStatusWnd 

    lgPageNo			= UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    lgStrPrevKey		= Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount			= 100
    lgSelectList		= Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList			= Request("lgTailList")                                 '�� : Orderby value
	LngMaxRow			= CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

	' ���Ѱ��� �߰� 
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

    If  iLoopCount < lgMaxCount Then                                   '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
        lgStrPrevKey = ""
    End If
  	
'	rs0.Close
'    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	'--------------- ������ coding part(�������,Start)----------------------------------------------------

	Redim UNIValue(2,5)

	UNISqlId(0) = "F3104MA102"
	UNISqlId(1) = "F3104MA103"	'�̿��ݾ� 
	UNISqlId(2) = "F3104MA103"	'�����հ� 

	'--------------- ������ coding part(�������,End)------------------------------------------------------
	UNIValue(0,0) = lgSelectList                                          '��: Select list
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
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

	'--------------- ������ coding part(�������,End)------------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
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
		Set lgADF = Nothing                                             '��: ActiveX Data Factory Object Nothing
		Exit Sub
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
		
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strBankCd			= Request("txtBankCd")
     strBankAcctNo		= Request("txtBankAcctNo")
     strDateFr			= UniConvDate(Request("txtDateFr"))
     strDateTo			= UniConvDate(Request("txtDateTo"))

	strWhere = ""

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
	strWhere	= strWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL



    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>


    With parent
         .ggoSpread.Source    = .frm1.vspdData2 

		.frm1.vspdData2.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '�� : Display data
		.lgStrPrevKey_B       = "<%=ConvSPChars(lgStrPrevKey)%>"     '�� :  set next data tag
		.lgPageNo_B           =  "<%=lgPageNo%>"                     '�� : Next next data tag
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
