<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3102mb1
'*  4. Program Name         : ��������ȸ 
'*  5. Program Desc         : Query of Deposit
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.11
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2003/06/12 Oh, Soo Min (MA�� C_SHEETMAXROWS_D ���� ����)
'=======================================================================================================


%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6                            '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgStrPrevKey
Dim lgTailList                                                              '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strBankCd, strBankAcctNo, strDateFr, strDateTo, strDocCur
Dim PreAmt, PreLocAmt, RcptAmt, RcptLocAmt, PaymAmt, PaymLocAmt
Dim strMsgCd, strMsg1, strMsg2
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1
Dim strWhere

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
    Call LoadInfTB19029B("Q","F","NOCOOKIE","QB")
    Call HideStatusWnd 

    lgPageNo			= UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    lgStrPrevKey		= Request("lgStrPrevKey")
    lgMaxCount			= 100
    lgSelectList		= Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList			= Request("lgTailList")                                 '�� : Orderby value
    lgDataExist			= "No"
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
     
    'rs0�� ���� ��� 
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
'			Call ServerMesgBox(lgstrData, vbInformation, I_MKSCRIPT)
            
        Else
			Call ServerMesgBox("lgPageNo : " & lgPageNo, vbInformation, I_MKSCRIPT)
        
            lgPageNo = lgPageNo + 1
            lgStrPrevKey = rs0(8)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
        lgStrPrevKey = ""
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
                    '��: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(6)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(6,6)

    UNISqlId(0) = "F3102MA101KO441"	'��������ȸ 
    UNISqlId(1) = "F3102MA102KO441"	'����� 
    UNISqlId(2) = "F3102MA103KO441"	'�̿��ݾ� 
    UNISqlId(3) = "F3102MA104KO441"	'�����հ� 
    UNISqlId(4) = "F3102MA105KO441"	'���¹�ȣ 
	UNISqlId(5) = "A_GETBIZ"
    UNISqlId(6) = "A_GETBIZ"
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1) = FilterVar(strBankCd, "''", "S") 
	UNIValue(0,2) = FilterVar(strBankAcctNo, "''", "S")
	UNIValue(0,3) = FilterVar(strDateFr, "''", "S")
	UNIValue(0,4) = FilterVar(strDateTo, "''", "S")
	UNIValue(0,5) = strWhere

	UNIValue(1,0) = FilterVar(strBankCd, "''", "S")

	UNIValue(2,0) = FilterVar(strBankCd, "''", "S")
	UNIValue(2,1) = FilterVar(strBankAcctNo, "''", "S")
	UNIValue(2,2) = FilterVar(strDateFr, "''", "S") 
	UNIValue(2,3) = strWhere

	UNIValue(3,0) = FilterVar(strBankCd, "''", "S") 
	UNIValue(3,1) = FilterVar(strBankAcctNo, "''", "S") 
	UNIValue(3,2) = FilterVar(strDateFr, "''", "S") 
	UNIValue(3,3) = FilterVar(strDateTo, "''", "S") 
	UNIValue(3,4) = strWhere
	
	UNIValue(4,0) = FilterVar(strBankCd, "''", "S") 
	UNIValue(4,1) = FilterVar(strBankAcctNo, "''", "S") 
	
	UNIValue(5,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(6,0)  = FilterVar(strBizAreaCd1, "''", "S")
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6 )
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strBankCd <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBankCd_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
				.txtBankCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
				.txtBankNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
			End With
		</Script>
<%	
	End If
	
	rs1.Close
	Set rs1 = Nothing
	
	If (rs4.EOF And rs4.BOF) Then
		If strMsgCd = "" And strBankAcctNo <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBankAcctNo_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
				.txtBankAcctNo.value = "<%=ConvSPChars(Trim(rs4(0)))%>"
			End With
		</Script>
<%	
	End If
	
	rs4.Close
	Set rs4 = Nothing
	
	If (rs5.EOF And rs5.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd.value = "<%=Trim(rs5(0))%>"
		.frm1.txtBizAreaNm.value = "<%=Trim(rs5(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs5.Close
	Set rs5 = Nothing
	
	
	If (rs6.EOF And rs6.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT1")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd1.value = "<%=Trim(rs6(0))%>"
		.frm1.txtBizAreaNm1.value = "<%=Trim(rs6(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs6.Close
	Set rs6 = Nothing
	
	
	
	PreAmt     = 0
	PreLocAmt  = 0
	RcptAmt    = 0
	RcptLocAmt = 0
	PaymAmt    = 0
	PaymLocAmt = 0
	
	If Not(rs2.EOF And rs2.BOF) Then
		If IsNull(rs2(0)) = False Then PreAmt    = rs2(0)
		If IsNull(rs2(1)) = False Then PreLocAmt = rs2(1)
	End If
	
	rs2.Close
	Set rs2 = Nothing
	
	If Not(rs3.EOF And rs3.BOF) Then
		If IsNull(rs3(0)) = False Then RcptAmt    = rs3(0)
		If IsNull(rs3(1)) = False Then RcptLocAmt = rs3(1)
		If IsNull(rs3(2)) = False Then PaymAmt    = rs3(2)
		If IsNull(rs3(3)) = False Then PaymLocAmt = rs3(3)
	End If

	rs3.Close
	Set rs3 = Nothing

    If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If

	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    strBankCd		= UCase(Trim(Request("txtBankCd")))
    strBankAcctNo	= UCase(Trim(Request("txtBankAcctNo")))
    strDateFr		= UniConvDate(Request("txtDateFr"))
    strDateTo		= UniConvDate(Request("txtDateTo"))

    If Trim(Request("txtDocCur")) = "" Then
		strDocCur = ""
	Else
		strDocCur = " AND B.DOC_CUR = " & Filtervar(UCase(Trim(Request("txtDocCur"))), "''", "S")
	End If
	
	strWhere	= strDocCur

	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))            '�����From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))            '�����To
	
	if strBizAreaCd <> "" then
		strWhere = strWhere & " AND B.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere = strWhere & " AND B.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	end if


	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND B.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	' ���Ѱ��� �߰� 
	strWhere	= strWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL


	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
With parent
	
	If "<%=lgDataExist%>" = "Yes" Then
	   If "<%=strDocCur%>" <> "" Then
		.frm1.txtPreAmt.Text     = "<%=UNINumClientFormat(PreAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtRcptAmt.Text    = "<%=UNINumClientFormat(RcptAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtPaymAmt.Text    = "<%=UNINumClientFormat(PaymAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtBalAmt.Text     = "<%=UNINumClientFormat(Cdbl(PreAmt) + Cdbl(RcptAmt) - Cdbl(PaymAmt), ggAmtOfMoney.DecPoint, 0)%>"	
	   Else 
	    .frm1.txtPreAmt.Text     = "<%=UNINumClientFormat(PreAmt, 2, 0)%>"
		.frm1.txtRcptAmt.Text    = "<%=UNINumClientFormat(RcptAmt, 2, 0)%>"
		.frm1.txtPaymAmt.Text    = "<%=UNINumClientFormat(PaymAmt, 2, 0)%>"
		.frm1.txtBalAmt.Text     = "<%=UNINumClientFormat(Cdbl(PreAmt) + Cdbl(RcptAmt) - Cdbl(PaymAmt), 2, 0)%>"	
	   End If	  	

		.frm1.txtPreLocAmt.Text  = "<%=UNINumClientFormat(PreLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtRcptLocAmt.Text = "<%=UNINumClientFormat(RcptLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtPaymLocAmt.Text = "<%=UNINumClientFormat(PaymLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtBalLocAmt.Text  = "<%=UNINumClientFormat(Cdbl(PreLocAmt) + Cdbl(RcptLocAmt) - Cdbl(PaymLocAmt), ggAmtOfMoney.DecPoint, 0)%>"

        .ggoSpread.Source    = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '�� : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		.lgStrPrevKey        = "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",4),parent.GetKeyPos("A",5),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",4),parent.GetKeyPos("A",7),   "A" ,"I","X","X")
	'	Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1 , -1 ,parent.GetKeyPos("A",4),parent.GetKeyPos("A",5),   "A" ,"Q","X","X")
	'	Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1 , -1, parent.GetKeyPos("A",4),parent.GetKeyPos("A",7),   "A" ,"Q","X","X")
		.frm1.vspdData.Redraw = True
		.DbQueryOk()
	End If
End with
	
</Script>	


