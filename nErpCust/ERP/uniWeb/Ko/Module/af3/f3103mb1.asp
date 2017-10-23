<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3103ma1
'*  4. Program Name         : �������ܰ���ȸ 
'*  5. Program Desc         : Query of Deposit Balance
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.11
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song, Mun Gil  ������ڵ�, �����ڵ� ���� Check
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
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5     '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgStrPrevKey
Dim lgTailList                                                              '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strBizAreaCd	'����� 
Dim strBizAreaCd1	'�����1
Dim strDpstFg		'�����ݱ��� 
Dim strDateMid		'�������� 
Dim strTransSts		'�ŷ����� 
Dim strBankCd		'���� 
Dim strDocCur		'��ȭ 
Dim strWhere		'Where ���� 
Dim strMsgCd, strMsg1, strMsg2

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


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    lgStrPrevKey   = Request("lgStrPrevKey")
	lgMaxCount     = 100
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

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
        Else
            lgPageNo = lgPageNo + 1
            lgStrPrevKey = rs0(3)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
        lgPageNo = ""  
        lgStrPrevKey = ""                                                '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(5)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(5,2)

    UNISqlId(0) = "F3103MA101"
    UNISqlId(1) = "F3103MA102"	'�����հ�, �ܾ� 
    UNISqlId(2) = "F3103MA103"	'������ڵ� 
    UNISqlId(3) = "F3103MA104"	'�����ڵ� 
    UNISqlId(4) = "F3103MA105"	'��ȭ�ڵ� 
    UNISqlId(5) = "F3103MA103"	'������ڵ�1

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1) = Trim(strWhere)
	
	
	UNIValue(1,0) = Trim(strWhere)

	UNIValue(2,0) = FilterVar(strBizAreaCd, "''", "S") 
	UNIValue(3,0) = FilterVar(strBankCd, "''", "S") 
	UNIValue(4,0) = FilterVar(strDocCur, "''", "S") 
	UNIValue(5,0) = FilterVar(strBizAreaCd1, "''", "S") 
	
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4,rs5)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	If Not(rs1.EOF And rs1.BOF) Then
%>
		<Script Language=vbscript>
			With parent.frm1
		   If "<%=strDocCur%>" <> "" Then				
			.txtRcptAmt.Text    = "<%=UNINumClientFormat(rs1(0), ggAmtOfMoney.DecPoint, 0)%>"
			.txtPaymAmt.Text    = "<%=UNINumClientFormat(rs1(2), ggAmtOfMoney.DecPoint, 0)%>"
			.txtBalAmt.Text     = "<%=UNINumClientFormat(rs1(4), ggAmtOfMoney.DecPoint, 0)%>"
		   Else	
			.txtRcptAmt.Text    = "<%=UNINumClientFormat(rs1(0), 2, 0)%>"
			.txtPaymAmt.Text    = "<%=UNINumClientFormat(rs1(2), 2, 0)%>"
			.txtBalAmt.Text     = "<%=UNINumClientFormat(rs1(4), 2, 0)%>"
		   End If
		   	
			.txtRcptLocAmt.Text = "<%=UNINumClientFormat(rs1(1), ggAmtOfMoney.DecPoint, 0)%>"
			.txtPaymLocAmt.Text = "<%=UNINumClientFormat(rs1(3), ggAmtOfMoney.DecPoint, 0)%>"
			.txtBalLocAmt.Text  = "<%=UNINumClientFormat(rs1(5), ggAmtOfMoney.DecPoint, 0)%>"
			
			End With
		</Script>
<%	
	End If
	
	rs1.Close
	Set rs1 = Nothing
	
	
	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strBizAreaCd <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBizAreaCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
			With parent.frm1
				.txtBizAreaCd.Value = "<%=ConvSPChars(Trim(rs2(0)))%>"
				.txtBizAreaNm.Value = "<%=ConvSPChars(Trim(rs2(1)))%>"
			End With
		</Script>
<%
	End If
	
	rs2.Close
	Set rs2 = Nothing
	
	If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And strBankCd <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBankCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
			With parent.frm1
				.txtBankCd.Value = "<%=ConvSPChars(Trim(rs3(0)))%>"
				.txtBankNm.Value = "<%=ConvSPChars(Trim(rs3(1)))%>"
			End With
		</Script>
<%
	End If
	
	rs3.Close
	Set rs3 = Nothing
	
	If (rs4.EOF And rs4.BOF) Then
		If strMsgCd = "" And strDocCur <> "" Then
			strMsgCd = "970000"
			strMsg1 = Request("txtDocCur_Alt")
		End If
	End If
	
	rs4.Close
	Set rs4 = Nothing
	
	If (rs5.EOF And rs5.BOF) Then
		If strMsgCd = "" And strBizAreaCd1 <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBizAreaCd_Alt1")
		End If
	Else
%>
		<Script Language=vbScript>
			With parent.frm1
				.txtBizAreaCd1.Value = "<%=ConvSPChars(Trim(rs5(0)))%>"
				.txtBizAreaNm1.Value = "<%=ConvSPChars(Trim(rs5(1)))%>"
			End With
		</Script>
<%
	End If
	
	rs5.Close
	Set rs5 = Nothing
	
	
    If (rs0.EOF And rs0.BOF) Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If

'	rs0.Close
'	Set rs0 = Nothing 
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
    strBizAreaCd	= UCase(Trim(Request("txtBizAreaCd")))
    strBizAreaCd1	= UCase(Trim(Request("txtBizAreaCd1")))
    strDpstType		= UCase(Trim(Request("cboDpstType")))
    strDateMid		= UniConvDate(Request("txtDateMid"))
    strTransSts		= UCase(Trim(Request("cboTransSts")))
    strBankCd		= UCase(Trim(Request("txtBankCd")))
    strDocCur		= UCase(Trim(Request("txtDocCur")))

	strWhere = ""
	
	if strBizAreaCd <> "" then
		strWhere = strWhere & " and A.biz_area_cd >=  " & FilterVar(strBizAreaCd , "''", "S") & ""
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere = strWhere & " and A.biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""
	end if
	
	strWhere = strWhere & " and A.start_dt <=  " & FilterVar(strDateMid , "''", "S") & " and E.trans_dt <=  " & FilterVar(strDateMid , "''", "S") & ""
	
	If strDpstType <> "" Then strWhere = strWhere & " and A.dpst_Type	=  " & FilterVar(strDpstType , "''", "S") & ""
    If strTransSts <> "" Then strWhere = strWhere & " and A.trans_sts	=  " & FilterVar(strTransSts , "''", "S") & ""
    If strBankCd   <> "" Then strWhere = strWhere & " and A.bank_cd		=  " & FilterVar(strBankCd , "''", "S") & ""
    If strDocCur   <> "" Then strWhere = strWhere & " and A.doc_cur		=  " & FilterVar(strDocCur , "''", "S") & ""

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND E.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND E.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND E.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND E.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	' ���Ѱ��� �߰� 
	strWhere	= strWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL



    '--------------- ������ coding part(�������,End)------------------------------------------------------


    
End Sub

%>

<Script Language=vbscript>
	With parent
	If "<%=lgDataExist%>" = "Yes" Then
        .ggoSpread.Source    = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '�� : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		.lgStrPrevKey        = "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",2),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",3),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
		.frm1.vspdData.Redraw = True
		.DbQueryOk()
	End If
	End with
	
</Script>	

<%
	Response.End 
%>
