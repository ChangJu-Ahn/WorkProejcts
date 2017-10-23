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
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6            '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strFromDt																'�� : ������ 
Dim strToDt																	'�� : ������ 
Dim strDealBpCd																'�� : �μ� 
Dim strPayBpCd																'�� : �ŷ�ó 
Dim strAcctCd																'�� : �����ڵ� 
Dim strDesc																	'�� : ��� 
Dim strApNo																	'�� : ä����ȣ 
Dim strRefNo																'�� : ������ȣ 
Dim strInvDocNo																'�� : �����ȣ 
Dim strCond
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1

Dim iPrevEndRow
Dim iEndRow	

Dim strMsgCd
Dim strMsg1

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- ������ coding part(��������,End)----------------------------------------------------------
	  
	Call HideStatusWnd()
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "A","NOCOOKIE","MB")              

	lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	    
	lgSelectList	= Request("lgSelectList")                               '�� : select ����� 
	lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgTailList		= Request("lgTailList")                                 '�� : Orderby value
	lgDataExist		= "No"
	iPrevEndRow		= 0
	iEndRow			= 0
	        
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

Sub  MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 30
    
    lgstrData = ""

    lgDataExist    = "Yes"

    If CInt(lgPageNo) > 0 Then
		iPrevEndRow =  C_SHEETMAXROWS_D * CDbl(lgPageNo)    
       rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo

    End If
       
    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
				
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
    	
	rs0.Close
    Set rs0 = Nothing                                                 '��: ActiveX Data Factory Object Nothing
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub  FixUNISQLData()

    Redim UNISqlId(6)
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(6,10)

    UNISqlId(0) = "A4114MA101"
	UNISqlId(1) = "ABPNM"
    UNISqlId(2) = "ABPNM"
    UNISqlId(3) = "A7116MA102"
    UNISqlId(4) = "A4114MA103"
    UNISqlId(5) = "A_GETBIZ"
    UNISqlId(6) = "A_GETBIZ"
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = FilterVar(strDealBpCd, "''", "S")
    UNIValue(2,0) = FilterVar(strPayBpCd, "''", "S")
    UNIValue(3,0) = FilterVar(strAcctCd, "''", "S") 
    
    UNIValue(4,0) = strCond
    UNIValue(5,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(6,0)  = FilterVar(strBizAreaCd1, "''", "S")
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub  QueryData()
    Dim iStr

	strMsgCd = ""
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)
    Set lgADF = Nothing   
        
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strDealBpCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDealBpCd_Alt")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent		
		.frm1.txtDealBpCd.value = "<%=Trim(rs1(0))%>"
		.frm1.txtDealBpNm.value = "<%=Trim(rs1(1))%>"					
	End With
	</Script>
<%
    End If
    
	rs1.Close
	Set rs1 = Nothing
    
    If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strPayBpCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtPayBpCd_Alt")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtPayBpCd.value = "<%=Trim(rs2(0))%>"
		.frm1.txtPayBpNm.value = "<%=Trim(rs2(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs2.Close
	Set rs2 = Nothing
    
    If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And strAcctCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtAcctCd_Alt")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtAcctNm.value = "<%=Trim(rs3(0))%>"					
	End With
	</Script>
<%
    End If
	
	rs3.Close
	Set rs3 = Nothing
	'//////
	If (rs4.EOF And rs4.BOF) Then	
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtTotApLocAmt.Text = "<%=UNINumClientFormat(Trim(rs4(0)), ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtTotClsLocAmt.Text = "<%=UNINumClientFormat(Trim(rs4(1)), ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtTotBalLocAmt.Text = "<%=UNINumClientFormat(Trim(rs4(2)), ggAmtOfMoney.DecPoint, 0)%>"
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
	
	
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, "", I_MKSCRIPT)
		rs0.Close
        Set rs0 = Nothing
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	   
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strFromDt		= UNIConvDate(Trim(Request("txtFromDt")))         '�������� 
     strToDt		= UNIConvDate(Trim(Request("txtToDt")))           '�������� 
     strDealBpCd	= Trim(UCase(Request("txtDealBpCd")))	                 '�μ��ڵ� 
     strPayBpCd		= Trim(UCase(Request("txtPayBpCd")))                  '�ŷ�ó�ڵ�     
     strAcctCd		= Trim(UCase(Request("txtAcctCd")))                   '�����ڵ�     
     strDesc		= Trim(UCase(Request("txtDesc")))                 '��� 
     strArNo		= Trim(UCase(Request("txtApNo")))                 'ä����ȣ 
     strRefNo		= Trim(UCase(Request("txtRefNo")))                '������ȣ 
     strInvDocNo	= Trim(UCase(Request("txtInvDocNo")))             '�����ȣ 
     strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))            '�����From
	 strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))            '�����To

	strCond= " AND AP_DT >= " & FilterVar(strFromDt, "''", "S") & " AND AP_DT <= " & FilterVar(strToDt, "''", "S")
		
	If strDealBpCd <> ""	Then strCond = strCond & " AND DEAL_BP_CD = "			& FilterVar(strDealBpCd , "''", "S") 
	If strPayBpCd <> ""		Then strCond = strCond & " AND PAY_BP_CD = "			& FilterVar(strPayBpCd , "''", "S") 
	If strAcctCd <> ""		Then strCond = strCond & " AND A_OPEN_AP.ACCT_CD = "	& FilterVar(strAcctCd , "''", "S") 
	If strDesc <> ""		Then strCond = strCond & " AND A_OPEN_AP.AP_DESC = "	& FilterVar(strDesc , "''", "S") 
	If strArNo <> ""		Then strCond = strCond & " AND A_OPEN_AP.AP_NO = "		& FilterVar(strArNo , "''", "S") 
	If strRefNo <> ""		Then strCond = strCond & " AND A_OPEN_AP.REF_NO = "		& FilterVar(strRefNo , "''", "S") 
	If strInvDocNo <> ""	Then strCond = strCond & " AND A_OPEN_AP.INV_DOC_NO = "	& FilterVar(strInvDocNo , "''", "S") 
	
	if strBizAreaCd <> "" then
		strCond = strCond & " AND A_OPEN_AP.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strCond = strCond & " AND A_OPEN_AP.BIZ_AREA_CD >= " & FilterVar("", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strCond = strCond & " AND A_OPEN_AP.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strCond = strCond & " AND A_OPEN_AP.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if
	

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND A_OPEN_AP.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND A_OPEN_AP.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND A_OPEN_AP.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND A_OPEN_AP.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	strCond		= strCond	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

    '--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
    
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists

       End If
    
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",4),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",5),"A", "Q" ,"X","X")
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo_A      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk(1)
    End If  
</Script>	

