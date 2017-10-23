<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3104mb1
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
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4          '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strBizAreaCd															'�� : ����� 
Dim strBizAreaCd1															'�� : �����1
Dim strDpstFg																'�� : �����ݱ��� 
Dim strDpstType																'�� : ���������� 
Dim strBankCd																'�� : ���� 
Dim strTransSts																'�� : �ŷ����� 
Dim strDocCur																'�� : ��ȭ 
Dim strWhere																'�� : Where ���� 
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
            lgStrPrevKey = rs0(2)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                   '��: Check if next data exists
        lgPageNo = ""              
        lgStrPrevKey = ""                                   '��: ���� ����Ÿ ����.
    End If
  	
'	rs0.Close
'    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(4,2)

    UNISqlId(0) = "F3104MA101"
    UNISqlId(1) = "F3104MA105"	'������ڵ� 
    UNISqlId(2) = "F3104MA106"	'�����ڵ� 
    UNISqlId(3) = "F3104MA107"	'�����ڵ� 
    UNISqlId(4) = "F3104MA105"	'������ڵ�1
    '--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = lgSelectList                                          '��: Select list
    
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = UCase(Trim(strWhere))
    UNIValue(1,0) = FilterVar(strBizAreaCd , "''", "S")
    UNIValue(2,0) = FilterVar(strBankCd , "''", "S") 
    UNIValue(3,0) = FilterVar(strDocCur , "''", "S")
    UNIValue(4,0) = FilterVar(strBizAreaCd1 , "''", "S")
    
	
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
	
	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strBizAreaCd <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBizAreaCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
		End With
		</Script>
<%		
	End If
	
	rs1.Close
	Set rs1 = Nothing
	
	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strBankCd <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBankCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBankCd.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
			.txtBankNm.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
		End With
		</Script>
<%
	End If
	
	rs2.Close
	Set rs2 = Nothing
	
	If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And strDocCur <> "" Then
			strMsgCd = "970000"
			strMsg1 = Request("txtDocCur_Alt")
		End If
	End If
	
	rs3.Close
	Set rs3 = Nothing
	
	If (rs4.EOF And rs4.BOF) Then
		If strMsgCd = "" And strBizAreaCd1 <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBizAreaCd_Alt1")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBizAreaCd1.value = "<%=ConvSPChars(Trim(rs4(0)))%>"
			.txtBizAreaNm1.value = "<%=ConvSPChars(Trim(rs4(1)))%>"
		End With
		</Script>
<%		
	End If
	
	rs4.Close
	Set rs4 = Nothing
	
    If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
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
	
	strBankCd		= UCase(Trim(Request("txtBankCd")))
	strDpstType		= UCase(Trim(Request("cboDpstType")))
	strTransSts		= UCase(Trim(Request("cboTransSts")))
	strDocCur		= UCase(Trim(Request("txtDocCur")))
	
	strWhere = ""
	
	if strBizAreaCd <> "" then
		strWhere = strWhere & " and A.biz_area_cd >=  " & FilterVar(strBizAreaCd , "''", "S") & ""
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere = strWhere & " and A.biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""
	end if
	
	If strBankCd   <> "" Then strWhere = strWhere & " and A.bank_cd   =  " & FilterVar(strBankCd , "''", "S") & " "
	If strDpstType <> "" Then strWhere = strWhere & " and A.dpst_type =  " & FilterVar(strDpstType , "''", "S") & " "
	If strTransSts <> "" Then strWhere = strWhere & " and A.trans_sts =  " & FilterVar(strTransSts , "''", "S") & " "
	If strDocCur   <> "" Then strWhere = strWhere & " and A.doc_cur   =  " & FilterVar(strDocCur , "''", "S") & " "
	
	' ���Ѱ��� �߰� 
'	If lgAuthBizAreaCd <> "" Then
'		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
'	End If
'	
'	If lgInternalCd <> "" Then
'		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
'	End If
'	
'	If lgSubInternalCd <> "" Then
'		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
'	End If
'	
'	If lgAuthUsrID <> "" Then
'		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
'	End If
'	
'	' ���Ѱ��� �߰� 
'	strWhere	= strWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	
	
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
    With parent
        .ggoSpread.Source     = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '�� : Display data
		.lgPageNo_A           =  "<%=lgPageNo%>"               '�� : Next next data tag
		.lgStrPrevKey_A       = "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",3),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
		.DbQueryOk()
		.frm1.vspdData.Redraw = True
	End with
</Script>	

