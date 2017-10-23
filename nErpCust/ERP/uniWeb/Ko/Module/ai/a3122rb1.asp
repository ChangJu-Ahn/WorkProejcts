<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2                     '�� : DBAgent Parameter ���� 
	Dim lgstrData                                                           '�� : data for spreadsheet data
	Dim lgStrPrevKey                                                        '�� : ���� �� 
	Dim lgMaxCount                                                          '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim iPrevEndRow
	Dim iEndRow	
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
	Dim strFrPrDt	                                                           
	Dim strToPrDt
	Dim strFrPrNo	                                                           
	Dim strToPrNo
	Dim strBpCd
	Dim strBp_Sts	
    Dim strRcptType	                                                           
		                                                           '�� : ������ 
	Dim strCond

	Dim strMsgCd
	Dim strMsg1
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","QB")
    Call HideStatusWnd 


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0
    
    Call TrimData()
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
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
		rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
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
    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "A3104RA1"
    UNISqlId(1) = "ABPNM"
    UNISqlId(2) = "Commonqry"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1)  = strCond

    UNIValue(1,0) = " " & FilterVar(strBpCd, "''", "S") & " " 
    UNIValue(2,0) = "Select JNL_CD, JNL_NM from a_jnl_item where jnl_cd= " & FilterVar(strRcptType, "''", "S") & "" 
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
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
      
	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strBPCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBpCd_alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.value = ""
		End With
		</Script>
<%
    Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtBpCd.value = "<%=Trim(rs1(0))%>"
			.frm1.txtBpNm.value = "<%=Trim(rs1(1))%>"
		End With
		</Script>
<%
    End If
    rs1.close
	Set rs1 = Nothing 
	
	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strRcptType <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtRcptType_alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtRcptTypeNm.value = ""
		End With
		</Script>
<%
    Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtRcptType.value = "<%=Trim(rs2(0))%>"
			.frm1.txtRcptTypeNm.value = "<%=Trim(rs2(1))%>"
		End With
		</Script>
<%
    End If
    rs2.close
	Set rs2 = Nothing 
	
    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    End If
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strFrPrNo	   = UCase(Trim(Request("txtFrPrNo")))
     strToPrNo     = UCase(Trim(Request("txtToPrNo")))
     strFrPrDt     = UCase(Trim(UNIConvDate(Request("txtFrPrDt"))))
     strToPrDt     = UCase(Trim(UNIConvDate(Request("txtToPrDt"))))
     strRcptType   = UCase(Trim(Request("txtRcptType")))
     strBpCd	   = UCase(Trim(Request("txtBpCd")))
	 strBp_Sts	   = UCase(Trim(Request("txtChk_Bp")))
	
	  strCond = strCond & " and A.RCPT_DT >=  " & FilterVar(strFrPrDt , "''", "S") & "and A.RCPT_DT <=  " & FilterVar(strToPrDt , "''", "S") & ""
     
     If strFrPrNo <> "" Then strCond = strCond & " and A.Rcpt_No >=  " & FilterVar(strFrPrNo , "''", "S") & ""	 
     
     If strToPrNo <> "" Then strCond = strCond & " and A.Rcpt_No <=  " & FilterVar(strToPrNo , "''", "S") & ""
     
     If strRcptType <> "" Then  strCond = strCond & " and A.Rcpt_Type =  " & FilterVar(strRcptType , "''", "S") & "" 

	 If  strBp_Sts <>"TRUE" then
		If strBpCd <> "" Then
		 	strCond = strCond & " and A.BP_CD =  " & FilterVar(strBpCd , "''", "S") & ""
		 end if
	 else
		strCond = strCond & " and (A.BP_CD IS NULL or A.BP_CD='')"
     End if 
     
     strCond = strCond & " and A.RCPT_FG = " & FilterVar("S", "''", "S") & " "
     strCond = strCond & " and A.RCPT_INPUT_TYPE = " & FilterVar("RT", "''", "S") & " "
     
    '--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			Parent.Frm1.htxtFrPrDt.Value    = Parent.Frm1.txtFrPrDt.Text                  'For Next Search
			Parent.Frm1.htxtToPrDt.Value    = Parent.Frm1.txtToPrDt.Text
			Parent.Frm1.htxtFrPrNo.Value	= Parent.Frm1.txtFrPrNo.Value
			Parent.Frm1.htxtToPrNo.Value    = Parent.Frm1.txtToPrNo.Value
			Parent.Frm1.htxtBPCd.Value		= Parent.Frm1.txtBPCd.Value
			Parent.Frm1.htxtRcptType.Value		= Parent.Frm1.txtRcptType.Value
       End If
       'Show multi spreadsheet data from this line
			Parent.ggoSpread.Source  = Parent.frm1.vspdData
			Parent.frm1.vspdData.Redraw = False
			Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
			Parent.frm1.vspdData.Redraw = True	
			Parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		
			Parent.DbQueryOk
    End If   

</Script>	


