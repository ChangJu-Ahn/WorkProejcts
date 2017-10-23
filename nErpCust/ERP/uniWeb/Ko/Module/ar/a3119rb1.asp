<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1                  '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
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

Dim strCond

Dim strMsgCd
Dim strMsg1

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd	' ����� 
Dim lgInternalCd	' ���κμ� 
Dim lgSubInternalCd	' ���κμ�(��������)
Dim lgAuthUsrID		' ���� 
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"

    iPrevEndRow = 0
    iEndRow = 0

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	    
    strFrPrNo	   = UCase(Trim(Request("txtFrPrNo")))                                                          
    strToPrNo     = UCase(Trim(Request("txtToPrNo")))     
    strFrPrDt     = UCase(Trim(UNIConvDate(Request("txtFrPrDt"))))
    strToPrDt     = UCase(Trim(UNIConvDate(Request("txtToPrDt"))))
    strBPCd		  = UCase(Trim(Request("txtBpCd")))  

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

    Redim UNISqlId(1)                                                    '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(1,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "A3119RA101"
	UNISqlId(1) = "ABPNM"
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = strCond
    UNIValue(1,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & "" )    
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
    
	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strBpCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBPCd_Alt")
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
			.frm1.txtBPCd.value = "<%=Trim(rs1(0))%>"
			.frm1.txtBPNm.value = "<%=Trim(rs1(1))%>"
		End With
		</Script>
<%
    End If
    
	Set rs1 = Nothing 
        
    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
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
 Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
   

	 strCond = strCond & " and A.CONF_FG = " & FilterVar("C", "''", "S") & "  and (A.RCPT_STS = " & FilterVar("O", "''", "S") & "  OR A.RCPT_NO IN(SELECT distinct RCPT_NO FROM A_RCPT_ADJUST)) "
	 
     If strFrPrNo <> "" Then
		strCond = strCond & " and A.RCPT_NO >=  " & FilterVar(strFrPrNo , "''", "S") & ""	 
     End If
     
     If strToPrNo <> "" Then
		strCond = strCond & " and A.RCPT_NO <=  " & FilterVar(strToPrNo , "''", "S") & ""
     End If
         
     If strToPrDt <> "" Then
		strCond = strCond & " and A.RCPT_DT <=  " & FilterVar(strToPrDt , "''", "S") & ""
     End If
     
     If strFrPrDt <> "" Then
		strCond = strCond & " and A.RCPT_DT >=  " & FilterVar(strFrPrDt , "''", "S") & ""
     End If  
     
     If strBpCd <> "" Then
		strCond = strCond & " and A.BP_CD =  " & FilterVar(strBPCd , "''", "S") & ""
     End If

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If        
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			Parent.Frm1.htxtFrPrDt.Value    = Parent.Frm1.txtFrPrDt.Text
			Parent.Frm1.htxtToPrDt.Value    = Parent.Frm1.txtToPrDt.Text
			Parent.Frm1.htxtFrPrNo.Value	= Parent.Frm1.txtFrPrNo.Value
			Parent.Frm1.htxtToPrNo.Value	= Parent.Frm1.txtToPrNo.Value
			Parent.Frm1.htxtBPCd.Value		= Parent.Frm1.txtBPCd.Value
       End If
			'Show multi spreadsheet data from this line
			Parent.ggoSpread.Source			= Parent.frm1.vspdData
			Parent.frm1.vspdData.Redraw = False
			Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",8),Parent.GetKeyPos("A",10),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",8),Parent.GetKeyPos("A",12),"A", "I" ,"X","X")
			Parent.frm1.vspdData.Redraw = True
			Parent.lgPageNo					=  "<%=lgPageNo%>"               '�� : Next next data tag

			Parent.DbQueryOk
    End If   

</Script>	

