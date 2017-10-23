
<%@ LANGUAGE=VBSCript %>


<%Option Explicit%>



<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/IncSvrNumber.inc" -->
<!-- #Include file="../inc/IncSvrDate.inc" -->
<!-- #Include file="../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServerAdoDb.asp" -->
<!-- #Include file="../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="LoadInfTB19029.asp" -->
<% 

Err.Clear
On Error Resume Next


Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                         '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPageNo
Dim lgtxtMaxRows
Dim iLoopCount, iEndRow
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strCond
Dim strHqBrchNo
Dim strMsgCd, strMsg1, strMsg2
Const C_SHEETMAXROWS_D =30
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

	lgPageNo       = Request("lgPageNo")
    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = C_SHEETMAXROWS_D                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgtxtMaxRows		= Request("txtMaxRows")
	
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""

	If Len(Trim(lgPageNo)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then          
          lgPageNo = CInt(lgPageNo)
       End If   
    Else   
       lgPageNo = 0
    End If   
    
    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   
	iLoopCount = 0

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '�� : Discard previous data
		iLoopCount = iRCnt
        rs0.MoveNext
    Next
	
	rs0.PageSize     = lgMaxCount
	rs0.AbsolutePage = lgPageNo + 1
	
    iRCnt = -1
	iEndRow = 0
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
		iEndRow = iEndRow + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
    If  iRCnt < lgMaxCount Then                                            '��: Check if next data exists
        lgStrPrevKey = ""          
        lgPageNo = ""                                        '��: ���� ����Ÿ ����.
    End If    
    
'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,2)

    UNISqlId(0) = "A5120RA201"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		strMsgCd = "900014"
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    strHqBrchNo  = Request("txtHqBrchNo")    

	strCond = ""
	
	strCond = strCond & " A.hq_brch_no = '" & strHqBrchNo & "' "			
	
End Sub

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData0 
         .ggoSpread.SSShowData "<%=lgstrData%>", "F"                            '��: Display data 
         
		 Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData0,<%=iLoopCount%>,<%=iLoopCount + iEndRow%>,.GetKeyPos("C",5),.GetKeyPos("C",3),"C", "Q" ,"X","X")
         
         .lgStrPrevKey        = "<%=lgStrPrevKey%>"    
         .lgPageNo_C		=  "<%=lgPageNo%>"               '�� : Next next data tag                   '��: set next data tag


         Call .DbQueryOk(0)
	End with
</Script>	

<%
	Response.End 
%>