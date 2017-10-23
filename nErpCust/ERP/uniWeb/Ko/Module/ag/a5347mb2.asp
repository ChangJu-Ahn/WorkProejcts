<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->

<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
    Call loadInfTB19029B("Q", "*","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "*", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
    Dim lgstrData                                                              '�� : data for spreadsheet data
    Dim lgStrPrevKey                                                           '�� : ���� �� 
    Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgDataExist
    Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
    Dim iPrevEndRow
    Dim iEndRow
    Dim strSql
    Dim strSql2
    Dim strSql3
	Dim lgtxtTransDt
	Dim lgtxtBizArea
	Dim lgtxtTransType
	Dim lgFlag
	Dim bAbNomal
    
'--------------- ������ coding part(��������,End)----------------------------------------------------------
    Call HideStatusWnd
    lgPageNo       = UNICInt(Trim(Request("lgPageNo_B")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList_B")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT_B"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList_B")                                 '�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0
    
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
Sub TrimData()
	Dim strSqlTemp,strSqlTemp2 

	lgFlag	= Trim(Request("RADIO"))
	lgtxtBatchNo		= FilterVar(Trim(Request("txtBatchNo")),"","S")
	lgtxtMaxRows		= Request("txtMaxRows")	
	strSql = strSql & "  WHERE  A.BATCH_NO  =  " & FilterVar(lgtxtBatchNo , "''", "S") & " "	
'	Call SvrMsgBox(strSql , vbInformation, I_MKSCRIPT)
	
End Sub
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 30     
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo

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
        iEndRow = iPrevEndRow + iLoopCount + 1
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
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

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,3)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

	UNISqlId(0) = "A5347MA102"
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList

'    Call SvrMsgBox(strSql , vbInformation, I_MKSCRIPT)

  	UNIValue(0,1)  = strSql	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    If  rs0.EOF And rs0.BOF Then
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData1
       Parent.frm1.vspdData1.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                  '�� : Display data       
       Parent.lgPageNo_B      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk2
       Parent.frm1.vspdData1.Redraw = True
    End If   

</Script>	

