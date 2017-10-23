<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<% 

Err.Clear
On Error Resume Next


Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")
Call LoadBNumericFormatB("Q", "A","NOCOOKIE","RB")
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo


Dim lgBatchNo
Dim lgSeq
Const C_SHEETMAXROWS_D = 30
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPlantCd																'�� : ���� 
Dim strFromDt																'�� : ������ 
Dim strToDt																	'�� : ������ 
Dim strItemCd																'�� : ǰ�� 
Dim strRoutNo																'�� : ����� 
'--------------- ������ coding part(��������,End)----------------------------------------------------------

    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    lgBatchNo = Trim(Request("txtBatchNo"))
	lgSeq = Trim(Request("txtSeq"))

    Call FixUNISQLData()
    Call QueryData()

'==========================================================================================
' Query Data
'==========================================================================================
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr

    lgstrData = ""
    lgDataExist    = "Yes"

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
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If

	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim lstrsql
    Redim UNISqlId(0)

	lstrsql = ""
	lstrsql = " A.BATCH_NO =  " & FilterVar(lgBatchNo, "''", "S") & "  "
	lstrsql = lstrsql & " And A.SEQ =  " & FilterVar(lgSeq  , "''", "S") & " "

	Redim UNIValue(0,3)

    UNISqlId(0) = "A5140RA103"


    UNIValue(0,0) = lgSelectList                                          '��: Select list
    UNIValue(0,1) = lstrsql


    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub
'==========================================================================================
' Query Data
'==========================================================================================
Sub QueryData()
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set lgADF = Nothing

    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If  rs0.EOF And rs0.BOF Then
		'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else
        Call  MakeSpreadSheetData()
    End If
End Sub

%>

<Script Language=vbscript>
	If "<%=lgDataExist%>" = "Yes" Then
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
		End If

		Parent.ggoSpread.Source  = Parent.frm1.vspdData2
		Parent.ggoSpread.SSShowData "<%=lgstrData%>"
		Parent.lgPageNo_B      =  "<%=lgPageNo%>"               '�� : Next next data tag
		Parent.DbQueryOk(2)
	End If
</Script>

