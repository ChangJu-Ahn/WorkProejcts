

<%@ LANGUAGE=VBSCript %>

<%Option Explicit%>
<!-- #Include file="../inc/incSvrMain.asp" -->
<!-- #Include file="../inc/incSvrDate.inc" -->
<!-- #Include file="../inc/incSvrNumber.inc" -->
<!-- #Include file="../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
<% 

Err.Clear
On Error Resume Next


Const C_SHEETMAXROWS_D  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo


Dim lgTempGlNo
Dim lgSeq

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPlantCd																'�� : ���� 
Dim strFromDt																'�� : ������ 
Dim strToDt																	'�� : ������ 
Dim strItemCd																'�� : ǰ�� 
Dim strRoutNo																'�� : ����� 
'--------------- ������ coding part(��������,End)----------------------------------------------------------

	Call LoadBasisGlobalInf()
	Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")
    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D										   '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    lgTempGlNo = Trim(Request("txtTempGlNo"))
	lgSeq = Trim(Request("txtSeq"))	
	
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

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,3)

    UNISqlId(0) = "a5130ra103"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = FilterVar(lgTempGlNo,"''","S")
    UNIValue(0,2) = lgSeq
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
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
          Parent.Frm1.htxtSchoolCd.Value  = Parent.Frm1.txtSchoolCd.Value                  'For Next Search
       End If
    
       Parent.ggoSpread.Source  = Parent.frm1.vspdData2
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
       Parent.lgPageNo_B      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk(2)
    End If   
</Script>	

