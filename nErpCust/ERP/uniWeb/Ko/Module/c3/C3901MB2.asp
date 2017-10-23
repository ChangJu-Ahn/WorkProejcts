<% Option Explicit%>

<%  Response.Expires = -1 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call loadInfTB19029B("Q", "C","NOCOOKIE","QB")
Call LoadBasisGlobalInf()

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2						'�� : DBAgent Parameter ���� 
Dim lgstrData, lgstrData_C                                                              '�� : data for spreadsheet data
Dim lgTailList, lgTailList_C                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList, lgSelectList_C
Dim lgSelectListDT, lgSelectListDT_C
Dim lgDataExist, lgDataExist_C
Dim lgPageNo, lgPageNo_C
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim lgAllcAmt
													'�� : ����� 
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"

    lgPageNo_C			= UNICInt(Trim(Request("lgPageNo_C")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList_C		= Request("lgSelectList_C")                               '�� : select ����� 
    lgSelectListDT_C	= Split(Request("lgSelectListDT_C"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList_C		= Request("lgTailList_C")                                 '�� : Orderby value
    lgDataExist_C		= "No"

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

    Const C_SHEETMAXROWS_D  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
    
    lgstrData = ""
    lgDataExist    = "Yes"

    If CInt(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CInt(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

Sub MakeSpreadSheetData_C()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr

    Const C_SHEETMAXROWS_C  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 

    lgstrData_C = ""
    lgDataExist_C    = "Yes"

    If CInt(lgPageNo_C) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_C * CInt(lgPageNo_C)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
    Do while Not (rs2.EOF Or rs2.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT_C) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT_C(ColCnt),rs2(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_C Then
            lgstrData_C      = lgstrData_C      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo_C = lgPageNo_C + 1
            Exit Do
        End If
        rs2.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_C Then                                            '��: Check if next data exists
        lgPageNo_C = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs2.Close
    Set rs2 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,4)

    UNISqlId(0) = "C3901MA02"	'sheet2 ȭ�� 
    UNISqlId(1) = "C3901MA02"	'Allc Amt
    UNISqlId(2) = "C3901MA04"	'sheet3 ȭ�� 

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    UNIValue(1,0) = "SUM(a.DIFF_AMT)"
    UNIValue(2,0) = lgSelectList_C
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = FilterVar(Trim(Request("txtYyyymm"))	,"' '","S")
    UNIValue(0,2) = FilterVar(Trim(Request("txtPlantCd"))	,"' '","S")
    UNIValue(0,3) = FilterVar(Trim(Request("txtItemCd"))	,"' '","S")

    UNIValue(1,1) = FilterVar(Trim(Request("txtYyyymm"))	,"' '","S")
    UNIValue(1,2) = FilterVar(Trim(Request("txtPlantCd"))	,"' '","S")
    UNIValue(1,3) = FilterVar(Trim(Request("txtItemCd"))	,"' '","S")

    UNIValue(2,1) = FilterVar(Trim(Request("txtYyyymm"))	,"' '","S")
    UNIValue(2,2) = FilterVar(Trim(Request("txtPlantCd"))	,"' '","S")
    UNIValue(2,3) = FilterVar(Trim(Request("txtItemCd"))	,"' '","S")

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNIValue(2,UBound(UNIValue,2)) = UCase(Trim(lgTailList_C))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

	lgAllcAmt		= 0

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        rs1.Close
        Set rs0 = Nothing
        Set rs1 = Nothing
    Else
		lgAllcAmt = rs1(0)

        Call  MakeSpreadSheetData()
    End If

    If  rs2.EOF And rs2.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs2.Close
        Set rs2 = Nothing
    Else
        Call  MakeSpreadSheetData_C()
    End If
End Sub

%>

		    
<Script Language=vbscript>
	If "<%=lgDataExist%>" = "Yes" Then

		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
		End If

		Parent.frm1.txtAllcAmt.text		= "<%=UniNumClientFormat(lgAllcAmt,ggAmtOfMoney.Decpoint,0)%>" 

		Parent.ggoSpread.Source  = Parent.frm1.vspdData1
		Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
		Parent.lgPageNo_B      =  "<%=lgPageNo%>"               '�� : Next next data tag

		Parent.ggoSpread.Source  = Parent.frm1.vspdData2
		Parent.ggoSpread.SSShowData "<%=lgstrData_C%>"                  '�� : Display data
		Parent.lgPageNo_C      =  "<%=lgPageNo_C%>"               '�� : Next next data tag

		Parent.DbQueryOk("2")
	End If   
</Script>	

