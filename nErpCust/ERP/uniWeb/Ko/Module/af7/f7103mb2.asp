<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Query)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2001/01/15
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : Kim Hee Jung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

                                                       '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
DIm lgMaxCount
Dim iPrevEndRow
Dim iEndRow	

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPrrcptNo																'�� : ���� 

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
	Call HideStatusWnd
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("*", "A", "NOCOOKIE", "MB")   'ggQty.DecPoint Setting... 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"													'�� : Orderby value
    iPrevEndRow = 0
    iEndRow = 0
        
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim RecordCnt
    Dim ColCnt
    Dim iLoopCount
    Dim iRowStr

    lgstrData = ""

    lgDataExist    = "Yes"

    If CInt(lgPageNo) > 0 Then
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

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,2)

    UNISqlId(0) = "a7103ma2"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = Trim(strPrrcptNo)
    
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
    
    iStr = Split(lgstrRetMsg,gColSep)
    Set lgADF = Nothing 

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
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
	 strPrrcptNo    = FilterVar(Request("txtPrrcptNo"), "''","S")	 
End Sub

%>

<Script Language=vbscript>
   If "<%=lgDataExist%>" = "Yes" Then
		With parent
			.ggoSpread.Source  = .frm1.vspdData2
			Parent.frm1.vspdData2.Redraw = False
			Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData2,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("B",1),"A", "Q" ,"X","X")
			Parent.frm1.vspdData2.Redraw = True
			.lgPageNo_B      =  "<%=lgPageNo%>"               '�� : Next next data tag
			.DbQueryOk("2")
		End with
	End If 
</Script>	

