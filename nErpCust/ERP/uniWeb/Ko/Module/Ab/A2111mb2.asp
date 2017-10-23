<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next
Err.Clear 

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT

    Call HideStatusWnd()
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")              

    lgSelectList   = Request("lgSelectList")								'�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)				'�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")									'�� : Orderby value
    
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  ColCnt
    Dim  iRowStr
    
    lgstrData   = ""

    Do While Not (rs0.EOF Or rs0.BOF)
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
				
		lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        rs0.MoveNext
	Loop

	rs0.Close
    Set rs0 = Nothing 
End Sub
    
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub  FixUNISQLData()
    Redim UNISqlId(0)														'��: SQL ID ������ ���� ����Ȯ�� 

    Redim UNIValue(0,2)

    UNISqlId(0) = "A2111MA101"

    UNIValue(0,0) = lgSelectList											'��: Select list
    UNIValue(0,1) = FilterVar(UCase(Request("txtGlCtrlFld")), "''", "S")
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"									'��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub  QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    Set lgADF = Nothing
       
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then        
        rs0.Close
        Set rs0 = Nothing
        Response.End														'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

%>

<Script Language=vbscript>
    With parent
        .ggoSpread.Source    = .frm1.vspdData2 
		Parent.frm1.vspdData.Redraw = False
		Parent.ggoSpread.SSShowData "<%=lgstrData%>"						'�� : Display data
		Parent.frm1.vspdData.Redraw = True
        .DbQueryOk(2)
	End with
</Script>	

