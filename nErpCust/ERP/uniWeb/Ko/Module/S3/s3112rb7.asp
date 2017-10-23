<%
'********************************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : 
'*  3. Program ID           : S3112RB7
'*  4. Program Name         : ���ֳ�����Ȳ(������Ȳ��ȸ����) 
'*  5. Program Desc         : ���ֳ�����Ȳ(������Ȳ��ȸ����) 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Kim Hyung suk
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "RB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "RB")
On Error Resume Next

Call HideStatusWnd

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgPageNo                                                               '�� : ���� �� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList													       '�� : select ����� 
Dim lgSelectListDT														   '�� : �� �ʵ��� ����Ÿ Ÿ��	
Const C_SHEETMAXROWS_D  = 30  

lgPageNo   = UNICInt(Trim(Request("lgPageNo")),0)   
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgTailList     = Request("lgTailList")                                 '�� : Order by value

Call TrimData()
Call FixUNISQLData()
Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
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

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
																		  '�Ʒ��� ���� ȭ��ܿ��� �־� �ִ� query�� where�������� �� �� �ִ�.	
    Redim UNISqlId(0)                                                        '��: SQL ID ������ ���� ����Ȯ�� 
																		  '��ȸȭ�鿡�� �ʿ��� query���ǹ����� ����(Statements table�� ����)
    Redim UNIValue(0,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 

    UNISqlId(0) = "S3112RA701"  ' main query(spread sheet�� �ѷ����� query statement)
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtHConSoNo")) Then
		strVal = " " & FilterVar(Request("txtHConSoNo"), "''", "S") & " "
	Else
		strVal = ""
	End If
	
    
    UNIValue(0,1) = strVal    '	UNISqlId(0)�� �ι�° ?�� �Էµ�	
        
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	Dim FalsechkFlg
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'��:ADO ��ü�� ���� 
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    FalsechkFlg = False
    
    iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End     

    Else   

        Call  MakeSpreadSheetData()

    End If
   
End Sub


%>

<Script Language=vbscript>
    With parent
  
        .ggoSpread.Source    = .frm1.vspdData
        .frm1.vspdData.Redraw = False
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"                            '��: Display data 
        
        Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,-1,-1,parent.frm1.txtHCur.value,Parent.GetKeyPos("A",9),"C", "Q" ,"X","X")
        Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,-1,-1,parent.frm1.txtHCur.value,Parent.GetKeyPos("A",10),"A", "Q" ,"X","X")                
        
        .lgPageNo = "<%=lgPageNo%>"			
		.frm1.vspdData.Redraw = True
   	End with
</Script>	
 	
<%
	Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>