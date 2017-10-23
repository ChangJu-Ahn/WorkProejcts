<%'======================================================
'********************************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5112RA7
'*  4. Program Name         : ����ä�ǳ�����Ȳ ���� 
'*  5. Program Desc         :
'*  6. Comproxy List        : S51128ListBillDtlSvr
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/05/03
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Kwak Eunkyoung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd ȭ�� layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� layout
'*                            -2001/12/19 : Date ǥ������ 
'*                            -2002/05/03 : ADO��ȯ 
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0	           '�� : DBAgent Parameter ���� 
   Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
   Dim lgMaxCount                                              '�� : Spread sheet �� visible row �� 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
   'Dim strItemNm

	Dim iFrPoint
	iFrPoint=0

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")

    Call HideStatusWnd 
    
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = 100							                       '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 

    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
 
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint		= CLng(lgMaxCount) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(0)
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
    UNISqlId(0) = "s5112ra701" 
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

'--------------- ������ coding part(�������,Start)----------------------------------------------------
	strVal = " "
				
	If Len(Request("txtConBillNo")) Then
		strVal = FilterVar(Trim(Request("txtConBillNo")), "''", "S")
	Else
		strVal = "''"
	End If
'--------------- ������ coding part(�������,End)----------------------------------------------------
    UNIValue(0,1) = strVal  

'================================================   
   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
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
       'Show multi spreadsheet data from this line
		With parent       
			.ggoSpread.Source  = .frm1.vspdData
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'�� : Display data
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",1),"C","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",2),"A","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",3),"A","Q","X","X")
			.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
 			.DbQueryOk
			.frm1.vspdData.Redraw = True        
		End with
    End If   

</Script>	
