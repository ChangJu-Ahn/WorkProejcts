<%'======================================================================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���԰��� 
'*  3. Program ID           : M5111MB1
'*  4. Program Name         : �������� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/03/21
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : Oh chang won
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date ǥ������ 
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim iTotstrData

Dim strPoType	                                                           '�� : �������� 
Dim strPoFrDt	                                                           '�� : ������ 
Dim strPoToDt	                                                           '�� :
Dim strSpplCd	                                                           '�� : ����ó 
Dim strPurGrpCd	                                                           '�� : ���ű׷� 
Dim strItemCd	                                                           '�� : ǰ�� 
Dim strTrackNo	                                                           '�� : Tracking No
Dim BlankchkFlg

Dim arrRsVal(5)								'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array

Dim iPrevEndRow
Dim iEndRow
    
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	
	iPrevEndRow = 0
    iEndRow = 0
    
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
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100 
        
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                 

    End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
            PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")

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
    Dim strVal
    Dim arrVal(3)															
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(1,2)

    UNISqlId(0) = "M5111ra101"									'* : ������ ��ȸ�� ���� SQL�� ���� 
 
    UNIValue(0,0) = lgSelectList                                          '��: Select list

	strVal = " "
    

	If Trim(Request("txtIvNo")) <> "" Then
		strVal = " AND A.IV_NO >= " & FilterVar(UCase(Request("txtIvNo")), "''", "S") & "  AND A.IV_NO <=  " & FilterVar(UCase(Request("txtIvNo")), "''", "S") & " "
	Else
		strVal = ""
	End If

    'strVal =strVal & " AND A.POSTED_FLG = 'Y' "

  	If Len(Trim(Request("txtFrIvDt"))) Then
		strVal = strVal & " AND A.IV_DT >= " & FilterVar(UNIConvDate(Request("txtFrIvDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToIvDt"))) Then
		strVal = strVal & " AND A.IV_DT <= " & FilterVar(UNIConvDate(Request("txtToIvDt")), "''", "S") & ""		
	End If

    If Trim(Request("hdnRefPoNo")) <> "" Then
		strVal = strVal & " AND B.PO_NO >= " & FilterVar(UCase(Request("hdnRefPoNo")), "''", "S") & "  AND B.PO_NO <=  " & FilterVar(UCase(Request("hdnRefPoNo")), "''", "S") & " "
    End If
    
    If Trim(Request("hdnCurr")) <> "" Then
        strVal = strVal & " AND A.IV_CUR = " & FilterVar(UCase(Request("hdnCurr")), "''", "S") & " "		
    ELSE
        strVal = strVal & " AND A.IV_CUR = " & FilterVar("zzzzzzzzz", "''", "S") & ""		
    End If
    
    If Trim(Request("hdnGroupCd")) <> "" Then
        strVal = strVal & " AND A.PUR_GRP = " & FilterVar(UCase(Request("hdnGroupCd")), "''", "S") & " "		
    ELSE
        strVal = strVal & " AND A.PUR_GRP = " & FilterVar("zzzzzzzzz", "''", "S") & ""		
    End If
	
    If Trim(Request("hdnSupplierCd")) <> "" Then
        strVal = strVal & " AND A.BP_CD = " & FilterVar(UCase(Request("hdnSupplierCd")), "''", "S") & " "		
    ELSE
        strVal = strVal & " AND A.BP_CD = " & FilterVar("zzzzzzzzz", "''", "S") & ""		
    End If

	' ���ְ������� �߰� 
	If Len(Request("txtSubcontraflg")) Then
		strVal = strVal & " AND C.SUBCONTRA_FLG = " & FilterVar(Trim(UCase(Request("txtSubcontraflg"))), " " , "S") & ""		
	End If

    UNIValue(0,1) = strVal   '---������ 

   
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0) '* : Record Set �� ���� ���� 
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
    
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	
End Sub


%>
<Script Language=vbscript>
    With parent
        .ggoSpread.Source    = .frm1.vspdData 
        Parent.frm1.vspdData.Redraw = False
        .ggoSpread.SSShowData "<%=iTotstrData%>", "F"                             '��: Display data 
        Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnCurr.value, Parent.GetKeyPos("A",8),"C", "I" ,"X","X")
        Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnCurr.value, Parent.GetKeyPos("A",9),"A", "I" ,"X","X")
        
        .lgPageNo				=  "<%=lgPageNo%>"				    '��: Next next data tag
		.frm1.hdnIvNo.value	    =  "<%=ConvSPChars(Request("txtIvNo"))%>" 	
  		.frm1.hdnFrIvDt.value   =  "<%=ConvSPChars(Request("txtFrIvDt"))%>" 	
  		.frm1.hdnToIvDt.value	=  "<%=ConvSPChars(Request("txtToIvDt"))%>" 	
        .DbQueryOk
        Parent.frm1.vspdData.Redraw = True
	End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

