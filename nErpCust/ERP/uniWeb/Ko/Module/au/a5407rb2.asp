<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>


<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2000/11/01
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

Response.Buffer = True                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("*", "A", "NOCOOKIE", "RB")
Call HideStatusWnd 


Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3               '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim txtFromDt
Dim txtToDt
Dim txtDocCur
Dim txtDeptCd
Dim txtCardCoCd
Dim txtCardNo
Dim txtGlNoSeq
Dim txtMaxRows
Dim strWHERESQL

Dim LngMaxRow
Dim iLoopCount

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
 
	lgPageNo = 0
    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                               '�� : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
     
          
     
    LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

    Call TrimData()  
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
   
   If Request("txtFromDt") <> "" Then
	txtFromDt		= Request("txtFromDt")
   End If
   	 	
	txtToDt			= Request("txtToDt")	
	txtDocCur		= Request("txtDocCur")
	txtDeptCd		= Request("txtDeptCd")
	txtCardCoCd		= Request("txtCardCoCd")
	txtCardNo		= Request("txtCardNo")
	txtGlNoSeq		= Trim(Request("txtGlNoSeq"))
	txtMaxRows		= Request("txtMaxRows")
	
	strWHERESQL = ""
	
	If txtFromDt <> "" Then     strWHERESQL = strWHERESQL & " AND A.GL_DT >=  " & FilterVar(txtFromDt , "''", "S") & " "
	If txtToDt   <> "" Then     strWHERESQL = strWHERESQL & " AND A.GL_DT <=  " & FilterVar(txtToDt , "''", "S") & " " 
	If txtDeptCd <> "" Then		strWHERESQL = strWHERESQL & " AND B.DEPT_CD =  " & FilterVar(txtDeptCd , "''", "S") & " "
	If txtCardNo <> "" Then		strWHERESQL = strWHERESQL & " AND D.CREDIT_NO =  " & FilterVar(txtCardNo , "''", "S") & " "
	If txtCardCoCd <> "" Then	strWHERESQL = strWHERESQL & " AND D.CARD_CO_CD =  " & FilterVar(txtCardCoCd , "''", "S") & " "
	

End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	
	Dim ii
	Dim iArrTemp
	Dim iArrTemp2
	Dim iStrWhere
	iStrWhere = ""
    
    Redim UNISqlId(0)                                                    '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    
    Redim UNIValue(0,6)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "a5407ra201"
     
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    
    'rs0�� ���� Value�� setting 
       
    UNIValue(0,0)	= lgSelectList  
'  	UNIValue(0,1)	= txtFromDt	
'	UNIValue(0,2)	= txtToDt
	UNIValue(0,1)	= FilterVar(txtDocCur,"''","S")
	UNIValue(0,2)	= strWHERESQL
	
	iArrTemp = split(txtGlNoSeq, gRowSep)		
	For ii = 0 To Ubound(iArrTemp,1) - 1
		iArrTemp2 = split(iArrTemp(ii),gColSep)			
		
		If Trim(iArrTemp2(0)) <> "" And Trim(iArrTemp2(1)) <> "" Then
			iStrWhere = iStrWhere & " (a.gl_no <>  " & FilterVar(iArrTemp2(0), "''", "S") & " or a.gl_seq <> " & iArrTemp2(1) & ") and "			 
		End If		
	Next
	
	If InStr(1,iStrWhere, "and") > 0 Then
		iStrWhere = Mid(iStrWhere,1,InStrRev(iStrWhere, "and") -1)	
		iStrWhere = "and	( " & iStrWhere & " ) "
	End If
	
	UNIValue(0,3)	= iStrWhere
	UNIValue(0,4)	= lgTailList
    '--------------- ������ coding part(�������,End)------------------------------------------------------
       
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode



End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMsgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------
' MakeSpreadSheetData
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If Len(Trim(lgPageNo)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          lgPageNo = CInt(lgPageNo)
       End If   
    Else   
       lgPageNo = 0
    End If      
    'rs0�� ���� ��� 
    rs0.PageSize     = lgMaxCount                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
		
        iLoopCount =  iLoopCount + 1
        iRowStr = ""     
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 		
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))'rs0(ColCnt)'
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

End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then		
       'Set condition data to hidden area
       With parent
			If "<%=lgPageNo%>" = "1" Or "<%=lgPageNo%>" = ""  Then   ' "1" means that this query is first and next data exists
					.Frm1.hFromDt.Value		= .Frm1.txtFromDt.text
					.Frm1.hToDt.Value		= .Frm1.txtToDt.text
					.Frm1.hDocCur.Value		= .Frm1.txtDocCur.Value
					.Frm1.hDeptCd.Value		= .Frm1.txtDeptCd.Value
					.Frm1.hCardCoCd.Value   = .Frm1.txtCardCoCd.Value
					.Frm1.hCardNo.Value     = .Frm1.txtCardNo.Value					
			End If
        'Show multi spreadsheet data from this line       
        .ggoSpread.Source	= .frm1.vspdData
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",7),parent.GetKeyPos("A",9),   "A" ,"I","X","X")
        .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag
		.DbQueryOk
		.frm1.vspdData.Redraw = True
	   End With
	End if

</Script>	
