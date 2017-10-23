<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1	 , rs2                 '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iPrevEndRow
Dim iEndRow	
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
dim strPrDt
dim strToPrDt
dim strDeptCd
dim strBpCd
dim strDocCur
	                                                           '�� : ������ 
Dim strCond
Dim strMsgCd
Dim strMsg1

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","A","NOCOOKIE","RB")
    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0
      
    Call TrimData()
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
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
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
            lgstrData  = lgstrData  & iRowStr & Chr(11) & Chr(12)
            
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

    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "A3109RB110"
    UNISqlId(1) = "ADEPTNM"
	UNISqlId(2) = "ABPNM"
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1)  = strCond
    
    UNIValue(1,0) = FilterVar(strDeptCd, "''", "S")
	UNIValue(1,1) = FilterVar(UCase(Request("hOrgChangeId")), "''", "S")   

    UNIValue(2,0) = FilterVar(strBpCd, "''", "S")

   '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1 , rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strDeptCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDeptCd_Alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.txtDeptNm.value = ""
		End With
		</Script>
<%
    Else
%>
		<Script Language=vbScript>
		With parent
			.txtDeptCd.value = "<%=Trim(ConvSPChars(rs1(0)))%>"
			.txtDeptNm.value = "<%=Trim(ConvSPChars(rs1(1)))%>"
		End With
		</Script>
<%
    End If
	Set rs1 = Nothing 
        
    If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strBpCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBpCd_Alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.txtBpNm.value = ""
		End With
		</Script>
<%
    Else
%>
		<Script Language=vbScript>
		With parent
			.txtBpCd.value = "<%=Trim(ConvSPChars(rs2(0)))%>"
			.txtBpNm.value = "<%=Trim(ConvSPChars(rs2(1)))%>"
		End With
		</Script>
<%
    End If
	Set rs2 = Nothing 

    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
		exit Sub
    End If

    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Trim Data
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

	strPrDt		= UCase(Trim(UNIConvDate(Request("txtPrDt"))))
	strToPrDt	= UCase(Trim(UNIConvDate(Request("txtToPrDt"))))
		
	strDeptCd	= UCase(Trim(Request("txtDeptcd")))                                                          
	strBpCd     = UCase(Trim(Request("txtBpcd")))
	strDocCur   = UCase(Trim(Request("txtDocCur")))	

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))
		 
		
	strCond = strCond & " AND fp.PRRCPT_STS  = " & FilterVar("O", "''", "S") & "  AND fp.CONF_FG = " & FilterVar("C", "''", "S") & "  AND fp.bal_amt <> 0 AND fp.Gl_NO <> '' "
		
	If Trim(Request("txtAllcDt")) <> "" Then strCond = strCond & " AND fp.PRRCPT_DT <= " & FilterVar(UNIConvDate(Trim(Request("txtAllcDt"))), "''", "S")
				
	If strPrDt <> "" Then strCond = strCond & " AND fp.PRRCPT_DT >= " & FilterVar(strPrDt, "''", "S")
		
	If strToPrDt <> "" Then strCond = strCond & " AND fp.PRRCPT_DT <= " & FilterVar(strToPrDt, "''", "S")
		
	IF strBpCd <> "" Then strCond = strCond & " AND fp.BP_CD  = " & FilterVar(strBpCd, "''", "S")
		
	IF strDocCur <> "" Then strCond = strCond & " AND fp.DOC_CUR  = " & FilterVar(strDocCur, "''", "S")
		
	IF strDeptCd <> "" Then strCond = strCond & " AND fp.DEPT_CD  = " & FilterVar(strDeptCd, "''", "S")

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND fp.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND fp.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND fp.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND fp.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	strCond = strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	
End Sub 

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then  
			Parent.htxtPrDt.Value		= Parent.txtPrDt.text
			Parent.htxtToPrDt.Value		= Parent.txtToPrDt.Text
			Parent.htxtDeptcd.Value		= Parent.txtDeptcd.Value
			Parent.htxtBpcd.Value		= Parent.txtBpcd.Value
			Parent.htxtDocCur.Value		= Parent.txtDocCur.Value
       End If
			'Show multi spreadsheet data from this line
			Parent.ggoSpread.Source  = Parent.vspdData
			Parent.vspdData.Redraw = False
			Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",10),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",12),"A", "I" ,"X","X")
			Parent.vspdData.Redraw = True		
			Parent.lgPageNo      =  "<%=lgPageNo%>"							'�� : Next next data tag

			Parent.DbQueryOk

    End If   

</Script>	

