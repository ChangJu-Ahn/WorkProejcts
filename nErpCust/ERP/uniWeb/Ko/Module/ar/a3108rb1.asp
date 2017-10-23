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

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1                         '�� : DBAgent Parameter ���� 
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
Dim strFrAllcDt	                                                           
Dim strToAllcDt
Dim strFrAllcNo	                                                           
Dim strToAllcNo
Dim strdeptcd

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
    Call LoadInfTB19029B("I","*","NOCOOKIE","QB")
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

    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(1,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "A3108RA101"
	UNISqlId(1) = "ADEPTNM"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & "" )
    UNIValue(1,1) = UCase(" " & FilterVar(UCase(Request("txtOrgChangeId")), "''", "S") & "" )     
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
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
			.Frm1.txtDeptNm.value = ""
		End With
		</Script>
<%
    Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtDeptCd.value = "<%=Trim(rs1(0))%>"
			.frm1.txtDeptNm.value = "<%=Trim(rs1(1))%>"
		End With
		</Script>
<%
    End If
    
	Set rs1 = Nothing 
        
    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
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
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
 
     strFrAllcDt     = UCase(Trim(UNIConvDate(Request("txtFrAllcDt"))))
     strToAllcDt     = UCase(Trim(UNIConvDate(Request("txtToAllcDt"))))
     strFrAllcNo	 = UCase(Trim(Request("txtFrAllcNo")))                                                          
     strToAllcNo     = UCase(Trim(Request("txtToAllcNo")))
     strdeptcd       = UCase(Trim(Request("txtdeptcd")))
     
	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))     
     
     strCond = "AND A.ALLC_TYPE = " & FilterVar("P", "''", "S") & "  "
     If strFrAllcDt <> "" Then
		strCond = strCond & " and A.ALLC_DT >=  " & FilterVar(strFrAllcDt , "''", "S") & ""
     End If
     
     If strToAllcDt <> "" Then
		strCond = strCond & " and A.ALLC_DT <=  " & FilterVar(strToAllcDt , "''", "S") & ""
     End If
     
     If strFrAllcNo <> "" Then
		strCond = strCond & " and A.ALLC_NO >=  " & FilterVar(strFrAllcNo , "''", "S") & ""
     End If
     
     If strToAllcNo <> "" Then
		strCond = strCond & " and A.ALLC_NO <=  " & FilterVar(strToAllcNo , "''", "S") & ""
     End If
     
     If strdeptcd <> "" Then
		strCond = strCond & " and A.dept_cd =  " & FilterVar(strdeptcd , "''", "S") & ""
     End If

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND C.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	strCond = strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			Parent.Frm1.htxtFrAllcDt.Value    = Parent.Frm1.txtFrAllcDt.Text
			Parent.Frm1.htxtToAllcDt.Value    = Parent.Frm1.txtToAllcDt.Text
			Parent.Frm1.htxtFrAllcNo.Value	= Parent.Frm1.txtFrAllcNo.Value
			Parent.Frm1.htxtToAllcNo.Value	= Parent.Frm1.txtToAllcNo.Value
			Parent.Frm1.htxtdeptcd.Value	= Parent.Frm1.txtdeptcd.Value
       End If
			'Show multi spreadsheet data from this line
			Parent.ggoSpread.Source  = Parent.frm1.vspdData
			Parent.frm1.vspdData.Redraw = False
			Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,"X",Parent.GetKeyPos("A",2),"A", "Q" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,"X",Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
			Parent.frm1.vspdData.Redraw = True	
			Parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
			Parent.DbQueryOk
    End If   

</Script>	

