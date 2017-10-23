
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("*", "M", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                         '�� : DBAgent Parameter ���� 
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
	Dim strFrApDt	                                                           
	Dim strToApDt
	Dim strFrApNo	                                                           
	Dim strToApNo
	Dim strdeptcd

	Dim strCond

	Dim strMsgCd
	Dim strMsg1
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    
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

    UNISqlId(0) = "A4101RA101"
    UNISqlId(1) = "ADEPTNM"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
    UNIValue(1,1) = UCase(" " & FilterVar(UCase(Request("txtOrgChangeId")), "''", "S") & " " )

    
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
%>
<Script Language=vbScript>
		With parent
			.frm1.txtDeptNm.value = ""
		End With
		</Script>
<%
		Else
%>
<Script Language=vbScript>
		With parent
			.frm1.txtDeptNm.value = ""
		End With
		</Script>
<%
		End If
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
		Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
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
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strFrApDt     =  UCase(Trim(UNIConvDate(Request("txtFrApDt"))))     
     strToApDt     =  UCase(Trim(UNIConvDate(Request("txtToApDt"))))
     strFrApNo	   =  UCase(Trim(Request("txtFrApNo")))                                                          
     strToApNo     =  UCase(Trim(Request("txtToApNo")))
     strdeptcd     =  UCase(Trim(Request("txtdeptcd")))
     
     strCond  = " and A.AP_TYPE = " & FilterVar("NT", "''", "S") & "  "
     
     If strFrApDt <> "" Then
		strCond = strCond & " and A.AP_DT >=  " & FilterVar(strFrApDt , "''", "S") & ""
     End If
     
     If strToApDt <> "" Then
		strCond = strCond & " and A.AP_DT <=  " & FilterVar(strToApDt , "''", "S") & ""
     End If
     
     If strFrApNo <> "" Then
		strCond = strCond & " and A.AP_NO >=  " & FilterVar(strFrApNo , "''", "S") & ""
     End If
     
     If strToApNo <> "" Then
		strCond = strCond & " and A.AP_NO <=  " & FilterVar(strToApNo , "''", "S") & ""
     End If
     
     If strdeptcd <> "" Then
		strCond = strCond & " and A.dept_cd =  " & FilterVar(strdeptcd , "''", "S") & ""
     End If     
     
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub
'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------
''2004.8.19 commentó�� 
'Function FilterVar(Byval str,Byval strALT)
'     Dim strL
'     strL = UCase(Trim(str))
'     If Len(strL) Then
'        FilterVar = " " & FilterVar(strL , "''", "S") & ""
'     Else
'        FilterVar = strALT   
'     End If
'End Function
%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.htxtFrApDt.Value      = Parent.Frm1.txtFrApDt.Text
          Parent.Frm1.htxtToApDt.Value         = Parent.Frm1.txtToApDt.Text
          Parent.Frm1.htxtFrApNo.Value  = Parent.Frm1.txtFrApNo.Value
          Parent.Frm1.htxtToApNo.Value    = Parent.Frm1.txtToApNo.Value
          Parent.Frm1.htxtdeptcd.Value    = Parent.Frm1.txtdeptcd.Value
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data

       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",6),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",14),"A", "Q" ,"X","X")
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk
    End If   

</Script>	

