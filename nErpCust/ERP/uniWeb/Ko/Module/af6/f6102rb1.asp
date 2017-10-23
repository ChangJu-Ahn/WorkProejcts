<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("*", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                         '�� : DBAgent Parameter ���� 
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
Dim strfrtempgldt	                                                           
Dim strtotempgldt
Dim strfrtempglno	                                                           
Dim strtotempglno
Dim strdeptcd
DIm strDealBpCd	                                                           '�� : ������ 

Dim strCond
Dim	strDeptNm
' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd	' ����� 
Dim lgInternalCd	' ���κμ� 
Dim lgSubInternalCd	' ���κμ�(��������)
Dim lgAuthUsrID		' ���� 
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
    iPrevEndRow = 0
    iEndRow = 0

    strfrtempgldt = UNIConvDate(Request("txtfrtempgldt"))
    strtotempgldt = UNIConvDate(Request("txttotempgldt"))
    strfrtempglno = UCase(Trim(Request("txtfrtempglno")))                                                         
    strtotempglno = UCase(Trim(Request("txttotempglno")))
    strdeptcd     = UCase(Trim(Request("txtdeptcd")))
    strDealBpCd   = UCase(Trim(Request("txtDealBpCd")))
    
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
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,2)

    UNISqlId(0) = "F6102RA101"
	UNISqlId(1) = "ADEPTNM"
	UNISqlId(2) = "ABPNM"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    'UNIValue(0,2) = UCase(Trim(strtotempgldt))
    UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
    UNIValue(1,1) = UCase(" " & FilterVar(GetGlobalInf(gChangeOrgId), "''", "S") & " " )
    UNIValue(2,0) = UCase(" " & FilterVar(strDealBpCd, "''", "S") & " " )
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing   
    
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

	If strDealBpCd <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtDealBpNm.value = "<%=Trim(rs2(1))%>"
			End With
		</Script>
<%		
		Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtDealBpNm.value = ""
			End With
		</Script>
<%
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDealBpCD_Alt")
			Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)		'No Data Found!!
	        rs2.Close
		    Set rs2 = Nothing
			Exit sub
		End IF
	rs2.Close
	Set rs2 = Nothing
	End If
	
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
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
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     
     strCond = strCond & " AND ((a.gl_no <> '' and A.prpaym_sts = " & FilterVar("O", "''", "S") & "  and A.conf_fg = " & FilterVar("C", "''", "S") & " ) OR A.prpaym_no IN (SELECT distinct prpaym_no FROM f_prpaym_sttl ))" 
     If strfrtempgldt <> "" Then strCond = strCond & " and A.PRPAYM_DT >= " & FilterVar(strfrtempgldt, "''", "S") 
   
     If strtotempgldt <> "" Then strCond = strCond & " and A.PRPAYM_DT <= " & FilterVar(strtotempgldt, "''", "S")  

     If strfrtempglno <> "" Then strCond = strCond & " and A.PRPAYM_NO >= " & FilterVar(strfrtempglno , "''", "S")
    
     If strtotempglno <> "" Then strCond = strCond & " and A.PRPAYM_NO <= " & FilterVar(strtotempglno , "''", "S")

     If strdeptcd <> "" Then strCond = strCond & " and a.dept_cd =  " & FilterVar(strdeptcd, "''", "S") & " "
     
     If strDealBpCd <> "" Then	strCond = strCond & " and A.bp_cd = " & FilterVar(strDealBpCd , "''", "S")      
     
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If     
	
'     strCond = strCond & " and A.PRPAYM_FG = 'PP'"
     
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub
%>

<Script Language=vbscript>

    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          
          Parent.Frm1.htxtfrtempgldt.Value  = Parent.Frm1.txtfrtempgldt.Text
          Parent.Frm1.htxttotempgldt.Value  = Parent.Frm1.txttotempgldt.Text          
          Parent.Frm1.htxtfrtempglNo.Value  = Parent.Frm1.txtfrtempglNo.Value
          Parent.Frm1.htxttotempglNo.Value  = Parent.Frm1.txttotempglNo.Value
          Parent.Frm1.htxtdeptcd.Value      = Parent.Frm1.txtdeptcd.Value
       	  Parent.Frm1.htxtDealBpCd.Value	= Parent.Frm1.txtDealBpCd.Value
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",8),Parent.GetKeyPos("A",12),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",8),Parent.GetKeyPos("A",14),"A", "Q" ,"X","X")
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk
    End If   

</Script>	