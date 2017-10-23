<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1, rs2, rs3                         '�� : DBAgent Parameter ���� 
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
Dim strFrArDt	                                                           
Dim strToArDt
Dim strFrArNo	                                                           
Dim strToArNo
DIm strDealBpCd
Dim strdeptcd
Dim stradjustNo
Dim strProject
	                                                           
	                                                           '�� : ������ 
Dim strCond

Dim strMsgCd
Dim strMsg1

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd	' ����� 
Dim lgInternalCd	' ���κμ� 
Dim lgSubInternalCd	' ���κμ�(��������)
Dim lgAuthUsrID		' ���� 
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("*", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
     strFrArDt     = UNIConvDate(Trim(Request("txtFrArDt")))
     strToArDt     = UNIConvDate(Trim(Request("txtToArDt")))
     strFrArNo	   = UCase(Trim(Request("txtFrArNo")))
     strToArNo     = UCase(Trim(Request("txtToArNo")))
     strDealBpCd   = UCase(Trim(Request("txtDealBpCd")))
     stradjustNo   = UCase(Trim(Request("txtadjustNo")))
     strdeptcd     = UCase(Trim(Request("txtdeptcd")))
     strProject    = UCase(Trim(Request("txtProject")))
    
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

    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(3,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "A3101RA101"
	UNISqlId(1) = "ADEPTNM"
	UNISqlId(2) = "ABPNM"
	UNISqlId(3) = "Commonqry"
    '--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & "" )
    UNIValue(1,1) = UCase(" " & FilterVar(UCase(Request("txtOrgChangeId")), "''", "S") & "" )  
    UNIValue(2,0) = UCase(" " & FilterVar(strDealBpCd, "''", "S") & "" ) 
    UNIValue(3,0) = "SELECT adjust_no FROM A_AR_ADJUST WHERE adjust_no=" & UCase(" " & FilterVar(stradjustNo, "''", "S") & "" )
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3)
    
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
			.frm1.txtDeptCd.value = "<%=Trim(ConvSPChars(rs1(0)))%>"
			.frm1.txtDeptNm.value = "<%=Trim(ConvSPChars(rs1(1)))%>"
		End With
		</Script>
<%
    End If
    
	Set rs1 = Nothing 
    
        
	If strDealBpCd <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtDealBpNm.value = "<%=Trim(ConvSPChars(rs2(1)))%>"
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
			Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs2.Close
		    Set rs2 = Nothing
			Exit sub
		End IF
	rs2.Close
	Set rs2 = Nothing
	End If
	
    If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And stradjustNo <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtadjustNo_Alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtadjustNo.focus
		End With
		</Script>
<%
	Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtadjustNo.value = "<%=Trim(ConvSPChars(rs3(0)))%>"
		End With
		</Script>
<%
    End If
    
    rs3.Close
	Set rs3 = Nothing 	
    
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
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
      
     strCond  = " AND A.CONF_FG = " & FilterVar("C", "''", "S") & "  AND A.GL_NO <>'' AND (A.AR_STS = " & FilterVar("O", "''", "S") & "  OR A.AR_NO IN(SELECT distinct AR_NO FROM A_AR_ADJUST))"
   
     If strFrArDt <> "" Then strCond = strCond & " and A.AR_DT >= " & FilterVar(strFrArDt ,null	,"S")
     
     If strToArDt <> "" Then strCond = strCond & " and A.AR_DT <= " & FilterVar(strToArDt ,null	,"S") 
        
     If strFrArNo <> "" Then strCond = strCond & " and A.AR_NO >= " & FilterVar(strFrArNo ,null	,"S")  
         
     If strToArNo <> "" Then strCond = strCond & " and A.AR_NO <= " & FilterVar(strToArNo ,null	,"S") 
      
     If strdeptcd <> "" Then strCond = strCond & " and A.dept_cd = " & FilterVar(strdeptcd ,null,"S")  
     
     If strDealBpCd <> "" Then	strCond = strCond & " and A.deal_bp_cd = " & FilterVar(strDealBpCd ,null,"S") 
	
	 If strProject <> "" Then	strCond = strCond & " AND A.project_no Like  " & FilterVar("%" & strProject & "%", "''", "S") & "" 
	
	 If stradjustNo <> "" Then	
			strCond = strCond & " AND A.AR_NO = (SELECT distinct AR_NO FROM A_AR_ADJUST WHERE adjust_no=" & FilterVar(stradjustNo ,null,"S") & ")" 
	 End if

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
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			Parent.Frm1.htxtFrArDt.Value    = Parent.Frm1.txtFrArDt.Text
			Parent.Frm1.htxtToArDt.Value    = Parent.Frm1.txtToArDt.Text
			Parent.Frm1.htxtFrArNo.Value	= Parent.Frm1.txtFrArNo.Value
			Parent.Frm1.htxtToArNo.Value	= Parent.Frm1.txtToArNo.Value
			Parent.Frm1.htxtDealBpCd.Value	= Parent.Frm1.txtDealBpCd.Value
			Parent.Frm1.htxtdeptcd.Value	= Parent.Frm1.txtdeptcd.Value
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

