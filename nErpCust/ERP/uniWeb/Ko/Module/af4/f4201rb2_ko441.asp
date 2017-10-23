<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

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
Err.Clear

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1	, rs2, rs3, rs4	       '�� : DBAgent Parameter ����
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

Dim strOpenType
Dim strFrOpenDt
Dim strToOpenDt
Dim strDocCur
Dim strOrgChangeId
Dim strDeptCd
Dim strBpCd
Dim strBizCd
Dim strBpCd2
Dim strAllcAmt
Dim strRefNo
Dim strAcctCd
Dim strGlNo
Dim strMgntCd1
Dim strMgntCd2
Dim strCardCoCd
Dim strCardNo
Dim strFrCardUserId
Dim strToCardUserId
Dim strChkLocalCur
Dim strCond
Dim strFrDueDt
Dim strToDueDt
Dim strParentGlNo

Dim PAmt
Dim strMsgCd
Dim strMsg1
Dim skip_rs3,skip_rs4,no_mgnt1,no_mgnt2

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
    lgMaxCount     = UNICInt(Request("lgMaxCount") ,0)                          '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�
    lgSelectList   = Request("lgSelectList")                               '�� : select �����
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ��
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow	   = 0
    iEndRow        = 0

    Call SubOpenDB(lgObjConn)                                               '��: Make a DB Connection
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    Call SubCloseDB(lgObjConn)                                              '��: Close DB Connection    

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
	
    If strAllcAmt <> 0 Then
		Do While Not (Rs0.EOF Or Rs0.BOF)
			PAmt = strAllcAmt 
			strAllcAmt = strAllcAmt - UNIConvNum(Rs0(21) ,0)

		    iRowStr = ""
			For ColCnt = 0 To UBound(lgSelectListDT) - 1
				If ColCnt = 21  Then '�����ұݾ� ���� 
					If strAllcAmt > 0 Then 
						iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
					Else
						iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),PAmt)
					End If
				Else					
					iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
				End If		
			Next
 
			lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
			iEndRow = iLoopCount
			
			iLoopCount = iLoopCount + 1

			If strAllcAmt <= 0 Then 
				Exit Do
			End If	
			        
			rs0.MoveNext
		Loop
	Else
		If CDbl(lgPageNo) > 0 Then
			iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)
			rs0.Move = iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		End If

		iLoopCount = -1
    
		Do While Not (rs0.EOF Or rs0.BOF)
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
	End if

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
	Dim temp_nm1,temp_nm2
	Dim stbl_id,scol_id,stbl_id2,scol_id2

    Redim UNISqlId(6)                                                     '��: SQL ID ������ ���� ����Ȯ��
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(6,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ����

	Select Case UCase(strOpenType)
		Case "AP"
			UNISqlId(0) = "F5150RA202KO441"
			UNISqlId(1) = "ADEPTNM"
			UNISqlId(2) = "ABPNM"
			UNISqlId(3) = "ABIZNM"
			UNISqlId(4) = "ABPNM"
			
			UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
			UNIValue(1,1) = UCase(" " & FilterVar(strOrgChangeId, "''", "S") & " " )
			UNIValue(2,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & " " )
			UNIValue(3,0) = UCase(" " & FilterVar(strBizCd, "''", "S") & " " )
			UNIValue(4,0) = UCase(" " & FilterVar(strBpCd2, "''", "S") & " " )						
	End Select 	

    UNIValue(0,0) = lgSelectList                                          '��: Select list
	UNIValue(0,1) = strCond

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

    Set lgADF = Server.CreateObject("prjPublic.cCtlTake")
    
    Select Case UCase(strOpenType)
		Case "AP"
			lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1 , rs2, rs3 , rs4)
	End Select	
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strDeptCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDeptCd_alt")
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


	Select Case UCase(strOpenType)
		Case "AP"
		    If (rs2.EOF And rs2.BOF) Then
				If strMsgCd = "" And strBpCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBpCd_alt")
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

	End Select	

	Set rs2 = Nothing 		

	Select Case UCase(strOpenType)
		Case "AP"
		    If (rs3.EOF And rs3.BOF) Then
				If strMsgCd = "" And strBizCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBizCd_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtBizNm.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.txtBizCd.value = "<%=Trim(ConvSPChars(rs3(0)))%>"
					.txtBizNm.value = "<%=Trim(ConvSPChars(rs3(1)))%>"
				End With
				</Script>
		<%
		    End If
 
		Case Else
	End Select	

	Set rs3 = Nothing 
	
	Select Case UCase(strOpenType)
		Case "AP"
		    If (rs4.EOF And rs4.BOF) Then
				If strMsgCd = "" And strBpCd2 <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBpCd2_alt")
				End If
		%>
				<Script Language=vbScript>
				With parent
					.txtBpNm2.value = ""
				End With
				</Script>
		<%
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.txtBpCd2.value = "<%=Trim(ConvSPChars(rs4(0)))%>"
					.txtBpNm2.value = "<%=Trim(ConvSPChars(rs4(1)))%>"
				End With
				</Script>
		<%
		    End If
		Case Else
		
	End Select		

	Set rs4 = Nothing 

	If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
		Response.End													'��: �����Ͻ� ���� ó���� ������
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
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

	strOpenType     = Trim(Request("txtOpenType"))
	strFrOpenDt		= Trim(Request("txtFrOpenDt"))
	strToOpenDt		= Trim(Request("txtToOpenDt"))
	strDocCur		= Trim(Request("txtDocCur"))
	strOrgChangeId	= Trim(Request("txtOrgChangeId"))
	strDeptCd		= Trim(Request("txtDeptCd"))
	strBpCd			= Trim(Request("txtBpCd"))
	strBizCd		= Trim(Request("txtBizCd"))
	strBpCd2		= Trim(Request("txtBpCd2"))
	strAllcAmt		= Trim(Request("txtAllcAmt"))
	strRefNo		= Trim(Request("txtRefNo"))
	strAcctCd		= Trim(Request("txtAcctCd"))
	strGlNo			= Trim(Request("txtGlNo"))
	strMgntCd1		= Trim(Request("txtMgntCd1"))
	strMgntCd2		= Trim(Request("txtMgntCd2"))
	strCardCoCd		= Trim(Request("txtCardCoCd"))
	strCardNo		= Trim(Request("txtCardNo"))
	strFrCardUserId = Trim(Request("txtFrCardUserId"))
	strToCardUserId = Trim(Request("txtToCardUserId"))
	strParentGlNo   = Trim(Request("txtParentGLNo"))
	
	strFrDueDt		= Trim(Request("txtFrDueDt"))
	strToDueDt  	= Trim(Request("txtToDueDt"))
	
	if strToDueDt="" then strToDueDt="2999-12-31"
	
	strChkLocalCur  = Trim(Request("chkLocalCur"))

	' ���Ѱ��� �߰�
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

	Select Case UCase(strOpenType)
		Case "AP"
			strCond = strCond & " AND A.ap_dt >= " & FilterVar(strFrOpenDt , "''", "S") & ""
			strCond = strCond & " AND A.ap_dt <= " & FilterVar(strToOpenDt , "''", "S") & ""

			If strChkLocalCur = "Y" Then 
				strCond = strCond & " AND A.doc_cur like " & FilterVar(gCurrency & "%","''","S") & ""
			Else
				strCond = strCond & " AND A.doc_cur like " & FilterVar(strDocCur & "%","''","S") & ""
			End If

			If strBpCd <> "" Then
				strCond = strCond & " AND A.pay_bp_cd = " & FilterVar(strBpCd , "''", "S") & ""
			End If			

			If strDeptCd <> "" Then
				strCond = strCond & " AND A.dept_cd = " & FilterVar(strDeptCd , "''", "S") & ""
				strCond = strCond & " AND A.org_change_id = " & FilterVar(strOrgChangeId , "''", "S") & ""
			End If

			If strBizCd <> "" Then
				strCond = strCond & " AND A.biz_area_cd =" & FilterVar(strBizCd , "''", "S") & ""
			End If

			If strBpCd2 <> "" Then
				strCond = strCond & " AND A.deal_bp_cd = " & FilterVar(strBpCd2 , "''", "S") & ""
			End If			

			If strRefNo <> "" Then
				strCond = strCond & " AND A.ref_no like " & FilterVar("%" & strRefNo , "''", "S") & ""
			End If

			If strAcctCd <>  "" Then
			   strCond = strCond & "  And a.acct_cd = " & FilterVar(strAcctCd , "''", "S") & ""  
			End If 
			
			If strParentGlNo <> "" Then
				strCond = strCond & " AND A.gl_no <> " & FilterVar(strParentGlNo , "''", "S") & ""
			End If			
			strCond = strCond & " AND A.ap_due_dt between " & FilterVar(strFrDueDt , "''", "S") & " and " & FilterVar(strToDueDt , "''", "S") & ""
			
			' ���Ѱ��� �߰�
			If lgAuthBizAreaCd <> "" Then
				strCond = strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S") & ""
			End If

			If lgInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S") & ""
			End If

			If lgSubInternalCd <> "" Then
				strCond = strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S") & ""
			End If

			If lgAuthUsrID <> "" Then
				strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S") & ""
			End If			
										
	End Select
End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			Parent.htxtOpenType.value = Parent.cboOpenType.value
			Parent.htxtFrOpenDt.value = Parent.txtFrOpenDt.Text
			Parent.htxtToOpenDt.value = Parent.txtToOpenDt.Text
			Parent.htxtDocCur.value   = Parent.txtDocCur.value
'			Parent.hOrgChangeId.value = Parent.hOrgChangeId.value
			Parent.htxtDeptCd.value   = Parent.txtDeptCd.value
			Parent.htxtBpCd.value     = Parent.txtBpCd.value
			Parent.htxtBizCd.value    = Parent.txtBizCd.value
			Parent.htxtBpCd2.value    = Parent.txtBpCd2.value
			Parent.htxtAllcAmt.value  = Parent.txtAllcAmt.Text
			Parent.htxtRefNo.value    = Parent.txtRefNo.value
			Parent.htxtAcctCd.value   = Parent.txtAcctCd.value
			Parent.htxtGlNo.value     = Parent.txtGlNo.value
			Parent.htxtMgntCd1.value  = Parent.txtMgntCd1.value
			Parent.htxtMgntCd2.value  = Parent.txtMgntCd2.value
			Parent.htxtCardCoCd.value = Parent.txtCardCoCd.value
			Parent.htxtCardNo.value   = Parent.txtCardNo.value
			Parent.hChkLocalCur.value = Parent.ChkLocalCur.value
       End If

       'Show multi spreadsheet data from this line
       
		Parent.ggoSpread.Source		= Parent.vspdData
		Parent.vspdData.Redraw = False
		Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data

		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",10),"A", "I" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",11),"A", "I" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",12),"A", "I" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",9),Parent.GetKeyPos("A",22),"A", "I" ,"X","X")
		Parent.vspdData.Redraw = True
		Parent.lgPageNo				=  "<%=lgPageNo%>"               '�� : Next next data tag
			       
		Parent.DbQueryOk
    End If   

</Script>	

