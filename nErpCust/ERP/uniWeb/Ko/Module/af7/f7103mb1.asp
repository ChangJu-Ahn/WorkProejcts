<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2002/08/23
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : Jeong Yong Kyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

                                                       '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5     '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
DIm lgMaxCount
Dim lgPageNo
Dim iPrevEndRow
Dim iEndRow	

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strFromDt																'�� : ������ 
Dim strToDt																	'�� : ������ 
Dim strDeptCd																'�� : �μ� 
Dim strBpCd																	'�� : �ŷ�ó 
DIm strPrrcptType
Dim strCond
Dim strBizAreaCd															'�� : ���ۻ���� 
Dim strBizAreaNm
Dim strBizAreaCd1															'�� : �������� 
Dim strBizAreaNm1
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


	Call HideStatusWnd
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "QB")   'ggQty.DecPoint Setting...
	Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QB")


	lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	    
	lgSelectList	= Request("lgSelectList")                               '�� : select ����� 
	lgMaxCount		= CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgTailList		= Request("lgTailList")                                 '�� : Orderby value
	lgDataExist		= "No"
	iChangeOrgId	= UCase(Trim(Request("OrgChangeId"))) 
	iPrevEndRow		= 0
	iEndRow			= 0

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

	Call TrimData()
	Call FixUNISQLData()
	Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim RecordCnt
    Dim ColCnt
    Dim iLoopCount
    Dim iRowStr

    lgstrData = ""

    lgDataExist    = "Yes"

    If CInt(lgPageNo) > 0 Then
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

    Redim UNISqlId(5)                                               '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(5,4)

    UNISqlId(0) = "A7103MA1"
    UNISqlId(1) = "ADEPTNM"
    UNISqlId(2) = "ABPNM"
    UNISqlId(3) = "Commonqry"
    UNISqlId(4) = "A_GETBIZ"
    UNISqlId(5) = "A_GETBIZ"
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = FilterVar(strFromDt,"''","S")
    UNIValue(0,2) = FilterVar(strToDt,"''","S")
    UNIValue(0,3) = UCase(Trim(strCond))
    
    
	UNIValue(1,0)  = " " & FilterVar(strDeptCd, "''", "S") & " "		
	UNIValue(1,1)  = " " & FilterVar(iChangeOrgId, "''", "S") & " "	
    
    UNIValue(2,0) = " " & UCase(FilterVar(Request("txtBPCd"), "''", "S")) & " "
    UNIValue(3,0) = " select jnl_nm from A_JNL_ITEM where jnl_cd = " & FilterVar(Request("txtPrrcptType"), "''", "S")  
    
    UNIValue(4,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(5,0)  = FilterVar(strBizAreaCd1, "''", "S")
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	Dim lgADF													'�� : ActiveX Data Factory ���� �������� 
	Dim StrMsg1,strMsg2,strMsg3,strEmpty                                                            
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    
    iStr = Split(lgstrRetMsg,gColSep)
    Set lgADF = Nothing  
    
    strEmpty = ""
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	If rs1.EOF And rs1.BOF Then
		strmsg1 = Trim(Request("txtDeptCd_Alt"))
		If UCase(Trim(Request("txtDeptCd"))) <> "" Then
			Call DisplayMsgBox("970000", vbOKOnly, strmsg1, "", I_MKSCRIPT)
			Response.End
		End If
		Response.Write "<Script Language=vbscript>"											  & vbCr
		Response.Write " parent.frm1.txtDeptNm.value = """ & strEmpty & """"				  & vbCr
 	    Response.Write "</Script>"															  & vbCr
	Else
		Response.Write "<Script Language=vbscript>"											  & vbCr
		Response.write " parent.frm1.txtDeptCd.value = """ & ConvSPChars(Trim(rs1(0))) & """" & vbCr	
		Response.write " parent.frm1.txtDeptNm.value = """ & ConvSPChars(Trim(rs1(1))) & """" & vbCr   
		Response.Write "</Script>"															  & vbCr	
	End If
    rs1.Close
	Set rs1 = Nothing
	
	If rs2.EOF And rs2.BOF Then
		strmsg2 = Trim(Request("txtBpCd_Alt"))
		If UCase(Trim(Request("txtBPCd"))) <> "" Then
			Call DisplayMsgBox("970000", vbOKOnly, strmsg2, "", I_MKSCRIPT)
			Response.End
		End If
		Response.Write "<Script Language=vbscript>"											& vbCr
		Response.write " parent.frm1.txtBPNm.value = """ & strEmpty & """"					& vbCr				
		Response.Write "</Script>"														    & vbCr
	Else
		Response.Write "<Script Language=vbscript>"											& vbCr
		Response.write " parent.frm1.txtBPCd.value = """ & ConvSPChars(Trim(rs2(0))) & """" & vbCr 
		Response.write " parent.frm1.txtBPNm.value = """ & ConvSPChars(Trim(rs2(1))) & """" & vbCr 
		Response.Write "</Script>"															& vbCr
	End If
	rs2.Close
	Set rs2 = Nothing

	IF rs3.eof and rs3.bof then
		strMsg3 = Trim(Request("txtPrrcptType_Alt"))
		If UCase(Trim(Request("txtPrrcptType"))) <> ""Then
			Call DisplayMsgBox("970000", vbOKOnly, strMsg3, "", I_MKSCRIPT)		'No Data Found!!
	    	Response.End 
		End If	    	
		Response.Write "<Script Language=vbScript>"										& vbCr
		Response.Write " parent.frm1.txtPrrcptTypeNm.value = """ & strEmpty & """"		& vbCr
        Response.Write "</Script>    "												    & vbCR			
    Else
		Response.Write "<Script Language=vbScript>"											        & vbCr
        Response.Write " parent.frm1.txtPrrcptTypeNm.value = """ & ConvSPChars(Trim(rs3(0))) & """" & vbCr
        Response.Write "</Script>    "														        & vbCr
	End If	
	rs3.Close
	Set rs3 = Nothing
	
If (rs4.EOF And rs4.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd.value = "<%=Trim(rs4(0))%>"
		.frm1.txtBizAreaNm.value = "<%=Trim(rs4(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs4.Close
	Set rs4 = Nothing   
    
    
If (rs5.EOF And rs5.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd1_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd1.value = "<%=Trim(rs5(0))%>"
		.frm1.txtBizAreaNm1.value = "<%=Trim(rs5(1))%>"					
	End With
	</Script>
<%
    End If
	
	If  Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, "", I_MKSCRIPT)
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    End If
    
	rs5.Close
	Set rs5 = Nothing 
		
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
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
	
	strFromDt		= UNIConvDate(Request("txtFromDt"))
	strToDt			= UNIConvDate(Request("txtToDt"))
	strDeptCd		= UCase(Trim(Request("txtDeptCd")))
	strBpCd			= UCase(Trim(Request("txtBpCd")))
	strPrrcptType	= UCase(Trim(Request("txtPrrcptType")))
	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))					'�����From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))					'�����To
	
	If strDeptCd <> "" Then 
		strCond = strCond & " and A.internal_cd = (SELECT internal_cd FROM b_acct_dept  WHERE org_change_id = "
		strCond = strCond & FilterVar(iChangeOrgId ,null,"S") & " AND dept_cd =  " & FilterVar(strDeptCd ,null,"S") & ")"
	End if
	
	If strBpCd <> "" Then strCond = strCond & " and A.bp_cd = " & FilterVar(strBpCd, "''", "S") 
	If strPrrcptType <> "" Then strCond = strCond & " and a.prrcpt_type = " & FilterVar(strPrrcptType , "''", "S")
	
	if strBizAreaCd <> "" then
		strCond = strCond & " AND a.BIZ_AREA_CD >= " & FilterVar(strBizAreaCd , "''", "S") 
	else
		strCond = strCond & " AND a.BIZ_AREA_CD >= " & FilterVar("", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strCond = strCond & " AND a.BIZ_AREA_CD <= " & FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strCond = strCond & " AND a.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	End if


	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND a.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND a.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND a.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND a.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	strCond		= strCond	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
		With parent
			.ggoSpread.Source  = .frm1.vspdData
			 Parent.frm1.vspdData.Redraw = False
			 Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
			 Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
			 Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",4),"A", "Q" ,"X","X")
			 Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",5),"A", "Q" ,"X","X")
			 Parent.frm1.vspdData.Redraw = True
			.lgPageNo_A      =  "<%=lgPageNo%>"               '�� : Next next data tag
'			.DbQueryOk("1")
       End with
    End If   
</Script>	
