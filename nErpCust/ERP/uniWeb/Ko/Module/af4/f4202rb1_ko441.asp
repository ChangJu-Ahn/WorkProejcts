<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Accounting - Treasury
'*  2. Function Name        : Loan
'*  3. Program ID           : f4202rb1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001.02.19
'*  7. Modified date(Last)  : 2001.11.10
'*  8. Modifier (First)     : Song, Mun Gil
'*  9. Modifier (Last)      : Oh, Soo Min
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")

Const C_SHEETMAXROWS_D = 30
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3                             '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                              '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strCond
Dim strLoanFrDt, strLoanToDt
Dim strDueFrDt, strDueToDt
Dim strBankCd, strLoanType
Dim strPgmId
Dim strDocCur
Dim strLoanfg   
Dim strLoanNo
Dim strMsgCd, strMsg1, strMsg2

Dim  iLoopCount
Dim  LngMaxRow

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
	Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	    
	lgMaxCount     = C_SHEETMAXROWS_D                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgDataExist    = "No"
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1
	    
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
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(lgMaxCount) * CDbl(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
     
    'rs0�� ���� ��� 
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
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
                    '��: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,2)

    UNISqlId(0) = "F4202ra101"
    UNISqlId(1) = "ABANKNM"
    UNISqlId(2) = "AMINORNM"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    UNIValue(1,0) = Filtervar(strBankCd	, "", "S")
    UNIValue(2,0) = Filtervar("F1000", "", "S")
    UNIValue(2,1) = Filtervar(strLoanType	, "", "S")
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    


	If rs1.EOF And rs1.BOF Then
		If strMsgCd = "" And strBankCd <> "" Then
			strMsgCd = "970000"
			strMsg1 = Request("txtBankLoanCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBankLoanCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			.txtBankLoanNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
		End With
		</Script>
<%
	End If
	
	If rs2.EOF And rs2.BOF Then
		If strMsgCd = "" And strLoanType <> "" Then
			strMsgCd = "970000"
			strMsg1 = Request("txtLoanType_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtLoanType.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
			.txtLoanTypeNm.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
		End With
		</Script>
<%
	End If
	
    If  rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Set lgADF = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If

	'rs0.Close
	'Set rs0 = Nothing 
	'Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
  
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    strLoanFrDt  = UniConvDate(Request("txtLoanFromDt"))
    strLoanToDt  = UniConvDate(Request("txtLoanToDt"))
    strDocCur	 = UCase(Request("txtDocCur"))
    strBankCd    = UCase(Request("txtBankLoanCd"))
    strDueFrDt   = Request("txtDueFromDt")
    strDueToDt   = Request("txtDueToDt")
    strLoanfg	 = Request("cboLoanFg")
    strLoanType  = Request("txtLoanType")
    strLoanNo	 = Request("txtLoanNo")
    strPgmId	 = Request("txtPgmId")
    
  	strCond = ""
	If  strPgmId = "F4201MA1" Then												'���Ա� TOTAL
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LN" , "''", "S") & " "		
		strCond = strCond & " and A.loan_bank_cd <>  " & FilterVar("", "''", "S") & " "
		strCond = strCond & " and A.loan_bank_cd = E.bank_cd "									
	Elseif strPgmId = "F4204MA1" Then											'����������Ա� 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LT" , "''", "S") & " "		
'		strCond = strCond & " and (A.loan_bank_cd = '" & "" & "' "
'		strCond = strCond & " or   A.loan_bank_cd  is null ) "												
'	Elseif strPgmId = "F4223MA1" Then											'���Աݻ�ȯ��ȹ���� 
	Elseif strPgmId = "F4207MA1" Then											'���Աݿ�û 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LQ" , "''", "S") & " "		
		
	Elseif strPgmId = "F4231MA1" Then											'������������ 
		strCond = strCond & " and A.int_votl =  " & FilterVar("F" , "''", "S") & " "
'		strCond = strCond & " and A.rdp_cls_fg = '" & "N" & "' "				'��ȯ�Ϸ�� ���� display������ 
	Elseif strPgmId = "F4234MA1" Then											'�������Աݸ��⿬�� 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LR" , "''", "S") & " "		
	End If
	
		strCond = strCond & " and A.loan_plc_type =  " & FilterVar("BK" , "''", "S") & " "
		strCond = strCond & " and A.loan_bank_cd = E.bank_cd "				
		
	If strLoanFrDt <> "" Then strCond = strCond & " and A.loan_dt >=  " & FilterVar(strLoanFrDt , "''", "S") & " "			'������ 
	If strLoanToDt <> "" Then strCond = strCond & " and A.loan_dt <=  " & FilterVar(strLoanToDt , "''", "S") & " "
	If strDocCur   <> "" Then strCond = strCond & " and A.doc_cur = " & Filtervar(strDocCur	, "''", "S")				'�ŷ���ȭ 
	If strDueFrDt  <> "" Then strCond = strCond & " and A.due_dt >=  " & FilterVar(UniConvDate(strDueFrDt), "''", "S") & " "				'������ 
	If strDueToDt  <> "" Then strCond = strCond & " and A.due_dt <=  " & FilterVar(UniConvDate(strDueToDt), "''", "S") & " "
	If strLoanNo   <> "" Then strCond = strCond & " and A.loan_no = " & Filtervar(strLoanNo	, "''", "S")				'���Թ�ȣ 
	If strLoanfg   <> "" Then strCond = strCond & " and A.loan_fg =  " & FilterVar(strLoanfg , "''", "S") & " "				'��ܱⱸ�� 
	If strLoanType <> "" Then strCond = strCond & " and A.loan_type = " & Filtervar(strLoanType	, "''", "S")			'���Կ뵵 
	If strBankCd   <> "" Then strCond = strCond & " and A.loan_bank_cd = " & Filtervar(strBankCd	, "''", "S")
	
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	strCond	= strCond	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub
'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------


%>

<Script Language=vbscript>
With parent
	If "<%=lgDataExist%>" = "Yes" Then
        .ggoSpread.Source    = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '�� : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",3),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
		.frm1.vspdData.Redraw = True

'         With .frm1
'			.hLoanFromDt.value = strLoanFrDt
'			.hLoanToDt.value   = strLoanToDt
'			.hDueFromDt.value  = strDueFrDt
'			.hDueToDt.value    = strDueToDt
'			.hBankLoanCd.value = strBankCd
'			.hLoanType.value   = strLoanType
 '        End With
         
	End If         
	.DbQueryOk()
End with
</Script>	

<%
	Response.End 
%>

