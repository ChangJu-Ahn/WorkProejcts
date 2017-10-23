<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : 
'*  3. Program ID        : A5461MB1
'*  4. Program �̸�      : �ΰ�����ǥ�ݾ�Ȯ�� 
'*  5. Program ����      : 
'*  6. Comproxy ����Ʈ   : 
'*  7. ���� �ۼ������   : 2003/06/17
'*  8. ���� ���������   : 
'*  9. ���� �ۼ���       : �ȵ��� 
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("Q", "B","NOCOOKIE","QB")
%>


<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next											' ��: 
Dim lgADF                                                       ' �� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                 ' �� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs2				' �� : DBAgent Parameter ���� 

Const C_SHEETMAXROWS_D  = 100 

Call HideStatusWnd												' ��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Dim LngMaxRow													' ���� �׸����� �ִ�Row
Dim LngRow

Dim ColSep, RowSep 
Dim Where01, Group01, Select01, InnerJoin01

Dim strMode														'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim lgPageNo												' Note NO ���� �� 
Dim lgMaxCount
Dim txtMaxRows3
Dim strData
Dim txtGlLocAmt2

'Const GroupCount = 30

strMode = Request("txtMode")									'�� : ���� ���¸� ���� 

	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)
	txtMaxRows3 = UNICInt(Request("txtMaxRows3"),0)


Call FixUNISQLData()
Call QueryData()

'#########################################################################################################
'												2.1 FixUNISQLData()
'##########################################################################################################	
Sub FixUNISQLData()

	Dim DispMeth, txtGlInputCd, txtFrDt, txtToDt, txtGlFrDt, txtGlToDt, txtShowDt, txtVatIoFg, txtVatTypeCd, txtBizAreaCd, txtTaxBizAreaCd, txtBpCd, txtShowBp
	
	DispMeth = Request("DispMeth")
	txtGlInputCd = UCase(Trim(Request("txtGlInputCd")))
	txtFrDt = Request("txtFrDt")
	txtToDt = Request("txtToDt")
	txtGlFrDt = Request("txtGlFrDt")
	txtGlToDt = Request("txtGlToDt")
	txtShowDt = UCase(Trim(Request("txtShowDt")))
	txtVatIoFg = UCase(Trim(Request("txtVatIoFg")))
	txtVatTypeCd = UCase(Trim(Request("txtVatTypeCd")))
	txtBizAreaCd = UCase(Trim(Request("txtBizAreaCd")))
	txtTaxBizAreaCd = UCase(Trim(Request("txtTaxBizAreaCd")))
	txtBpCd = UCase(Trim(Request("txtBpCd")))
	txtShowBp = UCase(Trim(Request("txtShowBp")))
	
    Where01 = ""
    If txtVatIoFg = "I" Then
    	Where01 = Where01 & " A.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  AND A.ACCT_TYPE = " & FilterVar("VP", "''", "S") & " "
    Else
    	Where01 = Where01 & " A.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  AND A.ACCT_TYPE = " & FilterVar("VR", "''", "S") & " "
    End If
    If txtVatIoFg		<> "" Then	Where01 = Where01 & " AND A.IO_FG = "			& Filtervar(txtVatIoFg	, "''", "S")
	If txtFrDt			<> "" Then  Where01 = Where01 & " AND A.ISSUED_DT >= "		& Filtervar(UniConvDate(txtFrDt)	, null, "S")
	If txtToDt			<> "" Then  Where01 = Where01 & " AND A.ISSUED_DT <= "		& Filtervar(UniConvDate(txtToDt)	, null, "S")
	If txtGlFrDt		<> "" Then  Where01 = Where01 & " AND A.GL_DT >= "			& Filtervar(UniConvDate(txtGlFrDt)	, null, "S")
	If txtGlToDt		<> "" Then  Where01 = Where01 & " AND A.GL_DT <= "			& Filtervar(UniConvDate(txtGlToDt)	, null, "S")
	If txtBizAreaCd		<> "" Then	Where01 = Where01 & " AND A.BIZ_AREA_CD = "		& Filtervar(txtBizAreaCd	, "''", "S")
	If txtTaxBizAreaCd	<> "" Then	Where01 = Where01 & " AND A.REPORT_BIZ_AREA_CD = "		& Filtervar(txtTaxBizAreaCd	, "''", "S")
    If txtGlInputCd		<> "" Then	Where01 = Where01 & " AND A.GL_INPUT_TYPE = "	& Filtervar(txtGlInputCd	, "''", "S")
    If txtVatTypeCd		<> "" Then	Where01 = Where01 & " AND A.VAT_TYPE = "		& Filtervar(txtVatTypeCd	, "''", "S")
    If txtBpCd			<> "" Then	Where01 = Where01 & " AND A.BP_CD = "			& Filtervar(txtBpCd			, "''", "S")
	If DispMeth = "True" Then
		Where01 = Where01 & " AND (ISNULL(A.GL_DT,'') <> ISNULL(A.ISSUED_DT ,'') "
		Where01 = Where01 & " OR E.REPORT_BIZ_AREA_CD <> A.REPORT_BIZ_AREA_CD) "
	End If
    
    Redim UNISqlId(1)                                                      '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "A5461MB201"
    UNISqlId(1) = "A5461MB202"

    Redim UNIValue(1,1)
    
	UNIValue(0,0) = Where01
	
	UNIValue(1,0) = Where01
			
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode	    
    
End Sub

'#########################################################################################################
'												2.2 QueryData()
'##########################################################################################################	
Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs2)
    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		rs2.Close:		Set rs2 = Nothing
		%><Script Language=vbscript>parent.frm1.txtFrDt2.Focus</Script><%
		Set lgADF = Nothing
	Else
		Call  MakeSpreadSheetData()
    End If						
		    
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing    	    
															'��: �����Ͻ� ���� ó���� ������ 
End Sub

'#########################################################################################################
'												2.4. HTML ��� ������ 
'##########################################################################################################		
Sub MakeSpreadSheetData()
	Dim intLoopCnt
	Dim iLoopCount
	Dim LngMaxRow       

	'-----------------------
	'Result data display area
	'-----------------------	
	intLoopCnt = rs0.recordcount
	LngMaxRow = txtMaxRows3										'Save previous Maxrow

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(lgMaxCount) * CDbl(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    If cint(intLoopCnt) <> 0 Then
		strData = ""
		txtGlLocAmt2 = 0
		iLoopCount = -1

		Do while Not (rs0.EOF Or rs0.BOF)
			iLoopCount =  iLoopCount + 1
			If  iLoopCount < lgMaxCount Then
				strData = strData & Chr(11) & ConvSPChars(rs0("BP_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("BP_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("VAT_TYPE"))
				strData = strData & Chr(11) & ConvSPChars(rs0("VAT_TYPE_NM"))
				strData = strData & Chr(11) & UNINumClientFormat(rs0("VAT_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
				strData = strData & Chr(11) & UNINumClientFormat(rs0("NET_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
				strData = strData & Chr(11) & ConvSPChars(rs0("GL_INPUT_TYPE"))
				strData = strData & Chr(11) & ConvSPChars(rs0("INPUT_TYPE_NM"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("GL_DT"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("ISSUED_DT"))
				strData = strData & Chr(11) & ConvSPChars(rs0("BIZ_AREA_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("BIZ_AREA"))
				strData = strData & Chr(11) & ConvSPChars(rs0("REPORT_BIZ_AREA_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("TAX_BIZ_AREA_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("TAX_BIZ_AREA_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("TAX_BIZ"))
				strData = strData & Chr(11) & LngMaxRow + iLoopCount
				strData = strData & Chr(11) & Chr(12)
			Else
			    lgPageNo = lgPageNo + 1
			    Exit Do
			End If
			rs0.MoveNext
		Loop
		
		If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
		    lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
		End If
		
		If NOT(rs2.EOF) And NOT(rs2.BOF) Then
			txtGlLocAmt2 = UNINumClientFormat(rs2("SUM_VAT_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
		End If
		
	End If
	rs0.close:			Set rs0 = Nothing	                                                    '��: ActiveX Data Factory Object Nothing
	rs2.close:			Set rs2 = Nothing

End Sub	

%>

<Script Language=vbscript>

		With parent
			.frm1.vspdData3.Redraw = False
			.ggoSpread.Source = .frm1.vspddata3
			.ggoSpread.SSShowData "<%=strData%>"
			.frm1.txtGlLocAmt2.Text = "<%=txtGlLocAmt2%>"
			.lgPageNo      =  "<%=lgPageNo%>"

'			If .lgStrPrevKey <> "" Then
'				.DbQuery
'			Else
				.DbQueryOK
'			End If
			.frm1.vspdData3.Redraw = True
			
		End With
</Script>


