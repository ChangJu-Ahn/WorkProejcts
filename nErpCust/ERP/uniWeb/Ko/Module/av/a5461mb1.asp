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
'* 13. History           : 1.5 revision �п��� �������� ������, �� ��⺰ uniCODE �۾������ؼ� 
'*                         Check Out �޾� Ȯ����.
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
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3				' �� : DBAgent Parameter ���� 

Call HideStatusWnd												' ��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Dim LngMaxRow													' ���� �׸����� �ִ�Row
Dim LngRow

Dim ColSep, RowSep 
Dim Where01, Group01, Select01, InnerJoin01, VatSelect, AGlSelect, VatGroup, AGlGroup

Dim strMode														'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim lgStrPrevKey												' Note NO ���� �� 
Dim txtMaxRows1, txtMaxRows2
Dim strData, strData2
Dim txtVatLocAmt1, txtGlLocAmt1

Const GroupCount = 30

	strMode = Request("txtMode")									'�� : ���� ���¸� ���� 

	lgStrPrevKey = "" & UCase(Trim(Request("lgStrPrevKey")))	
	txtMaxRows1 = UNICInt(Request("txtMaxRows1"),0)
	txtMaxRows2 = UNICInt(Request("txtMaxRows2"),0)
		

Call FixUNISQLData()
Call QueryData()

'#########################################################################################################
'												2.1 FixUNISQLData()
'##########################################################################################################	
Sub FixUNISQLData()

	Dim DispMeth, txtGlInputCd, txtFrDt, txtToDt, txtShowDt, txtVatIoFg, txtVatTypeCd, txtTaxBizAreaCd, txtBpCd, txtShowBp, txtShowBiz
	
	DispMeth	 = Request("DispMeth")
	txtGlInputCd = UCase(Trim(Request("txtGlInputCd")))
	txtFrDt		 = UniConvDAte(Request("txtFrDt"))
	txtToDt		 = UniConvDAte(Request("txtToDt"))
	txtShowDt	 = UCase(Trim(Request("txtShowDt")))
	txtVatIoFg	 = UCase(Trim(Request("txtVatIoFg")))
	txtVatTypeCd = UCase(Trim(Request("txtVatTypeCd")))
	txtTaxBizAreaCd = UCase(Trim(Request("txtTaxBizAreaCd")))
	txtBpCd		 = UCase(Trim(Request("txtBpCd")))
	txtShowBp	 = UCase(Trim(Request("txtShowBp")))
	txtShowBiz	 = UCase(Trim(Request("txtShowBiz")))
	
	Select01	= ""
	Group01		= ""
	InnerJoin01 = ""
	Where01		= ""

	If txtShowDt = "N" Then
		VatSelect = ", '' ISSUED_DT"
		AGlSelect = ", '' ISSUED_DT"
	Else
		VatSelect = ", A.ISSUED_DT"
		AGlSelect = ", H.CTRL_VAL ISSUED_DT"
		VatGroup  = ", A.ISSUED_DT "
		AGlGroup  = ", H.CTRL_VAL "
	End If
	
	If txtShowBp = "N" Then
		Select01 = Select01 & ", '' BP_CD, '' BP_NM "
	Else
		Select01 = Select01 & ", D.BP_CD, D.BP_NM"
		Group01  = Group01  & ", D.BP_CD, D.BP_NM"
		InnerJoin01 = " INNER JOIN B_BIZ_PARTNER D ON D.BP_CD = A.BP_CD"
	End If
	
	If txtShowBiz = "N" Then
		Select01 = Select01 & ", '' TAX_BIZ_AREA_CD, '' TAX_BIZ_AREA_NM "
	Else
		Select01 = Select01 & ", F.TAX_BIZ_AREA_CD, F.TAX_BIZ_AREA_NM"
		Group01  = Group01  & ", F.TAX_BIZ_AREA_CD, F.TAX_BIZ_AREA_NM"
		InnerJoin01 = InnerJoin01 & " INNER JOIN B_TAX_BIZ_AREA F ON F.TAX_BIZ_AREA_CD = A.REPORT_BIZ_AREA_CD"
	End If
	
	If txtTaxBizAreaCd <> "" Then	Where01 = Where01 & " AND A.REPORT_BIZ_AREA_CD = "	& Filtervar(txtTaxBizAreaCd	, "''", "S")
    If txtGlInputCd <> "" Then	Where01 = Where01 & " AND A.GL_INPUT_TYPE = "		& Filtervar(txtGlInputCd	, "''", "S")
    If txtVatTypeCd <> "" Then	Where01 = Where01 & " AND A.VAT_TYPE = "			& Filtervar(txtVatTypeCd	, "''", "S")
    If txtBpCd <> "" Then		Where01 = Where01 & " AND A.BP_CD = "				& Filtervar(txtBpCd			, "''", "S")
'    Where01 = Where01 & " AND A.CONF_FG='C' "
    
    Redim UNISqlId(3)                                                      '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "A5461MB101"
    UNISqlId(1) = "A5461MB102"
    UNISqlId(2) = "A5461MB103"
    UNISqlId(3) = "A5461MB104"
    
    Redim UNIValue(3,6)

'============== �ΰ��� ===================
	UNIValue(0,0) = VatSelect & Select01
	UNIValue(0,1) = InnerJoin01
	UNIValue(0,2) = FilterVar(UniConvDate(Request("txtFrDt")), "''", "S") 
	UNIValue(0,3) = FilterVar(UniConvDate(Request("txtToDt")), "''", "S") 
	UNIValue(0,4) = Where01 & " AND A.CONF_FG=" & FilterVar("C", "''", "S") & "  AND A.IO_FG = "	& FilterVar(txtVatIoFg	, "''", "S")
	UNIValue(0,5) = VatGroup & Group01
	UNIValue(0,6) = VatGroup & Group01
	
	UNIValue(2,0) = FilterVar(UniConvDate(Request("txtFrDt")), "''", "S") 
	UNIValue(2,1) = FilterVar(UniConvDate(Request("txtToDt")), "''", "S") 
	UNIValue(2,2) = Where01 & " AND A.CONF_FG=" & FilterVar("C", "''", "S") & "  AND A.IO_FG = "	& FilterVar(txtVatIoFg	, "''", "S")

    If txtVatIoFg = "I" Then	Where01 = Where01 & " AND A.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  AND A.ACCT_TYPE = " & FilterVar("VP", "''", "S") & " "
    If txtVatIoFg = "O" Then	Where01 = Where01 & " AND A.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  AND A.ACCT_TYPE = " & FilterVar("VR", "''", "S") & " "

'============== ��ǥ ===================	
	UNIValue(1,0) = AGlSelect & Select01
	UNIValue(1,1) = InnerJoin01
	UNIValue(1,2) = FilterVar(UniConvDate(Request("txtFrDt")), "''", "S") 
	UNIValue(1,3) = FilterVar(UniConvDate(Request("txtToDt")), "''", "S") 
	UNIValue(1,4) = Where01
	UNIValue(1,5) = AGlGroup & Group01
	UNIValue(1,6) = AGlGroup & Group01

	UNIValue(3,0) = FilterVar(UniConvDate(Request("txtFrDt")), "''", "S") 
	UNIValue(3,1) = FilterVar(UniConvDate(Request("txtToDt")), "''", "S") 
	UNIValue(3,2) = Where01
			
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode	    
    
End Sub

'#########################################################################################################
'												2.2 QueryData()
'##########################################################################################################	
Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If (rs0.EOF And rs0.BOF) AND (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		rs1.Close:		Set rs1 = Nothing
		rs2.Close:		Set rs2 = Nothing
		rs3.Close:		Set rs3 = Nothing
		%><Script Language=vbscript>parent.frm1.txtFrDt.Focus</Script><%
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
	Dim LngMaxRow       

	'-----------------------
	'Result data display area
	'-----------------------	
	intLoopCnt = rs0.recordcount
	LngMaxRow = txtMaxRows1										'Save previous Maxrow

    If cint(intLoopCnt) <> 0 Then
		strData = ""
		txtVatLocAmt1 = 0
		For LngRow = 1 To intLoopCnt
			strData = strData & Chr(11) & ConvSPChars(rs0("VAT_TYPE"))
			strData = strData & Chr(11) & ConvSPChars(rs0("VAT_TYPE_NM"))
			strData = strData & Chr(11) & UNINumClientFormat(rs0("VAT_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & UNINumClientFormat(rs0("NET_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & ConvSPChars(rs0("GL_INPUT_TYPE"))
			strData = strData & Chr(11) & ConvSPChars(rs0("INPUT_TYPE_NM"))
			strData = strData & Chr(11) & UNIDateClientFormat(rs0("ISSUED_DT"))
			strData = strData & Chr(11) & ConvSPChars(rs0("BP_CD"))
			strData = strData & Chr(11) & ConvSPChars(rs0("BP_NM"))
			strData = strData & Chr(11) & ConvSPChars(rs0("TAX_BIZ_AREA_CD"))
			strData = strData & Chr(11) & ConvSPChars(rs0("TAX_BIZ_AREA_NM"))
			strData = strData & Chr(11) & LngMaxRow + LngRow
			strData = strData & Chr(11) & Chr(12)
			rs0.MoveNext
		Next
		
		If NOT(rs2.EOF) And NOT(rs2.BOF) Then
			txtVatLocAmt1 = UNINumClientFormat(rs2("SUM_VAT_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
		End If
		
	End If
	rs0.close:			Set rs0 = Nothing	                                                    '��: ActiveX Data Factory Object Nothing
	rs2.close:			Set rs2 = Nothing


	intLoopCnt = rs1.recordcount
	LngMaxRow = txtMaxRows2										'Save previous Maxrow
    If cint(intLoopCnt) <> 0 Then
		strData2 = ""
		txtGlLocAmt1 = 0
		For LngRow = 1 To intLoopCnt
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("VAT_TYPE"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("VAT_TYPE_NM"))
			strData2 = strData2 & Chr(11) & UNINumClientFormat(rs1("VAT_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData2 = strData2 & Chr(11) & UNINumClientFormat(rs1("NET_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("GL_INPUT_TYPE"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("INPUT_TYPE_NM"))
			strData2 = strData2 & Chr(11) & UNIDateClientFormat(rs1("ISSUED_DT"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("BP_CD"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("BP_NM"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("TAX_BIZ_AREA_CD"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("TAX_BIZ_AREA_NM"))
			strData2 = strData2 & Chr(11) & LngMaxRow + LngRow
			strData2 = strData2 & Chr(11) & Chr(12)
			rs1.MoveNext
		Next

		If NOT(rs3.EOF) And NOT(rs3.BOF) Then
			txtGlLocAmt1 = UNINumClientFormat(rs3("SUM_VAT_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
		End If
		
	End If

	rs1.close:			Set rs1 = Nothing	                                                    '��: ActiveX Data Factory Object Nothing
	rs3.close:			Set rs3 = Nothing
	
End Sub	

%>

<Script Language=vbscript>

		With parent
			.frm1.vspdData1.Redraw = False
			.frm1.vspdData2.Redraw = False
			.ggoSpread.Source = .frm1.vspddata1
			.ggoSpread.SSShowData "<%=strData%>"
			.ggoSpread.Source = .frm1.vspddata2
			.ggoSpread.SSShowData "<%=strData2%>"
			.frm1.txtVatLocAmt1.Text = "<%=txtVatLocAmt1%>"
			.frm1.txtGlLocAmt1.Text = "<%=txtGlLocAmt1%>"

				.DbQueryOK

			.frm1.vspdData1.Redraw = True
			.frm1.vspdData2.Redraw = True
			
		End With
</Script>


