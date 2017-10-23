<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next								'☜: 

Call LoadBasisGlobalInf()

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3						'☜ : DBAgent Parameter 선언 
Dim lgstrData																'☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Const C_SHEETMAXROWS_D = 30

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim txtFrDt, txtToDt, txtVatIoFg, txtVatTypeCd, txtGlInputCd, txtIssuedDt, txtBpCd, txtBizAreaCd
Dim strData, strData2, txtVatLocAmt1, txtGlLocAmt1

Dim  iLoopCount
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
	Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"

    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim intLoopCnt

	'-----------------------
	'Result data display area
	'-----------------------	
	intLoopCnt = rs0.recordcount

    If cint(intLoopCnt) <> 0 Then
		strData = ""
		txtVatLocAmt1 = 0
		For iLoopCount = 1 To intLoopCnt
			strData = strData & Chr(11) & UNIDateClientFormat(rs0("ISSUED_DT"))
			strData = strData & Chr(11) & UNINumClientFormat(rs0("VAT_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & UNINumClientFormat(rs0("NET_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & ConvSPChars(rs0("BP_CD"))
			strData = strData & Chr(11) & ConvSPChars(rs0("BP_NM"))
			strData = strData & Chr(11) & ConvSPChars(rs0("REF_NO"))
			strData = strData & Chr(11) & ConvSPChars(rs0("REPORT_BIZ_AREA_CD"))
			strData = strData & Chr(11) & ConvSPChars(rs0("TAX_BIZ_AREA_NM"))
			strData = strData & Chr(11) & ConvSPChars(rs0("GL_NO"))
			strData = strData & Chr(11) & UNIDateClientFormat(rs0("GL_DT"))
			strData = strData & Chr(11) & ConvSPChars(rs0("VAT_NO"))
			strData = strData & Chr(11) & iLoopCount
			strData = strData & Chr(11) & Chr(12)
			rs0.MoveNext
		Next

		If NOT(rs2.EOF) And NOT(rs2.BOF) Then
			txtVatLocAmt1 = UNINumClientFormat(rs2(1), ggAmtOfMoney.DecPoint, 0)
		End If

	End If
	rs0.close:			Set rs0 = Nothing	                                                    '☜: ActiveX Data Factory Object Nothing
	rs2.close:			Set rs2 = Nothing

	intLoopCnt = rs1.recordcount
    If cint(intLoopCnt) <> 0 Then
		strData2 = ""
		txtGlLocAmt1 = 0
		For iLoopCount = 1 To intLoopCnt
			strData2 = strData2 & Chr(11) & UNIDateClientFormat(rs1("ISSUED_DT"))
			strData2 = strData2 & Chr(11) & UNINumClientFormat(rs1("ITEM_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData2 = strData2 & Chr(11) & UNINumClientFormat(rs1("NET_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("BP_CD"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("BP_NM"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("REF_NO"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("REPORT_BIZ_AREA_CD"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("TAX_BIZ_AREA_NM"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs1("GL_NO"))
			strData2 = strData2 & Chr(11) & UNIDateClientFormat(rs1("GL_DT"))
			strData2 = strData2 & Chr(11) & iLoopCount
			strData2 = strData2 & Chr(11) & Chr(12)
			rs1.MoveNext
		Next

		If NOT(rs3.EOF) And NOT(rs3.BOF) Then
			txtGlLocAmt1 = UNINumClientFormat(rs3(1), ggAmtOfMoney.DecPoint, 0)
		End If

	End If

	rs1.close:			Set rs1 = Nothing	                                                    '☜: ActiveX Data Factory Object Nothing
	rs3.close:			Set rs3 = Nothing
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim strWhere, strGroup
    Redim UNIValue(3,3)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

	txtFrDt			= Trim(Request("txtFrDt"))
	txtToDt			= Trim(Request("txtToDt"))
	txtVatIoFg		= UCase(Trim(Request("txtVatIoFg")))
	txtVatTypeCd	= UCase(Trim(Request("txtVatTypeCd")))
	txtGlInputCd	= UCase(Trim(Request("txtGlInputCd")))
	txtIssuedDt		= Trim(Request("txtIssuedDt"))
	txtBpCd			= UCase(Trim(Request("txtBpCd")))
	txtBizAreaCd	= UCase(Trim(Request("txtBizAreaCd")))

	strWhere = ""
	If txtGlInputCd <> "" Then strWhere = strWhere & " AND A.GL_INPUT_TYPE = "		& Filtervar(txtGlInputCd	, "''", "S")
	If txtVatTypeCd <> "" Then strWhere = strWhere & " AND A.VAT_TYPE = "			& Filtervar(txtVatTypeCd	, "''", "S")
	If txtIssuedDt	<> "" Then strWhere = strWhere & " AND A.ISSUED_DT = "			& Filtervar(UniConvDate(txtIssuedDt)	, null, "S")
	If txtBpCd		<> "" Then strWhere = strWhere & " AND A.BP_CD = "				& Filtervar(txtBpCd	, "''", "S")
	If txtBizAreaCd <> "" Then strWhere = strWhere & " AND A.REPORT_BIZ_AREA_CD= "	& Filtervar(txtBizAreaCd	, "''", "S")

    UNISqlId(0) = "a5461ra101"
    UNISQLID(1) = "a5461ra102"
    UNISqlId(2) = "a5461ra103"
    UNISQLID(3) = "a5461ra104"

    UNIValue(0,0) = Filtervar(UniConvDate(txtFrDt)	, "''", "S") 
    UNIValue(0,1) = Filtervar(UniConvDate(txtToDt)	, "''", "S") 
    UNIValue(0,2) = strWhere & " AND A.IO_FG = "	& Filtervar(txtVatIoFg	, "''", "S")

    UNIValue(2,0) = Filtervar(UniConvDate(txtFrDt)	, "''", "S") 
    UNIValue(2,1) = Filtervar(UniConvDate(txtToDt)	, "''", "S") 
    UNIValue(2,2) = strWhere & " AND A.IO_FG = "	& Filtervar(txtVatIoFg	, "''", "S")

	If txtVatIoFg = "I" Then strWhere = strWhere & " AND A.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  AND A.ACCT_TYPE = " & FilterVar("VP", "''", "S") & " "
	If txtVatIoFg = "O" Then strWhere = strWhere & " AND A.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  AND A.ACCT_TYPE = " & FilterVar("VR", "''", "S") & " "
    
    UNIValue(1,0) = Filtervar(UniConvDate(txtFrDt)	, "''", "S") 
    UNIValue(1,1) = Filtervar(UniConvDate(txtToDt)	, "''", "S") 
    UNIValue(1,2) = strWhere

    UNIValue(3,0) = Filtervar(UniConvDate(txtFrDt)	, "''", "S") 
    UNIValue(3,1) = Filtervar(UniConvDate(txtToDt)	, "''", "S") 
    UNIValue(3,2) = strWhere
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If  (rs0.EOF And rs0.BOF) AND (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close :		Set rs0 = Nothing
        rs1.Close :		Set rs1 = Nothing
        rs2.Close :		Set rs2 = Nothing
        rs3.Close :		Set rs3 = Nothing
        Exit Sub
    Else
        Call MakeSpreadSheetData()
    End If
    
End Sub
%>

<Script Language=vbscript>

		With parent
			.frm1.vspdData.Redraw = False
			.frm1.vspdData1.Redraw = False
			.ggoSpread.Source = .frm1.vspddata
			.ggoSpread.SSShowData "<%=strData%>"
			.ggoSpread.Source = .frm1.vspddata1
			.ggoSpread.SSShowData "<%=strData2%>"
			.frm1.txtVatLocAmt1.Text = "<%=txtVatLocAmt1%>"
			.frm1.txtGlLocAmt1.Text = "<%=txtGlLocAmt1%>"

			.DbQueryOK

			.frm1.vspdData.Redraw = True
			.frm1.vspdData1.Redraw = True
			
		End With
	 
</Script>
