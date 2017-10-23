<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% 

Err.Clear
On Error Resume Next


Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                         '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPageNo
Dim lgtxtMaxRows
Dim iLoopCount, iEndRow
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strCond
Dim strBatchNo, strRefNo
Dim strMsgCd, strMsg1, strMsg2
Const C_SHEETMAXROWS_D =30
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call HideStatusWnd 

	lgPageNo       = Request("lgPageNo")
    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgtxtMaxRows		= Request("txtMaxRows")

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()

'==========================================================================================
' Query Data
'==========================================================================================
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""

	If Len(Trim(lgPageNo)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          lgPageNo = CInt(lgPageNo)
       End If
    Else
       lgPageNo = 0
    End If

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If
    End If
	iLoopCount = 0

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
		iLoopCount = iRCnt
        rs0.MoveNext
    Next

	rs0.PageSize     = lgMaxCount
	rs0.AbsolutePage = lgPageNo + 1

    iRCnt = -1
	iEndRow = 0

    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
		iEndRow = iEndRow + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""          
        lgPageNo = ""                                        '☜: 다음 데이타 없다.
    End If

'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'==========================================================================================
' Set DB Agent arg
'==========================================================================================
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 

    Redim UNIValue(0,2)

    UNISqlId(0) = "A5140RA101"

    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    UNIValue(0,1) = " A.BATCH_NO =  " & FilterVar(strBatchNo, "''", "S") & "  "

	UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub
'==========================================================================================
' Query Data
'==========================================================================================
Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If  rs0.EOF And rs0.BOF Then
		strMsgCd = "900014"
    Else
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If

End Sub

'==========================================================================================
' Set default value or preset value
'==========================================================================================
Sub TrimData()

    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strBatchNo  = Request("txtBatchNo")
    strRefNo = Request("txtRefNo")

	strCond = ""

	strCond = strCond & " A.BATCH_NO =  " & FilterVar(strBatchNo, "''", "S") & "  "
	strMsg1 = Request("txtGlNo_Alt")

	'==============================================
	'Gl No, Ref No로 Gl Header 읽기 
	'==============================================
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,0)

    UNISqlId(0)   = "A5140RA102"
    UNIValue(0,0) = strCond

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If  rs1.EOF And rs1.BOF Then
		rs1.Close
		Set rs1 = Nothing
		Set lgADF = Nothing

		strMsgCd = "900014"
		Call DisplayMsgBox(strMsgCd, vbOKOnly, "", "", I_MKSCRIPT)

		Response.Write "<Script Language=vbScript>"	 & vbcr
		Response.Write "Parent.CancelClick()	  "	 & vbcr	
		Response.Write "</Script>				  "	 & vbcr	
		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else
		strBatchNo = Trim(rs1(0))
		strRefNo = Trim(rs1(2))
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtBatchNo.value		= "<%=ConvSPChars(Trim(rs1(0)))%>"		'전표번호 
				.frm1.txtInputType.value	= "<%=ConvSPChars(Trim(rs1(1)))%>"		'입력경로 
				.frm1.txtInputTypeNm.value	= "<%=ConvSPChars(Trim(rs1(2)))%>"		'입력경로 
				.frm1.txtGlDt.value			= "<%=UNIDateClientFormat(rs1(3))%>"	'전표일 
				.frm1.txtRefNo.value		= "<%=ConvSPChars(Trim(rs1(4)))%>"		'참조번호 
				.frm1.txtGlDesc.value		= "<%=ConvSPChars(Trim(rs1(5)))%>"		'비고 
				.frm1.txtBpcd1.value		= "<%=ConvSPChars(Trim(rs1(6)))%>"		'거래처 
				.frm1.txtBpcd1Nm.value		= "<%=ConvSPChars(Trim(rs1(7)))%>"
				.frm1.txtBpcd2.value		= "<%=ConvSPChars(Trim(rs1(8)))%>"		'입/출금처 
				.frm1.txtBpcd2Nm.value		= "<%=ConvSPChars(Trim(rs1(9)))%>"
				.frm1.txtBpcd3.value		= "<%=ConvSPChars(Trim(rs1(10)))%>"		'계산서발행처 
				.frm1.txtBpcd3Nm.value		= "<%=ConvSPChars(Trim(rs1(11)))%>"
				.frm1.txtInvDocNo.value		= "<%=ConvSPChars(Trim(rs1(12)))%>"		'입고번호 
				.frm1.txtInvDt.value		= "<%=UNIDateClientFormat(rs1(13))%>"
				.frm1.txtBlDocNo.value		= "<%=ConvSPChars(Trim(rs1(14)))%>"		'선하증권번호 
				.frm1.txtBlDt.value			= "<%=UNIDateClientFormat(rs1(15))%>"
				.frm1.txtLcDocNo.value		= "<%=ConvSPChars(Trim(rs1(16)))%>"		'L/C 번호 
				.frm1.txtLcDt.value			= "<%=UNIDateClientFormat(rs1(17))%>"
				.frm1.txtDueDt.value		= "<%=UNIDateClientFormat(rs1(18))%>"	'만기일자 

			End With
		</Script>
<%
    End If

	rs1.Close
	Set rs1 = Nothing
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData
         .ggoSpread.SSShowData "<%=lgstrData%>"              '☜: Display data
         .lgStrPrevKey        = "<%=lgStrPrevKey%>"    
		 .lgPageNo_A		=  "<%=lgPageNo%>"               '☜ : Next next data tag

         Call .DbQueryOk(1)
	End with
</Script>
<%
	Response.End 
%>
