<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->


<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I", "*","NOCOOKIE","QB")


'On Error Resume Next
'Err.clear

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag                              '☜ : DBAgent Parameter 선언Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim rs0, rs1, rs2, rs3, rs4, rs5
Dim lgPageNo                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strFromGlDt, strToGLDt, strDeptCd, strAcctCd, strInternalCd, strAcctCd2
Dim Fiscyyyymm00, Fiscyyyymm01,Fiscyyyymm02, Fiscyyyymm10
Dim TDrLocAmt,TCrLocAmt,NDrLocAmt,NCrLocAmt, Bal_Fg
Dim TDrSumAmt,NDrSumAmt,SDrSumAmt,TCrSumAmt,NCrSumAmt,SCrSumAmt,SDrAmt, SCrAmt
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1

Dim iChangeOrgId 
Dim strGlInputType
Dim strBpCd, strBpNm, StrDesc, strZeroFg
Dim strRefNo
Dim strAmtFr, strAmtTo
Dim txtStdYYYY
Dim txtStdMM
Dim txtBizAreaCd
Dim txtAcctCd
Dim txtInputType
Dim strInputTypeNM ,strInputType
Dim txtDrLocAmt, txtCrLocAmt

Const C_SHEETMAXROWS_D  = 30
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------


    Call HideStatusWnd 

    lgPageNo   = Request("lgPageNo")                               '☜ : Next key flag
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	iChangeOrgId   = Trim(Request("hOrgChangeId"))
	strGlInputType   = Trim(Request("cboGlInputType"))
	strRefNo   = Trim(Request("txtRefNo"))



    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0

    If Len(Trim(lgPageNo)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          iCnt = CInt(lgPageNo)
       End If
    End If

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1

    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""

		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

        If  iRCnt < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgPageNo = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If

'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "A5144RA101"	'계정보조부조회 
    UNISqlId(1) = "A5144RA102"	'계정코드 
	UNISqlId(2) = "ABIZNM"		'사업장 
	UNISqlId(3) = "A5144MA103"  '전표입력경로 
	UNISqlId(4) = "A5144RA103"  '전표입력경로 

	Redim UNIValue(5,3)
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = UCase(Trim(strWhere0))

	UNIValue(1,0) = FilterVar(txtAcctCd, "''", "S")

	UNIValue(2,0) = FilterVar(txtBizAreaCd, "''", "S") 
	UNIValue(3,0) = "" & FilterVar("A1001", "''", "S") & " "
	UNIValue(3,1) = FilterVar(strInputType, "''", "S")
	UNIValue(4,0) = UCase(Trim(strWhere0))

	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,2) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

    iStr = Split(lgstrRetMsg,gColSep)

	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strAcctCd <> "" Then 
			strMsgCd = "110100"												'Not Found	
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtAcctCd.value = "<%=ConvSPChars(txtAcctCd)%>"
			.txtAcctCdNm.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			End With
		</Script>
<%	
		If IsNull(rs1(1)) = False then Bal_fg    = ConvSPChars(Trim(rs1(1)))
	End If

	rs1.Close
	Set rs1 = Nothing

	If  rs2.EOF And rs2.BOF Then
		If strMsgCd = "" And txtBizAreaCd <> "" Then 
			strMsgCd = "124200"												'Not Found	
		end if
    Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(txtBizAreaCd)%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
			End With
		</Script>
<%
    End If

	rs2.Close
	Set rs2 = Nothing

    'rs3에 대한 결과 
    IF NOT (rs3.EOF or rs3.BOF) then
		strInputTypeNM	= Trim(rs3("minor_nm"))
		strInputType	= Trim(rs3("minor_cd"))
%>
		<Script Language=vbscript>
			With parent.frm1
				.txtInputType.value	= "<%=strInputType%>"
				.txtInputTypeNm.value	= "<%=strInputTypeNM%>"
			End With
		</Script>
<%
    END IF
    rs3.Close
    Set rs3 = Nothing

    'rs4에 대한 결과 
    IF NOT (rs4.EOF or rs4.BOF) then
		txtDrLocAmt	= UNINumClientFormat(rs4(0), ggAmtOfMoney.DecPoint, 0)
		txtCrLocAmt	= UNINumClientFormat(rs4(1), ggAmtOfMoney.DecPoint, 0)
%>
		<Script Language=vbscript>
			With parent.frm1
				.txtDrLocAmt.text	= "<%=txtDrLocAmt%>"
				.txtCrLocAmt.text	= "<%=txtCrLocAmt%>"
			End With
		</Script>
<%
    END IF
    rs4.Close
    Set rs4 = Nothing


	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, "", I_MKSCRIPT)
		rs0.Close
        Set rs0 = Nothing
        Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 

	Set lgADF = Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
	End If
	Set rs2 = Nothing


End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

	Dim txtFromDt, txtToDt
	txtFromDt		= Trim(Request("txtFromDt"))
	txtToDt			= Trim(Request("txtToDt"))
	txtBizAreaCd	= UCase(Trim(Request("txtBizAreaCd")))
	txtAcctCd		= UCase(Trim(Request("txtAcctCd")))
	strInputType		= UCase(Trim(Request("txtInputType")))

	strWhere0 = ""
	strWhere0 = strWhere0 & " A.Acct_cd = " & FilterVar(txtAcctCd, "''", "S")
	strWhere0 = strWhere0 & " AND b.gl_input_type = " & FilterVar(strInputType, "''", "S")
'	If Trim(txtStdMM) = "" Then
'		strFrDt = txtStdYYYY +  "0101"
'		strToDt = txtStdYYYY +  "1231"
		strWhere0 = strWhere0 & " AND b.gl_dt between  " & FilterVar(txtFromDt, "''", "S") & " and  " & FilterVar(txtToDt, "''", "S") & " "
'	Else
'		strFrDt = txtStdYYYY & "-" & txtStdMM
'		'strToDt = strDt + strDateMnth + "31"
'		
'		strWhere0 = strWhere0 & " AND convert(char(7),CONVERT(DATETIME,b.GL_DT ),121) = '" & strFrDt & "'"
'	End If


	If txtBizAreaCd <> "" Then
		'Call fnGetBpCd
		strWhere0 = strWhere0 & " AND b.biz_area_cd = " & FilterVar(txtBizAreaCd, "''", "S")
	End If

End Sub



%>

<Script Language=vbscript>
	With parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
		.lgPageNo =  "<%=ConvSPChars(lgPageNo)%>"                       '☜: set next data tag
		.DbQueryOk
	End with
</Script>

<%
	Response.End 
%>
