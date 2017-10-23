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


On Error Resume Next
Err.clear

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag                              '☜ : DBAgent Parameter 선언Dim lgstrData  
Dim rs0, rs1, rs2, rs3, rs4, rs5
Dim lgPageNo2                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strFromGlDt, strToGLDt, strDeptCd, strAcctCd, strInternalCd, strAcctCd2
Dim TDrLocAmt,TCrLocAmt,NDrLocAmt,NCrLocAmt, Bal_Fg
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1
Dim strInputType, strInputTypeNM
Dim strBpCd, strBpNm, StrDesc, strZeroFg
Dim strRefNo
Dim strAmtFr, strAmtTo
Dim lgtxtBizArea
Dim biz_area_cd, biz_area_nm
Dim strIssuedDt, strIssuedDt2
Dim strMoCd, strMoNm
Const C_SHEETMAXROWS_D  = 100
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------


    Call HideStatusWnd 

    lgPageNo2			= Request("lgPageNo2")                               '☜ : Next key flag
    lgSelectList		= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList			= Request("lgTailList")                                 '☜ : Orderby value
	strInputType		= Trim(Request("txtInputType"))
	strRefNo			= Trim(Request("txtRefNo"))
	lgtxtBizArea		= Trim(Request("txtBizAreaCd"))



    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'==========================================================================================
' Query Data
'==========================================================================================
Sub MakeSpreadSheetData()
	On Error Resume Next
	Err.clear

    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0

    If Len(Trim(lgPageNo2)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo2) Then
          iCnt = CInt(lgPageNo2)
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
            lgPageNo2 = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgPageNo2 = ""                                                  '☜: 다음 데이타 없다.
    End If

'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	On Error Resume Next
	Err.clear

    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 

    UNISqlId(0) = "A5144MA105"	'조회 
    UNISqlId(1) = "A5144MA106"	'차/대변합계 
	UNISqlId(2) = "ABIZNM"		'사업장 
	UNISqlId(3) = "A5144MA103"  '전표입력경로 
	UNISqlId(4) = "A5144MA103"  '모듈 

	Redim UNIValue(4,3)
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    UNIValue(0,1) = UCase(Trim(strWhere0))


	UNIValue(1,0) = strWhere0

	UNIValue(2,0) = FilterVar(lgtxtBizArea, "''", "S")
	UNIValue(3,0) = "" & FilterVar("A1001", "''", "S") & " "
	UNIValue(3,1) = FilterVar(strInputType, "''", "S")
	UNIValue(4,0) = "" & FilterVar("B0001", "''", "S") & " "
	UNIValue(4,1) = FilterVar(strMocd, "''", "S")
    UNIValue(0,2) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	On Error Resume Next
	Err.clear

    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If


	NDrLocAmt = 0
	NCrLocAmt = 0 

	If Not(rs1.EOF And rs1.BOF) Then
		If IsNull(rs1(0)) = False Then 
			TDrLocAmt    = rs1(0)
		Else
			TDrLocAmt    = 0
		End If
		If IsNull(rs1(1)) = False Then 
			TCrLocAmt    = rs1(1)
		Else
			TCrLocAmt    = 0
		End If
	End If

	rs1.Close
	Set rs1 = Nothing


    'rs1에 대한 결과 
    IF NOT (rs2.EOF or rs2.BOF) then
	    biz_area_nm = Trim(rs2("biz_area_nm"))
	    biz_area_cd = Trim(rs2("biz_area_cd"))
    END IF
    rs2.Close
    Set rs2 = Nothing

    'rs3에 대한 결과 
    IF NOT (rs3.EOF or rs3.BOF) then
		strInputTypeNM	= Trim(rs3("minor_nm"))
		strInputType	= Trim(rs3("minor_cd"))
    END IF
    rs3.Close
    Set rs3 = Nothing

    'rs3에 대한 결과 
    IF NOT (rs4.EOF or rs4.BOF) then
		strMoNm			= Trim(rs4("minor_nm"))
		strMoCd			= Trim(rs4("minor_cd"))
    END IF
    rs4.Close
    Set rs4 = Nothing

    %>

    <Script Language=vbscript>
		With parent
		.frm1.txtNDrAmt2.text		= "<%=UNINumClientFormat(TDrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtNCrAmt2.text		= "<%=UNINumClientFormat(TCrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"

		.frm1.txtBizAreaNm2.value	= "<%=biz_area_nm%>"
		.frm1.txtBizAreacd2.value	= "<%=biz_area_cd%>"

		.frm1.txtInputType2.value	= "<%=strInputType%>"
		.frm1.txtInputTypeNm2.value	= "<%=strInputTypeNM%>"


		End With
	</script>
	<%
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, "", I_MKSCRIPT)
		rs0.Close
        Set rs0 = Nothing
        Response.End
	End If

    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 

	Set lgADF = Nothing 

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
	End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

	strFromGlDt		= Request("txtFromGlDt")
	strToGLDt		= Request("txtToGlDt")
	strIssuedDt		= Request("txtIssuedDt")
	strIssuedDt2	= Request("txtIssuedDt2")
	strMocd			= Trim(Request("txtMocd"))

	strWhere0 = ""
	strWhere0 = strWhere0 & "  A.GL_DT BETWEEN  " & FilterVar(strFromGlDt, "''", "S") & " AND  " & FilterVar(strToGLDt, "''", "S") & " "

	If lgtxtBizArea <> "" Then
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD = " & FilterVar(lgtxtBizArea, "''", "S")
	End If
	If strInputType <> "" Then
		strWhere0 = strWhere0 & " AND A.GL_INPUT_TYPE = " & FilterVar(strInputType, "''", "S")
	End If

	If strIssuedDt <> "" Then
		strWhere0 = strWhere0 & " AND A.ISSUED_DT >=  " & FilterVar(strIssuedDt , "''", "S") & " "
	End If
	If strIssuedDt2 <> "" Then
		strWhere0 = strWhere0 & " AND A.ISSUED_DT <=  " & FilterVar(strIssuedDt2 , "''", "S") & " "
	End If
	If strMocd <> "" Then
		strWhere0 = strWhere0 & " AND G.REFERENCE= " & FilterVar(strMocd, "''", "S")
	End If

End Sub


'--------------------------------------------


%>

<Script Language=vbscript>
	With parent
		If "<%=lgPageNo2%>" = "1" Then   ' "1" means that this query is first and next data exists
		   .frm1.htxtFromGlDt2.value	= "<%=ConvSPChars(strFromGlDt)%>"
		   .frm1.htxtToGlDt2.value		= "<%=ConvSPChars(strToGLDt)%>"
		   .frm1.htxtIssuedDt21.value	= "<%=ConvSPChars(strIssuedDt)%>"
		   .frm1.htxtIssuedDt22.value	= "<%=ConvSPChars(strIssuedDt2)%>"
		   .frm1.htxtInputType.value	= "<%=ConvSPChars(strInputType)%>"
		   .frm1.htxtBizAreaCd2.value	= "<%=ConvSPChars(lgtxtBizArea)%>"
		End If
		.ggoSpread.Source = .frm1.vspdData3
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
		.lgPageNo2 =  "<%=ConvSPChars(lgPageNo2)%>"                       '☜: set next data tag
		.DbQueryOk2
	End with
</Script>
<%
	Response.End 
%>
