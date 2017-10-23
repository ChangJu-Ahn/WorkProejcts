<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->

<% 
Call LoadBasisGlobalInf()


On Error Resume Next
Err.clear

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag                              '☜ : DBAgent Parameter 선언Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim rs0, rs1, rs2, rs3, rs4, rs5
Dim lgPageNo2                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strAcctCd, strAcctCd2
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1
Dim strRefNo
Dim txtClassType
Dim biz_area_cd, biz_area_nm
Const C_SHEETMAXROWS_D  = 100
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------


    Call HideStatusWnd 

    lgPageNo2			= Request("lgPageNo2")                               '☜ : Next key flag
    lgSelectList		= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList			= Request("lgTailList")                                 '☜ : Orderby value

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

    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 


    UNISqlId(0) = "A5145MA102"	'조회 
    UNISqlId(1) = "A5124MA103"	'계정코드 
	UNISqlId(2) = "A5124MA103"
	UNISqlId(3) = "A5145MA103"

	Redim UNIValue(6,3)

    UNIValue(0,0) = lgSelectList                                          '☜: Select list

    UNIValue(0,1) = UCase(Trim(strWhere0))

	UNIValue(1,0) = FilterVar(strAcctCd, "''", "S")
	UNIValue(2,0) = FilterVar(strAcctCd2, "''", "S")
    UNIValue(3,0) = FilterVar(txtClassType, "''", "S")

    UNIValue(0,2) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub
'==========================================================================================
' Query Data
'==========================================================================================
Sub QueryData()
	On Error Resume Next
	Err.clear

    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If



	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strAcctCd <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtAcctCd_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtAcctCd21.value = "<%=ConvSPChars(strAcctCd)%>"
			.txtAcctNm21.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			End With
		</Script>
<%
		If IsNull(rs1(1)) = False then Bal_fg    = ConvSPChars(Trim(rs1(1)))
	End If

	rs1.Close
	Set rs1 = Nothing

	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strAcctCd2 <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtAcctCd_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtAcctCd22.value = "<%=ConvSPChars(strAcctCd2)%>"
			.txtAcctNm22.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
			End With
		</Script>
<%
'		If IsNull(rs2(1)) = False then Bal_fg    = ConvSPChars(Trim(rs2(1)))
	End If

	rs2.Close
	Set rs2 = Nothing


	If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" Then 
			strMsgCd = "110500"												'Not Found	
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtClassType2.value	= "<%=ConvSPChars(Trim(rs3(0)))%>"
			.txtClassTypeNm2.value	= "<%=ConvSPChars(Trim(rs3(1)))%>"
			End With
		</Script>
<%	

	End If

	rs3.Close
	Set rs3 = Nothing

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
End Sub

'==========================================================================================
' Set default value or preset value
'==========================================================================================
Sub TrimData()

	txtClassType	= UCase(Trim(Request("txtClassType")))
	strAcctCd		= UCase(Trim(Request("txtAcctCd")))
	strAcctCd2		= UCase(Trim(Request("txtAcctCd2")))

	strWhere0 = ""

	If strAcctCd <> "" Then
		strWhere0 = strWhere0 & " AND A.Acct_cd >=  " & FilterVar(strAcctCd, "''", "S") & "  "
	End If
	If strAcctCd2 <> "" Then
		strWhere0 = strWhere0 & " AND A.Acct_cd <=  " & FilterVar(strAcctCd2, "''", "S") & "  "
	End If

	strWhere0 = strWhere0 & " 	AND  a.ACCT_CD         NOT iN  (select f.ACCT_CD from A_CLSSFC_OF_ACCT f "
	strWhere0 = strWhere0 & "		where f.QUERY_TYPE_FG = " & FilterVar("Y", "''", "S") & "  AND f.CLASS_TYPE = " & FilterVar(txtClassType, "''", "S") & "  "
	strWhere0 = strWhere0 & "			group by f.ACCT_CD) "

End Sub

'==========================================================================================
%>

<Script Language=vbscript>
	With parent
		If "<%=lgPageNo2%>" = "1" Then   ' "1" means that this query is first and next data exists
		   .frm1.htxtClassType2.value		= "<%=ConvSPChars(txtClassType)%>"
		   .frm1.htxtAcctCd21.value			= "<%=ConvSPChars(strAcctCd)%>"
		   .frm1.htxtAcctCd22.value			= "<%=ConvSPChars(strAcctCd2)%>"
		End If
		.ggoSpread.Source = .frm1.vspdData2 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
		.lgPageNo2 =  "<%=ConvSPChars(lgPageNo2)%>"                       '☜: set next data tag
		.DbQueryOk2
	End with

</Script>
<%
	Response.End 
%>
