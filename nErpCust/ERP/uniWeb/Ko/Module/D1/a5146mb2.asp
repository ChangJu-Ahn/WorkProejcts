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
Dim UNISqlId, UNIValue, UNILock, UNIFlag                              '☜ : DBAgent Parameter 선언Dim lgstrData
Dim rs0, rs1, rs2, rs3, rs4, rs5
Dim lgPageNo                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strAcctCd, strAcctCd2
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1
Const C_SHEETMAXROWS_D  = 100
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call HideStatusWnd 

    lgPageNo			= Request("lgPageNo")                               '☜ : Next key flag
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

    If Len(Trim(lgPageNo)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          iCnt = CInt(lgPageNo)
       End If
    End If

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D
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
'==========================================================================================
' Set DB Agent arg
'==========================================================================================
Sub FixUNISQLData()
	On Error Resume Next
	Err.clear

    Redim UNISqlId(1) 


    UNISqlId(0) = "A5146MA102"	'조회 

	Redim UNIValue(1,3)

    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    UNIValue(0,1) = UCase(Trim(strWhere0))

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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If


	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, "", I_MKSCRIPT)
		rs0.Close
        Set rs0 = Nothing
        Response.End
	End If

    If  rs0.EOF And rs0.BOF Then
        'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
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

'==========================================================================================
' Set default value or preset value
'==========================================================================================
Sub TrimData()

	strAcctCd		= UCase(Trim(Request("txtAcctCd")))
	strAcctCd2		= UCase(Trim(Request("txtAcctCd2")))

	strWhere0 = ""
	strWhere0 = strWhere0 & "  A.Acct_cd =  " & FilterVar(strAcctCd, "''", "S") & "  "

'	If strAcctCd <> "" Then
'		strWhere0 = strWhere0 & " AND A.Acct_cd >= '" & FilterVar(strAcctCd,"","S") & "' "
'	End If
'	If strAcctCd2 <> "" Then
'		strWhere0 = strWhere0 & " AND A.Acct_cd <= '" & FilterVar(strAcctCd2,"","S") & "' "
'	End If

End Sub

'==========================================================================================
%>

<Script Language=vbscript>
	With parent
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
		   .frm1.htxtAcctCd2.value			= "<%=ConvSPChars(strAcctCd)%>"
		End If
		.ggoSpread.Source = .frm1.vspdData2 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
		.lgPageNo2 =  "<%=ConvSPChars(lgPageNo)%>"                       '☜: set next data tag
		.DbQueryOk2
	End with
</Script>

<%
	Response.End
%>
