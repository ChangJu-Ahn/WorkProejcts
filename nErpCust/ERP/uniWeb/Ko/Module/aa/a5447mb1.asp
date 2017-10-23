<%Option Explicit%>



<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<% 
	Call LoadBasisGlobalInf() 

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")


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
Dim strFrDt, strBizAreaCd, strFrAcctCd
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1
Dim biz_area_cd, biz_area_nm

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

	Redim UNISqlId(2) 


	UNISqlId(0) = "A5447MA101"	'조회 
	UNISqlId(1) = "A5124MA103"	'계정코드 
	UNISqlId(2) = "ABIZNM"		'사업장 

	Redim UNIValue(3,6)

	UNIValue(0,0) = lgSelectList                                          '☜: Select list
	UNIValue(0,1) = strWhere0
	UNIValue(0,5) = strWhere1

	UNIValue(1,0) = FilterVar(strFrAcctCd, "''", "S")
	UNIValue(2,0) = FilterVar(strBizAreaCd, "''", "S")

	UNIValue(0,6) = UCase(Trim(lgTailList))
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'==========================================================================================
' Query Data
'==========================================================================================
Sub QueryData()
	'On Error Resume Next
	'Err.clear

    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If


'rs1 --------------
	If (rs1.EOF And rs1.BOF) Then
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtFrAcctNm.value = ""
			End With
		</Script>
<%	
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtFrAcctCd.value = "<%=ConvSPChars(strFrAcctCd)%>"
			.txtFrAcctNm.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			End With
		</Script>
<%
	End If

	rs1.Close
	Set rs1 = Nothing



'rs2 --------------
    IF  (rs2.EOF or rs2.BOF) then
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaNm.value = ""
			End With
		</Script>
<%
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
			End With
		</Script>
<%
    END IF
    rs2.Close
    Set rs2 = Nothing



    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
		Set lgADF = Nothing
        Response.End
    Else
        Call  MakeSpreadSheetData()
    End If

	rs0.Close

	Set rs0 = Nothing 
	Set lgADF = Nothing
End Sub

'==========================================================================================
' Set default value or preset value
'==========================================================================================
Sub TrimData()
	On Error Resume Next
	Err.clear

	strFrDt			= UCase(Trim(Request("txtFrDt")))
	strBizAreaCd	= UCase(Trim(Request("txtBizAreaCd")))
	strFrAcctCd		= UCase(Trim(Request("txtFrAcctCd")))


	strWhere0 = ""
	strWhere1 = ""
	strWhere0 = strWhere0 & " WHERE HIS_DT <= DATEADD(DAY,-1,DATEADD(MONTH, +1,  CONVERT(DATETIME, " & FilterVar(strFrDt, "''", "S") & " + " & FilterVar("01", "''", "S") & " ,120)))"
	strWhere0 = strWhere0 & " AND	HIS_DT >= CONVERT(DATETIME, LEFT( " & FilterVar(strFrDt, "''", "S") & ",4)+" & FilterVar("0101", "''", "S") & " ,120)"

	If strFrAcctCd <> "" Then
		strWhere0 = strWhere0 & " AND ASST_ACCT_CD =  " & FilterVar(strFrAcctCd, "''", "S") & "  "
	End If



	If strBizAreaCd <> "" Then
		strWhere1 = strWhere1 & " WHERE A.BIZ_AREA_CD = " & FilterVar(strBizAreaCd, "''", "S") & "  "
	End If
End Sub

'==========================================================================================
%>

<Script Language=vbscript>
	With parent
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
		   .frm1.htxtFrDt.value				= "<%=ConvSPChars(strFrDt)%>"
		   .frm1.htxtBizAreaCd.value		= "<%=ConvSPChars(strBizAreaCd)%>"
		   .frm1.htxtFrAcctCd.value			= "<%=ConvSPChars(strFrAcctCd)%>"
		End If
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
		.lgPageNo =  "<%=ConvSPChars(lgPageNo)%>"                       '☜: set next data tag
		.DbQueryOk
	End with
</Script>

<%
	Response.End
%>
