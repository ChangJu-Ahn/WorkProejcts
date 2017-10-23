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
Dim UNISqlId, UNIValue, UNILock, UNIFlag                              '☜ : DBAgent Parameter 선언Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7
Dim lgPageNo                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strFromGlDt, strToGLDt, strDeptCd, strAcctCd, strInternalCd, strAcctCd2
Dim TDrLocAmt,TCrLocAmt,NDrLocAmt,NCrLocAmt, Bal_Fg
Dim TDrSumAmt,NDrSumAmt,SDrSumAmt,TCrSumAmt,NCrSumAmt,SCrSumAmt,SDrAmt, SCrAmt
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1
Dim strGetInternalNm
Dim Fiscyyyy,Fiscmm,Fiscdd, strGlDtYr, strGlDtMnth, strGlDtDt
Dim strCompFiscStartDt
Dim iChangeOrgId 
Dim strGlInputType
Dim strBpCd, strBpNm, StrDesc, strZeroFg,StrProject
Dim strRefNo
Dim strAmtFr, strAmtTo
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

Const C_SHEETMAXROWS_D  = 100
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

    Redim UNISqlId(7)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    if Request("txtFG") = "1" then
		UNISqlId(0) = "A5124MA101"	'계정보조부조회 
    Else
		UNISqlId(0) = "A5124MA201"	'계정보조부조회 
    End if
    UNISqlId(1) = "A5124MA103"	'계정코드 
	UNISqlId(2) = "A5124MA103"
    UNISqlId(3) = "A5124MA104"	'이월금액 
    UNISqlId(4) = "A5124MA105"	'발생금액차변 
    UNISqlId(5) = "A5124MA107"	'발생금액대변 
	UNISqlId(6) = "A_GETBIZ"
    UNISqlId(7) = "A_GETBIZ"
    
	Redim UNIValue(7,17)
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    
    UNIValue(0,1) = FilterVar(strFromGlDt, "''", "S")
	UNIValue(0,2) = FilterVar(strToGLDt, "''", "S")
	UNIValue(0,3) = FilterVar(strAcctCd, "''", "S")
	UNIValue(0,4) = FilterVar(strAcctCd2, "''", "S")
	
	IF strGlInputType <> "" THEN
		UNIValue(0,5) = FilterVar(strGlInputType, "''", "S")
	ELSE
		UNIValue(0,5) = FilterVar("%", "''", "S")
	END IF
	
	IF strBizAreaCd <> "" THEN
		' 권한관리 추가 
		If lgAuthBizAreaCd <> "" Then
			UNIValue(0,6) = FilterVar(lgAuthBizAreaCd, "''", "S")
		Else
			UNIValue(0,6) = FilterVar(strBizAreaCd, "''", "S")
		End If
	ELSE
		If lgAuthBizAreaCd <> "" Then
			UNIValue(0,6) = FilterVar(lgAuthBizAreaCd, "''", "S")
		Else
			UNIValue(0,6) = FilterVar("0", "''", "S")
		End If
	END IF
	
	IF strBizAreaCd1 <> "" THEN
		If lgAuthBizAreaCd <> "" Then
			UNIValue(0,7) = FilterVar(lgAuthBizAreaCd, "''", "S")
		Else
			UNIValue(0,7) = FilterVar(strBizAreaCd1, "''", "S")
		End If
	ELSE
		If lgAuthBizAreaCd <> "" Then
			UNIValue(0,7) = FilterVar(lgAuthBizAreaCd, "''", "S")
		Else
			UNIValue(0,7) = FilterVar("ZZZZZZZZZZ", "''", "S")
		End If
	END IF

	IF strInternalCd <> "" THEN
		UNIValue(0,8) = FilterVar(strInternalCd, "''", "S")
		
		If lgInternalCd <> "" Then
			UNIValue(0,8) = FilterVar(lgInternalCd, "''", "S")
		End If

		If lgSubInternalCd <> "" Then
			UNIValue(0,8) = FilterVar(lgSubInternalCd & "%", "''", "S")
 		End If
	ELSE
		UNIValue(0,8) = FilterVar("%", "''", "S")
		
		If lgInternalCd <> "" Then
			UNIValue(0,8) = FilterVar(lgInternalCd, "''", "S")
		End If

		If lgSubInternalCd  <> "" Then
			UNIValue(0,8) = FilterVar(lgSubInternalCd & "%", "''", "S")
 		End If
	END IF

	IF StrDesc <> "" THEN
		UNIValue(0,9) = FilterVar("%" & StrDesc & "%" , "''", "S")
	ELSE
		UNIValue(0,9) = FilterVar("%", "''", "S")
	END IF
	IF StrProject <> "" THEN
		UNIValue(0,10) = FilterVar("%" & StrProject & "%" , "''", "S")
	ELSE
		UNIValue(0,10) = FilterVar("%", "''", "S")
	END IF
	IF strBpCd <> "" THEN
		UNIValue(0,11) = FilterVar(strBpCd, "''", "S")
	ELSE
		UNIValue(0,11) = FilterVar("%", "''", "S")
	END IF
	IF strRefNo <> "" THEN
		UNIValue(0,12) = FilterVar(strRefNo & "%", "''", "S")
	ELSE
		UNIValue(0,12) = FilterVar("%", "''", "S")
	END IF
	
	IF strAmtFr <> "0" THEN
		UNIValue(0,13) = FilterVar(strAmtFr, "''", "SNM")
	ELSE
		UNIValue(0,13) = cdbl(-99999999999999)
	END IF
	IF strAmtTo <> "0" THEN
		UNIValue(0,14) = FilterVar(strAmtTo, "''", "SNM")
	ELSE
		UNIValue(0,14) = FilterVar("999999999999999", "''", "SNM")
	END IF	
	
	UNIValue(1,0) = FilterVar(strAcctCd, "''", "S")

	UNIValue(2,0) = FilterVar(strAcctCd2, "''", "S")
	
	UNIValue(3,0) = FilterVar(strAcctCd, "''", "S")
	UNIValue(3,1) = FilterVar(strAcctCd2, "''", "S")
	UNIValue(3,2) = Trim(strWhere1)
		
	UNIValue(4,0) = Trim(strWhere0)

	UNIValue(5,0) = Trim(strWhere0)
	
	UNIValue(6,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(7,0)  = FilterVar(strBizAreaCd1, "''", "S")
	
	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		UNIValue(6,0) = UNIValue(6,0) & " AND BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
		UNIValue(7,0) = UNIValue(7,0) & " AND BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,15) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7)
    
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    If Trim(strGetInternalNm) = "" Then
       If strMsgCd = "" And strDeptCd <> "" Then 
		  strMsgCd = "970000"												'Not Found
          strMsg1 = Request("txtDeptCd_Alt")
       End If
    %>	

   <%	

    Else
    %>	
    <Script Language=vbScript>
	  With parent
		.frm1.txtDeptCd.value = "<%=ConvSPChars(strDeptCd)%>"
		.frm1.txtDeptNm.value = "<%=ConvSPChars(Trim(strGetInternalNm))%>"
				
	  End With 
    </Script>
   
   <%	
    End If

'거래처 정보 
    %>	

    <Script Language=vbScript>
	  With parent
		.frm1.txtBpCd.value = "<%=ConvSPChars(strBpCd)%>"
		.frm1.txtBpNm.value = "<%=ConvSPChars(Trim(strBpNM))%>"
				
	  End With 
    </Script>
   
   <%	



	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strAcctCd <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtAcctCd_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtAcctCd.value = "<%=ConvSPChars(strAcctCd)%>"
			.txtAcctNm.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
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
			.txtAcctCd2.value = "<%=ConvSPChars(strAcctCd2)%>"
			.txtAcctNm2.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
			End With
		</Script>
<%	
'		If IsNull(rs2(1)) = False then Bal_fg    = ConvSPChars(Trim(rs2(1)))
	End If
'===


'===	rs2.Close
	Set rs2 = Nothing
	
	If (rs6.EOF And rs6.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd.value = "<%=Trim(rs6(0))%>"
		.frm1.txtBizAreaNm.value = "<%=Trim(rs6(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs6.Close
	Set rs6 = Nothing
	
	
	If (rs7.EOF And rs7.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT1")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd1.value = "<%=Trim(rs7(0))%>"
		.frm1.txtBizAreaNm1.value = "<%=Trim(rs7(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs7.Close
	Set rs7 = Nothing
	
	
	TDrLocAmt = 0
	TCrLocAmt = 0
	NDrLocAmt = 0
	NCrLocAmt = 0 
	
	If strDeptCd = "" Then 
		If Not(rs3.EOF And rs3.BOF) Then
			If IsNull(rs3(0)) = False Then 
				TDrLocAmt    = rs3(0)
			Else
				TDrLocAmt    = 0
			End If
			If IsNull(rs3(1)) = False Then 
				TCrLocAmt    = rs3(1)
			Else
				TCrLocAmt    = 0
			End If
		End If
	Else
		TDrLocAmt = 0
		TCrLocAmt = 0
	End If 
	
	rs3.Close
	Set rs3 = Nothing
	

	If Not(rs4.EOF And rs4.BOF) Then
		If IsNull(rs4(0)) = False Then 
			NDrLocAmt    = rs4(0)
		Else
			NDrLocAmt    = 0
		End If
	End If
	
	rs4.Close
	Set rs4 = Nothing
	
	If Not(rs5.EOF And rs5.BOF) Then
		If IsNull(rs5(0)) = False Then 
			NCrLocAmt    = rs5(0)
		Else
			NCrLocAmt    = 0
		End If
	End If

	rs5.Close
	Set rs5 = Nothing
	
	
    TDrSumAmt = cdbl(TDrLocAmt) - cdbl(TCrLocAmt)
    NDrSumAmt = cdbl(NDrLocAmt) - cdbl(NCrLocAmt)

    SDrSumAmt = cdbl(TDrLocAmt) - cdbl(TCrLocAmt) + cdbl(NDrLocAmt) - cdbl(NCrLocAmt)
    TCrSumAmt = cdbl(TCrLocAmt) - cdbl(TDrLocAmt)
    NCrSumAmt = cdbl(NCrLocAmt) - cdbl(NDrLocAmt)

    SCrSumAmt = cdbl(TCrLocAmt) - cdbl(TDrLocAmt) + cdbl(NCrLocAmt) - cdbl(NDrLocAmt)
    SDrAmt	  = cdbl(TDrLocAmt) + cdbl(NDrLocAmt)
    SCrAmt    = cdbl(TCrLocAmt) + cdbl(NCrLocAmt)

    %>
    
    <Script Language=vbscript>
		With parent
    	.frm1.txtTDrAmt.text		= "<%=UNINumClientFormat(TDrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
		.frm1.txtTCrAmt.text		= "<%=UNINumClientFormat(TCrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
				
		.frm1.txtNDrAmt.text		= "<%=UNINumClientFormat(NDrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
		.frm1.txtNCrAmt.text		= "<%=UNINumClientFormat(NCrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
		
		If "<%=ConvSPChars(Bal_fg)%>" = "DR" Then
			.frm1.txtTSumAmt.text	= "<%=UNINumClientFormat(round(TDrSumAmt,0), ggAmtOfMoney.DecPoint, 0)%>"       

			.frm1.txtNSumAmt.text	= "<%=UNINumClientFormat(NDrSumAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
			.frm1.txtSSumAmt.text	= "<%=UNINumClientFormat(SDrSumAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
		Else
			.frm1.txtTSumAmt.text	= "<%=UNINumClientFormat(TCrSumAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
			.frm1.txtNSumAmt.text	= "<%=UNINumClientFormat(NCrSumAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
			.frm1.txtSSumAmt.text	= "<%=UNINumClientFormat(SCrSumAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
		End If
		.frm1.txtSDrAmt.text		= "<%=UNINumClientFormat(SDrAmt, ggAmtOfMoney.DecPoint, 0)%>"        	
		.frm1.txtSCrAmt.text		= "<%=UNINumClientFormat(SCrAmt, ggAmtOfMoney.DecPoint, 0)%>"        	

		End With
	</script>
	<%
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
	                                                  '☜: ActiveX Data Factory Object Nothing
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		'Response.End 
	
	End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

		strFromGlDt = Request("txtFromGlDt")
	strToGLDt	= Request("txtToGlDt")
	strDeptCd	= UCase(Trim(Request("txtDeptCd")))
	strAcctCd	= UCase(Trim(Request("txtAcctCd")))
	strAcctCd2	= UCase(Trim(Request("txtAcctCd2")))

	strBpCd		= UCase(Trim(Request("txtBpCd")))
	StrDesc		= Trim(Request("txtDesc"))
	StrProject	= Trim(Request("txtProject"))
	
	strZeroFg   = Trim(Request("ZeroFg"))
	strAmtFr = UNIConvNum(Request("txtAmtFr"),0)
	strAmtTo = UNIConvNum(Request("txtAmtTo"),0)
	
	strBizAreaCd  = Trim(UCase(Request("txtBizAreaCd")))            '사업장From
	strBizAreaCd1 = Trim(UCase(Request("txtBizAreaCd1")))            '사업장To
	
	'권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	 
	Call fnGetCompStDt
	gFiscStart = GetGlobalInf("gFiscStart")
		
    Call ExtractDateFrom(gFiscStart,gServerDateFormat,gServerDateType,Fiscyyyy,Fiscmm,Fiscdd)
    Call ExtractDateFrom(strFromGlDt,gServerDateFormat,gServerDateType,strGlDtYr,strGlDtMnth,strGlDtDt)	

	strWhere0 = ""
	strWhere0 = strWhere0 & " A.Acct_cd >= " & FilterVar(strAcctCd, "''", "S")
	strWhere0 = strWhere0 & " AND A.Acct_cd <= " & FilterVar(strAcctCd2, "''", "S")
	
	If strBpCd <> "" Then
		Call fnGetBpCd
		strWhere0 = strWhere0 & " AND A.BP_CD = " & FilterVar(strBpCd, "''", "S")
	End If
	
	If StrDesc <> "" Then
		strWhere0 = strWhere0 & " AND A.ITEM_DESC LIKE " & FilterVar("%" & StrDesc & "%" , "''", "S") & ""
	End If
	
	If StrProject <> "" Then
		strWhere0 = strWhere0 & " AND A.PROJECT_NO LIKE " & FilterVar("%" & StrProject & "%" , "''", "S") & ""
	End If
	
	If strGlInputType <> "" Then
		strWhere0 = strWhere0 & " AND B.GL_INPUT_TYPE =  " & FilterVar(strGlInputType, "''", "S") & " "
	End If

	If strRefNo <> "" Then
		strWhere0 = strWhere0 & " AND B.REF_NO LIKE  " & FilterVar(strRefNo & "%", "''", "S") & ""
	End If

	if strBizAreaCd <> "" then
		strWhere0 = strWhere0 & " AND B.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND B.BIZ_AREA_CD >= " & FilterVar("0", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere0 = strWhere0 & " AND B.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND B.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if
'-------------------------
'금액 
'-----------------------

	If strAmtFr <> 0 or strAmtTo <> 0 Then
		If strAmtFr > 0 and strAmtTo <= 0 Then
			strWhere0 = strWhere0 & " and a.ITEM_LOC_AMT >= " & strAmtFr 
		ElseIf strAmtFr <= 0 and strAmtTo > 0 Then
			strWhere0 = strWhere0 & " and a.ITEM_LOC_AMT <= " & strAmtTo 
		Else
			strWhere0 = strWhere0 & " and a.ITEM_LOC_AMT between " & strAmtFr & " and " & strAmtTo
		End If
	End If


	If strDeptCd <> "" Then
		Call fnGetInternalCd
	
		If 	strInternalCd = "" Then										' Internal Code가 없는 경우 
			strWhere0 = strWhere0 & " and A.Org_Change_Id = "			' :dept_cd, orgchangeid로 조회   
			
			'strWhere0 = strWhere0 & FilterVar(iChangeOrgId,"" & FilterVar("X", "''", "S") & " ","S")			
			strWhere0 = strWhere0 & FilterVar(iChangeOrgId, "''", "S")
			
			strWhere0 = strWhere0 & " and A.Dept_Cd = "
			'strWhere0 = strWhere0 & FilterVar(strDeptCd,"" & FilterVar("X", "''", "S") & " ","S")			
			strWhere0 = strWhere0 & FilterVar(strDeptCd, "''", "S")
		Else
			strWhere0 = strWhere0 & " and A.internal_cd = "
			'strWhere0 = strWhere0 & FilterVar(strInternalCd,"" & FilterVar("X", "''", "S") & " ","S")
			strWhere0 = strWhere0 & FilterVar(strInternalCd, "''", "S")
		End If
	End If

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
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
	
	' 권한관리 추가 
	strWhere0 = strWhere0 & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

	strWhere1 = strWhere0
	strWhere1 = strWhere0 & " AND B.gl_dt <  " & FilterVar(strFromGlDt, "''", "S") 

	strWhere0 = strWhere0 & " AND B.gl_dt between  " & FilterVar(strFromGlDt, "''", "S") & " AND  " & FilterVar(strToGLDt, "''", "S") & " "

End Sub
'--------------------------------------------
'Company(start_Dt)/ 이월금액 
'--------------------------------------------
Sub fnGetCompStDt()
    Dim iStr

    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,1)

    ON error resume next
 
    UNISqlId(0) = "A5124MA108"

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)


    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing

		strMsgCd = "970000"
		strMsg1 = Request("txtDeptCd_Alt")
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)

        strCompFiscStartDt = "1900-01-01"

    Else    
        strCompFiscStartDt   = Trim(rs0(0))

    End If
End Sub 

'--------------------------------------------
'내부부서코드 select
'--------------------------------------------
Sub fnGetInternalCd()
    Dim iStr

    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,1)

    ON error resume next
	Err.clear
 
    UNISqlId(0) = "A5124MA102"

    'UNIValue(0,0) = FilterVar(strDeptCd,"" & FilterVar("X", "''", "S") & " ","S")    
    'UNIValue(0,1) = FilterVar(iChangeOrgId,"" & FilterVar("X", "''", "S") & " ","S")	
    
    UNIValue(0,0) = FilterVar(strDeptCd, "''", "S")		
	UNIValue(0,1) = FilterVar(iChangeOrgId, "''", "S")	
	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing
	
        strInternalCd = ""
    Else    
        strInternalCd   = Trim(rs0(0))
        strGetInternalNm = Trim(rs0(1))
    End If
End Sub 
'--------------------------------------------
'거래처코드 select
'--------------------------------------------
Sub fnGetBpCd()
    Dim iStr

    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,1)

    ON error resume next
	Err.clear
 
    UNISqlId(0) = "A5124MA109"

    UNIValue(0,0) = FilterVar(strBpCd, "''", "S")	

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing

        strBpCd = ""
	    Call DisplayMsgBox("126100", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		%>	

		<Script Language=vbScript>
			With parent
				.frm1.txtBpCd.value = ""
				.frm1.txtBpNm.value = ""
			End With 
		</Script>

		<%	
		response.end

    Else    
        strBpCd		= Trim(rs0(0))
        strBpNm		= Trim(rs0(1))
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
