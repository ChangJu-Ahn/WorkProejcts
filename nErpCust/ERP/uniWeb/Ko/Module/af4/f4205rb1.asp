<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Accounting - Treasury
'*  2. Function Name        : Loan
'*  3. Program ID           : f4205rb1
'*  4. Program Name         : 차입금번호팝업 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001.02.19
'*  7. Modified date(Last)  : 2003.04.29
'*  8. Modifier (First)     : Song, Mun Gil
'*  9. Modifier (Last)      : Oh, Soo Min
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","QB")

Const C_SHEETMAXROWS_D = 30
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3                             '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgStrPrevKey
Dim lgTailList                                                              '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strCond
Dim strLoanFrDt, strLoanToDt
Dim strDueFrDt, strDueToDt
Dim strBpCd, strLoanType
Dim strPgmId
Dim strDocCur
Dim strLoanfg   
Dim strLoanNo
Dim strMsgCd, strMsg1, strMsg2

Dim  iLoopCount
Dim  LngMaxRow

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
' 권한관리 추가 
Dim lgAuthBizAreaCd	' 사업장 
Dim lgInternalCd	' 내부부서 
Dim lgSubInternalCd	' 내부부서(하위포함)
Dim lgAuthUsrID		' 개인 

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

  
    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    lgStrPrevKey   = Request("lgStrPrevKey")
    lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(lgMaxCount) * CDbl(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
     
    'rs0에 대한 결과 
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
                    '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(2,2)

    UNISqlId(0) = "F4205ra101"
    UNISqlId(1) = "ABPNM"
    UNISqlId(2) = "AMINORNM"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    UNIValue(1,0) = Filtervar(UCase(Trim(strBpCd))	, "''", "S")
    UNIValue(2,0) = Filtervar("F1000", "''", "S")
    UNIValue(2,1) = Filtervar(strLoanType	, "", "S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim lgADF
    Dim lgstrRetMsg 
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	If strBpCd <> "" Then
			If Not (rs1.EOF OR rs1.BOF) Then
	%>
			<Script Language=vbScript>
				With parent
					.frm1.txtBpNm.value = "<%=Trim(rs1(1))%>"
				End With
			</Script>
	<%		
			Else
	%>
			<Script Language=vbScript>
				With parent
					.frm1.txtBpNm.value = ""
				End With
			</Script>
	<%		
				Call DisplayMsgBox("970000", vbOKOnly, Request("txtBpCd_Alt"), "", I_MKSCRIPT)		'No Data Found!!
		        rs1.Close
			    Set rs1 = Nothing
				Exit sub
			End IF
		rs1.Close
		Set rs1 = Nothing
		End If
	
	
	If rs2.EOF And rs2.BOF Then
		If strMsgCd = "" And strLoanType <> "" Then
			strMsgCd = "970000"
			strMsg1 = Request("txtLoanType_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtLoanType.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
			.txtLoanTypeNm.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
		End With
		</Script>
<%
	End If
	
    If  rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
    Else    
        Call  MakeSpreadSheetData()
    End If

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
  
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strLoanFrDt  = UniConvDate(Request("txtLoanFromDt"))
    strLoanToDt  = UniConvDate(Request("txtLoanToDt"))
    strDocCur	 = UCase(Request("txtDocCur"))
    strBpCd      = UCase(Request("txtBpCd"))
    strDueFrDt   = Request("txtDueFromDt")
    strDueToDt   = Request("txtDueToDt")
    strLoanfg	 = Request("cboLoanFg")
    strLoanType  = Request("txtLoanType")
    strLoanNo	 = Request("txtLoanNo")
    strPgmId	 = Request("txtPgmId")
    
	strCond = ""
	If  strPgmId = "F4205MA1" Then												'거래처차입금 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LN" , "''", "S") & " "												
	Elseif strPgmId = "F4206MA1" Then											'거래처기초차입금 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LT" , "''", "S") & " "		
														
	Elseif strPgmId = "F4223MA1" Then											'차입금상환계획변경 
		
	Elseif strPgmId = "F4231MA1" Then											'이자율변경등록 
'		strCond = strCond & " and (A.loan_fg = '" & "LL" & "' "
'		strCond = strCond & " or   A.loan_fg = '" & "LN" & "' ) "
		strCond = strCond & " and A.int_votl =  " & FilterVar("F" , "''", "S") & " "
'		strCond = strCond & " and A.rdp_cls_fg = '" & "N" & "' "				'상환완료된 건은 display만가능 
	Elseif strPgmId = "F4235MA1" Then											'은행차입금만기연장 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LR" , "''", "S") & " "		
	End If
	
		strCond = strCond & " and A.loan_plc_type =  " & FilterVar("BP" , "''", "S") & " "		
		strCond = strCond & " and A.bp_cd <>  " & FilterVar("", "''", "S") & " "
		strCond = strCond & " and A.bp_cd = E.bp_cd "		

	If strLoanFrDt <> "" Then strCond = strCond & " and A.loan_dt >=  " & FilterVar(strLoanFrDt , "''", "S") & " "			'차입일 
	If strLoanToDt <> "" Then strCond = strCond & " and A.loan_dt <=  " & FilterVar(strLoanToDt , "''", "S") & " "
	If strDocCur   <> "" Then strCond = strCond & " and A.doc_cur = " & Filtervar(strDocCur	, "''", "S")				'거래통화 
	If strDueFrDt  <> "" Then strCond = strCond & " and A.due_dt >=  " & FilterVar(UniConvDate(strDueFrDt), "''", "S") & " "				'만기일 
	If strDueToDt  <> "" Then strCond = strCond & " and A.due_dt <=  " & FilterVar(UniConvDate(strDueToDt), "''", "S") & " "
	If strLoanNo   <> "" Then strCond = strCond & " and A.loan_no = " & Filtervar(strLoanNo	, "''", "S")				'차입번호 
	If strLoanfg   <> "" Then strCond = strCond & " and A.loan_fg =  " & FilterVar(strLoanfg , "''", "S") & " "				'장단기구분 
	If strLoanType <> "" Then strCond = strCond & " and A.loan_type = " & Filtervar(strLoanType	, "''", "S")			'차입용도 
	If strBpCd     <> "" Then strCond = strCond & " and A.bp_cd = " & Filtervar(strBpCd	, "''", "S")				'차입거래처	

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If    

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub
'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------

%>

<Script Language=vbscript>
With parent
	If "<%=lgDataExist%>" = "Yes" Then
        .ggoSpread.Source    = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",3),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
         
'         With .frm1
'			.hLoanFromDt.value = strLoanFrDt
'			.hLoanToDt.value   = strLoanToDt
'			.hDueFromDt.value  = strDueFrDt
'			.hDueToDt.value    = strDueToDt
'			.hBankLoanCd.value = strBpCd
'			.hLoanType.value   = strLoanType
 '        End With
         
	End If
	.DbQueryOk()
End with
</Script>	

<%
	Response.End 
%>

