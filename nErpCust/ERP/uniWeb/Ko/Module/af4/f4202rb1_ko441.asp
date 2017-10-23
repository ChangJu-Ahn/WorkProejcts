<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Accounting - Treasury
'*  2. Function Name        : Loan
'*  3. Program ID           : f4202rb1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001.02.19
'*  7. Modified date(Last)  : 2001.11.10
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

'On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")

Const C_SHEETMAXROWS_D = 30
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3                             '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                              '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strCond
Dim strLoanFrDt, strLoanToDt
Dim strDueFrDt, strDueToDt
Dim strBankCd, strLoanType
Dim strPgmId
Dim strDocCur
Dim strLoanfg   
Dim strLoanNo
Dim strMsgCd, strMsg1, strMsg2

Dim  iLoopCount
Dim  LngMaxRow

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
	Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	    
	lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgDataExist    = "No"
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1
	    
	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

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

    UNISqlId(0) = "F4202ra101"
    UNISqlId(1) = "ABANKNM"
    UNISqlId(2) = "AMINORNM"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    UNIValue(1,0) = Filtervar(strBankCd	, "", "S")
    UNIValue(2,0) = Filtervar("F1000", "", "S")
    UNIValue(2,1) = Filtervar(strLoanType	, "", "S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    


	If rs1.EOF And rs1.BOF Then
		If strMsgCd = "" And strBankCd <> "" Then
			strMsgCd = "970000"
			strMsg1 = Request("txtBankLoanCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBankLoanCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			.txtBankLoanNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
		End With
		</Script>
<%
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
'		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Set lgADF = Nothing
'		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If

	'rs0.Close
	'Set rs0 = Nothing 
	'Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing

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
    strBankCd    = UCase(Request("txtBankLoanCd"))
    strDueFrDt   = Request("txtDueFromDt")
    strDueToDt   = Request("txtDueToDt")
    strLoanfg	 = Request("cboLoanFg")
    strLoanType  = Request("txtLoanType")
    strLoanNo	 = Request("txtLoanNo")
    strPgmId	 = Request("txtPgmId")
    
  	strCond = ""
	If  strPgmId = "F4201MA1" Then												'차입금 TOTAL
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LN" , "''", "S") & " "		
		strCond = strCond & " and A.loan_bank_cd <>  " & FilterVar("", "''", "S") & " "
		strCond = strCond & " and A.loan_bank_cd = E.bank_cd "									
	Elseif strPgmId = "F4204MA1" Then											'은행기초차입금 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LT" , "''", "S") & " "		
'		strCond = strCond & " and (A.loan_bank_cd = '" & "" & "' "
'		strCond = strCond & " or   A.loan_bank_cd  is null ) "												
'	Elseif strPgmId = "F4223MA1" Then											'차입금상환계획변경 
	Elseif strPgmId = "F4207MA1" Then											'차입금요청 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LQ" , "''", "S") & " "		
		
	Elseif strPgmId = "F4231MA1" Then											'이자율변경등록 
		strCond = strCond & " and A.int_votl =  " & FilterVar("F" , "''", "S") & " "
'		strCond = strCond & " and A.rdp_cls_fg = '" & "N" & "' "				'상환완료된 건은 display만가능 
	Elseif strPgmId = "F4234MA1" Then											'은행차입금만기연장 
		strCond = strCond & " and A.loan_basic_fg =  " & FilterVar("LR" , "''", "S") & " "		
	End If
	
		strCond = strCond & " and A.loan_plc_type =  " & FilterVar("BK" , "''", "S") & " "
		strCond = strCond & " and A.loan_bank_cd = E.bank_cd "				
		
	If strLoanFrDt <> "" Then strCond = strCond & " and A.loan_dt >=  " & FilterVar(strLoanFrDt , "''", "S") & " "			'차입일 
	If strLoanToDt <> "" Then strCond = strCond & " and A.loan_dt <=  " & FilterVar(strLoanToDt , "''", "S") & " "
	If strDocCur   <> "" Then strCond = strCond & " and A.doc_cur = " & Filtervar(strDocCur	, "''", "S")				'거래통화 
	If strDueFrDt  <> "" Then strCond = strCond & " and A.due_dt >=  " & FilterVar(UniConvDate(strDueFrDt), "''", "S") & " "				'만기일 
	If strDueToDt  <> "" Then strCond = strCond & " and A.due_dt <=  " & FilterVar(UniConvDate(strDueToDt), "''", "S") & " "
	If strLoanNo   <> "" Then strCond = strCond & " and A.loan_no = " & Filtervar(strLoanNo	, "''", "S")				'차입번호 
	If strLoanfg   <> "" Then strCond = strCond & " and A.loan_fg =  " & FilterVar(strLoanfg , "''", "S") & " "				'장단기구분 
	If strLoanType <> "" Then strCond = strCond & " and A.loan_type = " & Filtervar(strLoanType	, "''", "S")			'차입용도 
	If strBankCd   <> "" Then strCond = strCond & " and A.loan_bank_cd = " & Filtervar(strBankCd	, "''", "S")
	
	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
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

	strCond	= strCond	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	
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
		.frm1.vspdData.Redraw = True

'         With .frm1
'			.hLoanFromDt.value = strLoanFrDt
'			.hLoanToDt.value   = strLoanToDt
'			.hDueFromDt.value  = strDueFrDt
'			.hDueToDt.value    = strDueToDt
'			.hBankLoanCd.value = strBankCd
'			.hLoanType.value   = strLoanType
 '        End With
         
	End If         
	.DbQueryOk()
End with
</Script>	

<%
	Response.End 
%>

