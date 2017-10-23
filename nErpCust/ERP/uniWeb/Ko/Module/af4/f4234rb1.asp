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

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","QB")

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3                             '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                              '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Const C_SHEETMAXROWS_D = 30

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strCond1
Dim strCond2
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

    UNISqlId(0) = "F4234ra101"
    UNISqlId(1) = "ABANKNM"
    UNISqlId(2) = "AMINORNM"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strCond1
    UNIValue(1,0) = Filtervar(strBankCd	, "", "S")
    UNIValue(2,0) = Filtervar("F1000"	, "", "S")
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
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
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
		rs0.Close
		Set rs0 = Nothing
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
    strDueFrDt   = Request("txtDueFromDt")
    strDueToDt   = Request("txtDueToDt")
    strBankCd    = UCase(Request("txtBankLoanCd"))
    strLoanfg	 = Request("cboLoanFg")
    strLoanType  = Request("txtLoanType")
    strLoanNo	 = Request("txtLoanNo")   
    strPgmId	 = Request("txtPgmId")
    
	strCond1 = ""
'	strCond1 = strCond1 & " A.loan_basic_fg <> '" & "LR" & "' "				'Not Rollover차입금 
	strCond1 = strCond1 & " 	(A.cls_ro_fg = " & FilterVar("N", "''", "S") & "  or A.cls_ro_fg = '') "					'Not RollOver차입금 
	strCond1 = strCond1 & " AND (A.conf_fg =  " & FilterVar("C" , "''", "S") & " "					'Confirm 여부 check
	strCond1 = strCond1 & " OR   A.conf_fg =  " & FilterVar("E" , "''", "S") & " )"					'Confirm 여부 check
	strCond1 = strCond1 & " AND (A.rdp_cls_fg = " & FilterVar("N", "''", "S") & "  or A.rdp_cls_fg = '') "					'Not 상환완료된 차입금 
		
	If  strPgmId = "F4234MA1" Then												
		strCond1 = strCond1 & " AND A.loan_plc_type =  " & FilterVar("BK" , "''", "S") & " "			'은행차입금		
	Elseif strPgmId = "F4235MA1"  Then										
		strCond1 = strCond1 & " AND A.loan_plc_type =  " & FilterVar("BP" , "''", "S") & " "			'거래처차입금		
	End If
	
	If strLoanFrDt <> "" Then strCond1 = strCond1 & " and A.loan_dt >=  " & FilterVar(strLoanFrDt , "''", "S") & " "
	If strLoanToDt <> "" Then strCond1 = strCond1 & " and A.loan_dt <=  " & FilterVar(strLoanToDt , "''", "S") & " "
	If strDocCur   <> "" Then strCond1 = strCond1 & " and A.doc_cur = " & Filtervar(strDocCur	, "''", "S")				'거래통화	
	If strDueFrDt  <> "" Then strCond1 = strCond1 & " and A.due_dt >=  " & FilterVar(UNIConvDate(strDueFrDt), "''", "S") & " "
	If strDueToDt  <> "" Then strCond1 = strCond1 & " and A.due_dt <=  " & FilterVar(UNIConvDate(strDueToDt), "''", "S") & " "
	
	If  strPgmId = "F4234MA1" Then												
		If strBankCd   <> "" Then strCond1 = strCond1 & " and A.loan_bank_cd = " & Filtervar(strBankCd	, "''", "S")	'은행차입금		
	Elseif strPgmId = "F4235MA1"  Then										
		If strBankCd   <> "" Then strCond1 = strCond1 & " and A.bp_cd = " & Filtervar(strBankCd	, "''", "S")			'거래처차입금		
	End If
	
	If strLoanfg   <> "" Then strCond1 = strCond1 & " and A.loan_fg =  " & FilterVar(strLoanfg , "''", "S") & " "				'장단기구분 
	If strLoanType <> "" Then strCond1 = strCond1 & " and A.loan_type = " & Filtervar(strLoanType	, "''", "S")	
	If strLoanNo   <> "" Then strCond1 = strCond1 & " and A.loan_no = " & Filtervar(strLoanNo	, "''", "S")				'차입번호 

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		strCond1		= strCond1 & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond1		= strCond1 & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond1		= strCond1 & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond1		= strCond1 & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
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
			Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",9),parent.GetKeyPos("A",10),   "A" ,"I","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",9),parent.GetKeyPos("A",12),   "A" ,"I","X","X")
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
