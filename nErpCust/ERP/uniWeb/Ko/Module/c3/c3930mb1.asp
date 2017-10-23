<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3605mB1
'*  4. Program Name         : 가공비제공내역조회 
'*  5. Program Desc         : 가공비제공내역조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2002/03/25
'*  8. Modified date(Last)  : 2002/03/25
'*  9. Modifier (First)     : Jang Yoon Ki
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜: 

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")       

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2	,rs3 ,rs4						'☜ : DBAgent Parameter 선언 
Dim lgstrData																'☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim txtYyyyMm
Dim txtPlantCd
Dim txtPlantNm
Dim txtParentItemAcctCd
Dim txtParentItemAcctNm
Dim txtChildItemAcctCd
Dim txtChildItemAcctNm
Dim txtBasSum
Dim txtIssueSum
Dim txtRcptSum
Dim txtBalSum
Dim SetFocusFlag
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

	Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
'   lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
 
	txtYyyyMm = Trim(Request("txtYyyyMm"))
	txtPlantCd = Trim(Request("txtPlantCd"))
	txtParentItemAcctCd = Trim(Request("txtParentItemAcctCd"))
	txtChildItemAcctCd = Trim(Request("txtChildItemAcctCd"))

	
    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    
   	Const C_SHEETMAXROWS_D  = 100 
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time        

    If UniConvNumStringToDouble(lgPageNo,0) > 0 Then
       rs0.Move     = UniConvNumStringToDouble(lgMaxCount,0) * UniConvNumStringToDouble(lgPageNo,0)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
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
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	Dim strWhere
    Redim UNIValue(4,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "C3930MA101"	'spread sheet
    UNISqlId(1) = "commonqry"	'name
    UNISqlId(2) = "C3930MA102"	'sum
    UNISqlId(3) = "commonqry"	'name
	UNISqlId(4) = "commonqry"	'name
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    strWhere = " and A.YYYYMM = " & FilterVar(txtYyyyMm ,"''"       ,"S")
    
    if txtPlantCd <> "" then
		strWhere = strWhere & " and A.prnt_plant_cd = " & FilterVar(txtPlantCd   , "''", "S")
	end if
	
    if txtParentItemAcctCd <> "" then
		strWhere = strWhere & " and a.prnt_item_acct = " & FilterVar(txtParentItemAcctCd   , "''", "S")
	end if

    if txtChildItemAcctCd <> "" then
		strWhere = strWhere & " and b.child_item_acct = " & FilterVar(txtChildItemAcctCd   , "''", "S")
	end if

	UNIValue(0,1)  = strWhere

	UNIValue(1,0) = "select plant_nm from b_plant Where plant_cd= " & FilterVar(txtPlantCd, "''", "S") & " "

	UNIValue(2,0)  = strWhere

	UNIValue(3,0) = "select minor_nm from b_minor Where major_cd=" & FilterVar("P1001", "''", "S") & "  and minor_cd =  " & FilterVar(txtChildItemAcctCd, "''", "S") & " "
    UNIValue(4,0) = "select minor_nm from b_minor Where major_cd=" & FilterVar("P1001", "''", "S") & "  and minor_cd =  " & FilterVar(txtParentItemAcctCd, "''", "S") & " "
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'--------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF
                                                                      '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)   
   
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If 
   
	If txtPlantCd <> "" Then
		If Not (rs1.EOF OR rs1.BOF) Then
			txtPlantNm = Trim(rs1("Plant_Nm"))
		Else
			
			'SetFocusFlag = 1
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)		'공장이 존재하지 않습니다 
			txtPlantNm = ""
			rs1.Close
			Set rs1 = Nothing 
			Exit Sub
		End IF
		rs1.Close
		Set rs1 = Nothing 
	End If		

	If txtChildItemAcctCd <> "" Then
		If Not (rs3.EOF OR rs3.BOF) Then
			txtChildItemAcctNm = Trim(rs3("minor_Nm"))
		Else
			
			'SetFocusFlag = 1
			Call DisplayMsgBox("236022", vbOKOnly, "", "", I_MKSCRIPT)		'품목계정이 유효하지 않습니다.
			txtChildItemAcctNm = ""
			rs3.Close
			Set rs3 = Nothing 
			Exit Sub
		End IF
		rs3.Close
		Set rs3 = Nothing 
	End If		

	If txtParentItemAcctCd <> "" Then
		If Not (rs4.EOF OR rs4.BOF) Then
			txtParentItemAcctNm = Trim(rs4("minor_Nm"))
		Else
			
			'SetFocusFlag = 1
			Call DisplayMsgBox("236022", vbOKOnly, "", "", I_MKSCRIPT)		'품목계정이 유효하지 않습니다.
			txtChildItemAcctNm = ""
			rs4.Close
			Set rs4 = Nothing 
			Exit Sub
		End IF
		rs4.Close
		Set rs4 = Nothing 
	End If		

    If  rs0.EOF And rs0.BOF Then
		'SetFocusFlag = 2
		Call DisplayMsgBox("236043", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()

	    'rs2
		If Not (rs2.EOF OR rs2.BOF) Then
			txtBasSum = rs2("TOT_PRNT_BAS_AMT")
			txtIssueSum = rs2("TOT_PRNT_ISSUE_AMT")
			txtRcptSum = rs2("TOT_PRNT_RCPT_AMT")
			txtBalSum = rs2("TOT_PRNT_BAL_AMT")
		Else
			txtBasSum = 0
			txtIssueSum = 0
			txtRcptSum = 0
			txtBalSum = 0
		End IF
		rs2.Close
		Set rs2 = Nothing
    End If
    
End Sub

%>

<Script Language=vbscript>

With Parent
	.frm1.txtPlantNm.value				= "<%=ConvSPChars(txtPlantNm)%>"
	.frm1.txtParentItemAcctNm.value				= "<%=ConvSPChars(txtParentItemAcctNm)%>"
	.frm1.txtChildItemAcctNm.value				= "<%=ConvSPChars(txtChildItemAcctNm)%>"
	
	If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.Frm1.hYyyyMm.Value	  = "<%=txtYyyyMm%>"                'For Next Search
			.Frm1.hPlantCd.Value	  = "<%=txtPlantCd%>"                'For Next Search
			.Frm1.hParentItemAcctCd.Value	  = "<%=txtParentItemAcctCd%>"                'For Next Search			
			.Frm1.hChildItemAcctCd.Value	  = "<%=txtChildItemAcctCd%>"                'For Next Search			
		End If

						'rs1 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 
		.frm1.txtBasSum.text				= "<%=UNINumClientFormat(txtBasSum,ggAmtOfMoney.Decpoint,0)%>"				'기초금액 합계 
		
		.frm1.txtIssueSum.text				= "<%=UNINumClientFormat(txtIssueSum,ggAmtOfMoney.Decpoint,0)%>"			'출고금액 합계 
		
		.frm1.txtRcptSum.text				= "<%=UNINumClientFormat(txtRcptSum,ggAmtOfMoney.Decpoint,0)%>"				'입고금액 합계 
		
		.frm1.txtBalSum.text				= "<%=UNINumClientFormat(txtBalSum,ggAmtOfMoney.Decpoint,0)%>"				'기말금액 합계 
		
         
       'Show multi spreadsheet data from this line			
		
		.ggoSpread.Source  = Parent.frm1.vspdData
		.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		.DbQueryOk
		
    End If
    
End With
</Script>
