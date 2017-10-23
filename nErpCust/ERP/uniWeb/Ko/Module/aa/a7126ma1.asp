
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Capital Expense
'*  3. Program ID           : a7126ma1
'*  4. Program Name         : 자본적 지출내역 등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : AS1011,  
'							  AS1018	
'							  +B19029LookupNumericFormatF	
'*  7. Modified date(First) : 2002/11/01
'*  8. Modified date(Last)  : 2002/11/01
'*  9. Modifier (First)     : Seo Hyo Seok
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'=======================================================================================================
'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit    							'☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const gIsShowLocal = "Y"					'자국금액을 쓰는가에 대한 변수 
<%
Const gIsShowLocal = "Y"
%>

'@PGM_ID
Const BIZ_PGM_ID  = "a7126mb1.asp"  							'비지니스 로직 ASP명 

Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"			'환율정보 비지니스 로직 ASP명 

'@Grid_Column
Dim C_ChgNo
Dim C_ChgDesc
Dim C_BizPartnerCd
Dim C_BizPartnerPopup
Dim C_BizPartnerNm
Dim C_ChgAmt
Dim C_ChgLocAmt
Dim C_TaxTypeCd
Dim C_TaxTypePopup
Dim C_TaxTypeNm
Dim C_TaxRate
Dim C_TaxAmt
Dim C_TaxLocAmt
Dim C_ReportBizAreaCd
Dim C_ReportBizAreaPopup
Dim C_ReportBizAreaNm
Dim C_IssuedDt
Dim C_PayTypeCd
Dim C_PayTypeNm
Dim C_PayTypeDesc
Dim C_PayTypePopup

'미지급금추가 
Dim C_ApAcctCd
Dim C_ApAcctPopup
Dim C_ApAcctNm

Dim C_ApDueDt


Const C_SHEETMAXROWS = 30					'한 화면에 보여지는 최대갯수 


Dim IsOpenPop								'Popup

Dim lgMasterQueryFg							'자산Master의 query 여부 

'======================================================================================================
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'=======================================================================================================

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE				'Indicates that current mode is Create mode
 
    lgIntGrpCount = 0						'initializes Group View Size

    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""						'initializes Previous Key

	lgBlnFlgChgValue = False				'Indicates that no value changed

End Sub

Sub initSpreadPosVariables()
	'@Grid_Column
	C_ChgNo					=  1		'변동번호 
	C_ChgDesc				=  2		'지출내역 
	C_BizPartnerCd			=  3		'거래처코드 
	C_BizPartnerPopup		=  4		'거래처팝업 
	C_BizPartnerNm			=  5		'거래처이름 
	C_ChgAmt				=  6		'지출금액 
	C_ChgLocAmt				=  7		'지출금액(자국)
	C_TaxTypeCd				=  8		'부가세유형 
	C_TaxTypePopup			=  9		'부가세유형팝업 
	C_TaxTypeNm				= 10		'부가세이름 
	C_TaxRate				= 11		'부가세율 
	C_TaxAmt				= 12		'부가세금액 
	C_TaxLocAmt				= 13		'부가세금액(자국)
	C_ReportBizAreaCd		= 14		'세금신고사업장코드 
	C_ReportBizAreaPopup	= 15		'세금신고사업장팝업 
	C_ReportBizAreaNm		= 16		'세금신고사업장명 
	C_IssuedDt				= 17		'계산서발행일 
	C_PayTypeCd				= 18		'지출유형코드 
	C_PayTypeNm				= 19		'지출유형명 
	C_PayTypeDesc			= 20		'지출유형상세내역 
	C_PayTypePopup			= 21		'지출유형상세내역팝업 

	'미지급금추가 
	C_ApAcctCd				= 22		'미지급금계정코드 
	C_ApAcctPopup			= 23		'미지급금계정코드팝업 
	C_ApAcctNm				= 24		'미지급금계정코드이름 
	C_ApDueDt				= 25		'미지급금만기일자 
End Sub
'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()

	<%
	Dim svrDate
	svrDate = GetSvrDate
	%>

	frm1.txtChgDt.text		= UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,parent.gDateFormat)
	
	
'	frm1.txtDueDt.text		= UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,parent.gDateFormat)	
'	frm1.txtIssuedDt.text	= UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,parent.gDateFormat)	

	frm1.txtDocCur.value	= parent.gCurrency
	frm1.txtXchRate.text	= 1
	frm1.hOrgChangeId.value	 = parent.gChangeOrgID

	lgBlnFlgChgValue = False

End Sub


'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub


'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    

	With frm1.vspdData
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.Spreadinit "V20030218",,parent.gAllowDragDropSpread    

		.MaxCols = C_ApDueDt + 1                               '☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		
		'Hidden Column 설정 
    	.Col = .MaxCols											'공통콘트롤 사용 Hidden Column
    	.ColHidden = True

		.ReDraw = false	
		
        Call GetSpreadColumnPos("A")
		
		Call AppendNumberPlace("6","3","0")

   		'Col, Header, ColWidth, HAlign, FloatMax, FloatMin, FloatSeparator, FloatSepChar, FloatDecimalPlaces, FloatDeciamlChar
		
		ggoSpread.SSSetEdit		C_ChgNo,				"변동번호",			20,,,18
		ggoSpread.SSSetEdit		C_ChgDesc,				"지출내역",			10,,,30
		ggoSpread.SSSetEdit		C_BizPartnerCd,			"거래처코드",		12,,,10
		ggoSpread.SSSetButton	C_BizPartnerPopup
		ggoSpread.SSSetEdit		C_BizPartnerNm,			"거래처이름",		12,,,40
		ggoSpread.SSSetFloat	C_ChgAmt,				"총지출금액",		15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_ChgLocAmt,			"총지출금액(자국)",	15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_TaxTypeCd,			"부가세유형",		12,,,10
		ggoSpread.SSSetButton	C_TaxTypePopup
		ggoSpread.SSSetEdit		C_TaxTypeNm,			"부가세이름",		12,,,40
		ggoSpread.SSSetFloat	C_TaxRate,				"부가세율",			10, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z","0","100"
		ggoSpread.SSSetFloat	C_TaxAmt,				"부가세금액",		15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_TaxLocAmt,			"부가세금액(자국)",	15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_ReportBizAreaCd,		"신고사업장코드",	14,,,10
		ggoSpread.SSSetButton	C_ReportBizAreaPopup
		ggoSpread.SSSetEdit		C_ReportBizAreaNm,		"신고사업장명",		14,,,40
		ggoSpread.SSSetDate		C_IssuedDt,				"계산서발행일",		12,2,parent.gDateFormat  
		ggoSpread.SSSetCombo	C_PayTypeCd,			"지출유형코드",		15, 2, true
		ggoSpread.SSSetCombo	C_PayTypeNm,			"지출유형명",		18, 2, false
		ggoSpread.SSSetEdit		C_PayTypeDesc,			"지출유형상세내역",	16,,,40
		ggoSpread.SSSetButton	C_PayTypePopup
		
		'미지급금계정추가 
		ggoSpread.SSSetEdit		C_ApAcctCd,				"미지급금계정코드",		15,,,10
		ggoSpread.SSSetButton	C_ApAcctPopup
		ggoSpread.SSSetEdit		C_ApAcctNm,				"미지급금계정명",		13,,,40
		
		ggoSpread.SSSetDate		C_ApDueDt,				"미지급금만기일자",	16,2,parent.gDateFormat  

			Call ggoSpread.SSSetColHidden(C_PayTypeCd,C_PayTypeCd,True)

		If gIsShowLocal = "N" Then
			Call ggoSpread.SSSetColHidden(C_ChgLocAmt,C_ChgLocAmt,True)
			Call ggoSpread.SSSetColHidden(C_TaxLocAmt,C_TaxLocAmt,True)
		End If

		.ReDraw = true
		
		Call SetSpreadLock 
		
	End With
    
End Sub


'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
Dim RowCnt
    With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SpreadLock	C_ChgNo,			-1,	C_ChgNo
		ggoSpread.SSSetRequired	C_ChgDesc,			-1,	C_ChgDesc
		ggoSpread.SSSetRequired	C_BizPartnerCd,		-1,	C_BizPartnerCd
		ggoSpread.SpreadLock	C_BizPartnerNm,		-1,	C_BizPartnerNm
		ggoSpread.SSSetRequired	C_ChgAmt,			-1,	C_ChgAmt
		ggoSpread.SpreadLock	C_TaxTypeNm,		-1,	C_TaxTypeNm
		ggoSpread.SpreadLock	C_ReportBizAreaNm,	-1,	C_ReportBizAreaNm
		ggoSpread.SSSetRequired	C_PayTypeNm,		-1,	C_PayTypeNm
		
		'미지급금계정추가 
		ggoSpread.SpreadLock	C_ApAcctNm,		-1,	C_ApAcctNm
		
		For RowCnt = 1 To .vspdData.MaxRows								
			
			.vspdData.Col = C_PayTypeCd
			.vspdData.Row = RowCnt										
			
			If .vspdData.text = "AP" Then					
				
				ggoSpread.SpreadLock		C_PayTypeDesc ,     RowCnt,  RowCnt
				ggoSpread.SpreadLock		C_PayTypePopup ,     RowCnt,  RowCnt		
				
				'미지급금계정추가 
				ggoSpread.SpreadUnLock      C_ApAcctCd ,    RowCnt,  RowCnt
				ggoSpread.SSSetRequired     C_ApAcctCd ,    RowCnt,  RowCnt
				ggoSpread.SpreadUnLock      C_ApAcctPopup,    RowCnt,  RowCnt

				ggoSpread.SpreadUnLock      C_ApDueDt ,    RowCnt,  RowCnt
				ggoSpread.SSSetRequired     C_ApDueDt ,    RowCnt,  RowCnt
						
				
			ElseIf  .vspdData.text = "DP" or .vspdData.text = "NP" or .vspdData.text = "CP" or .vspdData.text = "NE"Then						
						
				ggoSpread.SpreadUnLock      C_PayTypeDesc ,    RowCnt,  RowCnt
				ggoSpread.SSSetRequired     C_PayTypeDesc ,    RowCnt,  RowCnt
				ggoSpread.SpreadUnLock      C_PayTypePopup ,    RowCnt,  RowCnt
				ggoSpread.SSSetRequired     C_PayTypePopup ,    RowCnt,  RowCnt				
				
				'미지급금계정추가				
				ggoSpread.SpreadLock		C_ApAcctCd ,     RowCnt,  RowCnt
				ggoSpread.SpreadLock		C_ApAcctPopup ,     RowCnt,  RowCnt
				
				ggoSpread.SpreadLock		C_ApDueDt ,     RowCnt,  RowCnt		
				
			Else
		
				ggoSpread.SpreadLock		C_PayTypeDesc ,     RowCnt,  RowCnt
				ggoSpread.SpreadLock		C_PayTypePopup ,	RowCnt,  RowCnt
				
				'미지급금계정추가				
				ggoSpread.SpreadLock		C_ApAcctCd ,     RowCnt,  RowCnt
				ggoSpread.SpreadLock		C_ApAcctPopup ,     RowCnt,  RowCnt
				
				ggoSpread.SpreadLock		C_ApDueDt ,			RowCnt,  RowCnt					
				
			End If 			
		
	Next     
		
		.vspdData.ReDraw = True
	End With

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal lRow,ByVal lRow2)
   	With frm1
		.vspdData.ReDraw = False
			
		ggoSpread.SSSetProtected	C_ChgNo,			lRow,	lRow2
		ggoSpread.SSSetRequired		C_ChgDesc,			lRow,	lRow2
		ggoSpread.SSSetRequired		C_BizPartnerCd,		lRow,	lRow2
		ggoSpread.SSSetProtected	C_BizPartnerNm,		lRow,	lRow2
		ggoSpread.SSSetRequired		C_ChgAmt,			lRow,	lRow2
		ggoSpread.SSSetProtected	C_TaxTypeNm,		lRow,	lRow2
		ggoSpread.SSSetProtected	C_ReportBizAreaNm,	lRow,	lRow2
		ggoSpread.SSSetRequired		C_PayTypeNm,		lRow,	lRow2
		ggoSpread.SSSetProtected	C_ApAcctNm,		lRow,	lRow2
			
		.vspdData.ReDraw = True
	End With
End Sub

'==========================================  OpenPoRef()  =================================================
'	Name : OpenMasterRef()
'	Description : 조건부 팝업 
'==========================================================================================================
Function OpenMasterRef()

	Dim strRet
	Dim arrParam(7)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	strRet = window.showModalDialog("a7103ra1.asp", Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet(0) = "" Then
		frm1.txtAsstNo.focus
		Exit Function
	Else
		Call SetMasterRef(strRet)
	End If	
		
End Function

'=========================================  SetPoRef()  ==================================================
'	Name : SetMasterRef()
'	Description : 조건부 팝업 
'========================================================================================================= 
Sub SetMasterRef(strRet)
    
	frm1.txtAsstNo.focus
	frm1.txtAsstNo.value     = strRet(0)
	frm1.txtAsstNm.value	 = strRet(3)
	
End Sub

'======================================================================================================
'   Function Name : OpenCapExpNo()
'   Function Desc : 
'=======================================================================================================
Function OpenCapExpNo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	Dim strAsstNo
	Dim IntRetCd

	If IsOpenPop = True Then Exit Function	

	strAsstNo  = Trim(frm1.txtAsstNo.value)
	
	If strAsstNo = "" then
		IntRetCD = DisplayMsgBox("117326","X","X","X")    '자산번호를 입력하십시오.
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = "자본적지출번호팝업"	
	arrParam(1) = "A_ASSET_CHG"				
	arrParam(2) = Trim(frm1.txtCapExpNo.Value)
	arrParam(3) = ""
	arrParam(4) = "ASST_CD = " & FilterVar(strAsstNo, "''", "S") & _
				" AND CHG_FG = " & FilterVar("01", "''", "S") & " "
	arrParam(5) = "자본적지출번호"			
	
	arrField(0) = "CAP_EXP_NO"
	arrField(1) = "ASSET_CHG_DESC"
    
	arrHeader(0) = "자본적지출번호"
	arrHeader(1) = "자본적지출내역"
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		frm1.txtCapExpNo.focus
		Exit Function
	Else
		Call SetCapExpNo(arrRet)
	End If
	
End Function

'=======================================================================================================
'   Function Name : SetCapExpNo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetCapExpNo(Byval arrRet)

	frm1.txtCapExpNo.focus
	frm1.txtCapExpNo.value  = arrRet(0)

End Function

'======================================================================================================
'   Function Name : OpenCapExpNo()
'   Function Desc : 
'=======================================================================================================
Function OpenApAcctCd(byVal IRow, byVal strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	Dim IntRetCd

	If IsOpenPop = True Then Exit Function	


	arrParam(0) = "미지급금계정 팝업"	
	arrParam(1) = "a_jnl_acct_assn a, a_acct b"
	arrParam(2) = Trim(strCode)
	arrParam(3) = ""
	arrParam(4) = "A.trans_type = " & FilterVar("AS002", "''", "S") & "  and A.Acct_cd = B.Acct_cd and Jnl_cd = " & FilterVar("AP", "''", "S") & " "
	arrParam(5) = "미지급금계정 코드"
	
    arrField(0) = "a.acct_cd"	
    arrField(1) = "b.acct_nm"
    
    arrHeader(0) = "미지급금계정 코드"		
    arrHeader(1) = "미지급금계정명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetApAcctCd(IRow, arrRet)
	End If	
	
End Function

'=======================================================================================================
'   Function Name : SetCapExpNo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetApAcctCd(byVal IRow, byVal arrRet)
	With frm1.vspdData
		.Row = iRow

		.Col = C_ApAcctCd
		.Text = arrRet(0)
				
		.Col = C_ApAcctNm
		.Text = arrRet(1)

	End With
   	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow iRow

End Function


'=======================================================================================================
'	Name : OpenNoteNo()
'	Description : Note No PopUp
'=======================================================================================================
Function OpenNoteNo(byVal IRow, byVal strCode, byVal strCard)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True	

 
	Select Case UCase(strCard)
	
	Case "DP"	
		arrParam(0) = "예적금코드 팝업"	' 팝업 명칭 
		arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"			' TABLE 명칭 
		arrParam(2) = strCode						' Code Condition
		arrParam(3) = ""							' Name Cindition
		arrParam(4) = "A.BANK_CD = B.BANK_CD "												' Where Condition'			
		arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "	
		arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "
		arrParam(4) = arrParam(4) & "AND C.DPST_FG IN (" & FilterVar("SV", "''", "S") & " ," & FilterVar("ET", "''", "S") & " ) "
		arrParam(4) = arrParam(4) & "AND C.DPST_TYPE IN (" & FilterVar("D1", "''", "S") & " ," & FilterVar("D2", "''", "S") & " ," & FilterVar("D3", "''", "S") & " ) "								
		arrParam(5) = "예적금코드"				' 조건필드의 라벨 명칭 
				
		arrField(0) = "B.BANK_ACCT_NO"				' Field명(2)
   		arrField(1) = "A.BANK_CD"					' Field명(1)
		arrField(2) = "A.BANK_NM"					' Field명(1)
	
		arrHeader(0) = "예적금코드"
		arrHeader(1) = "은행코드"						' Header명(1)
		arrHeader(2) = "은행명"						' Header명(1)
		
	Case "CP"
	
		arrParam(0) = "지불구매카드 팝업"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"				
		arrParam(2) = strCode
		arrParam(3) = ""
		
		arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("CP", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "지불구매카드번호"
		
	    arrField(0) = "A.NOTE_NO"		
	    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
	    arrField(2) = "C.BP_NM"	    
	    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
	    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
	    arrField(5) = "B.BANK_NM"	        
	    
	    arrHeader(0) = "지불구매카드번호"
	    arrHeader(1) = "금액"
		arrHeader(2) = "거래처"        		        	
		arrHeader(3) = "발행일"        		        
		arrHeader(4) = "만기일"        		        
		arrHeader(5) = "은행"       
	
	Case "NP"
	
		arrParam(0) = "지급어음번호 팝업"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"
		arrParam(2) = strCode
		arrParam(3) = ""
		
		arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D3", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "지급어음번호"
		
	    arrField(0) = "A.NOTE_NO"		
	    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
	    arrField(2) = "C.BP_NM"	    
	    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
	    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
	    arrField(5) = "B.BANK_NM"	        
	    
	    arrHeader(0) = "지급어음번호"
	    arrHeader(1) = "어음금액"
		arrHeader(2) = "거래처"        		        	
		arrHeader(3) = "발행일"        		        
		arrHeader(4) = "만기일"        		        
		arrHeader(5) = "은행"       
	
	Case else
	
		arrParam(0) = "배서어음번호 팝업"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"				
		arrParam(2) = strCode
		arrParam(3) = ""
		
		arrParam(4) = "A.NOTE_STS = " & FilterVar("ED", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "배서어음번호"
		
	    arrField(0) = "A.NOTE_NO"		
	    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
	    arrField(2) = "C.BP_NM"	    
	    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
	    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
	    arrField(5) = "B.BANK_NM"	        
	    
	    arrHeader(0) = "배서어음번호"
	    arrHeader(1) = "어음금액"
		arrHeader(2) = "거래처"
		arrHeader(3) = "발행일"        		        
		arrHeader(4) = "만기일"        		        
		arrHeader(5) = "은행"       
		
	End Select
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopup(arrRet,Irow, C_PayTypeDesc)
	End If	
	
End Function


'======================================================================================================
'   Function Name : OpenCapExpNo()
'   Function Desc : 
'=======================================================================================================

Function  OpenPopUp(Byval strCode, Byval iRow, Byval iCol)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrParamAdo(3)
	Dim strAsstNo
	Dim IntRetCd

	If IsOpenPop = True Then Exit Function	
	
	Select Case iCol

		Case C_BizPartnerPopup

			arrParam(0) = "거래처 팝업"
			arrParam(1) = "B_BIZ_PARTNER"				
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""			
			arrParam(5) = "거래처코드"
	
			arrField(0) = "BP_CD"	
			arrField(1) = "BP_NM"
			arrField(2) = "BP_RGST_NO"
    
			arrHeader(0) = "거래처 코드"		
			arrHeader(1) = "거래처 명"
			arrHeader(2) = "사업자등록번호"

		Case C_TaxTypePopup

			arrParam(0) = "부가세유형"						' 팝업 명칭 
			arrParam(1) = "B_Minor,b_configuration"				' TABLE 명칭 
			arrParam(2) = strCode		' Code Condition
			arrParam(3) = ""
			arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9001", "''", "S") & "  and B_Minor.minor_cd = b_configuration.minor_cd and " & _
			              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd = B_Minor.Major_Cd"	 
			arrParam(5) = "부가세유형"						' TextBox 명칭			

			arrField(0) = "B_Minor.MINOR_CD"							' Field명(0)
			arrField(1) = "B_Minor.MINOR_NM"							' Field명(1)
			arrField(2) = "b_configuration.REFERENCE"

			arrHeader(0) = "부가세유형"						' Header명(0)
			arrHeader(1) = "부가세명"						' Header명(1)
			arrHeader(2) = "부가세Rate"
		
		Case C_ReportBizAreaPopup

			arrParam(0) = "신고사업장 팝업"	
			arrParam(1) = "B_TAX_BIZ_AREA"				
			arrParam(2) = strCode
			arrParam(3) = "" 
			arrParam(4) = ""
			arrParam(5) = "신고사업장"
	
			arrField(0) = "TAX_BIZ_AREA_CD"	
			arrField(1) = "TAX_BIZ_AREA_NM"
    
			arrHeader(0) = "신고사업장코드"		
			arrHeader(1) = "신고사업장명"		

	End Select				
	
	IsOpenPop = True

	If iCol = C_BizPartnerPopup then		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=650px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	else 
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	end if
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iRow, iCol)
	End If

End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet, Byval iRow, Byval iCol)
	With frm1.vspdData

		.Row = iRow
	
		Select Case iCol
			Case C_BizPartnerPopup
				.Col = C_BizPartnerCd
				.Text = arrRet(0)
				
				.Col = C_BizPartnerNm
				.Text = arrRet(1)
			
			Case C_TaxTypePopup
				.Col = C_TaxTypeCd
				.Text = arrRet(0)
				
				.Col = C_TaxTypeNm
				.Text = arrRet(1)

				.Col = C_TaxRate
				.Text = arrRet(2)
			
			Case C_ReportBizAreaPopup
				.Col = C_ReportBizAreaCd
				.Text = arrRet(0)
				
				.Col = C_ReportBizAreaNm
				.Text = arrRet(1)

			Case C_PayTypeDesc
				.Col = C_PayTypeDesc
				.Text = arrRet(0)

		End Select				

	End With
   	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow iRow
End Function			



'==========================================  OpenPoRef()  ================================================
'	Name : OpenMasterRef1()
'	Description : 자료부 팝업 
'========================================================================================================= 
Function OpenMasterRef1()

	Dim strRet
	Dim arrParam(7)

	If IsOpenPop = True Then Exit Function
	If frm1.txtAsstNo1.className = parent.UCN_PROTECTED Then Exit Function

	IsOpenPop = True
	
	strRet = window.showModalDialog("a7103ra1.asp", Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet(0) = "" Then
		frm1.txtAsstNo1.focus
		Exit Function
	Else
		Call SetMasterRef1(strRet)
	End If	
		
End Function

'=============================================  SetPoRef()  ============================================
'	Name : SetMasterRef1()
'	Description : 자료부 팝업 
'=======================================================================================================

Sub SetMasterRef1(strRet)
    
	frm1.txtAsstNo1.focus
	frm1.txtAsstNo1.value     = strRet(0)
	frm1.txtAsstNm1.value	 = strRet(3)
	frm1.txtRegDt.text       = strRet(2)	

	frm1.txtAcctDeptNm.value = strRet(9)
	
	frm1.txtAcqQty.text     = strRet(7)	
	frm1.txtInvQty.text     = strRet(8)
	
	lgBlnFlgChgValue = True
	
End Sub

'============================================================
'지출부서코드 팝업 
'============================================================
Function OpenPopupDept(Byval strCode)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = parent.UCN_PROTECTED Then Exit Function
	
	arrParam(0) = strCode						'부서코드 
	arrParam(1) = frm1.txtChgDt.Text			'날짜(Default:현재일)
	arrParam(2) = "1"							'부서권한(lgUsrIntCd)
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/DeptPopupDt.asp", Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	End If

	frm1.txtDeptCd.focus
	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)

	Call txtDeptCd_onChange
	
	lgBlnFlgChgValue = True
End Function


'===========================================  OpenPoRef()  =============================================
'	Name : OpenDeptPopup()
'	Description : 지출부서 팝업 
'=======================================================================================================
Function OpenDeptPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strAsstNo
	Dim IntRetCD

	IsOpenPop = True
	
	strAsstNo = Trim(frm1.txtAsstNo1.value)
	
	If strAsstNo = "" then
		IntRetCD = DisplayMsgBox("117326","X","X","X")    '자산번호를 입력하십시오.
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = "지출부서팝업"	
	arrParam(1) = "B_ACCT_DEPT"
	arrParam(2) = Trim(frm1.txtDeptCd.value)
	arrParam(3) = ""
	arrParam(4) = " INTERNAL_CD IN (SELECT INTERNAL_CD FROM A_ASSET_INFORM_OF_DEPT WHERE ASST_NO =  " & FilterVar(frm1.txtAsstNo1.value, "''", "S") & " )"
	arrParam(5) = "지출부서"			
	
    arrField(0) = "DEPT_CD"	
    arrField(1) = "DEPT_NM"
    arrField(2) = "ORG_CHANGE_ID "
    
    arrHeader(0) = "지출부서코드"
    arrHeader(1) = "지출부서명"
    arrHeader(2) = "조직개편ID"
        
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDeptPopup(arrRet)
	End If	
		
End Function

'========================================  SetPoRef()  =================================================
'	Name : SetDeptPopup()
'	Description : 지출부서 팝업 
'=======================================================================================================
Sub SetDeptPopup(strRet)
    
	frm1.txtDeptCd.focus
	frm1.txtDeptCd.value     = Trim(strRet(0))
	frm1.txtDeptNm.value	 = strRet(1)
	frm1.hOrgChangeId.value	 = Trim(strRet(2))
	
	lgBlnFlgChgValue = True
End Sub

'===========================================================================
' Function Name : OpenDept
' Function Desc : OpenDeptCode Reference Popup
'===========================================================================
' jsk 20030826 부서 팝업 수정 
Function OpenDept()

	Dim arrRet
	Dim strAsstNo
	Dim IntRetCD
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	strAsstNo = Trim(frm1.txtAsstNo1.value)
	If strAsstNo = "" then
		IntRetCD = DisplayMsgBox("117326","X","X","X")    '자산번호를 입력하십시오.
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtDeptCd.value) 'strCode		            '  Code Condition
   	arrParam(1) = frm1.txtChgDt.Text

	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,"DeptCd")
	End If	
End Function

Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg

			case "DeptCd"
				.txtChgDt.text			= arrRet(3)
				.txtDeptCd.value        = arrRet(0)
				.txtDeptNm.value 		= arrRet(1)
				Call txtDeptCd_OnChange()

		End select	

		lgBlnFlgChgValue = True
	End With
'	msgbox frm1.hOrgChangeId.value
End Function

'===========================================  OpenPoRef()  =============================================
'	Name : OpenDocCurPopup()
'	Description : 거래통화 팝업 
'=======================================================================================================
Function OpenDocCurPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strAsstNo
	Dim IntRetCD

	arrParam(0) = "거래통화 팝업"
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래통화"
	
	arrField(0) = "CURRENCY"	
	arrField(1) = "CURRENCY_DESC"	
    
	arrHeader(0) = "거래통화"
	arrHeader(1) = "거래통화명"
        
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDocCur.focus
		Exit Function
	Else
		Call SetDocCurPopup(arrRet)
	End If	
		
End Function

'========================================  SetPoRef()  =================================================
'	Name : SetDocCurPopup()
'	Description : 거래통화 팝업 
'=======================================================================================================
Sub SetDocCurPopup(arrRet)
	frm1.txtDocCur.focus
	frm1.txtDocCur.value    = arrRet(0)		
	If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' 거래통화하고 Company 통화가 다를때 환율을 0으로 셋팅 
		frm1.txtXchRate.text	= 0                         ' 디폴트값인 1이 들어가 있으면 환율이 입력된 것으로 판단하여 
						                                        ' 환율정보를 읽지 않고 입력된 값으로 계산. 
	Else 
		frm1.txtXchRate.text	= 1
	End If							
	call txtDocCur_OnChangeASP()

	lgBlnFlgChgValue = True

End Sub

'=======================================================================================================
'Description : 결의전표 생성내역 팝업 
'=======================================================================================================
Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5130ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function
'=======================================================================================================
'Description : 회계전표 생성내역 팝업 
'=======================================================================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function



 
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

	ggoSpread.source = frm1.vspdData
	
	Call CommonQueryRs("MINOR_CD, MINOR_NM"," B_MINOR "," MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " & _ 
				"AND MINOR_CD not in(" & FilterVar("NR", "''", "S") & " , " & FilterVar("PP", "''", "S") & " , " & FilterVar("CR", "''", "S") & " , " & FilterVar("AP", "''", "S") & " , " & FilterVar("PR", "''", "S") & " )", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    lgF0 = "AP" & vbTab & replace(lgF0, Chr(11), vbTab) 
	ggoSpread.SetCombo lgF0, C_PayTypeCd
	
	lgF1 = "매입채무" & vbTab & replace(lgF1, Chr(11), vbTab)
	ggoSpread.SetCombo lgF1, C_PayTypeNm
	
End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			'@Grid_Column
			C_ChgNo					= iCurColumnPos(1)
			C_ChgDesc				= iCurColumnPos(2)
			C_BizPartnerCd			= iCurColumnPos(3)
			C_BizPartnerPopup		= iCurColumnPos(4)
			C_BizPartnerNm			= iCurColumnPos(5)
			C_ChgAmt				= iCurColumnPos(6)
			C_ChgLocAmt				= iCurColumnPos(7)
			C_TaxTypeCd				= iCurColumnPos(8)
			C_TaxTypePopup			= iCurColumnPos(9)
			C_TaxTypeNm				= iCurColumnPos(10)
			C_TaxRate				= iCurColumnPos(11)
			C_TaxAmt				= iCurColumnPos(12)
			C_TaxLocAmt				= iCurColumnPos(13)
			C_ReportBizAreaCd		= iCurColumnPos(14)
			C_ReportBizAreaPopup	= iCurColumnPos(15)
			C_ReportBizAreaNm		= iCurColumnPos(16)
			C_IssuedDt				= iCurColumnPos(17)
			C_PayTypeCd				= iCurColumnPos(18)
			C_PayTypeNm				= iCurColumnPos(19)
			C_PayTypeDesc			= iCurColumnPos(20)
			C_PayTypePopup			= iCurColumnPos(21)
			C_ApAcctCd				= iCurColumnPos(22)
			C_ApAcctPopup			= iCurColumnPos(23)
			C_ApAcctNm				= iCurColumnPos(24)
			C_ApDueDt				= iCurColumnPos(25)
    End Select
End Sub


'======================================================================================================
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     'Load table , B_numeric_format
    
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                         
                                                                            'Format Numeric Contents Field                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
'    frm1.txtAcqAmt.AllowNull =false
    'frm1.txtAcqLocAmt.AllowNull =false
'    frm1.txtTotalAmt.AllowNull =false
    'frm1.txtTotalLocAmt.AllowNull =false
'    frm1.txtApAmt.AllowNull =false
    'frm1.txtApLocAmt.AllowNull =false
    'frm1.txtVatAmt.AllowNull =false
    'frm1.txtVatLocAmt.AllowNull =false    
'    frm1.txtAcqQty.AllowNull =false
'    frm1.txtInvQty.AllowNull =false
    
       
    Call InitSpreadSheet                                                    'Setup the Spread sheet
    Call InitVariables                                                      'Initializes local global variables
	Call InitComboBox
    Call SetDefaultVal

    Call SetToolbar("1110111100100011")										'버튼 툴바 제어 
    frm1.txtAsstNo.focus

End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtChgDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgDt.Action = 7
    End If
End Sub


'=======================================================================================================
'   Event Name : txtPrpaymDt_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtChgDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2
	Dim IntRows
	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtChgDt.Text <> "") Then
		strSelect	=			 " Distinct org_change_id  "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtChgDt.Text, parent.gDateFormat,""), "''", "S") & "))"			

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
		If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(frm1.hOrgChangeId.value) Then
			'IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			
		End If
	End If

    lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )

End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col
		Select Case Col
			Case  C_ChgAmt
				Frm1.vspdData.Col = C_ChgLocAmt
				Frm1.vspdData.Text = ""
		End Select
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

   	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
End Sub




'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim i
    Dim tmpDrCrFG

    Call SetPopUpMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows <= 0 Then 
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col,lgSortKey
            lgSortKey = 1
        End If
        Exit Sub

    End If
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
 
	lgBlnFlgChgValue = True
	If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' 거래통화하고 Company 통화가 다를때 환율을 0으로 셋팅 
		frm1.txtXchRate.text	= 0                         ' 디폴트값인 1이 들어가 있으면 환율이 입력된 것으로 판단하여 
							                                        ' 환율정보를 읽지 않고 입력된 값으로 계산. 
	Else 
		frm1.txtXchRate.text	= 1
	End If			

    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							

		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    
End Sub


'==========================================================================================
'   Event Name : txtDocCur_OnChangeASP
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChangeASP()
 
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							

		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strCard
    Dim strCode
  
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		.Row = Row
		
		if Row > 0 And Col = C_BizPartnerPopup Then

			.Col = C_BizPartnerCd
			
			Call OpenPopup(Trim(.Text), Row, C_BizPartnerPopup)
		
		Elseif Row > 0 And Col = C_TaxTypePopup Then

			.Col = C_TaxTypeCd
			
			Call OpenPopup(Trim(.Text), Row, C_TaxTypePopup)
		
		Elseif Row > 0 And Col = C_ReportBizAreaPopup Then
			
			.Col = C_ReportBizAreaCd
			
			Call OpenPopup(Trim(.Text), Row, C_ReportBizAreaPopup)
			
		Elseif Row > 0 And Col = C_PayTypePopup Then
		
			.Col = C_PayTypeCd
			strCard = UCase(Trim(.Text))
			
			.Col = C_PayTypeDesc
			strCode = UCase(Trim(.Text))
			
			Call OpenNoteNo(Row, strCode, strCard)

		Elseif Row > 0 And Col = C_ApAcctPopup Then
		
			.Col = C_ApAcctCd
			strCode = UCase(Trim(.Text))
			
			Call OpenApAcctCd(Row, strCode)

		End If
	
	End With
End Sub


Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

		If Row >= NewRow Then
		    Exit Sub
		End If

    End With

End Sub



'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	 '----------  Coding part  -------------------------------------------------------------   
        
End Sub


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim varData

	With frm1.vspdData
		.ReDraw = False
		.Row = Row
    
		Select Case Col
			Case  C_PayTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_PayTypeCd
				.Value = intIndex
				varData = .text
		End Select
		
		ggoSpread.source = frm1.vspdData
		
		Select Case UCase(Trim(.text))
		
		Case "DP", "NP", "NE", "CP"								
		
			ggoSpread.SpreadUnLock		C_PayTypeDesc,		Row,	C_PayTypeDesc	,Row			
			ggoSpread.SSSetRequired		C_PayTypeDesc,		Row,	Row			
			ggoSpread.SpreadUnLock		C_PayTypePopup,		Row,	C_PayTypePopup			
			
			'미지급금계정추가				
			ggoSpread.SpreadLock		C_ApAcctCd,				Row, C_ApAcctCd			,Row			
			ggoSpread.SSSetProtected	C_ApAcctCd,				Row, Row	
			ggoSpread.SpreadLock		C_ApAcctPopup,			Row,	C_ApAcctPopup,			Row		
			
			ggoSpread.SpreadLock		C_ApDueDt,				Row, C_ApDueDt			,Row			
			ggoSpread.SSSetProtected	C_ApDueDt,				Row, Row	
		
		Case "CS", "CK"			
			
			ggoSpread.SpreadLock		C_PayTypeDesc,		Row,	C_PayTypeDesc,	Row
			ggoSpread.SpreadLock		C_PayTypePopup,		Row,	C_PayTypePopup,	Row
			
			'미지급금계정추가				
			ggoSpread.SpreadLock		C_ApAcctCd,				Row,	C_ApAcctCd,			Row		
			ggoSpread.SpreadLock		C_ApAcctPopup,			Row,	C_ApAcctPopup,			Row		
			
			ggoSpread.SpreadLock		C_ApDueDt,				Row,	C_ApDueDt,			Row		
			
		Case "AP"				
		
			ggoSpread.SpreadLock		C_PayTypeDesc,		Row,	C_PayTypeDesc,	Row
			ggoSpread.SpreadLock		C_PayTypePopup,		Row,	C_PayTypePopup,	Row
			
			'미지급금계정추가				
			ggoSpread.SpreadUnLock		C_ApAcctCd,			Row,	C_ApAcctCd,			Row
			ggoSpread.SSSetRequired		C_ApAcctCd,			Row,	Row					
			ggoSpread.SpreadUnLock		C_ApAcctPopup,		Row,	C_ApAcctPopup,			Row
			
			ggoSpread.SpreadUnLock		C_ApDueDt,			Row,	C_ApDueDt,			Row
			ggoSpread.SSSetRequired		C_ApDueDt,			Row,	Row					
		
		Case Else 				
		
			ggoSpread.SpreadLock		C_PayTypeDesc,		Row,	C_PayTypeDesc,	Row
			ggoSpread.SpreadLock		C_PayTypePopup,		Row,	C_PayTypePopup,	Row
			
			'미지급금계정추가				
			ggoSpread.SpreadLock		C_ApAcctCd,				Row,	C_ApAcctCd,			Row		
			ggoSpread.SpreadLock		C_ApAcctPopup,			Row,	C_ApAcctPopup,			Row		
			
			ggoSpread.SpreadLock		C_ApDueDt,				Row,	C_ApDueDt,			Row				
				
		End Select
		
		.Col = C_PayTypeDesc			
		.Text = ""
		.Col = C_ApAcctCd			
		.Text = ""
		.Col = C_ApAcctNm			
		.Text = ""
		.Col = C_ApDueDt			
		.Text = ""
		
		.ReDraw = True	
	
	End With

End Sub

'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================


'======================================================================================================
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'=======================================================================================================

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing

  '-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      	    Exit Function
    	End If
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    Call InitVariables                                                      'Initializes local global variables
    Call InitSpreadSheet																			'⊙: Initializes local global variables
    Call InitComboBox

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery															'Query db data
       
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False                                                          
	
	'-----------------------
	'Check previous data area
	'-----------------------
	
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	frm1.vspdData.MaxRows = 0
	Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  'Clear Contents  Field
	Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
	Call InitVariables                                                      'Initializes local global variables
    Call InitSpreadSheet
    Call InitComboBox

	Call SetDefaultVal
	call txtDocCur_OnChangeASP()   
	Call SetToolbar("1110111100100011")

	FncNew = True 
	
	'SetGridFocus                                                          

End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================

Function FncDelete() 
    Dim IntRetCD
	FncDelete = False
		
	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")   '삭제하시겠습니까?  
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	'-----------------------
	'Precheck area
	'-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        intRetCD = DisplayMsgBox("900002","x","x","x")                                
    	Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete                                                          '☜: Delete db data
    
    FncDelete = True

End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim varIssuedDt
	Dim iRow
	
	FncSave = False
	
	Err.Clear                                                               

    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer   

    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
      
    If Not chkField(Document, "2") Then               '⊙: Check required field(Single area)
       Exit Function
    End If

	if frm1.vspdData.MaxRows < 1 then  'fpDoubleSingle8
		IntRetCD = DisplayMsgBox("117991","X","X","X")  ''자산지출 금액을 입력하십시오.
		Exit Function
	end if

    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then              '⊙: Check required field(Multi area)
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtRegDt.text,frm1.txtChgDt.text,frm1.txtRegDt.Alt,frm1.txtChgDt.Alt, _
	    	               "970023",frm1.txtRegDt.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtChgDt.focus
	   Exit Function
	End If

	with frm1.vspdData

		For iRow = 1 to .Maxrows
			
			.Row = iRow
			.Col = C_IssuedDt
			
			if IsNull(.text) or Trim(.text) = "" then
				.text = frm1.txtChgDt.text		
			end if
			
			.Col = C_PayTypeCd			
			if  Trim(.text) = "AP" then			
				.Col = C_ApDueDt
				if IsNull(.text) or Trim(.text) = "" then				
					.text = frm1.txtChgDt.text
				End if
			Else			
				.Col = C_ApDueDt
				.text = ""
			End if
	    
		Next
		
	end with	
	
	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False Then
		   Exit Function
    End If		                                                '☜: Save db data	 
	FncSave = True
	
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()
   	frm1.vspdData.ReDraw = False
   	
    if frm1.vspdData.MaxRows < 1 then Exit Function
    	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	frm1.vspdData.Col = C_ChgNo
	frm1.vspdData.Text = ""

	call vspdData_ComboSelChange(C_PayTypeNm, frm1.vspdData.ActiveRow)
    
'	MaxSpreadVal frm1.vspdData.ActiveRow

    frm1.vspdData.ReDraw = True

End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 

    if frm1.vspdData.MaxRows < 1 then Exit Function

	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow(ByVal pvRowCnt)


    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                             '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
    Else
    imRow = AskSpdSheetAddRowCount()
    
    If imRow = "" Then
        Exit Function
		End If
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    if frm1.vspdData.MaxRows < 1 then Exit Function

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
'		lgBlnFlgChgValue = True
    End With
    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                              
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
    '완성이 되지 않았음 
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next
    '완성이 되지 않았음                                                    
End Function

'=======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)										
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                               
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()


    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()  
    Call InitComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
    Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete() 
    Dim strVal
    
    DbDelete = False														'⊙: Processing is NG 
 
 	Call LayerShowHide(1)  
	With frm1
        	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003						'☜: 
        	strVal = strVal     & "&txtAsstNo=" & Trim(.txtAsstNo.value)	'조회 조건 데이타 
        	strVal = strVal     & "&txtCapExpNo=" & Trim(.txtCapExpNo.value)	'조회 조건 데이타 
        	strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey
	End With

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()		

	lgBlnFlgChgValue = False         '삭제 성공후 실행 로직 
	frm1.txtCapExpNo.value = ""
'	Call FncNew()
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
    
	DbQuery = False                                                         
	
	Call LayerShowHide(1)
	
	Dim strVal
	
	With frm1
	
        	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'☜: 
        	strVal = strVal     & "&txtAsstNo=" & Trim(.txtAsstNo.value)	'조회 조건 데이타 
        	strVal = strVal     & "&txtCapExpNo=" & Trim(.txtCapExpNo.value)	'조회 조건 데이타 
        	strVal = strVal     & "&hOrgChangeId=" & Trim(.hOrgChangeId.value)	'조회 조건 데이타 
        	strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey

	End With
	
	Call RunMyBizASP(MyBizASP, strVal)										'비지니스 ASP 를 가동 
	
	DbQuery = True                                                          
    
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================================
Function DbQueryOk()													'조회 성공후 실행로직 
	Dim varData
	Dim iRow
	
	lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
	
	Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field	
	Call SetToolbar("1111111100111111")									'버튼 툴바 제어 

	Call InitData()			
	Call SetSpreadColor(-1,-1)

 	With frm1	
 	
		For iRow = 1 To frm1.vspdData.MaxRows
	
			.vspdData.Col = C_PayTypeCd		
			.vspdData.Row = iRow
			
			varData = frm1.vspdData.text
		Next
		
		.vspdData.Redraw = True			
	End With
	
	Call SetSpreadLock		
	
	call txtDocCur_OnChangeASP()
	Call ggoOper.SetReqAttr(frm1.txtCapExpNo1,	"Q")
	'SetGridFocus
	lgBlnFlgChgValue = False	
	
End Function

'=======================================================================================================
' Sub Name : InitData
' Sub Desc : 
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	dim temp
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row  = intRow
			
			.Col	 = C_PayTypeCd
			intIndex = .Value 			

			.Col     = C_PayTypeNm
			.Value   = intindex					
		Next	
	End With	
End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	
	Dim IntRows 
	Dim IntCols 
	
	Dim lGrpcnt 
	Dim strVal
	Dim strDel
	
	DbSave = False                                                          
	
	'On Error Resume Next                                                   
	
	Call LayerShowHide(1)
	
	'Call SetSumItem
	
	strVal = ""
	strDel = ""
	
	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
	End With
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	' Data 연결 규칙 
	' 0: Flag , 1: Row위치, 2~N: 각 데이타 
	
	lGrpCnt = 1
	
	With frm1.vspdData
	    
		For IntRows = 1 To .MaxRows
		
			.Row = IntRows
			.Col = 0

			If .Text = ggoSpread.DeleteFlag Then
				strDel = strDel & "D" & parent.gColSep & lGrpCnt & parent.gColSep		'D=Delete

			ElseIf .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & parent.gColSep & lGrpCnt & parent.gColSep		'U=Update

			ElseIF .Text = ggoSpread.InsertFlag Then				
				strVal = strVal & "C" & parent.gColSep & lGrpCnt & parent.gColSep		'C=Create, Sheet가 2개 이므로 구별			
			else
				strVal = strVal & "S" & parent.gColSep & lGrpCnt & parent.gColSep		'Update이고 single만 변한경우 처리를 위해서.			
			End If

			Select Case .Text
				Case ggoSpread.DeleteFlag
					.Col = C_ChgNo
					strDel = strDel & Trim(.Text) & parent.gRowSep				    '마지막 데이타는 Row 분리기호를 넣는다 
					
					lGrpcnt = lGrpcnt + 1             
								
				Case ggoSpread.UpdateFlag
					.Col = C_ChgNo								'2
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_ChgDesc							'3
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_BizPartnerCd						'4
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_ChgAmt								'5
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_ChgLocAmt
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxTypeCd							'7
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_TaxRate							'8
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxAmt								'9
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxLocAmt							'10
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_ReportBizAreaCd					'11
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_IssuedDt							'12
					strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gColSep
					.Col = C_PayTypeCd							'13
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_PayTypeDesc						'14
					strVal = strVal & Trim(.Text) & parent.gColSep		
					.Col = C_ApAcctCd						'14
					strVal = strVal & Trim(.Text) & parent.gColSep		
					.Col = C_ApDueDt							'15
					strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gRowSep					
					
					lGrpCnt = lGrpCnt + 1
					
				Case ggoSpread.InsertFlag
					.Col = C_ChgNo								'2
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_ChgDesc							'3
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_BizPartnerCd						'4
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_ChgAmt								'5
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_ChgLocAmt
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxTypeCd							'7
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_TaxRate							'8
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxAmt								'9
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxLocAmt							'10
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_ReportBizAreaCd					'11
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_IssuedDt							'12
					strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gColSep					
					.Col = C_PayTypeCd							'13
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_PayTypeDesc						'14
					strVal = strVal & Trim(.Text) & parent.gColSep		
					.Col = C_ApAcctCd						'14
					strVal = strVal & Trim(.Text) & parent.gColSep		
					.Col = C_ApDueDt							'15
					strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gRowSep				
									
					lGrpCnt = lGrpCnt + 1
					
				Case else			'Update이고 single만 변한경우 처리를 위해서.
					.Col = C_ChgNo								'2
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_ChgDesc							'3
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_BizPartnerCd						'4
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_ChgAmt								'5
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_ChgLocAmt
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxTypeCd							'7
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_TaxRate							'8
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxAmt								'9
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_TaxLocAmt							'10
					if Trim(.text) = "" then
						.text = 0
					end if 							'6
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_ReportBizAreaCd					'11
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_IssuedDt							'12
					strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gColSep
					.Col = C_PayTypeCd							'13
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_PayTypeDesc						'14
					strVal = strVal & Trim(.Text) & parent.gColSep		
					.Col = C_ApAcctCd						'14
					strVal = strVal & Trim(.Text) & parent.gColSep		
					.Col = C_ApDueDt							'15
					strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gRowSep					
					
					lGrpCnt = lGrpCnt + 1
					
			End Select
		Next
	End With

	frm1.txtMaxRows.value = lGrpCnt-1										'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value = strDel & strVal									'☜: Spread Sheet 내용을 저장 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'☜: 저장 비지니스 ASP 를 가동 

	DbSave = True                                                           
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

   	lgBlnFlgChgValue = false	

    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    Call InitVariables                                                      'Initializes local global variables
    Call InitSpreadSheet																			'⊙: Initializes local global variables
    Call InitComboBox
    
	Call DbQuery	

End Function

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
Sub txtDeptCd_onChange()

    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii
	if Trim(frm1.txtDeptCd.value = "") then		frm1.txtDeptNm.value = ""
	If Trim(frm1.txtDeptCd.value = "") Then		Exit sub
	If Trim(frm1.txtChgDt.Text = "") Then		Exit sub'
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtChgDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
'msgbox "a select " & strSelect  & " From " & strFrom& " where " & strWhere 
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
		End If
	lgBlnFlgChgValue = True
'	msgbox frm1.hOrgChangeId.value
End Sub


'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1

		ggoOper.FormatFieldByObjectOfCur .txtTotalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec

	End With

End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_ChgAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
		ggoSpread.SSSetFloatByCellOfCur C_TaxAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
		
	End With

End Sub
    
'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub




</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td NOWRAP background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자산지출내역등록</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
			        <TD WIDTH=10>&nbsp;</TD>								
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>자산번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstNo" SIZE=10 MAXLENGTH=18 TAG="12XXXU" ALT="자산번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnAsstNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMasterRef()">
										<INPUT TYPE="Text" NAME="txtAsstNm" SIZE=20 MAXLENGTH=40 tag="14X" ALT="자산명"></TD>
									<TD CLASS="TD5" NOWRAP>자본적지출번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCapExpNo" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="자본적지출번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnCapExpNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCapExpNo()"></TD>
								</TR>					
							</TABLE>        
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=20 valign=top>
						<TABLE <%=LR_SPACE_TYPE_50%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>자산번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstNo1" SIZE=10 MAXLENGTH=18 TAG="23XXXU" ALT="자산번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnAsstNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMasterRef1()">
									<INPUT TYPE="Text" NAME="txtAsstNm1" SIZE=20 MAXLENGTH=40 tag="24X" ALT="자산명"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>관리부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctDeptNm" SIZE=27 MAXLENGTH=40 tag="24XXXU" ALT="회계부서명"></TD>
								<TD CLASS="TD5" NOWRAP>취득일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a7126ma1_fpDateTime1_txtRegDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>취득수량</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a7126ma1_fpDoubleSingle0_txtAcqQty.js'></script>										
								</TD>										
								<TD CLASS="TD5" NOWRAP>재고수량</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a7126ma1_fpDoubleSingle1_txtInvQty.js'></script>										
								</TD>										
							</TR>
						</TABLE>								
					</TD>
				</TR>		
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>							
							<TR>
								<TD CLASS="TD5" NOWRAP>자본적지출번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCapExpNo1" SIZE=20 MAXLENGTH=18 tag="21XXXU" ALT="자본적지출번호"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>변동일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a7126ma1_fpDateTime2_txtChgDt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>지출부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="지출부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDept()">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 tag="24" ALT="지출부서명"></TD>
							</TR>
<%	If gIsShowLocal <> "N" Then	%>
							<TR>
								<TD CLASS="TD5" NOWRAP>거래통화</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" SIZE=10 tag="22XXXU" ><IMG SRC="../../image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDocCurPopup()"></TD>
								<TD CLASS="TD5" NOWRAP>환율</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a7126ma1_fpDoubleSingle5_txtXchRate.js'></script></TD>
							</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur"><INPUT TYPE=HIDDEN NAME="txtXchRate">
<%	End If %>
							<TR>
<%	If gIsShowLocal <> "N" Then	%>
								<TD CLASS="TD5" NOWRAP>총지출액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a7126ma1_fpDoubleSingle6_txtTotalAmt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>총지출액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a7126ma1_fpDoubleSingle7_txtTotalLocAmt.js'></script></TD>
<%	ELSE %>
								<TD CLASS="TD5" NOWRAP>총지출액</TD>
	                            <TD CLASS="TD656" NOWRAP COLSPAN=3><script language =javascript src='./js/a7126ma1_fpDoubleSingle6_txtTotalAmt.js'></script></TD>
<INPUT TYPE=HIDDEN NAME="txtTotalLocAmt">
<%	End If %>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>결의전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="결의전표번호"></TD>
								<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="회계전표번호"></TD>
							</TR>
							<TR>
								<TD WIDTH="80%" HEIGHT=100% COLSPAN=4>
									<script language =javascript src='./js/a7126ma1_fpSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=10>		
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1" ></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="34" TABINDEX = "-1" ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"    tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMode"         tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows"	  tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode"	  tag="34" TABINDEX = "-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

