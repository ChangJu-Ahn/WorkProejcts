<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>

<!--'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 본지점 출금반제 
'*  3. Program ID           : a8105ma1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : ap004mhq
'*  7. Modified date(First) : 2001/01/31
'*  8. Modified date(Last)  : 2001/01/31
'*  9. Modifier (First)     : Chang Sung Hee
'* 10. Modifier (Last)      : Chang Sung Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   ***************************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncServer.asp"  -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/AdoQuery.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE= VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "a8105mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a8105mb2.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID =  "a8105mb3.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Const C_ApNo = 1
Const C_AcctCd = 2
Const C_AcctNm = 3							
Const C_APBizCd = 4
Const C_APBizNm = 5
Const C_ApDt = 6
Const C_ApDueDt = 7
Const C_DocCur = 8
Const C_ApAmt = 9
Const C_ApRemAmt = 10
Const C_ApClsAmt = 11
Const C_ApClsLocAmt = 12
Const C_SHEETMAXROWS = 13

'vspddata1
Const C_BizCd = 1
Const C_BizPb = 2
Const C_BizNm = 3
Const C_HQDeptCd = 4
Const C_HQDeptPb = 5
Const C_HQDeptNm = 6
Const C_HqAllcAmt = 7
Const C_HqAllcLocAmt = 8

'@Global_Var
Dim  lgBlnFlgChgValue             'Variable is for Dirty flag
Dim  lgIntGrpCount                'Group View Size를 조사할 변수 
Dim  lgIntFlgMode                 'Variable is for Operation Status

Dim  lgStrPrevKey
Dim  lgStrPrevKey1
Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3
Dim  lgLngCurRows
Dim  strMode

Dim  intItemCnt					
Dim  IsOpenPop	
Dim  lgRetFlag	                'Popup
Dim  gSelframeFlg

Dim  lgCurrRow

'+++++++ 2002.01.10 수정사항;날짜 format ++++++++++++
<%
dim dtToday
dtToday = GetSvrDate
%>

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
Sub  InitVariables()

	lgIntFlgMode = OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
    lgStrPrevKey = ""                            'initializes Previous Key
    lgStrPrevKey1 = ""
    lgStrPrevKeyDtl = 0                         'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtAllcDt.text = UniConvDateAToB("<%=dtToday%>",gServerDateFormat,gDateFormat)
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029(gCurrency,"I","*") %>
<% Call LoadBNumericFormat("I", "*") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet()
        
    With frm1
	
		.vspdData.MaxCols = C_ApClsLocAmt + 1   
		.vspdData.Col = .vspdData.MaxCols
		.vspdData.ColHidden = true
	
		ggoSpread.Source = .vspdData
	
		ggoSpread.Spreadinit

		ggoSpread.SSSetEdit C_ApNo, "채무번호", 18,3		'1
		ggoSpread.SSSetEdit C_AcctCd,	"계정코드", 15,3	'2
		ggoSpread.SSSetEdit C_AcctNm, "계정코드명", 15,3	'3    
		ggoSpread.SSSetEdit C_APBizCd, "사업장", 10,3	'6
		ggoSpread.SSSetEdit C_APBizNm, "사업장명", 20,3	'7    
		ggoSpread.SSSetDate C_ApDt, "채무일자",10, 2, gDateFormat  
		ggoSpread.SSSetDate C_ApDueDt, "만기일자", 10, 2, gDateFormat  
		ggoSpread.SSSetEdit C_DocCur, "거래통화", 8, 3'10
		ggoSpread.SSSetFloat C_ApAmt, "채무액", 15, ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat C_ApRemAmt, "채무잔액", 15, ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat C_ApClsAmt, "반제금액",15, ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat C_ApClsLocAmt, "반제금액(자국)",15, ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
    
    
		.vspdData1.MaxCols = C_HqAllcLocAmt + 1   
		.vspdData1.Col = .vspdData1.MaxCols
		.vspdData1.ColHidden = true
	
		ggoSpread.Source = .vspdData1
	
		ggoSpread.Spreadinit

		ggoSpread.SSSetEdit C_BizCd, "사업장", 20,,,10,2	'1
		ggoSpread.SSSetButton    C_BizPb
		ggoSpread.SSSetEdit C_BizNm, "사업장명", 20,,,20,2	'3    
		ggoSpread.SSSetEdit C_HQDeptCd, "부서", 20,,,10,2	'6
		ggoSpread.SSSetButton    C_HQDeptPb
		ggoSpread.SSSetEdit C_HQDeptNm, "부서명", 20,,,20,2
		ggoSpread.SSSetFloat C_HqAllcAmt, "출금액", 15, ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat C_HqAllcLocAmt, "출금액(자국)", 15, ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		
    End With
	frm1.vspdData.ReDraw = true
	frm1.vspdData1.ReDraw = true
	
	intItemCnt = 0    
    
    SetSpreadLock "I", 0, 1, ""
    SetSpreadLock "I", 1, 1, ""
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2 )
       
    With frm1
		Select Case stsFg
			Case "Q"
				Select Case Index
					Case 0
						ggoSpread.Source = frm1.vspdData
						.vspdData.ReDraw = False
						ggoSpread.SpreadLock C_ApNo,-1, C_ApNo
						ggoSpread.SpreadLock C_AcctCd,-1, C_AcctCd
						ggoSpread.SpreadLock C_AcctNm,-1, C_AcctNm
						ggoSpread.SpreadLock C_APBizCd,-1, C_APBizCd
						ggoSpread.SpreadLock C_APBizNm,-1, C_APBizNm
						ggoSpread.SpreadLock C_ApDt,-1, C_ApDt
						ggoSpread.SpreadLock C_ApDueDt,-1, C_ApDueDt
						ggoSpread.SpreadLock C_DocCur,-1, C_DocCur
						ggoSpread.SpreadLock C_ApAmt,-1, C_ApAmt
						ggoSpread.SpreadLock C_ApRemAmt,-1, C_ApRemAmt    
						.vspdData.ReDraw = True   
					Case 1
						ggoSpread.Source = frm1.vspdData1
						.vspdData1.ReDraw = False
						ggoSpread.SpreadLock C_BizCd,-1, C_BizCd
						ggoSpread.SpreadLock C_BizPb,-1, C_BizPb
						ggoSpread.SpreadLock C_BizNm,-1, C_BizNm
						ggoSpread.SpreadLock C_HQDeptCd,-1, C_HQDeptCd
						ggoSpread.SpreadLock C_HQDeptPb,-1, C_HQDeptPb
						ggoSpread.SpreadLock C_HQDeptNm,-1, C_HQDeptNm
						.vspdData1.ReDraw = True   
				End Select				
			Case "I"
				Select Case Index
					case 0
						ggoSpread.Source = frm1.vspdData
						.vspdData.ReDraw = False
						ggoSpread.SpreadLock C_ApNo,-1, C_ApNo
						ggoSpread.SpreadLock C_AcctCd,-1, C_AcctCd
						ggoSpread.SpreadLock C_AcctNm,-1, C_AcctNm
						ggoSpread.SpreadLock C_APBizCd,-1, C_APBizCd
						ggoSpread.SpreadLock C_APBizNm,-1, C_APBizNm
						ggoSpread.SpreadLock C_ApDt,-1, C_ApDt
						ggoSpread.SpreadLock C_ApDueDt,-1, C_ApDueDt
						ggoSpread.SpreadLock C_DocCur,-1, C_DocCur
						ggoSpread.SpreadLock C_ApAmt,-1, C_ApAmt
						ggoSpread.SpreadLock C_ApRemAmt,-1, C_ApRemAmt    
						.vspdData.ReDraw = True   
					Case 1
						ggoSpread.Source = frm1.vspdData1
						.vspdData1.ReDraw = False						
						ggoSpread.SpreadLock C_BizNm,-1, C_BizNm
						ggoSpread.SpreadLock C_HQDeptNm, -1, C_HQDeptNm
						.vspdData1.ReDraw = True
						
				End Select	
		End Select		
    End With    
End Sub
'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor(Byval stsFg, Byval Index, ByVal lRow, ByVal lRow2)
    
    DIM objSpread
    Dim iTemp       
	
With frm1
    Select Case stsFg
			Case "Q"
				Select Case Index
					Case 0
						ggoSpread.Source = frm1.vspdData
						If lRow2 = "" Then 							
							Set objSpread = frm1.vspdData    
							lRow2 = objSpread.MaxRows
						END IF	        
    
						.vspdData.ReDraw = False
						ggoSpread.SSSetProtected C_ApNo, lRow, lRow2
						ggoSpread.SSSetProtected C_AcctCd, lRow, lRow2
						ggoSpread.SSSetProtected C_AcctNm, lRow, lRow2
						ggoSpread.SSSetProtected C_APBizCd,  lRow, lRow2
						ggoSpread.SSSetProtected C_APBizNm, lRow, lRow2
						ggoSpread.SSSetProtected C_ApDt, lRow, lRow2
						ggoSpread.SSSetProtected C_DocCur, lRow, lRow2
						ggoSpread.SSSetProtected C_ApDueDt, lRow, lRow2
						ggoSpread.SSSetProtected C_ApAmt, lRow, lRow2
						ggoSpread.SSSetProtected C_ApRemAmt, lRow, lRow2
						ggoSpread.SSSetRequired C_ApClsAmt, lRow, lRow2
						.vspdData.ReDraw = True   
					Case 1
						ggoSpread.Source = frm1.vspdData1
						If lRow2 = "" Then 							
							Set objSpread = frm1.vspdData1    
							lRow2 = objSpread.MaxRows
						END IF	            
						.vspdData1.ReDraw = False
						ggoSpread.SSSetProtected C_BizCd, lRow, lRow2
						ggoSpread.SSSetProtected C_BizNm, lRow, lRow2
						ggoSpread.SSSetProtected C_HQDeptCd, lRow, lRow2
						ggoSpread.SSSetProtected C_HQDeptNm, lRow, lRow2						
						ggoSpread.SSSetRequired C_HqAllcAmt, lRow, lRow2						
						.vspdData1.ReDraw = True   						
						
				End Select				
			Case "I"
				Select Case Index
					case 0
						ggoSpread.Source = frm1.vspdData
						If lRow2 = "" Then 							
							Set objSpread = frm1.vspdData    
							lRow2 = objSpread.MaxRows
						END IF	        
    
						.vspdData.ReDraw = False
						ggoSpread.SSSetProtected C_ApNo, lRow, lRow2
						ggoSpread.SSSetProtected C_AcctCd, lRow, lRow2
						ggoSpread.SSSetProtected C_AcctNm, lRow, lRow2
						ggoSpread.SSSetProtected C_APBizCd,  lRow, lRow2
						ggoSpread.SSSetProtected C_APBizNm, lRow, lRow2
						ggoSpread.SSSetProtected C_DocCur, lRow, lRow2
						ggoSpread.SSSetProtected C_ApDt, lRow, lRow2
						ggoSpread.SSSetProtected C_ApDueDt, lRow, lRow2
						ggoSpread.SSSetProtected C_ApAmt, lRow, lRow2
						ggoSpread.SSSetProtected C_ApRemAmt, lRow, lRow2
						ggoSpread.SSSetRequired C_ApClsAmt, lRow, lRow2
						.vspdData.ReDraw = True   
					Case 1
						ggoSpread.Source = frm1.vspdData1
						If lRow2 = "" Then 							
							Set objSpread = frm1.vspdData1    
							lRow2 = objSpread.MaxRows
						END IF	        
    
						.vspdData1.ReDraw = False						
						ggoSpread.SSSetRequired C_BizCd, lRow, lRow2
						ggoSpread.SSSetProtected C_BizNm, lRow, lRow2
						ggoSpread.SSSetRequired C_HQDeptCd, lRow, lRow2
						ggoSpread.SSSetProtected C_HQDeptNm, lRow, lRow2	
						ggoSpread.SSSetRequired C_HqAllcAmt, lRow, lRow2						
						.vspdData1.ReDraw = True   
						
						.vspddata1.Col = 1
						.vspddata1.Row = lRow2
						.vspddata1.Action = 0                         'SS_ACTION_ACTIVE_CELL
						.vspddata1.EditMode = True   
						
				End Select	
		End Select		            
	
end With    
End Sub
 '========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub  InitComboBox()
	
End Sub

 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefOpenAp()

	Dim arrRet
	Dim arrParam(4)	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(2) = frm1.txtDocCur.value					
	arrParam(3) = "Q"
    
	arrRet = window.showModalDialog("../ap/a4105ra1.asp", Array(arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpenAp(arrRet)
	End If
End Function

 '------------------------------------------  SetRefOpenAp()  --------------------------------------------------
'	Name : SetRefOpenAp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRefOpenAp(Byval arrRet)
	
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	
	With frm1
	
		.vspdData.focus		
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False	
	
		TempRow = .vspdData.MaxRows												'☜: 현재까지의 MaxRows
		'.vspdData.MaxRows = .vspdData.MaxRows + (Ubound(arrRet, 1) + 1)			'☜: Reference Popup에서 선택되어진 Row만큼 추가		

		For I = TempRow to TempRow + Ubound(arrRet, 1)
			sFindFg	= "N"
			For x = 1 to TempRow
				.vspdData.Row = x
				.vspdData.Col = C_ApNo				
				IF .vspdData.Text = arrRet(I - TempRow, 0) Then
					sFindFg	= "Y"
				End IF
			Next			
			IF 	sFindFg	= "N" Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
				.vspdData.Row = I + 1				
				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag
				FOR j = 0 to  C_ApRemAmt - 1
					.vspdData.Col = j + 1												'⊙: 첫번째 컬럼 
					.vspdData.text = arrRet(I - TempRow, j)				
				Next			
			ENd IF	
		Next	
		
		frm1.txtDocCur.Value = arrRet(0, 7)				
		frm1.txtbpCd.Value = arrRet(0, 11)				
		frm1.txtbpNm.Value = arrRet(0, 12)				

		SetSpreadLock "I",0, 1,""
		SetSpreadColor "I",0, 1,""
		
		.vspdData.ReDraw = True
		gSelframeFlg = Tab1
    End With
    
End Function
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere, Byval strCode1)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd
	Dim arrParamAdo(3)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	
		Case 0		
			
		Case 1
			If frm1.txtBpCd.className = "protected" Then Exit Function
			IsOpenPop = True
			arrParam(0) = "거래처팝업"
			arrParam(1) = "B_BIZ_PARTNER"				
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래처"			
	
			arrField(0) = "BP_CD"	
			arrField(1) = "BP_NM"	
    
			arrHeader(0) = "거래처"		
			arrHeader(1) = "거래처명"					' Header명(1)			
		
		case 2
			If frm1.txtDeptCd.className = "protected" Then Exit Function
			IsOpenPop = True
			arrParam(0) = "부서팝업"					' 팝업 명칭 
			arrParam(1) = "B_Acct_Dept"						' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtDeptCd.Value)		' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = "ORG_CHANGE_ID = " & FilterVar(gChangeOrgId, "''", "S")			' Where Condition
			arrParam(5) = "부서"			
	
			arrField(0) = "Dept_CD"							' Field명(0)
			arrField(1) = "Dept_NM"							' Field명(1)
			    
			arrHeader(0) = "부서"						' Header명(0)
			arrHeader(1) = "부서명"						' Header명(1)   			    		
			arrHeader(2) = "사업부"						' Header명(0)
			
		Case	 3		
			If frm1.txtDocCur.className = "protected" Then Exit Function
			IsOpenPop = True
			arrParam(0) = "거래통화팝업"				' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"						' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtDocCur.Value)		' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "거래통화"			
	
			arrField(0) = "CURRENCY"						' Field명(0)
			arrField(1) = "CURRENCY_DESC"					' Field명(1)
    
			arrHeader(0) = "거래통화"					' Header명(0)
			arrHeader(1) = "거래통화명"
			
		Case 4
			IsOpenPop = True
			arrParam(0) = "계정코드팝업"
			arrParam(1) = "A_Acct"				
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "계정코드"			
	
			arrField(0) = "ACCT_CD"	
			arrField(1) = "ACCT_NM"	
    
			arrHeader(0) = "계정코드"		
			arrHeader(1) = "계정코드명"						' Header명(1)				
			
		Case 5	
			If frm1.txtBankCd.className = "protected" Then Exit Function
			IsOpenPop = True
			arrParam(0) = "은행팝업"
			arrParam(1) = "B_BANK, F_DPST"				
			arrParam(2) = Trim(frm1.txtBankCd.Value)
			arrParam(3) = ""
			arrParam(4) = "B_BANK.BANK_CD = F_DPST.BANK_CD "
			arrParam(5) = "은행"			
	
			arrField(0) = "F_DPST.BANK_CD"	
			arrField(1) = "B_BANK.BANK_NM"	
    
			arrHeader(0) = "은행"		
			arrHeader(1) = "은행명"	
			   
		Case 6
			If frm1.txtBankAcct.className = "protected" Then Exit Function
			IsOpenPop = True
			arrParam(0) = "계좌번호팝업"
			arrParam(1) = "B_BANK, F_DPST"				
			arrParam(2) = Trim(frm1.txtBankAcct.Value)
			arrParam(3) = ""
			
			IF Trim(frm1.txtBankCd.Value) = "" Then
				strCd = "B_BANK.BANK_CD = F_DPST.BANK_CD "
			Else
				strCd = "B_BANK.BANK_CD = F_DPST.BANK_CD AND  F_DPST.BANK_CD = " & FilterVar(frm1.txtBankCd.Value, "''", "S")
			End IF		
			
			arrParam(4) = strCd
			arrParam(5) = "계좌번호"			
			
		    arrField(0) = "F_DPST.BANK_ACCT_NO"	
		    arrField(1) = "F_DPST.BANK_CD"	
		    arrField(2) = "B_BANK.BANK_NM"	
		    
		    arrHeader(0) = "계좌번호"		
		    arrHeader(1) = "은행"	
		    arrHeader(2) = "은행명"	
		    		
		Case 7
			If frm1.txtCheckCd.className = "protected" Then Exit Function
			IsOpenPop = True
			arrParam(0) = "어음번호팝업"				        ' 팝업 명칭 
			arrParam(1) = "f_note a,b_biz_partner b, b_bank c"		' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtCheckCd.Value)	        ' Code Condition
			arrParam(3) = ""								' Name Cindition
			IF frm1.txtInputType.value = "NR"  THEN
				'받을어음 
				arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D1", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"					' Where Condition
			ELSEIF frm1.txtInputType.value = "NP" THen
				'지급어음 
				arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D3", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"					' Where Condition
			END IF	
			arrParam(5) = "어음번호"					    		' 조건필드의 라벨 명칭 
	
			arrField(0) = "Note_no"							' Field명(0)
			arrField(1) =  "F2" & gColSep & "a.Note_amt"	' Field명(1)
			arrField(2) =  "DD" & gColSep & "a.Issue_dt"	' Field명(2)
			arrField(3) = "b.bp_nm"
			arrField(4) = "c.bank_nm"    	    
	
			arrHeader(0) = "어음번호"				' Header명(0)
			arrHeader(1) = "금액"					' Header명(1)
			arrHeader(2) = "발행일"					' Header명(1)	    
			arrHeader(3) = "거래처"					' Header명(1)
			arrHeader(4) = "은행"					' Header명(1)			
						    
		Case 8
			IsOpenPop = True
			arrParam(0) = "사업장팝업"					' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"						' TABLE 명칭 
			arrParam(2) = Trim(strCode)						' Code Condition
			arrParam(3) = ""								' Name Cindition			
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "사업장코드"			
	
			arrField(0) = "BIZ_AREA_CD"						' Field명(0)
			arrField(1) = "BIZ_AREA_NM"						' Field명(1)
			    
			arrHeader(0) = "사업장"						' Header명(0)
			arrHeader(1) = "사업장명"					' Header명(1)   			 								    
		    
		Case 9
			IsOpenPop = True
			arrParam(0) = "부서팝업"					' 팝업 명칭 
			arrParam(1) = "B_ACCT_DEPT A , B_COST_CENTER C, B_BIZ_AREA B"		' TABLE 명칭 
			arrParam(2) = Trim(strCode)						' Code Condition
			arrParam(3) = ""								' Name Cindition
			
			IF 	strCode1 <> "" Then			
				arrParam(4) = "A.ORG_CHANGE_ID = " & FilterVar(gChangeOrgId, "''", "S") & _
							  " AND B.BIZ_AREA_CD = " & FilterVar(strCode1, "''", "S") & _
							  " AND A.COST_CD = C.COST_CD " & _
							  " AND C.BIZ_AREA_CD = B.BIZ_AREA_CD "
			ELse
				arrParam(4) = "A.ORG_CHANGE_ID = " & FilterVar(gChangeOrgId, "''", "S") & _
							  " AND A.COST_CD = C.COST_CD " & _
							  " AND C.BIZ_AREA_CD = B.BIZ_AREA_CD"
			END IF	
				
			arrParam(5) = "부서"			
	
			arrField(0) = "A.Dept_CD"							' Field명(0)
			arrField(1) = "A.Dept_NM"							' Field명(1)
			arrField(2) = "B.BIZ_AREA_CD"					' Field명(2)
			arrField(3) = "B.BIZ_AREA_NM"					' Field명(3)
			    
			arrHeader(0) = "부서"						' Header명(0)
			arrHeader(1) = "부서명"						' Header명(1)   			    		
			arrHeader(2) = "사업장"						' Header명(2)
			arrHeader(3) = "사업장명"	
			
		CASE 10
			If frm1.txtInputType.className = "protected" Then Exit Function    
			
			arrParam(0) = "지급유형"					' 팝업 명칭						
			arrParam(1) = "B_MINOR,B_CONFIGURATION "
			arrParam(2) = Trim(frm1.txtInputType.value)		' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " & _
						  "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "	' Where Condition								
			arrParam(5) = "지급유형"					' TextBox 명칭 
		
			arrField(0) = "B_MINOR.MINOR_CD"				' Field명(0)
			arrField(1) = "B_MINOR.MINOR_NM"				' Field명(1)
	    
			arrHeader(0) = "지급유형"					' Header명(0)
			arrHeader(1) = "지급유형명"					' Header명(1)				 								    
		    	    
	End Select				
		
	IsOpenPop = True
	
	IF iwhere = 0 Then					
		arrRet = window.showModalDialog("a8105ra1.asp", Array(arrParamAdo), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	ELSEIF iwhere = 9 Then					
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	end if
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If

End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)

	With frm1
		Select Case iWhere
			Case 0		
				.txtAllcNo.value = arrRet(0)
				
			Case 1	
				.txtBpCd.value = arrRet(0)		
				.txtBpNm.value = arrRet(1)
			Case 2
				.txtDeptCd.value = arrRet(0)		
				.txtDeptNm.value = arrRet(1)
			Case 3
				.txtDocCur.value = arrRet(0)		
				
			Case 4
				'.vspdData1.Col = C_AcctCd
				'.vspdData1.Text = arrRet(0)
				'.vspdData1.Col = C_AcctShNm
				'.vspdData1.Text = arrRet(1)
			
				'Call vspdData1_Change(C_AcctCd, frm1.vspddata1.activerow )	 ' 변경이 읽어났다고 알려줌 
			case 5
				.txtBankCd.value = arrRet(0)		
				.txtBankNm.value = arrRet(1)			    		
			case 6
				.txtBankAcct.value = arrRet(0)		
				.txtBankCd.value = arrRet(1)		
				.txtBankNm.value = arrRet(2)	
				
			case 7	
				.txtCheckCd.value = arrRet(0)		
				
			case 8				
				.vspdData1.Col = C_BizCd
				.vspdData1.Text = arrRet(0)
				.vspdData1.Col = C_BizNm	
				.vspdData1.Text = arrRet(1)
				.vspdData1.Col = C_HQDeptCd
				.vspdData1.Text = ""
				.vspdData1.Col = C_HQDeptNM	
				.vspdData1.Text = ""
				
			case 9		
				.vspdData1.Col = C_HQDeptCd
				.vspdData1.Text = arrRet(0)
				.vspdData1.Col = C_HQDeptNM	
				.vspdData1.Text = arrRet(1)
				.vspdData1.Col = C_BizCd
				.vspdData1.Text = arrRet(2)
				.vspdData1.Col = C_BizNm	
				.vspdData1.Text = arrRet(3)
				
			Case 10
				.txtInputType.value = arrRet(0)		 	
				.txtInputTypeNm.value = arrRet(1)		 	
					
				call txtInputType_OnChange()	
	End Select				
	End With
	IF iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	end if	
	
End Function

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenCtrlPB()
'	Description : PopUp 관리항목 
'--------------------------------------------------------------------------------------------------------- 
Function OpenCtrlPB(Byval strTable, Byval strFld1 , Byval strFld2 , Byval strCode )
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "관리항목값팝업"				' 팝업 명칭 
	arrParam(1) = strTable	    					' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "관리항목값"					' 조건필드의 라벨 명칭 

	arrField(0) = strFld1	    			' Field명(0)
	arrField(1) = strFld2	    		' Field명(1)

	arrHeader(0) = "관리항목값"					' Header명(0)
	arrHeader(1) = "관리항목값명"

	
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCtrlPB(arrRet)
	End If	

End Function

Function SetCtrlPB(Byval arrRet)
	With frm1
		.vspdData2.Row =  .vspdData2.ActiveRow
		.vspdData2.Col =  C_CtrlVal
		.vspdData2.Text = arrRet(0)

		.vspdData2.Col =  C_CtrlValNm
		.vspdData2.Text = arrRet(1)
	End With

End Function

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
	
	If lgIntFlgMode <> OPMD_UMODE Then
	    Call SetToolbar("1110111100001111")										'⊙: 버튼 툴바 제어 
	Else    
	    Call SetToolbar("1111111100001111")										'⊙: 버튼 툴바 제어 
	End If
	    
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB2
	
	call SetSumItem()
	'Call SetToolBar()

End Function

'======================================================================================================
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'=======================================================================================================

'=======================================================================================================
'   Function Name : FindNumber
'   Function Desc : 
'=======================================================================================================
Function  FindNumber(ByVal objSpread, ByVal intCol)
Dim lngRows
Dim lngPrevNum
Dim lngNextNum

    FindNumber = 0

    lngPrevNum = 0
    lngNextNum = 0
    
    With frm1
        
        If objSpread.MaxRows = 0 Then
            Exit Function
        End If
        
        For lngRows = 1 To objSpread.MaxRows
            objSpread.Row = lngRows
            objSpread.Col = intCol
            lngNextNum = Clng(objSpread.Text)
            
            If lngNextNum > lngPrevNum Then
                lngPrevNum = lngNextNum
            End If
            
        Next
       
    End With        
    
    FindNumber = lngPrevNum
    
End Function
'=======================================================================================================
'   Function Name : CopyFromData
'   Function Desc : 
'=======================================================================================================
Function  CopyFromData(ByVal strItemSeq)
Dim lngRows
Dim boolExist
Dim iCols

    boolExist = False
    
    CopyFromData = boolExist
    
    With frm1

        Call SortHSheet()
                        
      '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = 1                

            If strItemSeq = .vspdData3.Text Then
                boolExist = True
                Exit For
            End If    
        Next
        
      '------------------------------------
        ' Show Data
        '------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            frm1.vspdData2.Redraw = False
            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows
                
                .vspdData3.Col = 1
                
                If strItemSeq <> .vspdData3.Text Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData2.MaxRows = .vspdData2.MaxRows + 1
                    .vspdData2.Row = .vspdData2.MaxRows
                    .vspdData2.Col = 0
                    .vspdData3.Col = 0
                    .vspdData2.Text = .vspdData3.Text
                  
                    For iCols = 1 To .vspdData3.MaxCols
                        .vspdData2.Col = iCols
                        .vspdData3.Col = iCols + 1
                        .vspdData2.Text = .vspdData3.Text
                    Next
                        
                End If   
                
                lngRows = lngRows + 1
                
            Wend
            
            ggoSpread.Source = frm1.vspdData2
            
            frm1.vspdData1.Row = lgCurrRow
            frm1.vspdData1.Col = frm1.vspdData1.MaxCols
            ggoSpread.Source = frm1.vspdData1
            
            frm1.vspdData2.Redraw = True
            
        End If
            
    End With        
    
    CopyFromData = boolExist
    
End Function

'=======================================================================================================
'   Function Name : CopyToHSheet
'   Function Desc : 
'=======================================================================================================
Sub  CopyToHSheet(ByVal Row)
Dim lRow
Dim iCols

	With frm1 
        
	    lRow = FindData

	    If lRow > 0 Then
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
        
            For iCols = 1 To .vspdData3.MaxCols
                .vspdData2.Col = iCols
                .vspdData3.Col = iCols + 1
                .vspdData3.Text = .vspdData2.Text
            Next
            
        End If

	End With
	
	'frm1.vspdData3.Row = 1
	'frm1.vspdData3.Col = 1
	
End Sub

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function  DeleteHSheet(ByVal strItemSeq)
Dim boolExist
Dim lngRows
 
    DeleteHSheet = False
    boolExist = False
    
    With frm1
    
        Call SortHSheet()
        
      '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = 1                

            If strItemSeq = .vspdData3.Text Then
                boolExist = True
                Exit For
            End If    
        Next
        
      '------------------------------------
        ' Data Delete
        '------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows
                .vspdData3.Col = 1
                
                If strItemSeq <> .vspdData3.Text Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Action = 5
                    .vspdData3.MaxRows = .vspdData3.MaxRows - 1
                End If   

            Wend
            
            ggoSpread.Source = frm1.vspdData2
            
            frm1.vspdData1.Row = lgCurrRow
            frm1.vspdData1.Col = frm1.vspdData1.MaxCols
            ggoSpread.Source = frm1.vspdData1
            
            frm1.vspdData2.Redraw = True
            
        End If
            
    End With
        
    DeleteHSheet = True
End Function    

'======================================================================================================
' Function Name : SortHSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function  SortHSheet()
    
    With frm1
        .vspdData3.BlockMode = True
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.SortBy = 0 'SS_SORT_BY_ROW
        
        .vspdData3.SortKey(1) = 1
        .vspdData3.SortKey(2) = 2
        
        .vspdData3.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(2) = 1 'SS_SORT_ORDER_ASCENDING
        
        .vspdData3.Col = 1
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 0
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.Action = 25 'SS_ACTION_SORT
        .vspdData3.BlockMode = False
    End With        
    
End Function

'=======================================================================================================
' Function Name : ShowHidden
' Function Desc : 
'=======================================================================================================
Sub  ShowHidden()
Dim strHidden
Dim lngRows
Dim lngCols
    
    With frm1.vspdData3
        For lngRows = 1 To .MaxRows
            .Row = lngRows
            For lngCols = 0 To .MaxCols
            .Col = lngCols  
                strHidden = strHidden & " | " & .Text
            Next
            strHidden = strHidden & vbCrLf
        Next
    End With        
    
'    msgbox strHidden    
End Sub

'======================================================================================================
' Function Name : SetSpreadFG
' Function Desc : This function set Muti spread Flag
'=======================================================================================================

Sub  SetSpreadFG( pobjSpread , ByVal pMaxRows )
    Dim lngRows 
    
    For lngRows = 1 To pMaxRows
        pobjSpread.Col = 0
        pobjSpread.Row = lngRows
        pobjSpread.Text = ""
    Next
    
End Sub

'======================================================================================================
' Function Name : SetSumItem
' Function Desc :
'=======================================================================================================
Function  SetSumItem()
    
    Dim DblTotClsAmt 
    Dim DblTotClsLocAmt 
    Dim DblTotDcLocAmt 
    Dim DblTotDcAmt 
    
    Dim lngRows 

	With frm1.vspdData 
	          
    If .MaxRows > 0 Then    
        For lngRows = 1 To .MaxRows
            .Row = lngRows
            .Col = C_ApClsAmt	'6
            If .Text = "" Then
                DblTotClsAmt = UniCDbl(DblTotClsAmt) + 0
            Else
                DblTotClsAmt = UniCDbl(DblTotClsAmt) + UniCDbl(.Text)
            End If
            
            .Col = C_ApClsLocAmt	'8
            If .Text = "" Then
                DblTotClsLocAmt = UniCDbl(DblTotClsLocAmt) + 0
            Else
                DblTotClsLocAmt = UniCDbl(DblTotClsLocAmt) + UniCDbl(.Text)
            End If                      
            
        Next 
    END IF     
    end with        
        
	frm1.txtPaymAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotClsAmt,gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")       
	frm1.txtPaymLocAmt.Text = 	UNIConvNumPCToCompanyByCurrency(DblTotClsLocAmt,gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")       	 
	
End Function

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
Sub  Form_Load()

    Call LoadInfTB19029                                                         'Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, gComNum1000, gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, gComNum1000, gComNumDec)    
                                     
                         
    Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
    Call InitSpreadSheet                                                        'Setup the Spread sheet
    Call InitVariables                                                          'Initializes local global variables
    Call SetDefaultVal
    Call InitComboBox
    
    Call SetToolbar("1110111100001111")										    '버튼 툴바 제어	
	frm1.txtAllcNo.focus
	   
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'=======================================================================================================
'   Event Name : txtInputType_Change()
'   Event Desc :  
'=======================================================================================================
Sub  txtInputType_onChange()
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    lgBlnFlgChgValue = True
	
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default

	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD = " & FilterVar(frm1.txtInputType.value, "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 			
			Select Case UCase(lgF0)
			Case "CS" & Chr(11)
				frm1.txtCheckCd.value   = ""
				frm1.txtBankCd.value   = ""
				frm1.txtBankNm.value   = ""
				frm1.txtBankAcct.value   = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
			Case "DP" & Chr(11)			' 예적금 
				frm1.txtCheckCd.value   = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
			Case "NO" & Chr(11)
				frm1.txtBankCd.value   = ""
				frm1.txtBankNm.value   = ""
				frm1.txtBankAcct.value   = ""				
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "N")	
			Case Else
				frm1.txtCheckCd.value   = ""
				frm1.txtBankCd.value   = ""
				frm1.txtBankNm.value   = ""
				frm1.txtBankAcct.value   = ""		
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
		End Select
	else
		frm1.txtCheckCd.value   = ""
		frm1.txtBankCd.value   = ""
		frm1.txtBankNm.value   = ""
		frm1.txtBankAcct.value   = ""		
		Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
		Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")	
	end if	
	

	'Select Case Ucase(frm1.txtInputType.value)
	'	Case "CS", "CK"	
	'		frm1.txtCheckCd.value   = ""
	'		frm1.txtBankCd.value   = ""
	'		frm1.txtBankNm.value   = ""
	'		frm1.txtBankAcct.value   = ""
	'		Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
	'		Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
	'		Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
	'	Case "DP" 			' 예적금 
	'		frm1.txtCheckCd.value   = ""
	'		Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
	'		Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
	'		Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
	'	Case "NP", "NR" 	' 지급어음/받을어음 
	'		frm1.txtBankCd.value   = ""
	'		frm1.txtBankNm.value   = ""
	'		frm1.txtBankAcct.value   = ""				
	'		Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
	'		Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
	'		Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "N")	
	'	Case Else
	'		frm1.txtCheckCd.value   = ""
	'		frm1.txtBankCd.value   = ""
	'		frm1.txtBankNm.value   = ""
	'		frm1.txtBankAcct.value   = ""		
	'		Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
	'		Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
	'		Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
	'End Select
	
	
End Sub

'======================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
Dim iFld1 
Dim istrCode

	'---------- Coding part -------------------------------------------------------------

	ggoSpread.Source = frm1.vspdData1

	With frm1.vspdData1
		If Row > 0 And Col = C_HQDeptPb Then
			
			.Row = Row
			.Col = Col - 1
			istrCode = .Text 

			.Col = C_BizCD
			iFld1 = .Text 
			
			Call OpenPopup(istrCode, 9, iFld1)
			
		ElseIF 	Row > 0 And Col = C_BizPb Then
			.Row = Row
			.Col = Col - 1
			
			istrCode = .Text 			
			
			Call OpenPopup(istrCode, 8, "")
		End If
		
	End With
	
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
    
End Sub


'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	
    gMouseClickStatus = "SP2C"	'Split 상태코드 
        
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData_EditChange(ByVal Col , ByVal Row )
    Dim DblNetAmt, DblVatAmt, DblNetLocAmt, DblVatLocAmt 

	With frm1.vspdData 

    End With
                
End Sub
'=======================================================================================================
'   Event Name : vspdData_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData_onfocus()
	gSelframeFlg = Tab1	
End Sub


'=======================================================================================================
'   Event Name : vspdData2_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_onfocus()
		gSelframeFlg = Tab2
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0             

End Sub

'======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_Change(ByVal Col, ByVal Row )

	Call CheckMinNumSpread(frm1.vspdData1, Col, Row)

    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
    
    frm1.vspdData1.Row = Row
    frm1.vspdData1.Col = 0             

End Sub

'======================================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspddata_KeyPress(KeyAscii )
     
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
Function  FncQuery() 
    Dim IntRetCD 
    Dim var1, var2
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData1
    var2 = ggoSpread.SSCheckChange
    
    
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then		
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO,"X","X")    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables
    frm1.vspdData.MaxRows = 0    
    frm1.vspdData1.MaxRows = 0        
  
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
           
    FncQuery = True																
   
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
    Dim var1, var2
	    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData1
    var2 = ggoSpread.SSCheckChange
  
  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then
        IntRetCD = DisplayMsgBox("900015", VB_YES_NO,"X","X")              
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables                                                      'Initializes local global variables
    
    frm1.vspdData.MaxRows = 0    
    frm1.vspdData1.MaxRows = 0    
    
    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus

    'SetGridFocus
    
    FncNew = True                                                          
    
End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncDelete() 
    Dim IntRetCD
    
    FncDelete = False                                                      
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                       
        
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", VB_YES_NO,"X","X")		            'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If					
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "1")                                  '⊙: Clear Condition Field
    'Call ggoOper.ClearField(Document, "2")  									'☜: Delete db data
    'frm1.vspdData.MaxRows = 0
    'frm1.vspdData1.MaxRows = 0
    
    FncDelete = True                                                        

End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
    Dim var1,var2
	
    FncSave = False                                                         
    
    Err.Clear                                                               

    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData1
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False And var2 = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
		Exit Function		
    End If    

    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                          'Check contents area
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Exit Function
    End If    
    
    ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Exit Function
    End If    

    If Not chkAllcDate() Then
		Exit Function
    End If  
    
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                             '☜: Save db data
    
    FncSave = True                                                       
End Function


Function chkAllcDate()
	Dim intI
	
	chkAllcDate = True
	With frm1
		For intI = 1 To .vspdData.Maxrows
			.vspdData.row = intI
			.vspdData.col = C_ApDt		'Const변수로 값이 설정되었는지 확인요.

			If CompareDateByFormat(.vspdData.Text,.txtAllcDt.Text,"채무일자",.txtAllcDt.Alt, _
		    	               "970025",.txtAllcDt.UserDefinedFormat,gComDateType, true) = False Then
			   .txtAllcDt.focus
			   chkAllcDate = False
			   Exit Function
			End If
		Next
	
	End With
End Function
'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
	Dim  IntRetCD
	
	If frm1.vspdData.Maxrows < 1 Then Exit Function 
	 
	frm1.vspdData.ReDraw = False
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", VB_YES_NO,"X","X")	'⊙: "Will you destory previous data"
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	
	With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData1	
		ggoSpread.CopyRow
		
		SetSpreadColor "I",1, frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow
    
		.vspdData.ReDraw = True
	End With
	
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 

	If gSelframeFlg = TAB1 Then
		if frm1.vspdData.Maxrows < 1 Then Exit Function
		With frm1.vspdData
		    .Row = .ActiveRow
		    .Col = 0
		    
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.EditUndo                     
		End With   
	Else
		if frm1.vspdData1.Maxrows < 1 Then Exit Function
		With frm1.vspdData1
		    .Row = .ActiveRow
		    .Col = 0
		    
		    ggoSpread.Source = frm1.vspdData1
		    ggoSpread.EditUndo                     
		End With   
	END IF
        
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow() 
	With frm1.vspdData1
		intItemCnt = .MaxRows
        	
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.InsertRow
		
		
		SetSpreadColor "I",1, frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow
		gSelframeFlg = Tab2
	 End With    
End Function
'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows 
    
	If gSelframeFlg = TAB1 Then
		if frm1.vspdData.Maxrows < 1 Then Exit Function
		ggoSpread.Source = frm1.vspdData
	else
		if frm1.vspdData1.Maxrows < 1 Then Exit Function
		ggoSpread.Source = frm1.vspdData1
	end if	
    lDelRows = ggoSpread.DeleteRow
    
End Function
'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next  
    parent.FncPrint()                                             
End Function
'=======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function  FncPrev() 
    On Error Resume Next                                               
End Function
'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function  FncNext() 
    On Error Resume Next                                               
End Function
'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function  FncFind() 
    Call parent.FncFind(C_SINGLEMULTI , True)                          
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function  FncExcel() 
	Call parent.FncExport(C_SINGLEMULTI)
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = 5
    
    If gMouseClickStatus = "SPCRP" Then
       ACol = Frm1.vspdData.ActiveCol
       ARow = Frm1.vspdData.ActiveRow

       If ACol > iColumnLimit Then
		  Frm1.vspdData.Col = iColumnLimit : frm1.vspdData.Row = 0  	 	
          iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
          Exit Function  
       End If   
    
       Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData.Col = ACol
       Frm1.vspdData.Row = ARow
    
       Frm1.vspdData.Action = 0    
    
       Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
    End If   

    If gMouseClickStatus = "SP2CRP" Then
       ACol = Frm1.vspdData1.ActiveCol
       ARow = Frm1.vspdData1.ActiveRow

       If ACol > iColumnLimit Then
		  Frm1.vspdData1.Col = iColumnLimit : frm1.vspdData1.Row = 0  	
          iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData1.Text), "X")
          Exit Function  
       End If   
    
       Frm1.vspdData1.ScrollBars = SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData1
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData1.Col = ACol
       Frm1.vspdData1.Row = ARow
    
       Frm1.vspdData1.Action = 0    
    
       Frm1.vspdData1.ScrollBars = SS_SCROLLBAR_BOTH
    End If   
End Function

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    If lgBlnFlgChgValue = True  Then   
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"	
		If IntRetCD = vbNo Then
			Exit Function
		End If		
    ELSE    
		ggoSpread.Source = frm1.vspdData    
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"	
			If IntRetCD = vbNo Then
				Exit Function
			End If		
		ELSE
			ggoSpread.Source = frm1.vspdData1        
			If ggoSpread.SSCheckChange = True Then
				IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"	
				If IntRetCD = vbNo Then
					Exit Function
				End If
			End If
		END IF
	END IF		
    
    FncExit = True
    
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function  DbDelete() 

    DbDelete = False														
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & UID_M0003
    strVal = strVal & "&txtAllcNo=" & Trim(frm1.txtAllcNo.value)				'☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         

End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()												        '삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables                                                      'Initializes local global variables
    
    frm1.vspdData.MaxRows = 0    
    frm1.vspdData1.MaxRows = 0    
    
    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus
    
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbQuery() 
    
    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    Dim strVal
    
    with frm1
        
		If lgIntFlgMode = OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & UID_M0001					'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.htxtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl			
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtMaxRows1=" & .vspdData1.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & UID_M0001					'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.txtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtMaxRows1=" & .vspdData1.MaxRows
		End If
    
    End With

	Call RunMyBizASP(MyBizASP, strVal)										    '☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                              
    
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk()
	
	With frm1
		.vspdData1.Col = 1:    intItemCnt = .vspddata1.MaxRows
	    SetSpreadLock "Q", 0, 1, ""
	    SetSpreadLock "Q", 1, 1, ""
	    SetSpreadColor "Q", 0, 1, ""
	    SetSpreadColor "Q", 1, 1, ""
    
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = OPMD_UMODE												'Indicates that current mode is Update mode
        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
        Call SetToolbar("1111111100001111")										'버튼 툴바 제어        
        
    End With
    
    Call txtInputType_onChange
    
    'SetGridFocus
    
    lgBlnFlgChgValue = False
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    Dim pAP010M 
    Dim lngRows 
    Dim lGrpcnt
    DIM strVal 
    Dim strDel

    DbSave = False                                                          
    Call LayerShowHide(1)
    'On Error Resume Next                                                   

    'Call SetSumItem

	With frm1
		
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = gUsrID
		.txtInsrtUserId.value  = gUsrID
		
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 

    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData	    
		For lngRows = 1 To .MaxRows
		    .Row = lngRows
			.Col = 0
				
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else

					strVal = strVal & "C" & gColSep  					'☜: C=Create, Row위치 정보 
			        .Col = C_ApNo								'1
			        strVal = strVal & Trim(.Text) & gColSep
			        .Col = C_AcctCd
			        strVal = strVal & Trim(.Text) & gColSep
			        .Col = C_ApDt
			        strVal = strVal & Trim(.Text) & gColSep
			        .Col = C_DocCur
			        strVal = strVal & Trim(.Text) & gColSep
			        .Col = C_ApClsAmt			        
			        strVal = strVal & Trim(.Text) & gColSep
			        .Col = C_ApClsLocAmt		            
			        strVal = strVal & Trim(.Text) & gRowSep
			            
			        lGrpCnt = lGrpCnt + 1	
			End Select		        
		Next
	End With	
	
	frm1.txtMaxRows.value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value =  strDel & strVal									'Spread Sheet 내용을 저장 
    
    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    ggoSpread.Source = frm1.vspdData1
	With frm1.vspdData1	    
		For lngRows = 1 To .MaxRows
		    .Row = lngRows
			.Col = 0
				
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else

					strVal = strVal & "C" & gColSep & lngRows & gColSep 					'☜: C=Create, Row위치 정보 
					.Col = C_BizCd								'1
			        strVal = strVal & Trim(.Text) & gColSep
			        .Col = C_HQDeptCd
			        strVal = strVal & Trim(.Text) & gColSep
			        .Col = C_HqAllcAmt
			        strVal = strVal & Trim(.Text) & gColSep
			        .Col = C_HqAllcLocAmt
			        strVal = strVal & Trim(.Text) & gRowSep
			            
			        lGrpCnt = lGrpCnt + 1	
			End Select		        
		Next
	End With	
	
    frm1.txtMaxRows1.value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread1.value =  strDel & strVal									'Spread Sheet 내용을 저장 
    
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'저장 비지니스 ASP 를 가동 
        
    DbSave = True                                                           
    
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk(ByVal AllcNo)													'☆: 저장 성공후 실행 로직 
   
    ggospread.SSDeleteFlag 1
    
    If lgIntFlgMode = OPMD_CMODE Then
		  frm1.txtAllcNo.value = AllcNo
	End If	  
	
	Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables
    frm1.vspdData.MaxRows = 0    
    frm1.vspdData1.MaxRows = 0 
    
    Dbquery()
	
End Function

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'=======================================================================================================
'   Event Name : txtAllcDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAllcDt.Action = 7                        
    End If
End Sub

'=======================================================================================================
'   Event Name : txtAllcDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
	
	Frm1.vspdData1.Row = 1
	Frm1.vspdData1.Col = 1
	Frm1.vspdData1.Action = 1	

End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<!--
 '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### --> 


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
			    <TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><IMG height=23 src="../../image/table/seltab_up_left.gif" width=9></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>본지점출금반제</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>								
					<TD WIDTH=* align=right><A href="vbscript:OpenRefOpenAp()">채무발생정보</A></TD>								
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">		
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>반제번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAllcNo" ALT="반제번호" MAXLENGTH=18 tag ="12XXXU"><IMG align=top name=btnCalType src="../../image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript: Call OpenPopup(frm1.txtAllcNo.value,0, '')"></TD>								
								</TR>						
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>		
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>출금일</TD>
								<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/a8105ma1_fpDateTime1_txtAllcDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="23NXXU" ALT="거래처"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBpCd.value,1, '')"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="거래처명"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10  tag=23NXXU" ALT="부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDeptCd.value,2, '')"> <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="부서명"></TD>
								<TD CLASS=TD5 NOWRAP>지급유형</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtInputType" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU" ><IMG SRC="../../image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtInputType.value, 10, '')">
													   <INPUT TYPE=TEXT NAME="txtInputTypeNm" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>																	   
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>은행</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBankCd" SIZE=10 MAXLENGTH=10 tag="21NXXU" ALT="은행"><IMG SRC="../../image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankCd.value,5, '')"> <INPUT TYPE=TEXT NAME="txtBankNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="은행명"></TD>											
								<TD CLASS=TD5 NOWRAP>어음번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCheckCd" SIZE=20 MAXLENGTH=20  tag="21NXXU" ALT="어음번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCheckCd.value,7, '')"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계좌번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT  TYPE=TEXT NAME="txtBankAcct" SIZE=30 MAXLENGTH=30  tag="21XXXU" ALT="계좌번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankAcct.value,6, '')"></TD>																						
								<TD CLASS=TD5 NOWRAP>전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=20  tag="24XXXU" ALT="전표번호"> </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="23NXXU"  ALT="거래통화"><IMG SRC="../../image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtDocCur.value,3, '')"></TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/a8105ma1_I168126285_txtXchRate.js'></script></TD>											
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>출금액</TD>
								<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/a8105ma1_I853062635_txtPaymAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>출금액(자국통화)</TD>
								<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/a8105ma1_I204228925_txtPaymLocAmt.js'></script></TD>
							</TR>												
							<TR HEIGHT="50%">
								<TD WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/a8105ma1_I351554030_vspdData.js'></script>
								</TD>											
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/a8105ma1_I750856792_vspdData1.js'></script>
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
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> SRC="../../blank.htm" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>	
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA><TEXTAREA class=hidden name=txtSpread1 tag="24"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread2 tag="24"></TEXTAREA><TEXTAREA class=hidden name=txtSpread3 tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24"><INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24"><INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24"><INPUT TYPE=hidden NAME="txtMaxRows1" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24"><INPUT TYPE=hidden NAME="htxtAllcNo" tag="24">
<INPUT TYPE=hidden NAME="hItemSeq" tag="24"><INPUT TYPE=hidden NAME="hAcctCd" tag="24"><INPUT TYPE=hidden NAME="txtMaxRows3" tag="24">
<script language =javascript src='./js/a8105ma1_I797864276_vspdData3.js'></script>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
