
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : F4212MA1
'*  4. Program Name         : 차입금현황조회 
'*  5. Program Desc         : Query of Loan State
'*  6. Comproxy List        : DB AGENT
'*  7. Modified date(First) : 2002.04.17
'*  8. Modified date(Last)  : 2003.05.19
'*  9. Modifier (First)     : Park, Joon Won
'* 10. Modifier (Last)      : Ahn, do hyun
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   ***************************************************
'	기능: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "f4212mb1_ko441.asp"                              '☆: Biz Logic ASP Name


'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 5					                          '☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                          

Dim lgSelectList                                         '☜: SpreadSheet의 초기  위치정보관련 변수                                     
Dim lgSelectListDT                                       '☜: SpreadSheet의 초기  위치정보관련 변수 


Dim lgSortFieldNm                                        '☜: Orderby popup용 데이타(필드설명)                       
Dim lgSortFieldCD                                        '☜: Orderby popup용 데이타(필드코드)      

Dim lgMaxFieldCount

Dim lgPopUpR                                             '☜: Orderby,Groupby default 값               
Dim lgKeyPos                                             '☜: Key위치 
Dim lgKeyPosVal                                          '☜: Key위치 Value 
Dim lgCookValue 

Dim IsOpenPop          

Dim lgSaveRow


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
  
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Function SetDefaultVal()
	Dim StartDate, FristDate

	StartDate	= "<%=GetSvrDate%>"
	FristDate	= UNIGetFirstDay("<%=GetSvrDate%>",parent.gServerDateFormat)

	frm1.txtLoanDtFr.Text  = UniConvDateAToB(FristDate, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtLoanDtTo.Text  = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
End Function

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","MA") %>

End Sub


'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1012", "''", "S") & "  AND MINOR_CD IN(" & FilterVar("U", "''", "S") & " ," & FilterVar("C", "''", "S") & " ) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboApSts ,lgF0  ,lgF1  ,Chr(11))


End Sub
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

'**********************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** 

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtLoanPlcCd.className) = "PROTECTED" Then Exit Function

	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanPlcCd.focus
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If

End Function
 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

   If IsOpenPop = True Then Exit Function
    Select case iWhere
		case 0
			If frm1.txtLoanPlcCd.className = parent.UCN_PROTECTED Then Exit Function	
			If frm1.txtLoanPlcfg1.Checked = true Then
				arrParam(0) = "은행팝업"
				arrParam(1) = "B_BANK A"
				arrParam(2) = strCode
				arrParam(3) = ""
				arrParam(4) = ""
				arrParam(5) = "은행코드"

				arrField(0) = "A.BANK_CD"
				arrField(1) = "A.BANK_NM"
						    
				arrHeader(0) = "은행코드"
				arrHeader(1) = "은행명"
			Else
				Call OpenBp(strCode, iWhere)
				exit function
			End If
        
        Case 1	
			arrParam(0) = "차입용도팝업"			' 팝업 명칭 
			arrParam(1) = "b_minor" 				    ' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "major_cd=" & FilterVar("f1000", "''", "S") & " "	        ' Where Condition
			arrParam(5) = "차입용도"				' 조건필드의 라벨 명칭 

			arrField(0) = "minor_cd"						' Field명(0)
			arrField(1) = "minor_nm"						' Field명(1)
    
			arrHeader(0) = frm1.txtLoanType.Alt				' Header명(0)
			arrHeader(1) = frm1.txtLoanTypeNm.Alt				    ' Header명(1)
		Case 2
			arrParam(0) = "거래통화팝업"								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' 조건필드의 라벨 명칭 

		    arrField(0) = "CURRENCY"										' Field명(0)
		    arrField(1) = "CURRENCY_DESC"									' Field명(1)

		    arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화코드명"									' Header명(1)


		case 3,4

			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition

			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "사업장 코드"			

			arrField(0) = "BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "BIZ_AREA_NM"					' Field명(1)

			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)

		Case Else
			Exit Function
	End Select
        
    IsOpenPop = True    
        
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' 거래처 
				frm1.txtLoanPlcCd.focus
			Case 1		'차입용도 
				frm1.txtLoanType.focus
			Case 2
				frm1.txtDocCur.focus
			Case 3
				frm1.txtBizAreaCd.focus
			Case 4
				frm1.txtBizAreaCd1.focus
		End Select
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetReturnPopUp()  --------------------------------------------------
'	Name : SetReturnPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnPopUp(Byval arrRet, Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' 거래처 
				frm1.txtLoanPlcCd.value		= arrRet(0)
				frm1.txtLoanPlcNm.value		= arrRet(1)
				frm1.txtLoanPlcCd.Focus
			Case 1		'차입용도 
				frm1.txtLoanType.value		= arrRet(0)
				frm1.txtLoanTypeNm.value	= arrRet(1)
				frm1.txtLoanType.focus
			Case 2
				frm1.txtDocCur.value		= arrRet(0)
				frm1.txtDocCur.focus
			Case 3
				frm1.txtBizAreaCd.Value		= arrRet(0)
				frm1.txtBizAreaNm.Value		= arrRet(1)
				frm1.txtBizAreaCd.focus
			Case 4
				frm1.txtBizAreaCd1.Value	= arrRet(0)
				frm1.txtBizAreaNm1.Value	= arrRet(1)
				frm1.txtBizAreaCd1.focus
				
				
		End Select

	End With
	
End Function


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("F4212MA1","S","A","V200302021",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")

End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
        With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.ReDraw = True
        End With
    End if
End Sub


'===========================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function



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
'   Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
'	Call initMinor()
End Sub


'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   

    'lgMaxFieldCount =  UBound(parent.gFieldNM)                      



    'Call parent.MakePopData(parent.gDefaultT,parent.gFieldNM,parent.gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,parent.C_MaxSelList)
	

	
	Call InitVariables                                                     '⊙: Initializes local global variables													
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call InitComboBox	
	Call txtLoanPlcfg_onchange()
	Call FncSetToolBar("New")
	frm1.txtLoanDtFr.focus 
	Set gActiveElement = document.activeElement 

	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing
		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
													
	
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD  

    FncQuery = False                                                            '⊙: Processing is NG                                        
    
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")                                       '⊙: Clear Contents  Field 					
    Call InitVariables                                                           '⊙: Initializes local global variables											
    '-----------------------
    'Check condition area
    '-----------------------
       
    If Not chkField(Document, "1") Then	                                         '⊙: This function check indispensable field  						
       Exit Function
    End If


    '네패스 요청으로 필수항목에서 선택항목으로 변경...20090901...kbs	
    If frm1.txtLoanDtFr.Text  = "" Then
	frm1.txtLoanDtFr.Text = "2001-01-01"
    End If

    If frm1.txtLoanDtTo.Text  = "" Then
	frm1.txtLoanDtTo.Text = "2999-12-31"
    End If

    
    If (frm1.txtPaymDtFr.Text <> "") And (frm1.txtPaymDtTo.Text <> "") Then
		If CompareDateByFormat(frm1.txtPaymDtFr.Text, frm1.txtPaymDtTo.Text, frm1.txtPaymDtFr.Alt, frm1.txtPaymDtTo.Alt, _
					"970025", frm1.txtPaymDtFr.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtPaymDtFr.focus											
			Exit Function
		End if	
	End If

    If (frm1.txtLoanDtFr.Text <> "") And (frm1.txtLoanDtTo.Text <> "") Then
		If CompareDateByFormat(frm1.txtLoanDtFr.Text, frm1.txtLoanDtTo.Text, frm1.txtLoanDtFr.Alt, frm1.txtLoanDtTo.Alt, _
					"970025", frm1.txtLoanDtFr.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtLoanDtFr.focus											
			Exit Function
		End if	
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If UCase(Trim(frm1.txtBizAreaCd.value)) > UCase(Trim(frm1.txtBizAreaCd1.value)) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
	if frm1.txtBizAreaCd.value <> "" then
	  If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  	Call DisplayMsgBox("970000","X",frm1.txtBizAreaCd.alt,"X")            '☜ : No data is found. 
	  	frm1.txtBizAreaNm.value = ""
	  	frm1.txtBizAreaCd.focus
 	  	Exit Function
	  End If
	End If
	  
	if frm1.txtBizAreaCd1.value <> "" then
	  If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd1.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  	Call DisplayMsgBox("970000","X",frm1.txtBizAreaCd1.alt,"X")            '☜ : No data is found.
	  	frm1.txtBizAreaNm1.value = ""
	  	frm1.txtBizAreaCd1.focus
 	  	Exit Function
	  End If
	End If
	
	Call FncSetToolBar("New")
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()															'☜: Query db data

    FncQuery = True													

End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal
	Dim txtLoanPlcfg
    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
    
	If frm1.txtLoanPlcfg1.checked Then
		txtLoanPlcfg = frm1.txtLoanPlcfg1.value
	ElseIf frm1.txtLoanPlcfg2.checked Then
		txtLoanPlcfg = frm1.txtLoanPlcfg2.value
	End if

    With frm1
		If lgIntFlgMode <> parent.OPMD_UMODE Then										'This means that it is first search

        	strVal = BIZ_PGM_ID & "?cboLoanFg="		& Trim(.cboLoanFg.value)
			strVal = strVal & "&cboConfFg="			& Trim(.cboConfFg.value)
			strVal = strVal & "&cboApSts="			& Trim(.cboApSts.value)
			strVal = strVal & "&txtDocCur="			& Trim(.txtDocCur.value)
			strVal = strVal & "&txtLoanPlcFg="		& Trim(txtLoanPlcFg)
			strVal = strVal & "&txtLoanPlcCd="		& Trim(.txtLoanPlcCd.value)
			strVal = strVal & "&txtLoanType="		& Trim(.txtLoanType.value)
			strVal = strVal & "&txtLoanDtFr="		& Trim(.txtLoanDtFr.Text)   
			strVal = strVal & "&txtLoanDtTo="		& Trim(.txtLoanDtTo.Text)
			strVal = strVal & "&txtPaymDtFr="		& Trim(.txtPaymDtFr.Text)
			strVal = strVal & "&txtPaymDtTo="		& Trim(.txtPaymDtTo.Text)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
		Else
        	strVal = BIZ_PGM_ID & "?cboLoanFg="		& Trim(.hLoanFg.value)
			strVal = strVal & "&cboConfFg="			& Trim(.hConfFg.value)
			strVal = strVal & "&cboApSts="			& Trim(.hApSts.value)
			strVal = strVal & "&txtDocCur="			& Trim(.hDocCur.value)
			strVal = strVal & "&txtLoanPlcFg="		& Trim(.hLoanPlcFg.value)
			strVal = strVal & "&txtLoanPlcCd="		& Trim(.hLoanPlcCd.value)
			strVal = strVal & "&txtLoanType="		& Trim(.hLoanType.value)
			strVal = strVal & "&txtLoanDtFr="		& Trim(.hLoanDtFr.value)   
			strVal = strVal & "&txtLoanDtTo="		& Trim(.hLoanDtTo.value)
			strVal = strVal & "&txtPaymDtFr="		& Trim(.hPaymDtFr.value)
			strVal = strVal & "&txtPaymDtTo="		& Trim(.hPaymDtTo.value)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.htxtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.htxtBizAreaCd1.value)
		End if
			
 '          Call CopyPopupInfABT()
            strVal = strVal	& "&lgPageNo="			& lgPageNo                          '☜: Next key tag
 			strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))

			' 권한관리 추가 
			strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
			strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
			strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
			strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												


    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call FncSetToolBar("Query")
	Call txtLoanPlcfg_onchange()
	Call CurFormatNumericOCX()
	
	'SetGridFocus
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtLoanDtFr.focus
	End If
	Set gActiveElement = document.activeElement 

End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	Dim intRetCD
	Dim strByCurrency
	With frm1

		If Trim(.txtDocCur.value) = "" Then
            intRetCD = CommonQueryRs("top 1 currency"," b_numeric_format "," decimals  = (select max(decimals) from b_numeric_format where data_type = 2 ) and data_type=2 and form_type = " & FilterVar("Q", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            
			If intRetCD = True Then	
				strByCurrency = Trim(Replace(lgF0,Chr(11),""))
			Else
				strByCurrency = parent.gCurrency
			End If
			ggoOper.FormatFieldByObjectOfCur .txtLoan,	strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtInt,	strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtRdp,	strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBal,	strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		Else
			ggoOper.FormatFieldByObjectOfCur .txtLoan,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtInt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtRdp,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBal,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		End If
	End With

End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'========================================================================================================
Function OpenGroupPopup()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(parent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = parent.gMethodText
  
	For ii = 0 to parent.C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
      
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(lgSortFieldCD,lgSortFieldNm,arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to parent.C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function

'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
End Function



'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
Sub txtLoanDtFr_DblClick(Button)
	if Button = 1 then
		frm1.txtLoanDtFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoanDtFr.Focus
	End if
End Sub

Sub txtLoanDtTo_DblClick(Button)
	if Button = 1 then
		frm1.txtLoanDtTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoanDtTo.Focus
	End if
End Sub

Sub txtPaymDtFr_DblClick(Button)
	if Button = 1 then
		frm1.txtPaymDtFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPaymDtFr.Focus
	End if
End Sub

Sub txtPaymDtTo_DblClick(Button)
	if Button = 1 then
		frm1.txtPaymDtTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPaymDtTo.Focus
	End if
End Sub

Sub txtLoanDtFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtLoanDtTo.focus
	   Call MainQuery
	End If   
End Sub

Sub txtLoanDtTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtLoanDtFr.focus
	   Call MainQuery
	End If   
End Sub

Sub txtPaymDtFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtLoanDtFr.focus
	   Call MainQuery
	End If   
End Sub

Sub txtPaymDtTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtLoanDtFr.focus
	   Call MainQuery
	End If   
End Sub

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    If frm1.vspdData.MaxRows = 0 then
        Exit Sub
    End If
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
        Exit Sub
    End If
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgCookValue = ""
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub


'======================================================================================================
'   Event Name : txtLoanPlcfg_onchange
'   Event Desc : 
'=======================================================================================================
Function txtLoanPlcfg_onchange()
	If frm1.txtLoanPlcfg0.checked = true then
		Call ggoOper.SetReqAttr(frm1.txtLoanPlcCd, "Q")
		frm1.txtLoanPlcCd.value = ""
		frm1.txtLoanPlcNm.value = ""
	Else
		Call ggoOper.SetReqAttr(frm1.txtLoanPlcCd, "D")
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
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

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>차입일자</TD>  
									<!--네패스 요청으로 필수항목에서 선택항목으로 변경...20090901...kbs	-->
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanDtFr name=txtLoanDtFr CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="시작차입일자"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanDtTo name=txtLoanDtTo CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="종료차입일자"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>상환만기일자</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpPaymDtFr name=txtPaymDtFr CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="시작만기일자"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpPaymDtTo name=txtPaymDtTo CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="종료만기일자"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCd.Value, 3)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
									<TD CLASS="TD5" NOWRAP>장단기구분</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLoanFg" ALT="장단기구분" STYLE="WIDTH: 135px" tag="11"><OPTION VALUE=""></OPTION></SELECT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCd1.Value, 4)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>차입용도</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoanType" MAXLENGTH="18" SIZE=10  ALT ="차입용도코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanType.Value, 1)">
									                     <INPUT NAME="txtLoanTypeNm" MAXLENGTH="40" SIZE=20  ALT ="차입용도명" tag="14"></TD>  
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입처구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg0 VALUE="BK" Checked tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg0>은행+거래처</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg1 VALUE="BK" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg1>은행</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg2 VALUE="BP" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg2>거래처</LABEL></TD>
									<TD CLASS="TD5" NOWRAP>차입처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanPlcCd" ALT="차입처" SIZE="10" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanPlcCd.Value, 0)">
															<INPUT NAME="txtLoanPlcNm" ALT="차입처명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래통화</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.Value, 2)">
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<!--TR>
									<TD CLASS=TD5 NOWRAP>차입은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10   tag="11" ALT="차입은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 0)">
														 <INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="14X" ALT="차입은행명"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>					 
								</TR-->
								<TR>
									<TD CLASS="TD5">승인상태</TD>
									<TD CLASS="TD6"><SELECT ID="cboConfFg" NAME="cboConfFg" ALT="승인상태" STYLE="WIDTH: 135px" tag="1XN"><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS="TD5">진행상황</TD>
									<TD CLASS="TD6"><SELECT ID="cboApSts" NAME="cboApSts" ALT="진행상황" STYLE="WIDTH: 135px" tag="1XN"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=* WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
				  <TR>
					<TD HEIGHT=40 WIDTH=25%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입총액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtLoan" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="차입총액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtLoanLoc" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="차입총액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>이자지급총액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtInt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="이자지급총액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtIntLoc" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="이자지급총액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>원금상환총액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtRdp" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="원금상환총액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtRdpLoc" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="원금상환총액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>차입잔액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtBal" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="차입잔액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtBalLoc" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="차입잔액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT>
									</TD>
								</TR>

							</TABLE>
						</FIELDSET>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=bizsize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hLoanFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hConfFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hApSts" tag="24">

<INPUT TYPE=HIDDEN NAME="hDocCur" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanPlcFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanPlcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanType" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanDtFr" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="hPaymDtFr" tag="24">
<INPUT TYPE=HIDDEN NAME="hPaymDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>

