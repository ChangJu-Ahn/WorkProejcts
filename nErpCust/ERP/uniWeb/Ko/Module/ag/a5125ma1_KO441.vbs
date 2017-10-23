'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID			= "a5125mb1_ko441.asp"
Const JUMP_PGM_ID_TAX_REP	= "a6114ma1"

'=                       4.2 Constant variables 
'========================================================================================================
Const C_GLINPUTTYPE = "TG"
Const MENU_NEW	=	"1110010000011111"
Const MENU_CRT	=	"1110111100111111"
Const MENU_UPD	=	"1111111100111111"
Const MENU_PRT	=	"1110000000011111"		

'=                       4.3 Common variables 
'========================================================================================================

'=                       4.4 User-defind Variables
'========================================================================================================

'⊙: Grid Columns
Dim  C_ItemSeq
Dim  C_deptcd
Dim  C_deptPopup
Dim  C_deptnm
Dim  C_AcctCd
Dim  C_AcctPopup
Dim  C_AcctNm
Dim  C_DrCrFg
Dim  C_DrCrNm
Dim  C_DocCur
Dim  C_DocCurPopup
Dim  C_ExchRate	
Dim  C_ItemAmt
Dim  C_ItemLocAmt
Dim  C_IsLAmtChange
Dim  C_ItemDesc	
Dim  C_VatType
Dim  C_VatNm
Dim  C_AcctCd2

Dim lgCurrRow
Dim lgStrPrevKeyDtl
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgBlnExecDelete
Dim lgFormLoad
Dim lgQueryOk
Dim lgstartfnc
Dim lgTempRate
Dim intItemCnt		
Dim IsOpenPop

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################
Sub initSpreadPosVariables()
	C_ItemSeq		= 1 
	C_deptcd		= 2 
	C_deptPopup		= 3 
	C_deptnm		= 4	
	C_AcctCd		= 5 
	C_AcctPopup		= 6 
	C_AcctNm		= 7 
	C_DrCrFg		= 8 
	C_DrCrNm		= 9 
	C_DocCur		= 10
	C_DocCurPopup	= 11
	C_ExchRate		= 12
	C_ItemAmt		= 13
	C_ItemLocAmt	= 14
	C_IsLAmtChange	= 15
	C_ItemDesc		= 16
	C_VatType		= 17
	C_VatNm			= 18
	C_AcctCd2		= 19
End Sub

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0

    frm1.txtTempGlNo.focus 
    Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
    frm1.txttempGLDt.text = UniConvDateAToB(iDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)
    frm1.hCongFg.value = ""

    frm1.cboConfFg.value = "U"
'-- eWare Inf Begin 
	frm1.cboAppFg.value  = "R" 
'-- eWare Inf End           
    frm1.cboGlType.value = "03"

    frm1.txtCommAndMode.value = "CREATE"
    frm1.cboGlInputType.value = C_GLINPUTTYPE

	frm1.txtDeptCd.value		= parent.gDepart
	frm1.hOrgChangeId.value 	= parent.gChangeOrgId
	
	
	IF CommonQueryRs( "ORG_NM" , "Z_USR_ORG_MAST  " , "ORG_CD = " & FilterVar(parent.gDepart, "''", "S") & " AND ORG_TYPE = 'DP' AND USE_YN = 'Y'", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		arrVal = Split(lgF0, Chr(11))
		frm1.txtLoginDeptNm.value = arrVal(0)	'기표부서(로그인부서)>>air
	End If	


    frm1.vspdData3.MaxCols = 16

	Call GetCheckAcct

    lgBlnFlgChgValue = False
End Sub

'========================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
    
	With frm1.vspdData
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.Spreadinit "V20030218",,parent.gAllowDragDropSpread    
	
		.MaxCols = C_AcctCd2 + 1
		.Col = .MaxCols
		.ColHidden = True
		.MaxRows = 0

		.ReDraw = False

        Call GetSpreadColumnPos("A")
'		Call AppEndNumberPlace("6","3","0")
        ggoSpread.SSSetFloat  C_ItemSeq,    " ", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_deptcd,     "귀속/예산부서",   10, , , 10, 2
        ggoSpread.SSSetButTon C_deptpopup
        ggoSpread.SSSetEdit   C_deptnm,     "부서명",     17, , , 30
		ggoSpread.SSSetEdit   C_AcctCd,     "계정코드", 15, , , 18
		ggoSpread.SSSetButTon C_AcctPopup
		ggoSpread.SSSetEdit   C_AcctNm,     "계정코드명", 20, , , 30
		ggoSpread.SSSetCombo  C_DrCrFg,     " ", 8
	    ggoSpread.SSSetCombo  C_DrCrNm,     "차대구분", 10
		ggoSpread.SSSetEdit   C_DocCur,     "거래통화",   10, , , 10, 2
        ggoSpread.SSSetButTon C_DocCurPopup
		ggoSpread.SSSetFloat  C_ExchRate,   "환율",		15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_ItemAmt,    "금액",       15, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt, "금액(자국)",	 15, parent.ggAmTofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_IsLAmtChange,   "",     30, , , 128
		ggoSpread.SSSetEdit   C_ItemDesc,   "비  고", 30, , , 128
		ggoSpread.SSSetCombo  C_VATTYPE,     "", 8
	    ggoSpread.SSSetCombo  C_VATNM,     "계산서유형",   20
		ggoSpread.SSSetEdit   C_AcctCd2,   "",     30, , , 128

		Call ggoSpread.MakePairsColumn(C_deptcd,C_deptpopup)
		Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPopup)
		Call ggoSpread.MakePairsColumn(C_DocCur,C_DocCurPopup)
		Call ggoSpread.MakePairsColumn(C_DrCrFg,C_DrCrNm,"1")
		Call ggoSpread.MakePairsColumn(C_VATTYPE,C_VATNM,"1")

		Call ggoSpread.SSSetColHidden(C_ItemSeq,C_ItemSeq,True)
		Call ggoSpread.SSSetColHidden(C_DrCrFg,C_DrCrFg,True)
		Call ggoSpread.SSSetColHidden(C_VatType,C_VatType,True)
		Call ggoSpread.SSSetColHidden(C_VATNM,C_VATNM,True)		
		Call ggoSpread.SSSetColHidden(C_IsLAmtChange,C_IsLAmtChange,True)
		Call ggoSpread.SSSetColHidden(C_AcctCd2,C_AcctCd2,True)
		
		
		.ReDraw = True                
    End With
    
    SetSpreadLock "I", 0, 1, ""
End Sub

'=======================================================================================================
Sub SetSpreadLock(ByVal stsFg, ByVal Index, ByVal lRow  , ByVal lRow2 )
    With frm1
		ggoSpread.Source = .vspdData
		lRow2 = .vspdData.MaxRows
		.vspdData.Redraw = False

		Select Case Index
			Case 0
				ggoSpread.SpreadUnLock		C_deptcd		, -1    , C_deptcd
				ggoSpread.SSSetRequired		C_deptcd		, -1    , C_deptcd
				ggoSpread.SpreadUnLock		C_deptpopup		, -1    , C_deptpopup
				ggoSpread.SpreadLock		C_deptnm		, -1    , C_deptnm
				ggoSpread.SpreadLock		C_AcctCd		, -1    , C_AcctCd
				ggoSpread.SpreadLock		C_AcctPopup		, -1    , C_AcctPopup
				ggoSpread.SpreadLock		C_AcctNm		, -1    , C_AcctNm
				ggoSpread.SpreadUnLock		C_DrCrNm		, -1    , C_DrCrNm
				ggoSpread.SSSetRequired		C_DrCrNm		, -1    , C_DrCrNm
				ggoSpread.SpreadUnLock		C_DocCur		, -1    , C_DocCur
				ggoSpread.SSSetRequired		C_DocCur		, -1    , C_DocCur
				ggoSpread.SpreadUnLock		C_DocCurPopup	, -1    , C_DocCurPopup
				ggoSpread.SpreadUnLock		C_ExchRate		, -1    , C_ExchRate
				ggoSpread.SpreadUnLock		C_ItemAmt		, -1    , C_ItemAmt
				ggoSpread.SSSetRequired		C_ItemAmt		, -1    , C_ItemAmt
				ggoSpread.SpreadUnLock		C_ItemLocAmt	, -1    , C_ItemLocAmt
				ggoSpread.SpreadUnLock		C_ItemDesc		, -1    , C_ItemDesc
				ggoSpread.SSSetRequired		C_ItemDesc		, -1    , C_ItemDesc
				ggoSpread.SpreadUnLock		C_VATNM			, -1    , C_VATNM
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			Case 1
				ggoSpread.SpreadLock C_deptcd		, -1    , C_deptcd
				ggoSpread.SpreadLock C_ItemSeq		, -1	, C_ItemSeq 
				ggoSpread.SpreadLock C_deptpopup	, -1	, C_deptpopup
				ggoSpread.SpreadLock C_ItemLocAmt	, -1	, C_ItemLocAmt
				ggoSpread.SpreadLock C_ItemDesc		, -1	, C_ItemDesc
				ggoSpread.SpreadLock C_AcctPopup	, -1	, C_AcctPopup
				ggoSpread.SpreadLock C_DrCrNm		, -1	, C_DrCrNm
				ggoSpread.SpreadLock C_DocCur		, -1	, C_DocCur
				ggoSpread.SpreadLock C_DocCurPopup	, -1	, C_DocCurPopup
				ggoSpread.SpreadLock C_ExchRate		, -1	, C_ExchRate
				ggoSpread.SpreadLock C_ItemAmt		, -1	, C_ItemAmt
				ggoSpread.SpreadLock C_VATNM		, -1	, C_VATNM
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		End Select

		.vspdData.Redraw = True
    End With
End Sub

'=======================================================================================================
Sub SetSpread2Lock(ByVal stsFg, ByVal Index, ByVal lRow  , ByVal lRow2 )
    With frm1
		ggoSpread.Source = .vspdData2
		lRow2 = .vspdData2.MaxRows
		.vspdData2.Redraw = False

		Select Case Index
			Case 0
			Case 1
				ggoSpread.SpreadLock 1, lRow, .vspdData2.MaxCols, lRow2	
		End Select

		.vspdData2.Redraw = True
	End With
End Sub

'=======================================================================================================
Sub SetSpreadColor(ByVal stsFg, ByVal Index, ByVal lRow, ByVal lRow2)
    With frm1
		If  lRow2 = "" Then	lRow2 = lRow
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_ItemSeq	, lRow, lRow2
		ggoSpread.SSSetProtected C_deptNm	, lRow, lRow2
		ggoSpread.SSSetProtected C_AcctNm	, lRow, lRow2
		ggoSpread.SSSetRequired  C_deptcd	, lRow, lRow2
		
		Select Case stsFg
			Case "I"
				ggoSpread.SSSetRequired C_AcctCd, lRow, lRow2
			Case "Q"
				ggoSpread.SSSetProtected C_AcctCd, lRow, lRow2
		End Select

		If  frm1.cboGlType.value = "01" Or frm1.cboGlType.value = "02" Then
			ggoSpread.SSSetProtected C_DrCrNm, lRow, lRow2
		ELSE
			ggoSpread.SSSetRequired C_DrCrNm, lRow, lRow2
		End If
		
		ggoSpread.SSSetRequired C_DocCur, lRow, lRow2		
		ggoSpread.SSSetRequired C_ItemAmt, lRow, lRow2
		ggoSpread.SSSetRequired C_ItemDesc, lRow, lRow2	'>>air	
		.vspdData.ReDraw = True
    End With
End Sub

'============================================================================================================
Function InitComboBoxGrid()
    ggoSpread.Source = frm1.vspdData

	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1012", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1

	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DrCrFg
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DrCrNm

    Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("B9001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_VatType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_VatNm
End Function

'============================================================================================================
Function OpenPopUp(ByVal strCode, ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 2
			arrParam(0) = "통화코드 팝업"	
			arrParam(1) = "B_Currency"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "통화코드"

			arrField(0) = "Currency"
			arrField(1) = "Currency_desc"

			arrHeader(0) = "통화코드"	
			arrHeader(1) = "통화코드명"
			
'			arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=B_CURRENCY_00", Array(Array(Trim(strCode))), _
'			           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	    	    			
			
		Case 3
			arrParam(0) = "계정코드팝업"
			arrParam(1) = "A_Acct, A_ACCT_GP"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD And A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "
			arrParam(5) = "계정코드"

			arrField(0) = "A_ACCT.Acct_CD"
			arrField(1) = "A_ACCT.Acct_NM"
    		arrField(2) = "A_ACCT_GP.GP_CD"
			arrField(3) = "A_ACCT_GP.GP_NM"

			arrHeader(0) = "계정코드"
			arrHeader(1) = "계정코드명"
			arrHeader(2) = "그룹코드"	
			arrHeader(3) = "그룹명"

'			arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=A_ACCT_00", Array(Array(Trim(strCode))), _
'			           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	    
	End Select

    If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	ElseIf iWhere = 3 Then
		arrRet = window.showModalDialog("../../comasp/a5101ma1_ko441_Popup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	
	Else		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	IsOpenPop = False

	If arrRet(0) <> "" Then     
		Call SetPopup(arrRet, iWhere)	
	End If

	Call FocusAfterPopup(iWhere)
End Function

'========================================================================================================= 
Function SetPopUp(ByRef arrRet, ByVal iWhere)
	With frm1
		Select Case iWhere	
			Case 2
				.vspdData.Row = .vspdData.ActiveRow 
				ggoSpread.Source = .vspdData
				ggoSpread.UpdateRow .vspdData.ActiveRow 
				.vspdData.Col  = C_ItemLocAmt
				.vspdData.Text = ""				
				.vspdData.Col  = C_DocCur
				.vspdData.Text = UCase(Trim(arrRet(0)))
				If Trim(.vspdData.Text) = parent.gCurrency Then
					.vspdData.Col  = C_ExchRate
					.vspdData.Text = 1
				Else
					Call FindExchRate(UniConvDateToYYYYMMDD(.txttempGLDt.text,parent.gDateFormat,""), UCase(Trim(arrRet(0))),.vspdData.ActiveRow)
				End If
				
				Call DocCur_OnChange(.vspdData.ActiveRow,.vspdData.ActiveRow)
			Case 3
				.vspdData.Row  = .vspdData.ActiveRow
				.vspdData.Col  = C_AcctCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_AcctNm
				.vspdData.Text = arrRet(1)
                Call vspdData_Change(C_AcctCd, .vspddata.activerow)
		End Select
	End With
End Function

'=======================================================================================================
Function FocusAfterPopup(ByVal iWhere)
	With frm1
		Select Case iWhere
			Case 2 
				Call SetActiveCell(.vspdData,C_DocCur,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				Call SetActiveCell(.vspdData,C_AcctCD,.vspdData.ActiveRow ,"M","X","X")
		End Select    
	End With
End Function

'========================================================================================================= 
Function OpenReftempgl()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(8)	                           '권한관리 추가 (3 -> 4)

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5101ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INForMATION, "a5101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	Call CookiePage("TEMP_GL_POPUP")

	arrParam(4)	= lgAuthorityFlag              '권한관리 추가	
	
	' 권한관리 추가
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = ""  Then
		frm1.txttempGlNo.focus
		Exit Function
	Else
		Call SetRefTempGl(arrRet)
	End If
End Function

'========================================================================================================= 
Function SetRefTempGl(ByRef arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	

	With frm1
		.txttempGlNo.value = UCase(Trim(arrRet(0)))
		.txttempGlNo.focus
    End With
	
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================= 
Function OpenDept(ByVal strCode, ByVal iWhere)
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtDeptCd.readOnly = true Then
		IsOpenPop = False
		Exit Function
	End If
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INForMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = strCode									'  Code Condition
   	arrParam(1) = frm1.txtTempGLDt.Text
	arrParam(2) = lgUsrIntCd								' 자료권한 Condition  

	If lgIntFlgMode = parent.OPMD_UMODE Then
		arrParam(3) = "T"									' 결의일자 상태 Condition  
	Else
		arrParam(3) = "F"									' 결의일자 상태 Condition  
	End If

	' 권한관리 추가
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetDept(arrRet, iWhere)
	End If

	Call FocusAfterDeptPopup (  iWhere)
End Function

'========================================================================================================= 
Function OpenUnderDept(ByVal strCode, ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim field_fg

	IsOpenPop = True
	If RTrim(LTrim(frm1.txtDeptCd.value)) <> "" Then
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " And A.COST_CD = B.COST_CD And B.BIZ_AREA_CD = ( Select B.BIZ_AREA_CD"
		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " And A.COST_CD = B.COST_CD And A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ")"
		arrParam(5) = "부서코드"

		arrField(0) = "A.DEPT_CD"
		arrField(1) = "A.DEPT_Nm"
		arrField(2) = "B.BIZ_AREA_CD"

		arrHeader(0) = "부서코드"
		arrHeader(1) = "부서코드명"
		arrHeader(2) = "사업장코드"
		
'		arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=B_ACCT_DEPT_01", _
'		   Array(Array(Trim(strCode)),Array("3",Trim(frm1.hOrgChangeId.value),Trim(frm1.txtDeptCd.value),Trim(frm1.hOrgChangeId.value))), _
'		   "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	 		
	Else
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A"
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID = (Select distinct org_change_id"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt = ( Select max(org_change_dt)"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		arrParam(5) = "부서코드"
		arrField(0) = "A.DEPT_CD"
		arrField(1) = "A.DEPT_Nm"
		arrHeader(0) = "부서코드"
		arrHeader(1) = "부서코드명"

'		arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=B_ACCT_DEPT_00", Array(Array(Trim(strCode))), _
'		           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	 
	End If

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetDept(arrRet, iWhere)
	End If

	Call FocusAfterDeptPopup(iWhere)
End Function

'========================================================================================================= 
Function SetDept(ByRef arrRet, ByVal iWhere)
	With frm1
		Select Case iWhere
		     Case "0"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtInternalCd.value = arrRet(2)
				If lgQueryOk <> True Then
					.txtTempGLDt.text = arrRet(3)
				Else 
				End If
				call txtDeptCd_OnChange()
             Case "1"
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
				.vspdData.Col  = C_deptcd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_deptnm
				.vspdData.Text = arrRet(1)

				Call deptCd_underChange(arrRet(0))
             Case Else
        End Select
	End With
End Function

'=======================================================================================================
Function FocusAfterDeptPopup(ByVal iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtDeptCd.focus
			Case 1 
				Call SetActiveCell(.vspdData,C_deptcd,.vspdData.ActiveRow ,"M","X","X")
		End Select    
	End With
End Function

'=======================================================================================================
Function SetSumItem()
    Dim DblTotDrAmt 
    DIm DblTotLocDrAmt
    Dim DblTotCrAmt 
    DIm DblTotLocCrAmt
    Dim lngRows 

	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		If .MaxRows > 0 Then
		    For lngRows = 1 To .MaxRows
		        .Row = lngRows
		        .Col = 0
		        If .text <> ggoSpread.DeleteFlag Then
			        .col = C_DrCrFg
			       If .text = "DR" Then
			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
			            Else
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
			            End If
			       ElseIf .text = "CR" Then
			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + 0
			            Else
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + UNICDbl(.Text)
			            End If
				   End If
				End If				
		    Next
		End If

        frm1.txtDrLocAmt.text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")
        frm1.txtCrLocAmt.text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")
	End With
	
'	If frm1.cboGlType.value = "01" Then
'		frm1.txtDrLocAmt.text = frm1.txtCrLocAmt.text
'	ElseIf frm1.cboGlType.value = "02" Then
'		frm1.txtCrLocAmt.text = frm1.txtDrLocAmt.text
'	End If
End Function

'========================================================================================================
Function CookiePage(ByVal Kubun)
	Dim strTemp
	Dim strNmwhere
	Dim arrVal
	Dim IntRetCD

	Select Case Kubun
		Case "ForM_LOAD"
			strTemp = ReadCookie("TEMP_GL_NO")
			Call WriteCookie("TEMP_GL_NO", "")

			If strTemp = "" Then Exit Function

			frm1.txtTempGlNo.value = strTemp

			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("TEMP_GL_NO", "")
				Exit Function 
			End If

			Call MainQuery()
		Case JUMP_PGM_ID_TAX_REP
			ggoSpread.Source = frm1.vspdData

			If frm1.vspddata.MaxRows	< 1  Then
				Exit Function
			End If

			frm1.vspddata.row = frm1.vspddata.ActiveRow	
			frm1.vspddata.Col = C_VatType

			If frm1.vspddata.Value	=	"" Then
				Exit Function
			End If

			frm1.vspddata.Col = C_ItemSeq

			strNmwhere = " TEMP_GL_NO  = " & FilterVar(frm1.txtTempGlNo.value , "''", "S")
			strNmwhere = strNmwhere & " And TEMP_ITEM_SEQ = " & frm1.vspddata.text & " "

			If CommonQueryRs( "VAT_NO" , "A_VAT" ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
				arrVal = Split(lgF0, Chr(11))  
				strTemp = arrVal(0)
			End If

			Call WriteCookie("VAT_NO", strTemp)	
		Case "TEMP_GL_POPUP"
			Call WriteCookie("PGMID", "A5101MA1")
		Case Else
			Exit Function
	End Select
End Function

'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD

    ggoSpread.Source = frm1.vspdData
        
    If (lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True ) And C_GLINPUTTYPE = frm1.cboGlInputType.value Then
		IntRetCD = DisplayMsgBox("990027", "X", "X", "X")
        Exit Function
    End If

	Select Case strPgmId
		Case JUMP_PGM_ID_TAX_REP
			ggoSpread.Source = frm1.vspdData

			If frm1.vspddata.MaxRows < 1 Then
				IntRetCD = DisplayMsgBox ("900002", "X","X","X")	
				Exit Function
			End If

			frm1.vspddata.row = frm1.vspddata.ActiveRow	
			frm1.vspddata.Col = C_VatType

			If frm1.vspddata.Value	=	"" Then
				IntRetCD = DisplayMsgBox ("205600", "X","X","X")	
				Exit Function
			End If
	End Select
	
	Call CookiePage(strPgmId)
	Call PgmJump(strPgmId)
End Function

'========================================================================================================
'	Desc : 입출금 화면에 따른 Grid의 Protect변환
'========================================================================================================
Sub CboGLType_ProtectGrid(ByVal GlType)
	ggoSpread.Source = frm1.vspdData
	Select Case GlType
		Case "01"
'		ggoSpread.SSSetProtected C_DocCur, 1, frm1.vspddata.maxrows
'		ggoSpread.SSSetProtected C_DocCurPopup, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows
		Case "02"
'		ggoSpread.SSSetProtected C_DocCur, 1, frm1.vspddata.maxrows
'		ggoSpread.SSSetProtected C_DocCurPopup, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows
		Case "03"
			ggoSpread.SSSetRequired C_DocCur, 1, frm1.vspddata.maxrows
			ggoSpread.SpreadUnLock C_DocCurPopup, 1, frm1.vspddata.maxrows
			ggoSpread.SpreadUnLock C_DrCrfg, 1, C_DrCrNm, frm1.vspddata.maxrows
			ggoSpread.SSSetRequired C_DrCrfg, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetRequired C_DrCrNm, 1, frm1.vspddata.maxrows
	End Select
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemSeq			= iCurColumnPos(1)
			C_deptcd			= iCurColumnPos(2)
			C_deptPopup			= iCurColumnPos(3)
			C_deptnm	   		= iCurColumnPos(4)
			C_AcctCd			= iCurColumnPos(5)
			C_AcctPopup			= iCurColumnPos(6)
			C_AcctNm			= iCurColumnPos(7)
			C_DrCrFg			= iCurColumnPos(8)
			C_DrCrNm			= iCurColumnPos(9)
			C_DocCur			= iCurColumnPos(10)
			C_DocCurPopup		= iCurColumnPos(11)
			C_ExchRate			= iCurColumnPos(12)
			C_ItemAmt			= iCurColumnPos(13)
			C_ItemLocAmt		= iCurColumnPos(14)
			C_IsLAmtChange		= iCurColumnPos(15)
			C_ItemDesc			= iCurColumnPos(16)
			C_VatType			= iCurColumnPos(17)
			C_VatNm				= iCurColumnPos(18)
			C_AcctCd2			= iCurColumnPos(19)
    End Select
End Sub

'=======================================================================================================
Sub vspdData_onfocus()
	lgCurrRow = frm1.vspdData.ActiveRow

'-- eWare If Begin		
	If Trim(parent.gEware) <> "" Then
		If lgIntFlgMode <> OPMD_UMODE Then                                        
			If frm1.hCongFg.value = "C" OR frm1.cboAppFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
				Call SetToolbar(MENU_PRT) 									'버튼 툴바 제어
			Else
 				Call SetToolbar(MENU_CRT)
			End If
		Else
			If frm1.hCongFg.value = "C" OR frm1.cboAppFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
				Call SetToolbar(MENU_PRT) 						   
			Else      
				Call SetToolbar(MENU_UPD)                                   '버튼 툴바 제어
			End If
		End If    	
'-- eWare If End
	Else
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
				Call SetToolbar(MENU_PRT)
			Else
				Call SetToolbar(MENU_CRT)	
			End If
		Else
			If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
				Call SetToolbar(MENU_PRT)
			Else
				Call SetToolbar(MENU_UPD)
			End If
		End If
	End If		
End Sub

'=======================================================================================================
Sub txttempGLDt_DblClick(ButTon)
    If ButTon = 1 Then
        frm1.txttempGLDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txttempGLDt.focus
    End If
End Sub

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

	ggoSpread.Source = frm1.vspdData
	frm1.vspddata.row = frm1.vspddata.ActiveRow	
 	frm1.vspdData.Col = C_AcctCd
	
    If Len(frm1.vspdData.Text) < 1 Then
        ggoSpread.Source = frm1.vspdData2
        ggoSpread.ClearSpreadData
	End If
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
Sub vspdData_MouseDown(ButTon, ShIft, X, Y)
	If ButTon = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq

            .hItemSeq.value = .vspdData.Text
            .vspdData2.MaxRows = 0
        End With

        frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End If

		lgCurrRow = NewRow
		Call DbQuery2(lgCurrRow)
    End If
End Sub

'==========================================================================================
Sub vspdData_ButTonClicked(ByVal Col, ByVal Row, ByVal ButTonDown)
	Dim Ifld1 
	Dim Ifld2
	Dim iTable
	Dim istrCode
	
	With frm1.vspdData
		If Row > 0 And Col = C_AcctPopUp Then
			.Col = Col - 1
			.Row = Row

			Call OpenPopUp(.text, 3)
		End If

		If Row > 0 And Col = C_deptPopup Then
			.Col = Col - 1
			.Row = Row
			Call OpenUnderDept(.Text, 1)
			Call SetActiveCell(frm1.vspdData,C_DeptCD,frm1.vspdData.ActiveRow ,"M","X","X")

    	End If
    	
		If Row > 0 And Col = C_DocCurPopup Then
			.Col = Col - 1
			.Row = Row
			Call OpenPopUp(.Text, 2)
		End If
	End With
End Sub

'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	Dim tmpDrCrFG
	Dim IntRetCD
	Dim TempExchRate
	Dim TempAmt

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    frm1.vspdData.Row = Row 
    frm1.vspdData.Col = Col  
	frm1.vspdData.action=0
				
    Select Case Col
	    Case   C_DeptCd
			frm1.vspdData.Col = C_DeptCd
			Call DeptCd_underChange(frm1.vspdData.text)
	    Case   C_AcctCd
		    frm1.vspdData.Col = 0
			If  frm1.vspdData.Text = ggoSpread.InsertFlag Then
				frm1.vspdData.Col = C_ItemSeq
				frm1.hItemSeq.value = frm1.vspdData.Text
				frm1.vspdData.Col = C_AcctCd

				If Len(frm1.vspdData.Text) > 0 Then
					frm1.vspdData.Row = Row
					frm1.vspdData.Col = C_ItemSeq
					DeleteHsheet frm1.vspdData.Text

					frm1.vspdData.Col = C_DrCrFg
					tmpDrCrFG = frm1.vspdData.text
					frm1.vspdData.Col = C_AcctCd

					If AcctCheck(frm1.vspdData.text,frm1.cboGlType.value, tmpDrCrFG) = True Then
						Call Dbquery3(Row)
						Call InputCtrlVal(Row)
					End If
				Else
					frm1.vspdData.Col = C_AcctNm
					frm1.vspdData.Text = ""
				End If
			End If
	  	Case	C_DrCrFg
    		Call SetSumItem	
    	Case	C_DrCrNm
    		Call vspdData_ComboSelChange(Col,Row)
			Call SetSumItem	
    	Case   C_ItemAmt
	    	frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_ItemAmt
			
    		TempAmt = UNICDbl(frm1.vspdData.text)
    		    		
    		frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_DocCur
    	
    		If UCase(Trim(frm1.vspdData.Text)) = parent.gCurrency Then
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = C_ItemLocAmt
				frm1.vspdData.Text = TempAmt
				
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = C_IsLAmtChange
				frm1.vspdData.Text = "Y"
			Else
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = C_ItemLocAmt
				frm1.vspdData.Text = ""
			End If
			
    		Call SetSumItem()
		Case   C_ItemLocAmt

	    	frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_ItemAmt
			
    		TempAmt = UNICDbl(frm1.vspdData.text)
    		    		
    		frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_DocCur
    	
    		If UCase(Trim(frm1.vspdData.Text)) = parent.gCurrency Then
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = C_ItemLocAmt
				frm1.vspdData.Text = TempAmt
			End If

			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_IsLAmtChange
			frm1.vspdData.Text = "Y"
			Call SetSumItem()
		Case	C_ExchRate
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_DocCur
			If UCase(Trim(frm1.vspdData.Text)) = parent.gCurrency Then
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = C_ExchRate
				frm1.vspdData.Text = 1
			End If
		Case	C_DocCur
			frm1.vspdData.Row  = Row
			frm1.vspdData.Col  = C_ItemLocAmt
			frm1.vspdData.Text = ""
			frm1.vspdData.Col  = C_DocCur
			If UCase(Trim(frm1.vspdData.Text)) = parent.gCurrency Then
				frm1.vspdData.Col = C_ExchRate
				frm1.vspdData.Text = 1
			Else
				Call FindExchRate(UniConvDateToYYYYMMDD(frm1.txttempGLDt.text,parent.gDateFormat,""), UCase(Trim(frm1.vspdData.Text)),frm1.vspdData.ActiveRow)
			End If
			
			Call DocCur_OnChange(Row,Row)
    End Select
End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim tmpDrCrFg
	Dim ii
	Dim iChkAcctForVat

	With frm1
		.vspddata.Row = Row
		Select Case Col
			Case C_DrCrNm
				.vspddata.Col = Col
				intIndex = .vspddata.Value
				.vspddata.Col = C_DrCrFg
				.vspddata.Value = intIndex
				tmpDrCrFg = .vspddata.text
				Call SetSpread2Color
			Case C_VatNm
				.vspddata.Col = Col
			    intIndex = .vspddata.Value
				.vspddata.Col = C_VatType
				.vspddata.Value = intIndex
			    Call InputCtrlVal(Row)'
		End Select
	End With
End Sub

'==========================================================================================
Sub txtTempGlNo_OnKeyPress()
	If window.event.keycode = 39 Then										'Single quotation mark 입력불가
		window.event.keycode = 0
	End If
End Sub

'==========================================================================================
Sub txtTempGlNo_OnKeyUp()
	If Instr(1,frm1.txtTempGlNo.value,"'") > 0 Then
		frm1.txtTempGlNo.value = Replace(frm1.txtTempGlNo.value, "'", "")
	End If
End Sub

'==========================================================================================
Sub txtTempGlNo_onpaste()
	Dim iStrTempGlNo
	
	iStrTempGlNo = window.clipboardData.getData("Text")
	iStrTempGlNo = RePlace(iStrTempGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrTempGlNo)
End Sub

'==========================================================================================
Sub DocCur_OnChange(FromRow, ToRow)
	Dim ii
	
    lgBlnFlgChgValue = True
    
	For ii = FromRow	To	ToRow
		frm1.vspdData.Row	= ii
		frm1.vspdData.Col	= C_DocCur
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.vspdData.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			Call CurFormatNumSprSheet(ii)
			Call SetSumItem
		End If
	Next  
End Sub

'==========================================================================================
Sub txtDeptCd_OnChange()
    Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtTempGLDt.Text = "") Or Trim(frm1.txtDeptCd.value) = "" Then
		Exit Sub
    End If
    
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "
	strFrom		=			 " b_acct_dept(NOLOCK) "
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " And org_change_id = (Select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( Select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

'		' 권한관리 추가
'		If lgInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
'		End If
'	
'		If lgSubInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
'		End If
	
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then

		IntRetCD = DisplayMsgBox ("124600","X","X","X")
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hOrgChangeId.value = ""
	Else 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			frm1.hOrgChangeId.value = Trim(arrVal2(2))
		Next
	End If
End Sub

'==========================================================================================
Sub QueryDeptCd_OnChange()
    Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD
	Dim arrVal1
	Dim arrVal2
	Dim ii

	If Trim(frm1.txtTempGLDt.Text = "") Then
		Exit Sub
    End If

    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "
	strFrom		=			 " b_acct_dept(NOLOCK) "
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " And org_change_id = (Select distinct org_change_id "
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( Select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"	

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hOrgChangeId.value = ""
	Else
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		
		For ii = 0 To Ubound(arrVal1,1) - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			frm1.hOrgChangeId.value = Trim(arrVal2(2))
		Next
	End If
End Sub

'==========================================================================================
Sub DeptCd_underChange(ByVal strCode)
    Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD 

    If Trim(frm1.txtTempGLDt.Text = "") Then
		Exit Sub
    End If
    
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "
	strFrom		=			 " b_acct_dept(NOLOCK) "
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
	strWhere	= strWhere & " And org_change_id = (Select distinct org_change_id "
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( Select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox ("124600","X","X","X")  

		frm1.vspdData.Col = C_deptcd
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.text = ""
		frm1.vspdData.Col = C_deptnm
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = ""
	End If	
End Sub

'==========================================================================================
Sub txttempGLDt_Change()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim IntRetCD
	Dim ii
	Dim arrVal1
	Dim arrVal2

	If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True

			With frm1
				If LTrim(RTrim(.txtDeptCd.value)) <> "" And Trim(.txtTempGLDt.Text <> "") Then
					strSelect	=			 " dept_cd, org_change_id, internal_cd "
					strFrom		=			 " b_acct_dept(NOLOCK) "
					strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
					strWhere	= strWhere & " And org_change_id = (Select distinct org_change_id "			
					strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( Select max(org_change_dt)"
					strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
						IntRetCD = DisplayMsgBox ("124600","X","X","X")
						.txtDeptCd.value = ""
						.txtDeptNm.value = ""
						.hOrgChangeId.value = ""
						
						If .vspdData.MaxRows <> 0 Then
							For ii = 1 To .vspdData.MaxRows
							.vspdData.Col = C_deptcd
						    .vspdData.Row = ii
						    .vspdData.text = ""
						    .vspdData.Col = C_deptnm
						    .vspdData.text = ""
							Next
						End If
					Else
						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))

						For ii = 0 To Ubound(arrVal1,1) - 1
							arrVal2 = Split(arrVal1(ii), chr(11))
							frm1.hOrgChangeId.value = Trim(arrVal2(2))
						Next
					End If 
				End If
			End With
		End If
	End If
End Sub

'==========================================================================================
Sub cboGLType_OnChange()
	Dim	i
	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData

	Select Case frm1.cboGlType.value 
		Case "01"			
			'입금전표로 바꾸면 차변이 입력되거나 현금계정이 입력되었는지 check한다.
			For i = 1 To  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				If  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )
					IntRetCD = DisplayMsgBox ("113106", "X", "X", "X")
					Exit Sub
				End If

				frm1.vspddata.col = C_DrCrFg
				If  Trim(frm1.vspddata.value) = "2" Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )
					IntRetCD = DisplayMsgBox ("113104", "X", "X", "X")
					Exit Sub
				End If
			Next

			For i = 1 To  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				If Trim(frm1.vspddata.value) <> "1"  Then
					frm1.vspdData.value	= "1"
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.value	= "1"
				End If
				
				Call vspdData_ComboSelChange(C_DrCrNm,i)
				frm1.vspddata.col = C_DocCur
				frm1.vspddata.text = parent.gCurrency
			Next

			Call CboGLType_ProtectGrid(frm1.cboGlType.value )
		Case "02"
			'출금전표로 바꾸면 대변이 입력되거나 현금계정이 입력되었는지 check한다.	
			For i = 1 To  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				If  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )
					IntRetCD = DisplayMsgBox ("113106", "X", "X", "X")
					Exit Sub
				End If

				frm1.vspddata.col = C_DrCrFg
				If  Trim(frm1.vspddata.value) = "1" Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )
					IntRetCD = DisplayMsgBox ("113105", "X", "X", "X")
					Exit Sub
				End If
			Next

			For i = 1 To  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				If Trim(frm1.vspddata.value) <> "2"  Then
					frm1.vspdData.value	= "2"
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.value	= "2"
				End If
				
				Call vspdData_ComboSelChange(C_DrCrNm,i)				
				frm1.vspddata.col = C_DocCur
				frm1.vspddata.text = parent.gCurrency
			Next
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )
		Case "03"
		'대체로 바꾸면 Protect를 풀어준다.
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )

	End Select	

	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim RetFlag

    lgstartfnc = True
    FncQuery = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox ("900013", parent.VB_YES_NO, "X", "X")
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

    If Not chkFieldByCell(frm1.txtTempGlNo,"A",1) Then Exit Function

	'-----------------------
    'Erase contents area
    '-----------------------
'    Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables

    If  DbQuery = False Then
		Exit Function
	End If
		
'    If frm1.vspddata.maxrows = 0 Then	
'		frm1.txtTempGlNo.value = ""
'    End If
   
    FncQuery = True	
    lgstartfnc = False
End Function

'========================================================================================
Function FncNew() 
	Dim IntRetCD 
	Dim var1, var2

	FncNew = False
	lgstartfnc = True

    Err.Clear
    On Error Resume Next

    If (lgBlnFlgChgValue = True Or var1 = True Or var2 = True) And lgBlnExecDelete <> True Then
        IntRetCD = DisplayMsgBox ("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	lgBlnExecDelete = False

    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")

    Call FormatDATEField(frm1.txttempGLDt)
    Call LockObjectField(frm1.txttempGLDt,"R")

    Call LockHTMLField(frm1.cboGlType,"R")
    Call LockHTMLField(frm1.cboConfFg,"P")
    Call LockHTMLField(frm1.cboGlInputType,"P")            

    Call FormatDoubleSingleField(frm1.txtDrLocAmt)
    Call LockObjectField(frm1.txtDrLocAmt,"P")

    Call FormatDoubleSingleField(frm1.txtCrLocAmt)
    Call LockObjectField(frm1.txtCrLocAmt,"P")

    Call LockHTMLField(frm1.txtDeptCd,"R")
    Call LockHTMLField(frm1.txtdesc,"R")  

'	Call ggoOper.LockField(Document, "N")

    Call InitComboBoxGrid
    frm1.txtTempGlNo.focus
    Call SetToolbar(MENU_NEW)
    
'	Call ggoOper.SetReqAttr(frm1.txtDeptCd,   "N")
'	Call ggoOper.SetReqAttr(frm1.txtTempGlDt, "N")
'	Call ggoOper.SetReqAttr(frm1.txtdesc,   "D")

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData

	SetGridFocus()
    SetGridFocus2()

	Call SetDefaultVal
	Call InitVariables
	Call SetSumItem()
    Set gActiveElement = document.ActiveElement

    lgBlnFlgChgValue = False

    FncNew = True
    lgFormLoad = True
    lgQueryOk = False
    lgstartfnc = False
End Function

'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
    Dim strNmwhere, arrVal, strTemp
    
    FncDelete = False
    Err.Clear
    On Error Resume Next
    lgBlnExecDelete = True

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then	
		intRetCd = DisplayMsgBox ("990008", parent.VB_YES_NO, "X", "X")
		If intRetCd = VBNO Then
			Exit Function
		End If
    Else
		IntRetCD = DisplayMsgBox ("900038", parent.VB_YES_NO, "X", "X")
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If


   	strNmwhere = " TEMP_GL_NO  = " & FilterVar(frm1.txtTempGlNo.value , "''", "S")
	If CommonQueryRs( "hq_brch_no" , "a_temp_gl" ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		arrVal = Split(lgF0, Chr(11))  
		strTemp = arrVal(0)
	End If
			
	If strTemp <> "" Then
		IntRetCD = DisplayMsgBox ("1a0513", "X", "X", "X")
		Exit Function
	End If	

    If  DbDelete = False Then
    	Exit Function
    End If
    FncDelete = True
End Function

'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim strNmwhere, arrVal, strTemp
    
    FncSave = False
    
    Err.Clear
    On Error Resume Next

	With frm1
	    ggoSpread.Source = .vspdData
	    If lgBlnFlgChgValue = False  And ggoSpread.SSCheckChange = False Then
	        IntRetCD = DisplayMsgBox ("900001", "X", "X", "X")
	        Exit Function
	    End If

	    If CheckSpread3 = False Then
			IntRetCD = DisplayMsgBox ("110420", "X", "X", "X")
	        Exit Function
	    End If

		If frm1.vspdData.MaxRows < 1 Then
			IntRetCD = DisplayMsgBox ("114100", "X", "X", "X")
			Exit Function
		End If

	    ggoSpread.Source = .vspdData

	    If Not chkFieldByCell(.txttempGLDt, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboGlType, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.txtDeptCd, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboConfFg, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboGlInputType, "A", "1") Then Exit Function
	    
	    If Not chkFieldByCell(.txtDesc, "A", "1") Then Exit Function	 	'>>air   
	    If Not ChkFieldLengthByCell(.txtDesc, "A", "1") Then Exit Function    

	'    If Not chkField(Document, "2") Then
	'		Exit Function
	'    End If
	
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strNmwhere = " TEMP_GL_NO  = " & FilterVar(frm1.txtTempGlNo.value , "''", "S")

			If CommonQueryRs( "hq_brch_no" , "a_temp_gl" ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
				arrVal = Split(lgF0, Chr(11))  
				strTemp = arrVal(0)
			End If
			
			If strTemp <> "" Then
				IntRetCD = DisplayMsgBox ("1a0513", "X", "X", "X")
				Exit Function
			End If
		End If

	    If Not ggoSpread.SSDefaultCheck Then
			Exit Function
	    End If

	    If  DbSave	= False Then
			Exit Function
	    End If
	    
	    FncSave = True
	End With	    
End Function

'========================================================================================
Function FncCopy() 
	Dim  IntRetCD

	With frm1
		.vspdData.ReDraw = False
		If .vspdData.MaxRows < 1 Then Exit Function
	
		ggoSpread.Source = .vspdData
		ggoSpread.CopyRow
		SetSpreadColor "I",0, .vspdData.ActiveRow, .vspdData.ActiveRow
		MaxSpreadVal .vspdData, C_ItemSeq, .vspdData.ActiveRow

		Call ReFormatSpreadCellByCellByCurrency(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,C_DocCur,C_ItemAmt, "A" ,"I","X","X")
        Call ReFormatSpreadCellByCellByCurrency(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,C_DocCur,C_ExchRate,"D" ,"I","X","X")
		Call vspdData_Change(C_AcctCd, .vspddata.activerow)
		Call SetSumItem()
	End With		
End Function

'========================================================================================================
Function FncCancel() 
    Dim iItemSeq
    Dim RowDocCur

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	If  frm1.vspdData.MaxRows = 1 Then  Call LockHTMLField(frm1.cboGlType,"R")	'ggoOper.SetReqAttr(frm1.cboGlType,   "N")

    With frm1.vspdData
        .Row = .ActiveRow
        .Col = 0

        If .Text = ggoSpread.InsertFlag Then
			.Col = C_AcctCd
			If len(Trim(.text)) > 0 Then 
				.Col = C_ItemSeq
				DeleteHSheet(.Text)
			End If
        End If
        
        ggoSpread.EditUndo
        ggoSpread.Source = frm1.vspdData
        
        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DocCur,C_ItemAmt, "A" ,"I","X","X")
        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DocCur,C_ExchRate,"D" ,"I","X","X")

		If .MaxRows = 0 Then
			Call SetToolbar(MENU_NEW)
			Exit Function
		End If

        InitData

        .Row = .ActiveRow
        .Col = 0
        
		If .row = 0 Then
			Exit Function
		End If

        If .Text = ggoSpread.InsertFlag Then
		    .Col = C_AcctCd
            If Len(.Text) > 0 Then
				.Col = C_ItemSeq
				frm1.hItemSeq.value = .Text
	            frm1.vspdData2.MaxRows = 0
		        Call DbQuery3(.ActiveRow)
            End If
        Else
			.Col = C_ItemSeq
            frm1.hItemSeq.value = .Text
			ggoSpread.Source=frm1.vspdData2
			ggoSpread.ClearSpreadData()
		    Call DbQuery2(.ActiveRow)
        End If
    End With
    Call SetSumItem()
End Function

'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim imRow2
	Dim iCurRowPos

    On Error Resume Next
    Err.Clear
    
    FncInsertRow = False

	With frm1
	    If Not chkFieldByCell(.txttempGLDt, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboGlType, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.txtDeptCd, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboConfFg, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboGlInputType, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.txtDesc, "A", "1") Then Exit Function	 	'>>air   	    	
	    
	'    If Not chkField(Document, "2") Then 
	'        Exit Function
	'    End If

	    If IsNumeric(Trim(pvRowCnt)) Then
	        imRow = CInt(pvRowCnt)
	    Else
	        imRow = AskSpdSheetAddRowCount()
			If imRow = "" Then
	            Exit Function
	        End If
	    End If

        .vspdData.ReDraw = False
		.vspdData.focus
        ggoSpread.Source = .vspdData
		iCurRowPos = .vspdData.ActiveRow

        For imRow2 = 1 To imRow 
            ggoSpread.InsertRow ,1
            .vspdData.row = .vspdData.ActiveRow
           .vspdData.col = C_deptcd
            .vspddata.text	= UCase(.txtDeptCd.Value)

            .vspdData.col = C_deptnm
            .vspddata.text	= .txtDeptNm.Value

            .vspdData.col = C_DocCur
            .vspddata.text	= parent.gCurrency

            .vspdData.col = C_ExchRate
            .vspddata.text	= "1"

            .vspdData.col = C_ItemDesc
            .vspddata.text	= .txtDesc.Value

			If .cboGlType.value = "01" Then
                .vspdData.col = C_DrCrNm
                .vspdData.value	= 1
                .vspdData.col = C_DrCrFg
                .vspdData.value	= 1
            ELSEIf .cboGlType.value = "02" Then
                .vspdData.col = C_DrCrNm
                .vspdData.value	= 2
                .vspdData.col = C_DrCrFg
                .vspdData.value	= 2
            End If
            
            SetSpreadColor "I", 0, .vspdData.ActiveRow, .vspdData.ActiveRow
            MaxSpreadVal .vspdData, C_ItemSeq, .vspdData.ActiveRow
        Next

        Call ReFormatSpreadCellByCellByCurrency(.vspdData,iCurRowPos + 1,iCurRowPos + imRow,C_DocCur,C_ItemAmt, "A" ,"I","X","X")
        Call ReFormatSpreadCellByCellByCurrency(.vspdData,iCurRowPos + 1,iCurRowPos + imRow,C_DocCur,C_ExchRate,"D" ,"I","X","X")

        .vspdData.ReDraw = True
    End With
    
    If Err.number = 0 Then
		FncInsertRow = True
    End If
    
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows
	Dim iDelRowCnt, i
    Dim DelItemSeq

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData 

		.Row = .ActiveRow
		.Col = 0 

		If frm1.vspdData.MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function
		.Col = 1 
		DelItemSeq = .Text

		lDelRows = ggoSpread.DeleteRow
    End With

    DeleteHsheet DelItemSeq
    Call SetSumItem()
End Function

'========================================================================================
Function FncPrint() 
    On Error Resume Next    
    parent.FncPrint()
End Function
'========================================================================================
Function FncExcel() 
    On Error Resume Next    
    Call parent.FncExport(parent.C_MULTI)
End Function
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
End Function
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopResToreSpreadColumnInf()
	Dim indx

	On Error Resume Next
	Err.Clear 		

	ggoSpread.Source = gActiveSpdSheet
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call PrevspdDataResTore(gActiveSpdSheet)
			Call ggoSpread.ResToreSpreadInf()
			Call InitSpreadSheet()
            Call InitComboBoxGrid
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
	        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, 1, -1 ,C_DocCur,C_ItemAmt, "A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, 1, -1 ,C_DocCur,C_ExchRate,"D" ,"I","X","X")	        

			If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			    Call SetSpreadLock("Q", 1, 1, "")
			    Call SetSpread2Lock("",1,"","")
			Else
                Call SetSpreadColor("Q", 0,1, frm1.vspdData.MaxRows)
                Call SetSpread2Color()
			End If
		Case "VSPDDATA2"
			Call PrevspdData2ResTore(gActiveSpdSheet)
			Call ggoSpread.ResToreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			    Call SetSpread2Lock("",1,"","")
			Else
			    Call SetSpread2Color()
			End If
	End Select

	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If
	
	Call SetSumItem()
End Sub

'=======================================================================================================
Sub PrevspdDataResTore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData.MaxRows
        frm1.vspdData.Row    = indx
        frm1.vspdData.Col    = 0
		
		If frm1.vspdData.Text <> "" Then
			Select Case frm1.vspdData.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData.Col = C_ItemSeq
					Call DeleteHsheet(frm1.vspdData.Text)
				Case ggoSpread.UpdateFlag
					For indx1 = 0 To frm1.vspdData3.MaxRows
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case frm1.vspdData3.Text 
							Case ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1
								If UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)
									Call fncResToreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtTempGlNo.Value)
								End If
						End Select
					Next
				Case ggoSpread.DeleteFlag
					Call fncResToreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtTempGLNo.Value)
			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

'=======================================================================================================
Sub PrevspdData2ResTore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 To frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
					        ggoSpread.EditUndo
						End If
					Next
				Case ggoSpread.UpdateFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 To frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncResToreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.htxtTempGLNo.Value)
						End If
					Next
				Case ggoSpread.DeleteFlag
			End Select
		End If
	Next	
	ggoSpread.Source = pActiveSheetName
End Sub

'========================================================================================================
Function fncResToreDbQuery2(Row, CurrRow, ByVal pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal
'	Dim tmpDrCrFG

	On Error Resume Next
	Err.Clear

	fncResToreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)

	With frm1
		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text

'	    .vspdData.Col = C_DrCrFg
'		frm1.vspdData.Col = C_DrCrFg
'		tmpDrCrFG = frm1.vspdData.text

	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If

		Call LayerShowHide(1)

		DbQuery2 = False

		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq

		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , LTrim(ISNULL(C.CTRL_VAL,'')), '',"
		strSelect = strSelect & " Case  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " Case WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  And  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
'		strSelect = strSelect & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'		strSelect = strSelect & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  		
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  And  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  And  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " End	, " & strItemSeq & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_TEMP_GL_DTL C (NOLOCK), A_TEMP_GL_ITEM D (NOLOCK) "

		strWhere =			  " D.TEMP_GL_NO = " & FilterVar(UCase(pInvalue1), "''", "S")
		strWhere = strWhere & " And D.ITEM_SEQ = " & strItemSeq & " "
		strWhere = strWhere & " And D.TEMP_GL_NO  =  C.TEMP_GL_NO  "
		strWhere = strWhere & " And D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	And D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " And C.CTRL_CD *= B.CTRL_CD "
		strWhere = strWhere & " And C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "


		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)

					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " And MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If

		If Row = CurrRow Then
			Call CopyFromData (strItemSeq)
		End If

		Call LayerShowHide(0)
		Call ResToreToolBar()
	End With

	If Err.number = 0 Then
		fncResToreDbQuery2 = True
	End If
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD
	Dim var1,var2

	FncExit = False

	ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True OR ggoSpread.SSCheckChange = True Then  
		IntRetCD = DisplayMsgBox ("900016", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Function FncBtnPreview() 
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId, varLoginDeptNm, varLoginUsrId
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
	Dim ObjName

    If Not chkFieldByCell(frm1.txtTempGlNo,"A",1) Then Exit Function
    
'    If Not chkField(Document, "1") Then
'		Exit Function
'    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"
	StrUrl = StrUrl & "|gUsrId|" & parent.gUsrId
	StrUrl = StrUrl & "|LoginDeptNm|" & parent.gDepart
	'StrUrl = StrUrl & "|LoginUsrId|" & varLoginUsrId	

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
End Function

'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId, varLoginDeptNm, varLoginUsrId
    Dim StrEbrFile
    Dim intRetCd
	Dim ObjName

    If Not chkFieldByCell(frm1.txtTempGlNo,"A",1) Then Exit Function
    
'	If Not chkField(Document, "1") Then
'       Exit Function
'    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)

    lngPos = 0

	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"
	StrUrl = StrUrl & "|gUsrId|" & parent.gUsrId
	StrUrl = StrUrl & "|LoginDeptNm|" & parent.gDepart
	'StrUrl = StrUrl & "|LoginUsrId|" & varLoginUsrId

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
End Function

'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)
	Dim intRetCd

	StrEbrFile = "a5101ma1_lko441"
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txttempGlDt.Text, parent.gDateFormat, parent.gServerDateType)	
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txttempGlDt.Text, parent.gDateFormat, parent.gServerDateType)

	' 회계전표의 key는 GL_NO이기 때문에 GL_NO만 넘긴다.	
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	VarTempGlNoFr = Trim(frm1.txttempGlNo.value)
	VarTempGlNoTo = Trim(frm1.txttempGlNo.value)
	varOrgChangeId = Trim(frm1.hOrgChangeId.value)
	'varLoginDeptNm = Trim(frm1.txtLoginDeptNm.value)
	
'	IF CommonQueryRs( "USR_NM" , "z_usr_mast_rec" , "USR_ID = " & FilterVar(parent.gUsrId, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
'		arrVal = Split(lgF0, Chr(11))
'		varLoginUsrId = arrVal(0)	'로그인유저명>>air
'	End If	
		
	'MsgBox varLoginDeptNm	'>>air
End Sub

'========================================================================================
' Function Name : FncBtnCalc
' Function Desc : This function calculate local amt from amt of multi
'========================================================================================
Function FncBtnCalc() 
	Dim ii
	Dim tempAmt, tempLocAmt, tempExch, TempSep, tempDoc
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strDate
	Dim strExchFg
	Dim IntRetCD
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

	With frm1
		strSelect	= "b.minor_cd"
		strFrom		= "b_company a, b_minor b"
		strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  And	a.xch_rate_fg = b.minor_cd"
		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchFg =  arrTemp(0)
		End If

		strDate = UniConvDateToYYYYMMDD(frm1.txttempGLDt.text,parent.gDateFormat,"")
		If .vspdData.MaxRows <> 0 Then
			For ii = 1 To .vspdData.MaxRows
				.vspdData.Row	=	ii
				.vspdData.Col	=	C_DocCur
				tempDoc			=	UCase(Trim(.vspdData.text))
				.vspdData.Col	=	C_ItemAmt
				tempAmt			=	UNICDbl(.vspdData.text)
				.vspdData.Col	=	C_ExchRate
				tempExch		=	UNICDbl(.vspdData.text)

				If tempDoc	<> "" And tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox ("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "Top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep = arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox ("121500", "X", "X", "X")
						End If
					End If
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt	=	tempAmt * TempExch
					Else
						tempLocAmt	=	tempAmt / TempExch
					End If
					.vspdData.Col	= C_ItemLocAmt
					.vspdData.text	= UNIConvNumPCToCompanyByCurrency(tempLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")

				ElseIf tempDoc = parent.gCurrency Then
					.vspdData.Col	= C_ItemLocAmt
					.vspdData.text	= UNIConvNumPCToCompanyByCurrency(tempAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")

				End If
			Next
		End If
	End With

	Call SetSumItem	
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim RetFlag

    Err.Clear

    DbQuery = False
    Call LayerShowHide(1)

    ggospread.Source=frm1.vspdData3
    ggoSpread.ClearSpreadData()

    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtTempGlNo=" & UCase(Trim(.htxtTempGlNo.value))
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 
			strVal = strVal & "&txtTempGlNo=" & UCase(Trim(.txtTempGlNo.value))
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
		End If

		' 권한관리 추가
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd		' 사업장
		strVal = strVal & "&lgInternalCd="		& lgInternalCd			' 내부부서
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd		' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID			' 개인
   
		Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True
End Function

'=======================================================================================================
Function DbQueryOk()
	Dim ii

	With frm1
        lgIntFlgMode = parent.OPMD_UMODE
		.txtCommAndMode.value = "UPDATE"
		Call InitData()

		Call ggoOper.SetReqAttr(frm1.txtTempGlDt,	"Q")

		'-- eWare If Begin
		If Trim(parent.gEware) <> "" Then
			If frm1.hCongFg.value = "C" OR frm1.cboConfFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
				Call SetToolbar(MENU_PRT)
				Call SetSpreadLock("Q", 1, 1, "")
				Call ggoOper.SetReqAttr(frm1.txtDeptCd,	"Q")
				Call ggoOper.SetReqAttr(frm1.txtdesc,   "Q")
				Call ggoOper.SetReqAttr(frm1.cboGlType,	"Q")
				Call ggoOper.SetReqAttr(frm1.cboConfFg,	"Q")
			Else
				Call SetToolbar(MENU_UPD)									'버튼 툴바 제어
				Call SetSpreadLock("Q", 0, 1, "")
				Call SetSpreadColor("Q", 0,1, .vspddata.MaxRows)
				Call ggoOper.SetReqAttr(frm1.txtDeptCd,	"N")
				Call ggoOper.SetReqAttr(frm1.txtdesc,	"D")
			End If
		Else
		'-- eWare If End
			If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
				Call SetToolbar(MENU_PRT)
				Call SetSpreadLock("Q", 1, 1, "")

				Call ggoOper.SetReqAttr(frm1.txtDeptCd,	"Q")
				Call ggoOper.SetReqAttr(frm1.txtdesc,   "Q")
				Call ggoOper.SetReqAttr(frm1.cboGlType,		"Q")
			Else
				Call SetToolbar(MENU_UPD)
				Call SetSpreadLock("Q", 0, 1, "")
				Call SetSpreadColor("Q", 0,1, .vspddata.MaxRows)

				Call ggoOper.SetReqAttr(frm1.txtDeptCd,	"N")
				Call ggoOper.SetReqAttr(frm1.txtdesc,	"D")
			End If
		End If
		
		If .vspdData.MaxRows > 0 Then
			.vspdData.Row = 1
			.vspdData.Col = C_ItemSeq
			.hItemSeq.Value = .vspdData.Text
			Call DbQuery2(1)
		End If
    End With
    
    Call QueryDeptCd_OnChange()
	Call SetGridFocus()
    Call SetGridFocus2()
    
    lgBlnFlgChgValue = False
End Function

'=======================================================================================================
Function DbQuery2(ByVal Row)
	Dim strVal	
	Dim lngRows

	Dim strSelect
	Dim strFrom
	Dim strWhere

	Dim strTableid
	Dim strColid
	Dim strColNm
	Dim strMajorCd
	Dim strNmwhere
	Dim i
	Dim arrVal
	Dim arrTemp
	Dim Indx1
'	Dim tmpDrCrFG	

	With frm1
	    .vspdData.row = Row
	    .vspdData.col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

'	    .vspdData.Col = C_DrCrFg
'		frm1.vspdData.Col = C_DrCrFg
'		tmpDrCrFG = frm1.vspdData.text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If

	    .vspdData2.ReDraw = False

        If CopyFromData(.hItemSeq.Value) = True Then
'			If .hCongFg.value = "C" Or .cboGlInputType.Value <> C_GLINPUTTYPE Then
'-- eWare If change
			If frm1.hCongFg.value = "C" Or frm1.cboConfFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then			
				Call SetSpread2Lock("",1,1,"")
			Else
				Call SetSpread2Color()
			End  If
			
			.vspdData2.ReDraw = True
            Exit Function
        End If

		Call LayerShowHide(1)

		DbQuery2 = False

		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq

		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , LTrim(ISNULL(C.CTRL_VAL,'')), '',"
		strSelect = strSelect & " Case  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " Case WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  And  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
'		strSelect = strSelect & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'		strSelect = strSelect & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  		
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  And  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  And  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " End	, " & .hItemSeq.Value & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_TEMP_GL_DTL C (NOLOCK), A_TEMP_GL_ITEM D (NOLOCK) "

		strWhere =			  " D.TEMP_GL_NO = " & FilterVar(UCase(.htxtTempGlNo.value), "''", "S")
		strWhere = strWhere & " And D.ITEM_SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " And D.TEMP_GL_NO  =  C.TEMP_GL_NO  "
		strWhere = strWhere & " And D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	And D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " And C.CTRL_CD *= B.CTRL_CD "
		strWhere = strWhere & " And C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = .vspdData2
			arrTemp =  Split(lgF2By2,Chr(12))

			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next

			lgF2By2 = Join(arrTemp,Chr(12))
			ggoSpread.SSShowData lgF2By2

			For lngRows = 1 To .vspdData2.Maxrows
				.vspddata2.row = lngRows	
				.vspddata2.col = C_Tableid 
				If Trim(.vspddata2.text) <> "" Then
					.vspddata2.col = C_Tableid
					strTableid = .vspddata2.text
					.vspddata2.col = C_Colid
					strColid = .vspddata2.text
					.vspddata2.col = C_ColNm
					strColNm = .vspddata2.text	
					.vspddata2.col = C_MajorCd
					strMajorCd = .vspddata2.text

					.vspddata2.col = C_CtrlVal

					strNmwhere = strColid & " =  " & FilterVar(UCase(.vspddata2.text), "''", "S")
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " And MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") 
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						.vspddata2.col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						.vspddata2.text = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & .hItemSeq.Value

				.vspdData2.Col = C_DtlSeq
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlCd
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlNm
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlVal
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlPB
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlValNm
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_Seq
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_Tableid
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_Colid
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_ColNm
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_Datatype
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_DataLen
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_DRFg
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_MajorCd
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_MajorCd + 1

				strVal = strVal & Chr(11) & lngRows

				strVal = strVal & Chr(11) & Chr(12)
			Next

			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If

		intItemCnt = .vspddata.MaxRows
        
'		If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
'-- eWare If change
		If frm1.hCongFg.value = "C" Or frm1.cboConfFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetSpread2Lock("",1,1,"")
		Else
			Call SetSpread2Color()
		End  If

		.vspdData2.ReDraw = True
	End With		

	Call LayerShowHide(0)

	DbQuery2 = True
	lgQueryOk = True
End Function

'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim intIndex2 

	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow

			.Col = C_DrCrFg
			intIndex = .value
			.col = C_DrCrNm
			.value = intindex
			
			.Col = C_VatType
			intIndex2 = .value
			.col = C_VatNm
			.value = intIndex2
		Next
	End With
End Sub

'========================================================================================================
Function DbSave() 
    Dim pAP010M 
    Dim lngRows , itemRows
    Dim lGrpcnt
    DIM strVal 
    Dim tempItemSeq
	Dim	intRetCd
	Dim ii	
    Dim strNote
    Dim strItemDesc
    strNote = ""
    DbSave = False
    
    Call LayerShowHide(1)
    
    On Error Resume Next
	Err.Clear 

	With frm1
		.txtFlgMode.value     = lgIntFlgMode
		.txtUpdtUserId.value  = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtMode.value        = parent.UID_M0002
		.txtAuthorityFlag.value     = lgAuthorityFlag               '권한관리 추가
	End With

    ' Data 연결 규칙
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타

    lGrpCnt = 1
    strVal = ""
 
    ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
		For lngRows = 1 To .MaxRows
    		.Row = lngRows
			.Col = 0

			If .Text <> ggoSpread.DeleteFlag Then
				strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별
			    .Col = C_ItemSeq	'1
			    strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_deptcd	    '2
			    strVal = strVal & Trim(.Text) & parent.gColSep
			    .Col = C_AcctCd		'3
			    strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_DrCrFG		'4
			    strVal = strVal & Trim(.Text) & parent.gColSep
			    .Col = C_ItemAmt	'5
			    strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep
				.Col = C_IsLAmtChange	
  				.Col = C_ItemLocAmt	'6
				strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep
			    .Col = C_ItemDesc	'7
			    strItemDesc = Trim(.Text)
			    
			    If TrIM(strItemDesc) = "" Or isnull(strItemDesc) Then
					 ggoSpread.Source = frm1.vspdData3
					 .Col = C_ItemSeq
					 tempItemSeq = .Text  
					 strNote = ""
					 With frm1.vspdData3
							For itemRows = 1 To .MaxRows
								.Row = itemRows
								.Col = 1
								
								If .Text =  tempItemSeq Then
									.Col= 9 'C_Tableid	+ 1
									If 	.Text = "B_BIZ_PARTNER" Or .Text = "B_BANK" Or .Text = "F_DPST" Then
										.Col = 7 'C_CtrlValNm + 1 
									Else
										.Col = 5 'C_CtrlVal + 1 
									End If	
									strNote = strNote & C_NoteSep & Trim(.Text)
								End If		    
							Next
							strNote = Mid(strNote,2)
					 End With
					 
					 strVal = strVal & strNote & parent.gColSep
					 ggoSpread.Source = frm1.vspdData
			    Else
					strVal = strVal & strItemDesc & parent.gColSep
			    End If

				.Col = C_ExchRate	'8
			    strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep

			    .Col = C_VatType	'9
			    strVal = strVal & Trim(.Text) & parent.gColSep

			    .Col = C_DocCur		'10
			    strVal = strVal & UCase(Trim(.Text)) & parent.gRowSep

			    lGrpCnt = lGrpCnt + 1
			End If
		Next
    End With
	
    frm1.txtMaxRows.value = lGrpCnt-1								'Spread Sheet의 변경된 최대갯수
    frm1.txtSpread.value  = strVal									'Spread Sheet 내용을 저장    

	If frm1.txtSpread.value = "" Then
		intRetCd = DisplayMsgBox ("990008", parent.VB_YES_NO, "X", "X")
		If intRetCd = VBNO Then
			Exit Function
		End If
		Call DbDelete
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.ClearSpreadData
		Call InitVariables
		Exit Function
	End If

    lGrpCnt = 1
    strVal = ""

    ggoSpread.Source = frm1.vspdData3

    With frm1.vspdData3      ' Dtl 저장 
		For itemRows = 1 To frm1.vspdData.MaxRows 
 		    frm1.vspdData.Row = itemRows
		    frm1.vspdData.Col = 0

		    If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then	
				frm1.vspdData.Col = C_ItemSeq
			    tempItemSEq = frm1.vspdData.Text  

			    For lngRows = 1 To .MaxRows
					.Row = lngRows
					.Col = 1

					If .text = tempitemseq Then
						.Col = 0 
						strVal = strVal & "C" & parent.gColSep
						.Col = 1 		 			'ItemSEQ	
						strVal = strVal & tempitemseq & parent.gColSep
						.Col =  2 'C_DtlSeq + 1   				'Dtl SEQ
						strVal = strVal & Trim(.Text) & parent.gColSep
						.Col =  3 'C_CtrlCd + 1		 		'관리항목코드
						strVal = strVal & Trim(.Text) & parent.gColSep
						.Col = 5 'C_CtrlVal + 1				'관리항목 Value 
						strVal = strVal & UCase(Trim(.Text)) & parent.gRowSep	
						lGrpCnt = lGrpCnt + 1
					End If
		    	Next
		   End If
   		Next
    End With

    frm1.txtMaxRows3.value = lGrpCnt-1					'Spread Sheet의 변경된 최대갯수
    frm1.txtSpread3.value  = strVal						'Spread Sheet 내용을 저장
    
    frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
    frm1.txthInternalCd.value =  lgInternalCd
    frm1.txthSubInternalCd.value = lgSubInternalCd
    frm1.txthAuthUsrID.value = lgAuthUsrID



    Call CommonQueryRs("PROJECT_NO","A_TEMP_GL"," TEMP_GL_NO =  " & FilterVar(Trim(frm1.txtTempGlNo.value), "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
         frm1.txthpjt_no.value = Replace(lgF0, Chr(11), "")

    alert frm1.txthpjt_no.value
    
    Call ExecMyBizASP(frm1, BIZ_PGM_ID)					'저장 비지니스 ASP 를 가동

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk(ByVal TempGlNo)
	lgBlnFlgChgValue = false

	frm1.txtTempGlNo.value = UCase(Trim(TempGlNo))
    frm1.txtCommAndMode.value = "UPDATE"

	Call ggoOper.ClearField(Document, "2")
    Call InitVariables
	Call DbQuery()
End Function

'========================================================================================
Function DbDelete()
	Dim strVal
	
    Err.Clear
    Call LayerShowHide(1)
	DbDelete = False

	frm1.hOrgChangeId.value = parent.gChangeOrgId

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtTempGlNo=" & UCase(Trim(frm1.txtTempGlNo.value))
    strVal = strVal & "&txtDeptCd=" & UCase(Trim(frm1.txtDeptCd.value))
	strVal = strVal & "&txTorgChangeId=" & Trim(frm1.hOrgChangeId.value)
	strVal = strVal & "&txtTempGlDt=" & Trim(frm1.txttempgldt.text)

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function

'=======================================================================================================
Function DbDeleteOk()
	Call FncNew()	
End Function

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet(Row)
	With frm1
		ggoSpread.Source = frm1.vspdData
		.vspdData.Row	= Row
		.vspdData.Col	= C_DocCur
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_DocCur,C_ItemAmt, "A" ,"I","X","X")         
        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_DocCur,C_ExchRate,"D" ,"I","X","X")		
	End With
End Sub
    
'=======================================================================================================    
Sub SetGridFocus()	    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
End Sub
'=======================================================================================================  
Sub InputCtrlVal(ByVal Row)
	Dim strAcctCd
	Dim ii

	lgBlnFlgChgValue = True

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Col = C_AcctCd
	frm1.vspdData.Row = Row
	strAcctCd	= Trim(frm1.vspdData.text)

	frm1.vspdData.Col = C_deptcd
	frm1.vspdData.Row = Row

	Call AuToInputDetail(strAcctCd, Trim(frm1.vspdData.text), frm1.txttempGLDt.text, Row)

	For ii = 1 To frm1.vspdData2.MaxRows
		frm1.vspddata2.col = C_CtrlVal
		frm1.vspddata2.row = ii

		If Trim(frm1.vspddata2.text) <> "" Then
			Call CopyToHSheet2(frm1.vspdData.ActiveRow,ii)

			'frm1.vspddata2.col = C_HItemSeq
			'Call CopyToHSheet2(frm1.vspdData2.ActiveRow,ii)
		End If
	Next
End Sub
