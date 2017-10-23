'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID			= "a5104mb1.asp"			'☆: 비지니스 로직 ASP명
Const JUMP_PGM_ID_TAX_REP	= "a6114ma1"

'                       4.2 Constant variables 
'========================================================================================================
Const C_GLINPUTTYPE = "GL"

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
Dim lgFormLoad
Dim lgQueryOk
Dim lgstartfnc
Dim intItemCnt
Dim lgBlnExecDelete
Dim IsOpenPop

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'                        5.1 Common Method-1
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
     
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtGLDt.text = UniConvDateAToB(iDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)

    frm1.txtCommandMode.Value = "CREATE"
    frm1.cboGlInputType.Value = C_GLINPUTTYPE

	frm1.cboGlType.Value = "03"

	frm1.txtDeptCd.Value	= parent.gDepart
	frm1.hOrgChangeId.Value = parent.gChangeOrgId 

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

'		Call AppendNumberPlace("6","3","0")
        Call GetSpreadColumnPos("A")
        ggoSpread.SSSetFloat  C_ItemSeq,    " ", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_deptcd,     "부서코드",   10, , , 10, 2
        ggoSpread.SSSetButton C_deptpopup
        ggoSpread.SSSetEdit   C_deptnm,     "부서명",     17, , , 30
		ggoSpread.SSSetEdit   C_AcctCd,     "계정코드",   15, , , 18
		ggoSpread.SSSetButton C_AcctPopup
		ggoSpread.SSSetEdit   C_AcctNm,     "계정코드명", 20, , , 30
		ggoSpread.SSSetCombo  C_DrCrFg,     "", 8
	    ggoSpread.SSSetCombo  C_DrCrNm,     "차대구분",   11
		ggoSpread.SSSetEdit   C_DocCur,     "거래통화",   10, , , 10, 2
        ggoSpread.SSSetButton C_DocCurPopup
		ggoSpread.SSSetFloat  C_ExchRate,   "환율", 15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_ItemAmt,    "금액",       15, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt, "금액(자국)", 15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_IsLAmtChange,   "",     30, , , 128
		ggoSpread.SSSetEdit   C_ItemDesc,   "비  고",     30, , , 128
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

	end with
    SetSpreadLock "I", 0, 1, ""
        
End Sub

'=======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )

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
				ggoSpread.SpreadUnLock		C_VATNM			, -1    , C_VATNM
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			Case 1
				ggoSpread.SpreadLock C_deptcd		, -1    , C_deptcd
				ggoSpread.SpreadLock C_ItemSeq		, -1	, C_ItemSeq
				ggoSpread.SpreadLock C_deptpopup	, -1	, C_deptpopup
				ggoSpread.SpreadLock C_ItemLocAmt	, -1	, C_ItemLocAmt
				ggoSpread.SpreadLock C_DrCrNm		, -1    , C_DrCrNm
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

'========================================================================================
Sub SetSpread2Lock(Byval stsFg,Byval Index,ByVal lRow  ,ByVal lRow2 )

    With frm1
		ggoSpread.Source = .vspdData2
		If lRow = "" Then
			lRow = 1
		End If	
		If lRow2 = "" Then
			lRow2 = .vspdData2.MaxRows
		End If

		.vspdData2.Redraw = False
		Select Case Index
			Case 0
			Case 1
				ggoSpread.SpreadLock 1, lRow, .vspdData2.MaxCols, lRow2	
		End Select
		.vspdData2.Redraw = True

    End With
End Sub

'========================================================================================
Sub SetSpreadColor(Byval stsFg, Byval Index, ByVal lRow, ByVal lRow2)
    With frm1

		if  lRow2 = "" THEN	lRow2 = lRow

		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_ItemSeq, lRow, lRow2
		ggoSpread.SSSetProtected C_deptNm,    lRow, lRow2
		ggoSpread.SSSetProtected C_AcctNm, lRow, lRow2
		ggoSpread.SSSetRequired  C_deptcd,    lRow, lRow2

		Select Case stsFg
		Case "I"
			ggoSpread.SSSetRequired C_AcctCd, lRow, lRow2
		CASE "Q"
			ggoSpread.SSSetProtected C_AcctCd, lRow, lRow2	
		End Select

		IF  frm1.cboGlType.Value = "01" Or frm1.cboGlType.Value = "02" Then
			ggoSpread.SSSetProtected C_DrCrNm, lRow, lRow2
		ELSE
			ggoSpread.SSSetRequired C_DrCrNm, lRow, lRow2
		END IF

		ggoSpread.SSSetRequired  C_DocCur, lRow, lRow2
		ggoSpread.SSSetRequired C_ItemAmt, lRow, lRow2
		.vspdData.ReDraw = True
    End With
End Sub



'=========================================================================================================
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

'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet				'권한관리 추가
	
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
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "
			arrParam(5) = "계정코드"

			arrField(0) = "A_ACCT.Acct_CD"
			arrField(1) = "A_ACCT.Acct_NM"
   			arrField(2) = "A_ACCT_GP.GP_CD"
			arrField(3) = "A_ACCT_GP.GP_NM"

			arrHeader(0) = "계정코드"	
			arrHeader(1) = "계정코드명"
			arrHeader(2) = "그룹코드"	
			arrHeader(3) = "그룹명"
			
'		arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=A_ACCT_00", Array(Array(Trim(strCode))), _
'		           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	 
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If	

	Call FocusAfterPopup (iWhere)
End Function

'=======================================================================================================
Function FocusAfterPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 2
				Call SetActiveCell(.vspdData,C_DocCur,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				Call SetActiveCell(.vspdData,C_AcctCd,.vspdData.ActiveRow ,"M","X","X")
		End Select    
	End With
End Function

'========================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 2
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 

				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
				.vspdData.Col  = C_ItemLocAmt
				.vspdData.Text = ""
				.vspdData.Col  = C_DocCur 
				.vspdData.Text = UCase(Trim(arrRet(0)))
				If Trim(.vspdData.Text) = parent.gCurrency Then
					.vspdData.Col  = C_ExchRate
					.vspdData.Text = 1
				Else
					call FindExchRate(UniConvDateToYYYYMMDD(frm1.txtGLDt.text,parent.gDateFormat,""), UCase(Trim(arrRet(0))),frm1.vspdData.ActiveRow)
				End IF

				Call DocCur_OnChange(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)
			Case 3
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 

				.vspdData.Col  = C_AcctCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_AcctNm
				.vspdData.Text = arrRet(1)
                Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)
		End Select
	End With
End Function

'========================================================================================
Function OpenRefGL()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(8)	                           '권한관리 추가 (3 -> 4)

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5104ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5104ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	IsOpenPop = True
	Call CookiePage("GL_POPUP")
	arrParam(4)	= lgAuthorityFlag 
	
	' 권한관리 추가
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = ""  Then
		frm1.txtGLNo.focus 
		Exit Function
	Else
		Call SetRefGL(arrRet)
	End If
End Function

Function SetRefGL(ByRef arrRet)
	
	With frm1
		.txtGlNo.Value = UCase(Trim(arrRet(0)))
    End With

	frm1.txtGLNo.focus 
End Function

'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo,varOrgChangeId)
	Dim intRetCd

	StrEbrFile = "a5121ma1"

	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtGlDt.Text, parent.gDateFormat,"")	
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtGlDt.Text, parent.gDateFormat,"")	
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	varGlNoFr = Trim(frm1.txtGlNo.Value)
	varGlNoTo = Trim(frm1.txtGlNo.Value)
	varOrgChangeId = Trim(frm1.hOrgChangeId.Value)	
End Sub

'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId
    Dim StrEbrFile
    Dim intRetCd
	Dim ObjName
	
    If Not chkFieldByCell(frm1.txtGlNo,"A",1) Then Exit Function
    	
'    If Not chkField(Document, "1") Then	
'       Exit Function
'    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId)

    lngPos = 0

	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|GlNoFr|" & varGlNoFr
	StrUrl = StrUrl & "|GlNoTo|" & varGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)

End Function

'========================================================================================
Function FncBtnPreview() 
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
	Dim ObjName

    If Not chkFieldByCell(frm1.txtGlNo,"A",1) Then Exit Function

'    If Not chkField(Document, "1") Then
'       Exit Function
'    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId)

    StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|GlNoFr|" & varGlNoFr
	StrUrl = StrUrl & "|GlNoTo|" & varGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
End Function

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
		strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrTemp = Split(lgF0, chr(11))
			strExchFg =  arrTemp(0)
		End If

		strDate = UniConvDateToYYYYMMDD(frm1.txtGLDt.text,parent.gDateFormat,"")
		If .vspdData.MaxRows <> 0 Then
			For ii = 1 To .vspdData.MaxRows
				.vspdData.Row	=	ii
				.vspdData.Col	=	C_DocCur
				tempDoc			=	UCase(Trim(.vspdData.text))
				.vspdData.Col	=	C_ItemAmt
				tempAmt			=	UNICDbl(.vspdData.text)
				.vspdData.Col	=	C_ExchRate
				tempExch		=	UNICDbl(.vspdData.text)

				If tempDoc	<> "" and tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
						End If
					End If
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt		=	tempAmt * TempExch
					Else
						tempLocAmt		=	tempAmt / TempExch
					End If
					.vspdData.Col	=	C_ItemLocAmt
					.vspdData.text	=	UNIConvNumPCToCompanyByCurrency(tempLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")

				ElseIf tempDoc = parent.gCurrency Then
					.vspdData.Col	=	C_ItemLocAmt
					.vspdData.text	=UNIConvNumPCToCompanyByCurrency(tempAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
					
				End If
			Next
		End If
	End With

	Call SetSumItem	
End Function

'========================================================================================
Function OpenDept(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(8)
	
	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.readOnly = true then
		IsOpenPop = False
		Exit Function
	End If
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	IsOpenPop = True

	arrParam(0) = strCode									'  Code Condition
   	arrParam(1) = frm1.txtGLDt.Text
	arrParam(2) = lgUsrIntCd								' 자료권한 Condition  
	If lgIntFlgMode = parent.OPMD_UMODE then
		arrParam(3) = "T"									' 결의일자 상태 Condition  
	Else
		arrParam(3) = "F"									' 결의일자 상태 Condition  
	End If

	' 권한관리 추가
	'arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	'arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
    If arrRet(0) <> "" Then
		Call SetDept(arrRet, iWhere)
	End If

	Call FocusAfterDeptPopup (  iWhere)

End Function

'========================================================================================

Function OpenUnderDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

	IsOpenPop = True
	If RTrim(LTrim(frm1.txtDeptCd.Value)) <> "" 	Then
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.Value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND B.BIZ_AREA_CD = ( SELECT B.BIZ_AREA_CD"
		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.Value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.Value , "''", "S") & ")"

		' 권한관리 추가
		If lgInternalCd <>  "" Then
			arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD =" & FilterVar(lgInternalCd, "''", "S")			' Where Condition
		End If

		If lgSubInternalCd <>  "" Then
			arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
		End If
		
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
		arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

		' 권한관리 추가
		If lgInternalCd <>  "" Then
			arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD =" & FilterVar(lgInternalCd, "''", "S")			' Where Condition
		End If

		If lgSubInternalCd <>  "" Then
			arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
		End If
	
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

	Call FocusAfterDeptPopup (  iWhere)
End Function

'========================================================================================
Function SetDept(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case "0"
				.txtDeptCd.Value = arrRet(0)
				.txtDeptNm.Value = arrRet(1)
				.txtInternalCd.Value = arrRet(2)
				If lgQueryOk <> True Then
					.txtGLDt.text = arrRet(3)
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
Function FocusAfterDeptPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtDeptCd.focus
			Case 1 
				Call SetActiveCell(.vspdData,C_deptcd,.vspdData.ActiveRow ,"M","X","X")
		End Select    
	End With

End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++

'=======================================================================================================
Function SetSumItem()
    Dim DblTotDrAmt 
    Dim DblTotLocDrAmt
    Dim DblTotCrAmt
    Dim DblTotLocCrAmt

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
			            .Col = C_ItemAmt	'6
			            If .Text = "" Then
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
			            Else
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
			            End If

			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
			            Else
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
			            End If
		            Elseif .text = "CR" then
			            .Col = C_ItemAmt	'6
			            If .Text = "" Then
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
			            Else
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
			            End If

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

		frm1.txtDrLocAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
		frm1.txtCrLocAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	End With	
	
'    If frm1.cboGlType.value = "01" Then
'		frm1.txtDrLocAmt.text = frm1.txtCrLocAmt.text
'	ElseIF frm1.cboGlType.value = "02" Then
'		frm1.txtCrLocAmt.text = frm1.txtDrLocAmt.text
'	End If	
End Function

'========================================================================================================
Function CookiePage(ByVal Kubun)
	Dim strTemp
	Dim strNmwhere
	Dim arrVal
	
	Select Case Kubun
		Case "FORM_LOAD"
			strTemp = ReadCookie("GL_NO")

			Call WriteCookie("GL_NO", "")

			If strTemp = "" then Exit Function

			frm1.txtGlNo.Value = strTemp

			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("GL_NO", "")
				Exit Function 
			End If

			Call FncQuery()
		Case JUMP_PGM_ID_TAX_REP
			ggoSpread.Source = frm1.vspdData

			If frm1.vspddata.MaxRows	< 1  Then
				Exit Function
			End IF

			frm1.vspddata.row = frm1.vspddata.ActiveRow	
			frm1.vspddata.Col = C_VatType

			If frm1.vspddata.Value	=	"" Then
				Exit Function
			End IF

			frm1.vspddata.Col = C_ItemSeq

			strNmwhere = " GL_NO  = " & FilterVar(frm1.txtGlNo.Value, "''", "S")
			strNmwhere = strNmwhere & " AND ITEM_SEQ = " & frm1.vspddata.text & " "

			IF CommonQueryRs( "VAT_NO" , "A_VAT" ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
				arrVal = Split(lgF0, Chr(11))
				strTemp = arrVal(0)
			End IF

			Call WriteCookie("VAT_NO", strTemp)	
		Case "GL_POPUP"
			Call WriteCookie("PGMID", "A5104MA1")

		Case Else
			Exit Function
	End Select
End Function

'========================================================================================================
'	Desc : 화면이동
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD

    ggoSpread.Source = frm1.vspdData    
    If (lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True ) And C_GLINPUTTYPE = frm1.cboGlInputType.Value Then
		IntRetCD = DisplayMsgBox("990027", "X", "X", "X")
        Exit Function
    End If

	Select Case strPgmId
		Case JUMP_PGM_ID_TAX_REP

			ggoSpread.Source = frm1.vspdData

			If frm1.vspddata.MaxRows	< 1  Then
				IntRetCD = DisplayMsgBox("900002", "X","X","X")	
				Exit Function
			End IF


			frm1.vspddata.row = frm1.vspddata.ActiveRow	
			frm1.vspddata.Col = C_VatType	

			If frm1.vspddata.Value	=	"" Then
				IntRetCD = DisplayMsgBox("205600", "X","X","X")	
				Exit Function
			End IF
	End Select

	Call CookiePage(strPgmId)
	Call PgmJump(strPgmId)
End Function

'========================================================================================================
'	Desc : 입출금 화면에 따른 Grid의 Protect변환
'========================================================================================================
Sub CboGLType_ProtectGrid(Byval GlType)
	ggoSpread.Source = frm1.vspdData
	Select Case GlType
		case "01"
'			ggoSpread.SSSetProtected C_DocCur, 1, frm1.vspddata.maxrows
'			ggoSpread.SSSetProtected C_DocCurPopup, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows
		Case "02"
'			ggoSpread.SSSetProtected C_DocCur, 1, frm1.vspddata.maxrows
'			ggoSpread.SSSetProtected C_DocCurPopup, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows
		Case "03"
			ggoSpread.SSSetRequired C_DocCur, 1, frm1.vspddata.maxrows
			ggoSpread.SpreadUnLock C_DocCurPopup, 1, frm1.vspddata.maxrows
			ggoSpread.SpreadUnLock C_DrCrfg, 1, C_DrCrNm, frm1.vspddata.maxrows
			ggoSpread.SSSetRequired C_DrCrfg, 1, frm1.vspddata.maxrows
			ggoSpread.SSSetRequired C_DrCrNm, 1, frm1.vspddata.maxrows
	END Select
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
			 C_ItemSeq			= iCurColumnPos(1)
			 C_deptcd			= iCurColumnPos(2)
			 C_deptPopup		= iCurColumnPos(3)
			 C_deptnm	   		= iCurColumnPos(4)
			 C_AcctCd			= iCurColumnPos(5)
			 C_AcctPopup		= iCurColumnPos(6)
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
			 C_VatNm			= iCurColumnPos(18)
			 C_AcctCd2			= iCurColumnPos(19)
    End Select    
End Sub

'=======================================================================================================
Sub vspdData_onfocus()
	lgCurrRow = frm1.vspdData.ActiveRow
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolbar(MENU_CRT)
        If frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetToolbar(MENU_PRT)
		Else
			Call SetToolbar(MENU_CRT)
		End if
    Else
        If frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetToolbar(MENU_PRT)
		Else
			Call SetToolbar(MENU_UPD)
		End If
    End If  
End Sub

'=======================================================================================================
Sub txtGLDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGLDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtGLDt.focus   
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
	end if
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    If Row <= 0 Then
    End If
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then

        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text
            ggoSpread.Source = frm1.vspdData2
            ggoSpread.ClearSpreadData
        End With

		frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End if

		lgCurrRow = NewRow
        Call DbQuery2(lgCurrRow)
    End If
End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim iFld1
	Dim iFld2
	Dim iTable
	Dim istrCode

	With frm1.vspdData
		If Row > 0 And Col = C_AcctPopUp Then
			.Col = Col - 1
			.Row = Row

			Call OpenPopUp(.Text, 3)
		End If

		If Row > 0 And Col = C_deptPopup Then
			.Col = Col - 1
			.Row = Row
			Call OpenUnderDept(.Text, 1)
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
				frm1.hItemSeq.Value = frm1.vspdData.Text
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
			End IF
		Case	C_DocCur
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_ItemLocAmt
			frm1.vspdData.Text = ""		
			frm1.vspdData.Col = C_DocCur
			If UCase(Trim(frm1.vspdData.Text)) = parent.gCurrency Then
				frm1.vspdData.Col = C_ExchRate
				frm1.vspdData.Text = 1
			Else
				Call FindExchRate(UniConvDateToYYYYMMDD(frm1.txtGLDt.text,parent.gDateFormat,""), UCase(Trim(frm1.vspdData.Text)),frm1.vspdData.ActiveRow)
			End IF
			
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
				SetSpread2Color

			Case C_VatNm
				.vspddata.Col = Col
			    intIndex = .vspddata.Value
				.vspddata.Col = C_VatType
				.vspddata.Value = intIndex
				Call InputCtrlVal(Row)
		End Select
	End With
End Sub

'==========================================================================================
Sub txtGlNo_OnKeyPress()	
	If window.event.keycode = 39 then
		window.event.keycode = 0	
	End If
End Sub

'==========================================================================================
Sub txtGlNo_OnKeyUp()	
	If Instr(1,frm1.txtGlNo.Value,"'") > 0 then
		frm1.txtGlNo.Value = Replace(frm1.txtGlNo.Value, "'", "")
	End if
End Sub

'==========================================================================================
Sub txtGlNo_onpaste()
	Dim iStrGlNo
	iStrGlNo = window.clipboardData.getData("Text")
	iStrGlNo = RePlace(iStrGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrGlNo)
End Sub

'==========================================================================================
Sub DocCur_OnChange(FromRow, ToRow)
	Dim ii
    lgBlnFlgChgValue = True

	For ii = FromRow	to	ToRow
		frm1.vspdData.Row	= ii
		frm1.vspdData.Col	= C_DocCur

		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.vspdData.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			Call CurFormatNumericOCX(ii)
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

	If Trim(frm1.txtGLDt.Text = "") Or Trim(frm1.txtDeptCd.value) = "" Then
		Exit sub
    End If

    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "
	strFrom		=			 " b_acct_dept(NOLOCK) "
	strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"	

'		' 권한관리 추가
'		If lgInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
'		End If
'	
'		If lgSubInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
'		End If


	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then

		IntRetCD = DisplayMsgBox("124600","X","X","X")
		frm1.txtDeptCd.Value = ""
		frm1.txtDeptNm.Value = ""
		frm1.hOrgChangeId.Value = ""
	Else
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			frm1.hOrgChangeId.Value = Trim(arrVal2(2))
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
	Dim jj

	If Trim(frm1.txtGLDt.Text = "") Then
		Exit sub
    End If
    
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "
	strFrom		=			 " b_acct_dept(NOLOCK) "
	strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S") 
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

'		' 권한관리 추가
'		If lgInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
'		End If
'	
'		If lgSubInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
'		End If


	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		frm1.txtDeptCd.Value = ""
		frm1.txtDeptNm.Value = ""
		frm1.hOrgChangeId.Value = ""
	Else
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			frm1.hOrgChangeId.Value = Trim(arrVal2(2))
		Next
	End If
End Sub

'==========================================================================================
Sub DeptCd_underChange(Byval strCode)
    Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD

    If Trim(frm1.txtGLDt.Text = "") Then
		Exit sub
    End If
    
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "
	strFrom		=			 " b_acct_dept(NOLOCK) "
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S") 
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

'		' 권한관리 추가
'		If lgInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
'		End If
'	
'		If lgSubInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
'		End If


	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  

		frm1.vspdData.Col = C_deptcd
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.text = ""
		frm1.vspdData.Col = C_deptnm
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = ""
	End If
End Sub

'==========================================================================================
Sub txtGLDt_Change()
	If lgstartfnc = False Then
		If lgFormLoad = True Then
			Dim strSelect,strFrom,strWhere
			Dim IntRetCD
			Dim ii,jj
			Dim arrVal1,arrVal2

			lgBlnFlgChgValue = True

			With frm1
				If LTrim(RTrim(.txtDeptCd.Value)) <> "" and Trim(.txtGLDt.Text <> "") Then
					strSelect	=			 " dept_cd, org_change_id, internal_cd "
					strFrom		=			 " b_acct_dept(NOLOCK) "
					strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(.txtDeptCd.Value)), "''", "S") 
					strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "
					strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
					strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"
	
					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
						IntRetCD = DisplayMsgBox("124600","X","X","X")
						.txtDeptCd.Value = ""
						.txtDeptNm.Value = ""
						.hOrgChangeId.Value = ""
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
						jj = Ubound(arrVal1,1)
						For ii = 0 to jj - 1
							arrVal2 = Split(arrVal1(ii), chr(11))
							frm1.hOrgChangeId.Value = Trim(arrVal2(2))
						Next
					End If 
				End If
			End With
		End If
	End IF
End Sub

'==========================================================================================
Sub cboGLType_OnChange()
	Dim	i
	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData

	Select Case frm1.cboGlType.Value 
		Case "01"
			'입금전표로 바꾸면 차변이 입력되거나 현금계정이 입력되었는지 check한다.
			For i = 1 To frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				IF  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.Value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.Value )
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")
					Exit sub
				End IF

				frm1.vspddata.col = C_DrCrFg
				IF  Trim(frm1.vspddata.Value) = "2" Then
					frm1.cboGlType.Value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.Value )
					IntRetCD = DisplayMsgBox("113104", "X", "X", "X")
					Exit sub
				End IF
			Next

			For i = 1 To frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				IF Trim(frm1.vspddata.Value) <> "1"  Then
					frm1.vspdData.Value	= "1"
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.Value	= "1"
				END IF
				
				Call vspdData_ComboSelChange(C_DrCrNm,i)
				frm1.vspddata.col = C_DocCur
				frm1.vspddata.text = parent.gCurrency
			Next
			
			Call CboGLType_ProtectGrid(frm1.cboGlType.Value )
		Case "02"
			'출금전표로 바꾸면 대변이 입력되거나 현금계정이 입력되었는지 check한다.	
			For i = 1 To frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				If  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.Value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.Value )
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")
					Exit sub
				End If

				frm1.vspddata.col = C_DrCrFg
				If  Trim(frm1.vspddata.Value) = "1" Then
					frm1.cboGlType.Value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.Value )
					IntRetCD = DisplayMsgBox("113105", "X", "X", "X")
					Exit sub
				End If
			Next

			For i = 1 To  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				If Trim(frm1.vspddata.Value) <> "2"  Then
					frm1.vspdData.Value	= "2"
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.Value	= "2"
				End If
				
				Call vspdData_ComboSelChange(C_DrCrNm,i)
				frm1.vspddata.col = C_DocCur
				frm1.vspddata.text = parent.gCurrency
			Next
			
			Call CboGLType_ProtectGrid(frm1.cboGlType.Value )
		Case "03"
		'대체로 바꾸면 Protect를 풀어준다.
			Call CboGLType_ProtectGrid(frm1.cboGlType.Value )
	End Select	

	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub

'#########################################################################################################
'												4. Common Function부
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수
'######################################################################################################### 



'#########################################################################################################
'												5. Interface부
'	기능: Interface
'######################################################################################################### 


'========================================================================================
Function FncQuery() 
    Dim IntRetCD
    Dim RetFlag
    lgstartfnc = True
    FncQuery = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData

    If Not chkFieldByCell(frm1.txtGlNo,"A",1) Then Exit Function

'    If Not chkField(Document, "1") Then
'       Exit Function
'    End If
    
    If lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True Then
		IntRetCD = ggoOper.DisplayMsgBox ("900013", parent.VB_YES_NO, "X", "X")
    	If IntRetCD = vbNo Then
      	Exit Function
     	End If
    End If

'    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    Call InitVariables

    If frm1.txtDeptCd.Value = "" Then
		frm1.txtDeptNm.Value = ""
    End If

    IF DbQuery = False Then	
		Exit Function
    End If

'    If frm1.vspddata.maxrows = 0 Then
'		frm1.txtGlNo.Value = ""
'    End If

    FncQuery = True
    lgstartfnc = False
End Function

'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    Dim var1, var2

    lgstartfnc = True
    FncNew = False

    On Error Resume Next
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    If (lgBlnFlgChgValue = True Or var1 = True Or var2 = True) And lgBlnExecDelete <> True Then
        IntRetCD = ggoOper.DisplayMsgBox ("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    lgBlnExecDelete = False
    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData

'    Call ggoOper.LockField(Document,  "N")
    SetGridFocus()
    SetGridFocus2()

	Call SetDefaultVal
    Call InitVariables
	Call SetSumItem()

    Call SetToolbar(MENU_NEW)
	Call ggoOper.SetReqAttr(frm1.txtGlDt,   "N")
	Call ggoOper.SetReqAttr(frm1.txtDeptCd, "N")
	Call ggoOper.SetReqAttr(frm1.cboGlType, "N")	
	Call ggoOper.SetReqAttr(frm1.txtdesc,	"D")

    lgBlnFlgChgValue = False

    FncNew = True
    lgFormLoad = True
    lgQueryOk = False
    lgstartfnc = False
End Function

'========================================================================================
Function FncDelete() 
	Dim IntRetCD 

    FncDelete = False
    Err.Clear
    lgBlnExecDelete = True
    On Error Resume Next

	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
		intRetCd = ggoOper.DisplayMsgBox ("990008", parent.VB_YES_NO, "X", "X")
		If intRetCd = VBNO Then
			Exit Function
		End IF
    Else
		IntRetCD = ggoOper.DisplayMsgBox ("900038", parent.VB_YES_NO, "X", "X")
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

    IF  DbDelete = False Then
		Exit Function
	END IF

    FncDelete = True
End Function

'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 

    FncSave = False

    Err.Clear
    
	With frm1
		ggoSpread.Source = .vspdData

		If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then
		    IntRetCD = ggoOper.DisplayMsgBox ("900001", "X", "X", "X")
		    Exit Function
		End If

		if CheckSpread3 = False then
		IntRetCD = ggoOper.DisplayMsgBox ("110420", "X", "X", "X")
		    Exit Function
		end if

		If frm1.vspdData.MaxRows < 1 Then
			IntRetCD = ggoOper.DisplayMsgBox ("113100", "X", "X", "X")
			Exit Function
		End If

		ggoSpread.Source = .vspdData
    
	    If Not chkFieldByCell(.txtGLDt, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboGlType, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.txtDeptCd, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboGlInputType, "A", "1") Then Exit Function
	    
	    If Not ChkFieldLengthByCell(.txtDesc, "A", "1") Then Exit Function        
    
'		If Not chkField(Document, "2") Then
'			Exit Function
'		End If

		If Not ggoSpread.SSDefaultCheck Then
		   Exit Function
		End If

		ggoSpread.Source = .vspdData3
		
		If Not ggoSpread.SSDefaultCheck Then
			Exit Function
		End If

		IF DbSave = False Then
			Exit Function
		End If

		FncSave = True
	End With		
End Function

'========================================================================================
Function FncCopy() 
	Dim  IntRetCD
	
	frm1.vspdData.ReDraw = False
	If frm1.vspdData.MaxRows < 1 Then Exit Function	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor "I", 0, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    MaxSpreadVal frm1.vspdData, C_ItemSeq, frm1.vspdData.ActiveRow
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DocCur,C_ItemAmt, "A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DocCur,C_ExchRate,"D" ,"I","X","X")
	Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)
    Call SetSumItem()
End Function

'========================================================================================================
Function FncCancel() 
    Dim iItemSeq
    Dim RowDocCur

	If frm1.vspdData.MaxRows < 1 Then 	Exit Function

	if  frm1.vspdData.MaxRows = 1 Then  Call ggoOper.SetReqAttr(frm1.cboGlType,   "N")	
    With frm1.vspdData
        .Row = .ActiveRow
        .Col = 0
        If .Text = ggoSpread.InsertFlag Then
			.Col = C_AcctCd
			If Len(Trim(.text)) > 0 Then 
				.Col = C_ItemSeq
				DeleteHSheet(.Text)
			end if
        End if

        ggoSpread.Source = frm1.vspdData
        ggoSpread.EditUndo

        Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_DocCur,C_ItemAmt,  "A" ,"I","X","X")
        Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_DocCur,C_ExchRate, "D" ,"I","X","X")

		If .MaxRows = 0 Then
			Call SetToolbar(MENU_NEW)
			Exit Function
		End If

        InitData

        .Row = .ActiveRow
        .Col = 0
		if .row = 0 then 
			Exit Function
		end if

        If .Text = ggoSpread.InsertFlag Then
            .Col = C_AcctCd
            If Len(.Text) > 0 Then
				.Col = C_ItemSeq
				frm1.hItemSeq.Value = .Text
	            frm1.vspdData2.MaxRows = 0
		        Call DbQuery3(.ActiveRow)
            End If
        Else
            .Col = C_ItemSeq
            frm1.hItemSeq.Value = .Text
            frm1.vspdData2.MaxRows = 0
		    Call DbQuery2(.ActiveRow)
        End if
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

	With frm1
		FncInsertRow = False

	    If Not chkFieldByCell(.txtGLDt, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboGlType, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.txtDeptCd, "A", "1") Then Exit Function
	    If Not chkFieldByCell(.cboGlInputType, "A", "1") Then Exit Function
	    
'		If Not chkField(Document, "2") Then 
'		    Exit Function
'		End If

		If IsNumeric(Trim(pvRowCnt)) then
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

        For imRow2 = 1 to imRow 
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
            IF  frm1.cboGlType.value = "01" Then
                .vspdData.col = C_DrCrNm
                .vspdData.value	= 1
                .vspdData.col = C_DrCrFg
                .vspdData.value	= 1
            ELSEIF frm1.cboGlType.value = "02" Then
                .vspdData.col = C_DrCrNm
                .vspdData.value	= 2
                .vspdData.col = C_DrCrFg
                .vspdData.value	= 2
            END IF
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
Sub PopRestoreSpreadColumnInf()
	Dim indx

	On Error Resume Next
	Err.Clear

	ggoSpread.Source = gActiveSpdSheet

    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet()
            Call InitComboBoxGrid
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,1,-1,C_DocCur,C_ItemAmt,  "A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,1,-1,C_DocCur,C_ExchRate, "D" ,"I","X","X")			

			If frm1.cboGlInputType.Value <> C_GLINPUTTYPE then
			    Call SetSpreadLock("Q", 1, 1, "")
			    Call SetSpread2Lock("",1,"","")
			Else
                Call SetSpreadColor("Q", 0,1, .vspddata.MaxRows)
                Call SetSpread2Color()
			End if

		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			If frm1.cboGlInputType.Value <> C_GLINPUTTYPE then
			    Call SetSpread2Lock("",1,"","")
			Else
                Call SetSpread2Color()
			End if
	End Select

	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If
	Call SetSumItem()
End Sub

'=======================================================================================================
Sub PrevspdDataRestore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 to frm1.vspdData.MaxRows
        frm1.vspdData.Row    = indx
        frm1.vspdData.Col    = 0

		If frm1.vspdData.Text <> "" Then
			Select Case frm1.vspdData.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData.Col = C_ItemSeq
					Call DeleteHsheet(frm1.vspdData.Text)
				Case ggoSpread.UpdateFlag
					For indx1 = 0 to frm1.vspdData3.MaxRows
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case frm1.vspdData3.Text 
							Case ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1
								If UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)
									Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtGLNo.Value)
								End If
						End Select
					Next
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtGLNo.Value)
			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

'=======================================================================================================
Sub PrevspdData2Restore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 to frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
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
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.htxtGLNo.Value)
						End If
					Next
				Case ggoSpread.DeleteFlag
			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

'========================================================================================================
' Name : fncRestoreDbQuery2
' Desc : This function is data query and display
'========================================================================================================
Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal
'	Dim tmpDrCrFG	

	on Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)

	With frm1
		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text
		
	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If

		Call LayerShowHide(1)

		DbQuery2 = False

		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq

		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
'		strSelect = strSelect & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'		strSelect = strSelect & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  		
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & strItemSeq & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_GL_DTL C (NOLOCK), A_GL_ITEM D (NOLOCK) "

		strWhere =			  " D.GL_NO = " & FilterVar(UCase(pInvalue1), "''", "S")
		strWhere = strWhere & " AND D.ITEM_SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.GL_NO  =  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
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
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
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
		Call RestoreToolBar()
	End With

	If Err.number = 0 Then
		fncRestoreDbQuery2 = True
	End If
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True OR ggoSpread.SSCheckChange = True Then
		IntRetCD = ggoOper.DisplayMsgBox ("900016", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    end if
    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim RetFlag

    DbQuery = False
    Call LayerShowHide(1)
    frm1.vspdData3.MaxRows = 0 

    Err.Clear

    With frm1
	    If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtGlNo=" & UCase(Trim(.htxtGlNo.Value))
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 
			strVal = strVal & "&txtGlNo=" & UCase(Trim(.txtGlNo.Value))
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
		End If
		
		' 권한관리 추가
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인
	
		Call RunMyBizASP(MyBizASP, strVal)
    End With    
    DbQuery = True
End Function


'=======================================================================================================
Function DbQueryOk()
	Dim ii

	With frm1
        lgIntFlgMode = parent.OPMD_UMODE

		Call ggoOper.SetReqAttr(frm1.txtGLDt,	"Q")
		Call ggoOper.SetReqAttr(frm1.cboGlType,	"Q")

        If frm1.cboGlInputType.Value <> C_GLINPUTTYPE then
			Call SetToolbar(MENU_PRT) 
			Call SetSpreadLock("Q", 1, 1, "")
			Call ggoOper.SetReqAttr(frm1.txtDeptCd,	"Q")
			Call ggoOper.SetReqAttr(frm1.txtdesc,   "Q")
		Else
			Call SetToolbar(MENU_UPD)
			Call SetSpreadLock("Q", 0, 1, "")
			Call SetSpreadColor("Q", 0,1, .vspddata.MaxRows)
			Call ggoOper.SetReqAttr(frm1.txtDeptCd,	"N")
			Call ggoOper.SetReqAttr(frm1.txtdesc,	"D")
		End if

        .txtCommandMode.Value = "UPDATE"

        Call InitData()

        For ii= 1 To .vspdData.MaxRows	
			CurFormatNumSprSheet(ii)
		Next

        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text
            Call DbQuery2(1)
        End If
    End With

	SetGridFocus()
    SetGridFocus2()
	Call QueryDeptCd_OnChange()
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
	
	on error resume next
	
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

        If CopyFromData(.hItemSeq.Value) = True Then
			If frm1.cboGlInputType.Value <> C_GLINPUTTYPE then
				Call SetSpread2Lock("",1,"","")
			Else
				Call SetSpread2Color()
			End If
            Exit Function
        End If

		Call LayerShowHide(1)

		DbQuery2 = False

		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq

		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , LTrim(ISNULL(C.CTRL_VAL,'')), '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
'		strSelect = strSelect & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'		strSelect = strSelect & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  		
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & .hItemSeq.Value & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_GL_DTL C (NOLOCK), A_GL_ITEM D (NOLOCK) "

		strWhere =			  " D.GL_NO = " & FilterVar(UCase(.txtGLNo.Value), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.GL_NO  =  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "

		frm1.vspdData2.ReDraw = False

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			ggoSpread.Source = frm1.vspdData2
			arrTemp =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp,Chr(12))
			ggoSpread.SSShowData lgF2By2

			For lngRows = 1 To frm1.vspdData2.Maxrows
				frm1.vspddata2.row = lngRows
				frm1.vspddata2.col = C_Tableid
				IF Trim(frm1.vspddata2.text) <> "" Then

					frm1.vspddata2.col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspddata2.col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspddata2.col = C_ColNm
					strColNm = frm1.vspddata2.text
					frm1.vspddata2.col = C_MajorCd
					strMajorCd = frm1.vspddata2.text

					frm1.vspddata2.col = C_CtrlVal

					strNmwhere = strColid & " =  " & FilterVar(UCase(frm1.vspddata2.text), "''", "S")

					IF Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") 
					End IF
					
					IF CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))
						frm1.vspddata2.text = arrVal(0)
					End IF
				End IF

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
			NEXT

			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal	
		END IF

		intItemCnt = .vspddata.MaxRows
		If frm1.cboGlInputType.Value <> C_GLINPUTTYPE then
			Call SetSpread2Lock("",1,"","")
		Else
			Call SetSpread2Color()
		End If
	End With

	Call LayerShowHide(0)
	frm1.vspdData2.ReDraw = True
	DbQuery2 = True
	lgQueryOk = True
End Function

Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim intIndex2 

	With frm1.vspdData

		For intRow = 1 To .MaxRows
			.Row = intRow
			.Col = C_DrCrFg
			intIndex = .Value
			.col = C_DrCrNm
			.Value = intindex
			.Col = C_VatType
			intIndex2 = .Value
			.col = C_VatNm
			.Value = intIndex2
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
    Dim strNote
    Dim strItemDesc

    DbSave = False
    Call LayerShowHide(1)

    Call SetSumItem

	With frm1
		.txtFlgMode.Value = lgIntFlgMode
		.txtUpdtUserId.Value = parent.gUsrID
		.txtInsrtUserId.Value  = parent.gUsrID
		.txtMode.Value = parent.UID_M0002
		.txtAuthorityFlag.Value     = lgAuthorityFlag               '권한관리 추가
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

		        .Col = C_ItemAmt		'5
		        strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep

				.Col = C_ItemLocAmt	'6
				strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep

		        .Col = C_ItemDesc	'7
				strItemDesc = Trim(.Text)

				If Trim(strItemDesc) = "" Or isnull(strItemDesc) Then
					 ggoSpread.Source = frm1.vspdData3
					 frm1.vspdData.Col = C_ItemSeq
					 tempItemSeq = frm1.vspdData.Text  
					 strNote = ""
					 With frm1.vspdData3
							For itemRows = 1 to frm1.vspdData3.MaxRows
								.Row = itemRows
								.Col = 1

								if .Text =  tempItemSeq then 
									.Col= 9 'C_Tableid	+ 1
									IF 	.Text = "B_BIZ_PARTNER" OR .Text = "B_BANK" OR .Text = "F_DPST" THEN
										.Col = 7 'C_CtrlValNm + 1 
									ELSE
										.Col = 5 'C_CtrlVal + 1 
									END IF
									strNote = strNote & C_NoteSep & Trim(.Text)
								end if
							Next
							strNote = Mid(strNote,2)
					 End With

					 strVal = strVal & strNote & parent.gColSep

					 ggoSpread.Source = frm1.vspdData
				Else
					strVal = strVal & strItemDesc & parent.gColSep		'8
				End if

				.Col = C_ExchRate	'9
		        strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep

		        .Col = C_VatType	'10
		        strVal = strVal & Trim(.Text) & parent.gColSep

		        .Col = C_DocCur		'11
		        strVal = strVal & UCase(Trim(.Text)) & parent.gRowSep

		        lGrpCnt = lGrpCnt + 1

		End If
    Next
    End With

    frm1.txtMaxRows.Value = lGrpCnt-1								'Spread Sheet의 변경된 최대갯수
    frm1.txtSpread.Value =  strVal									'Spread Sheet 내용을 저장

	IF frm1.txtSpread.Value = "" Then
		intRetCd = ggoOper.DisplayMsgBox ("990008", parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분
		If intRetCd = VBNO Then
			Exit Function
		End IF	
		IF  DbDelete = False Then
			Exit Function
		End IF
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.ClearSpreadData
		Call InitVariables
		Exit Function
	END IF

    lGrpCnt = 1
    strVal = ""

    ggoSpread.Source = frm1.vspdData3

    With frm1.vspdData3      ' Dtl 저장 

	For itemRows = 1 To frm1.vspdData.MaxRows 
 	    frm1.vspdData.Row = itemRows
	    frm1.vspdData.Col = 0

		if frm1.vspdData.Text <> ggoSpread.DeleteFlag then	

	        frm1.vspdData.Col = C_ItemSeq
		    tempItemSeq = frm1.vspdData.Text  

		    For lngRows = 1 To .MaxRows

				.Row = lngRows
				.Col = 1

				IF .text = tempitemseq THEN
	                .Col = 0 

					strVal = strVal & "C" & parent.gColSep

					.Col = 1 		 					'ItemSEQ	
					strVal = strVal & tempitemseq & parent.gColSep

					.Col =  2 'C_DtlSeq + 1   				'Dtl SEQ
					strVal = strVal & Trim(.Text) & parent.gColSep

					.Col =  3 'C_CtrlCd + 1		 		'관리항목코드
					strVal = strVal & Trim(.Text) & parent.gColSep

					.Col = 5 'C_CtrlVal + 1				'관리항목 Value 
					strVal = strVal & UCase(Trim(.Text)) & parent.gRowSep	

					lGrpCnt = lGrpCnt + 1
				End IF
	    	Next
		End If
    Next

    End With

    frm1.txtMaxRows3.Value = lGrpCnt-1							'Spread Sheet의 변경된 최대갯수
    frm1.txtSpread3.Value  = strVal								'Spread Sheet 내용을 저장
    
    frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
    frm1.txthInternalCd.value =  lgInternalCd
    frm1.txthSubInternalCd.value = lgSubInternalCd
    frm1.txthAuthUsrID.value = lgAuthUsrID    
    
    Call ExecMyBizASP(frm1, BIZ_PGM_ID)							'저장 비지니스 ASP 를 가동

    DbSave = True
End Function

'========================================================================================

Function DbSaveOk(Byval GlNo)
	frm1.txtGlNo.Value = UCase(Trim(GlNo))
    frm1.txtCommandMode.Value = "UPDATE"
	Call ggoOper.ClearField(Document, "2")
    Call InitVariables
	DbQuery
End Function

'========================================================================================

Function DbDelete()
	Dim strVal
	
    Err.Clear
    Call LayerShowHide(1)
	DbDelete = False

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtGlNo=" & UCase(Trim(frm1.txtGlNo.Value))
    strVal = strVal & "&txtGlDt=" & ggoOper.RetFormat(frm1.txtGLDt.Text, "yyyy-MM-dd")
    strVal = strVal & "&txtDeptCd=" & UCase(Trim(frm1.txtDeptCd.Value))
	strVal = strVal & "&txtOrgChangeId=" & Trim(frm1.hOrgChangeId.Value)
    strVal = strVal & "&txtGlinputType=" & Trim(frm1.txtGlinputType.Value)

	Call RunMyBizASP(MyBizASP, strVal)
    DbDelete = True
End Function

'=======================================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function


'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX(Row)
	With frm1
	End With
End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet(Row)
	With frm1
		ggoSpread.Source = frm1.vspdData
		.vspdData.Row	= Row
		.vspdData.Col	= C_DocCur
       Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,Row,Row,C_DocCur ,C_ItemAmt ,"A" ,"I","X","X")
       Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,Row,Row,C_DocCur ,C_ExchRate,"D" ,"I","X","X")	       	
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

		Call AutoInputDetail(strAcctCd, Trim(frm1.vspdData.text), frm1.txtGLDt.text, Row)
		For ii = 1 To frm1.vspdData2.MaxRows
			frm1.vspddata2.col = C_CtrlVal
			frm1.vspddata2.row = ii

			If Trim(frm1.vspddata2.text) <> "" Then
				Call CopyToHSheet2(frm1.vspdData.ActiveRow,ii)
			'	frm1.vspddata2.col = C_HItemSeq
			'	Call CopyToHSheet2(frm1.vspdData2.ActiveRow,ii)
			End if
		next

End Sub
