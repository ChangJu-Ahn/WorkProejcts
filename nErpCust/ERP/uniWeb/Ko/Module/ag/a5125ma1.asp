
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5125ma1
'*  4. Program Name         : 결의전표수정 
'*  5. Program Desc         : 결의전표내역을  수정,  조회 
'*  6. Component List       : PAGG005.dll
'*  7. Modified date(First) : 2003/01/10
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : Kim Ho Young
'* 10. Modifier (Last)      : Lim YOung Woon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="Acctctrl_ko441_1.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->	
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID      = "a5125mb1.asp"			'☆: 비지니스 로직 ASP명 
'Const BIZ_PGM_ID2      = "a5101mb2.asp"			'☆: 비지니스 로직 ASP명 
Const JUMP_PGM_ID_TAX_REP = "a6114ma1"

'=                       4.2 Constant variables 
'========================================================================================================
Const C_GLINPUTTYPE = "TG"
Const MENU_NEW	=	"1100000000011111"					 
Const MENU_UPD	=	"1110100100011111"					 
Const MENU_PRT	=	"1110000000011111"

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

'=========================================================================================================
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

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'=========================================================================================================
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
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
    Call ggoOper.ClearField(Document, "1") 
    frm1.hCongFg.value = ""
    frm1.txttempGLDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
    frm1.hCongFg.value = "" 
    frm1.cboConfFg.value = "U"    
    frm1.txtCommandMode.value = "CREATE"
    frm1.cboGlInputType.value = C_GLINPUTTYPE
	frm1.txtDeptCd.value	= parent.gDepart
	frm1.vspdData3.MaxRows = 0
    frm1.vspdData3.MaxCols = 16    
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	frm1.txtTempGlNo.focus	    	
    lgBlnFlgChgValue = False   
End Sub

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
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
		Call AppendNumberPlace("6","3","0")
        ggoSpread.SSSetFloat  C_ItemSeq,    " ", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_deptcd,     "부서코드",   10, , , 10, 2
        ggoSpread.SSSetButton C_deptpopup
        ggoSpread.SSSetEdit   C_deptnm,     "부서명",     17, , , 30
		ggoSpread.SSSetEdit   C_AcctCd,     "계정코드", 15, , , 18
		ggoSpread.SSSetButton C_AcctPopup
		ggoSpread.SSSetEdit   C_AcctNm,     "계정코드명", 20, , , 30
		ggoSpread.SSSetCombo  C_DrCrFg,     " ", 8
	    ggoSpread.SSSetCombo  C_DrCrNm,     "차대구분", 10
		ggoSpread.SSSetEdit   C_DocCur,     "거래통화",   10, , , 10, 2
        ggoSpread.SSSetButton C_DocCurPopup
		ggoSpread.SSSetFloat  C_ExchRate,   "환율", 15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_ItemAmt,    "금액",       15, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt, "금액(자국)", 15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_IsLAmtChange,   "",     30, , , 128
		ggoSpread.SSSetEdit   C_ItemDesc,   "비  고", 30, , , 128
		ggoSpread.SSSetCombo  C_VATTYPE,     "", 8
	    ggoSpread.SSSetCombo  C_VATNM,     "계산서유형",   20	    				
		ggoSpread.SSSetEdit   C_AcctCd2,   "",     30, , , 128

		Call ggoSpread.MakePairsColumn(C_deptcd,C_deptpopup)
		Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPopup)
		Call ggoSpread.MakePairsColumn(C_DrCrFg,C_DrCrNm,"1")
		Call ggoSpread.MakePairsColumn(C_VATTYPE,C_VATNM,"1")

		Call ggoSpread.SSSetColHidden(C_ItemSeq,C_ItemSeq,True)
		Call ggoSpread.SSSetColHidden(C_DrCrFg,C_DrCrFg,True)
		Call ggoSpread.SSSetColHidden(C_VatType,C_VatType,True)
		Call ggoSpread.SSSetColHidden(C_VatNm,C_VatNm,True)
		Call ggoSpread.SSSetColHidden(C_IsLAmtChange,C_IsLAmtChange,True)
		Call ggoSpread.SSSetColHidden(C_AcctCd2,C_AcctCd2,True)

		.ReDraw = True
                
    end with

    SetSpreadLock "I", 0, 1, ""        
End Sub
'=======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    Dim objSpread
    
    With frm1
		ggoSpread.Source = .vspdData
		Set objSpread = .vspdData
		lRow2 = objSpread.MaxRows
		objSpread.Redraw = False
		Select Case Index
			Case 0			
				ggoSpread.SpreadUnLock		C_deptcd		, -1    , C_deptcd
'				ggoSpread.SSSetRequired		C_deptcd		, -1    , C_deptcd
				ggoSpread.SpreadUnLock		C_deptpopup		, -1    , C_deptpopup
				ggoSpread.SpreadLock		C_deptnm		, -1    , C_deptnm
				ggoSpread.SpreadLock		C_AcctCd		, -1    , C_AcctCd
				ggoSpread.SpreadLock		C_AcctPopup		, -1    , C_AcctPopup
				ggoSpread.SpreadLock		C_AcctNm		, -1    , C_AcctNm
				ggoSpread.SpreadUnLock		C_DrCrNm		, -1    , C_DrCrNm
				ggoSpread.SSSetRequired		C_DrCrNm		, -1    , C_DrCrNm
				ggoSpread.SpreadLock		C_DocCur		, -1    , C_DocCur
				ggoSpread.SSSetRequired		C_DocCur		, -1    , C_DocCur
				ggoSpread.SpreadLock		C_DocCurPopup	, -1    , C_DocCurPopup
				ggoSpread.SpreadLock		C_ExchRate		, -1    , C_ExchRate
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
				ggoSpread.SpreadLock C_ItemDesc		, -1	, C_ItemDesc
				ggoSpread.SpreadLock C_AcctPopup	, -1	, C_AcctPopup
				ggoSpread.SpreadLock C_DrCrNm		, -1	, C_DrCrNm
				ggoSpread.SpreadLock C_DocCur		, -1	, C_DocCur    ', lRow2
				ggoSpread.SpreadLock C_DocCurPopup	, -1	, C_DocCurPopup    ', lRow2
				ggoSpread.SpreadLock C_ExchRate		, -1	, C_ExchRate    ', lRow2				
				ggoSpread.SpreadLock C_ItemAmt		, -1	, C_ItemAmt
				ggoSpread.SpreadLock C_VATNM		, -1	, C_VATNM
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		End Select

		objSpread.Redraw = True
		Set objSpread = Nothing
    
    End With
    
End Sub


'=======================================================================================================

Sub SetSpread2Lock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
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
Sub SetSpreadColor(Byval stsFg, Byval Index, ByVal lRow, ByVal lRow2)
    With frm1
		if  lRow2 = "" THEN	lRow2 = lRow
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_ItemSeq	, lRow, lRow2   ' 
		ggoSpread.SSSetProtected C_AcctNm	, lRow, lRow2   ' 계정코드명	
		ggoSpread.SSSetProtected C_AcctCd, lRow, lRow2	' 계정코드				

		Select Case stsFg
			Case "I"						
'				ggoSpread.SSSetRequired  C_deptcd,    lRow, lRow2	   ' 부서코드 
'				ggoSpread.SSSetRequired C_AcctCd, lRow, lRow2	' 계정코드 
			CASE "Q"					
				ggoSpread.SSSetProtected  C_deptcd,    lRow, lRow2	   ' 부서코드 
				ggoSpread.SSSetProtected  C_ItemDesc,    lRow, lRow2	   ' 부서코드 
		End Select	
		ggoSpread.SSSetProtected C_ExchRate, lRow, lRow2	' 차대구분 

		ggoSpread.SSSetProtected C_DrCrNm, lRow, lRow2	' 차대구분 
		ggoSpread.SSSetProtected  C_DocCur, lRow, lRow2	   ' 통화 

		ggoSpread.SSSetProtected C_DrCrNm, lRow, lRow2	' 차대구분				
		ggoSpread.SSSetProtected C_ItemAmt, lRow, lRow2	' 금액 
		ggoSpread.SSSetProtected C_ItemLocAmt, lRow, lRow2	' 금액 

		.vspdData.ReDraw = True
		
    End With

End Sub

'============================================================================================================
Sub InitComboBox()
	
	Err.clear
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1013", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboGlType ,lgF0  ,lgF1  ,Chr(11))
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlInputType ,lgF0  ,lgF1  ,Chr(11))
		
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1007", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))

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
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet				'권한관리 추가   							  
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	'//frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
	
		Case 1
			If frm1.txtDeptCd.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If
			
			arrStrRet =  AutorityMakeSql("DEPT",frm1.hORGCHANGEID.value, "","","","")	'권한관리 추가   							  
	
			arrParam(0) = "부서 팝업"
			arrParam(1) = arrstrRet(0)										'권한관리 추가   							  
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = arrstrRet(1)										'권한관리 추가   							  
						
			arrParam(5) = "부서코드"

			arrField(0) = "DEPT_CD"	
			arrField(1) = "DEPT_NM"	
    
			arrHeader(0) = "부서코드"
			arrHeader(1) = "부서명"	
			
			
		Case 2

			arrParam(0) = "통화코드 팝업"								' 팝업 명칭			
			arrParam(1) = "B_Currency"	    								' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "통화코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "Currency"	    								' Field명(0)
			arrField(1) = "Currency_desc"	    							' Field명(1)
    
			arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화코드명"									' Header명(1)
			

		Case 3
			arrParam(0) = "계정코드팝업"	
			arrParam(1) = "A_Acct, A_ACCT_GP"
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "												' Where Condition
			arrParam(5) = "계정코드"

			arrField(0) = "A_ACCT.Acct_CD"		
			arrField(1) = "A_ACCT.Acct_NM"
    		arrField(2) = "A_ACCT_GP.GP_CD"	
			arrField(3) = "A_ACCT_GP.GP_NM"	
			
			arrHeader(0) = "계정코드"	
			arrHeader(1) = "계정코드명"	
			arrHeader(2) = "그룹코드"
			arrHeader(3) = "그룹명"	

		Case 4
		
			arrStrRet =  AutorityMakeSql("DEPT_ITEM",frm1.hORGCHANGEID.value, frm1.txtDeptCd.value,"","","")'권한관리 추가 
			
			arrParam(0) = "부서 팝업"	
			arrParam(1) = arrstrRet(0)								'권한관리 추가 
			arrParam(2) = strCode	
			arrParam(3) = ""							
  			arrParam(4) = arrstrRet(1)								'권한관리 추가   							  

			arrParam(5) = "부서코드"

			arrField(0) = "A.DEPT_CD"
			arrField(1) = "A.DEPT_NM"
    
			arrHeader(0) = "부서코드"	
			arrHeader(1) = "부서명"
		
	End Select
    
    If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If	
	Call FocusAfterPopup (iWhere)

End Function

'========================================================================================================= 
'	Description : CtrlItem Popup에서 Return되는 값 setting
'========================================================================================================= 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
	
		Select Case iWhere	
			Case 1
				.txtDeptCd.value = UCase(Trim(arrRet(0)))
				.txtDeptNm.value = arrRet(1)
				Call txtDeptCd_OnChange()
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
					call FindExchRate(UniConvDateToYYYYMMDD(frm1.txttempGLDt.text,parent.gDateFormat,""), UCase(Trim(arrRet(0))),frm1.vspdData.ActiveRow)
				End IF
				
				call DocCur_OnChange(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)

			Case 3
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 
			
				.vspdData.Col  = C_AcctCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_AcctNm
				.vspdData.Text = arrRet(1)
				
                call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)
			Case 4
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 
			
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
				
				.vspdData.Col  = C_deptcd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_deptnm
				.vspdData.Text = arrRet(1)
		End Select

	End With
End Function	
'=======================================================================================================
Function FocusAfterPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1  
				.txtDeptCd.focus
			Case 2 
				Call SetActiveCell(.vspdData,C_DocCur,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				Call SetActiveCell(.vspdData,C_AcctCD,.vspdData.ActiveRow ,"M","X","X")
			Case 4
				Call SetActiveCell(.vspdData,C_deptcd,.vspdData.ActiveRow ,"M","X","X")
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
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5101ra1", "X")
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
	
	If arrRet(0) <> ""  Then	
		Call SetRefTempGl(arrRet)
	End If
	frm1.txttempGlNo.focus
	
End Function
'========================================================================================================= 
Function SetRefTempGl(ByRef arrRet)	
	With frm1
		.txttempGlNo.value = UCase(Trim(arrRet(0)))
    End With    
   
End Function

'========================================================================================================= 

Function OpenDept(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtTempGLDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

' T : protected F: 필수 
	If lgIntFlgMode = parent.OPMD_UMODE then
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
Function OpenUnderDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg   	

	IsOpenPop = True
	If RTrim(LTrim(frm1.txtDeptCd.value)) <> "" 	Then
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"				
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND B.BIZ_AREA_CD = ( SELECT B.BIZ_AREA_CD"
		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ")"

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
	Else
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A"				
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			

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
	End IF

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetDept(arrRet, iWhere)
	End If

	Call FocusAfterDeptPopup (  iWhere)
	

End Function
'========================================================================================================= 
Function SetDept(ByRef arrRet, Byval iWhere)
		
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
         '      .vspdData.Col = C_Dept_cd                         'spread
         '      .vspdData.Text = arrRet(1)
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
                    if .text <> ggoSpread.DeleteFlag then

		            .col = C_DrCrFg
			    
		            if .text = "DR" then		
		
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
	
		            elseif .text = "CR" then
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
		
			    end if	
		     end if	            
	        Next 
       End If                

        frm1.txtDrLocAmt.text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
        frm1.txtCrLocAmt.text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")

		If frm1.cboGlType.value = "01" Then
			frm1.txtDrLocAmt.text = frm1.txtCrLocAmt.text
		ElseIF frm1.cboGlType.value = "02" Then
			frm1.txtCrLocAmt.text = frm1.txtDrLocAmt.text
		End If
	End With

End Function

'========================================================================================================
'	Desc : Cookie Setting
'========================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp
	Dim strNmwhere
	Dim arrVal
	Dim IntRetCD
	
	Select Case Kubun		
	Case "FORM_LOAD"
	
		strTemp = ReadCookie("TEMP_GL_NO")
		Call WriteCookie("TEMP_GL_NO", "")

		If strTemp = "" then Exit Function
					
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
		End IF
		
		frm1.vspddata.row = frm1.vspddata.ActiveRow	
		frm1.vspddata.Col = C_VatType
		
		If frm1.vspddata.Value	=	"" Then		
			Exit Function
		End IF

		frm1.vspddata.Col = C_ItemSeq
		
		strNmwhere = " TEMP_GL_NO  = " & FilterVar(frm1.txtTempGlNo.value , "''", "S")
		strNmwhere = strNmwhere & " AND TEMP_ITEM_SEQ = " & frm1.vspddata.text & " "		
					
		IF CommonQueryRs( "VAT_NO" , "A_VAT" ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			
			arrVal = Split(lgF0, Chr(11))  
			strTemp = arrVal(0)
		End IF			
		
		Call WriteCookie("VAT_NO", strTemp)	
		
	Case "TEMP_GL_POPUP"
		Call WriteCookie("PGMID", "A5101MA1")
	
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
    If (lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True ) And C_GLINPUTTYPE = frm1.cboGlInputType.value Then
		IntRetCD = DisplayMsgBox("990027", "X", "X", "X")                          'No data changed!!
        Exit Function
    End If

	Select Case strPgmId			
	Case JUMP_PGM_ID_TAX_REP
		
		ggoSpread.Source = frm1.vspdData
		
		If frm1.vspddata.MaxRows < 1 Then
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
			ggoSpread.SSSetProtected C_DocCur, 1, frm1.vspddata.maxrows	' 통화 
			ggoSpread.SSSetProtected C_DocCurPopup, 1, frm1.vspddata.maxrows	' 통화팝업	
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows	' 차대구분 
		Case "02"			
			ggoSpread.SSSetProtected C_DocCur, 1, frm1.vspddata.maxrows	' 통화 
			ggoSpread.SSSetProtected C_DocCurPopup, 1, frm1.vspddata.maxrows	' 통화팝업 
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows	' 차대구분 
		Case "03"			
			ggoSpread.SSSetRequired C_DocCur, 1, frm1.vspddata.maxrows	' 통화 
			ggoSpread.SpreadUnLock C_DocCurPopup, 1, frm1.vspddata.maxrows	' 통화팝업 
			ggoSpread.SpreadUnLock C_DrCrfg, 1, C_DrCrNm, frm1.vspddata.maxrows
			ggoSpread.SSSetRequired C_DrCrfg, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetRequired C_DrCrNm, 1, frm1.vspddata.maxrows	' 차대구분 
	END Select 				
end Sub
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

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call InitSpreadSheet  
    Call InitCtrlSpread()
    Call InitCtrlHSpread()
    Call InitComboBox
    Call InitComboBoxGrid    
    Call SetAuthorityFlag                                               '권한관리 추가    
    Call SetToolbar(MENU_NEW)										'⊙: 버튼 툴바 제어    
    Call SetDefaultVal    
    Call InitVariables                         '⊙: Initializes local global variables
	Call CookiePage("FORM_LOAD")  	
	
	'Call GetAcctForVat		' -- 2006-07-18 함수 미정의로 리마킹함:choe0tae

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
	
End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'=======================================================================================================
Sub vspdData_onfocus()

	lgCurrRow = frm1.vspdData.ActiveRow

	If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value = C_GLINPUTTYPE Then
	   Call SetToolbar(MENU_PRT) 						   
	Else
	   Call SetToolbar(MENU_UPD)							
	End if

End Sub

'=======================================================================================================
Sub txttempGLDt_DblClick(Button)
    If Button = 1 Then
        frm1.txttempGLDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txttempGLDt.focus
    End If
End Sub


'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim i
    Dim tmpDrCrFG
	
	If lgIntFlgMode = parent.OPMD_UMODE Then
	    Call SetPopUpMenuItemInf("0001111111")
	Else
	    Call SetPopUpMenuItemInf("0000111111")	
	End If
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
        'frm1.vspdData2.MaxRows = 0
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

            .hItemSeq.value = .vspdData.Text
            .vspdData2.MaxRows = 0
        End With

        frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End if
        		
		lgCurrRow = NewRow     		
			'Call CopyFromData(frm1.hItemSeq.value)       
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

			Call OpenPopUp(.text, 3)
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
    
    Select Case Col
	    Case   C_DeptCd
			frm1.vspdData.Col = C_DeptCd
			Call DeptCd_underChange(frm1.vspdData.text)
			
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
	If window.event.keycode = 39 then	'Single quotation mark 입력불가 
		window.event.keycode = 0	
	End If
End Sub

'==========================================================================================
Sub txtTempGlNo_OnKeyUp()	
	If Instr(1,frm1.txtTempGlNo.value,"'") > 0 then
		frm1.txtTempGlNo.value = Replace(frm1.txtTempGlNo.value, "'", "")		
	End if
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
	For ii = FromRow	to	ToRow
		frm1.vspdData.Row	= ii
		frm1.vspdData.Col	= C_DocCur

		IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.vspdData.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			Call CurFormatNumSprSheet(ii)
			Call SetSumItem
		END IF	  
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
		Exit sub
    End If
    lgBlnFlgChgValue = True

		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
		' 권한관리 추가 
'		If lgInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
'		End If
'	
'		If lgSubInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
'		End If

		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
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
	Dim jj

	If Trim(frm1.txtTempGLDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"	

		' 권한관리 추가 
'		If lgInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
'		End If
'	
'		If lgSubInternalCd <> "" Then
'			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
'		End If


		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
'			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)

			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next

		End If

End Sub


'==========================================================================================
Sub DeptCd_underChange(Byval strCode)
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 

    If Trim(frm1.txtTempGLDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

		' 권한관리 추가 
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
Sub txttempGLDt_Change()

   If lgstartfnc = False Then
    If lgFormLoad = True Then
		Dim strSelect
		Dim strFrom
		Dim strWhere 	
		Dim IntRetCD 
		Dim ii
		Dim arrVal1
		Dim arrVal2
		Dim jj


		lgBlnFlgChgValue = True
		With frm1
		
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtTempGLDt.Text <> "") Then
			'----------------------------------------------------------------------------------------
			strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
			strFrom		=			 " b_acct_dept(NOLOCK) "		
			strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
			strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
			strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
			strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtTempGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"
	
				If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
					IntRetCD = DisplayMsgBox("124600","X","X","X")
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
					jj = Ubound(arrVal1,1)
							
					For ii = 0 to jj - 1
						arrVal2 = Split(arrVal1(ii), chr(11))			
						frm1.hOrgChangeId.value = Trim(arrVal2(2))
					Next	
				End If 
			End If
		End With
	End If
  End IF
End Sub

'==========================================================================================
Sub cboGLType_OnChange()
	
	dim	i		
	Dim IntRetCD	
	
	ggoSpread.Source = frm1.vspdData
	
	SELECT CASE frm1.cboGlType.value 
		CASE "01"			
			'입금전표로 바꾸면 차변이 입력되거나 현금계정이 입력되었는지 check한다.
			FOR i = 1 TO  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				IF  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")					
					Exit sub
				End IF
																			
				frm1.vspddata.col = C_DrCrFg
				IF  Trim(frm1.vspddata.value) = "2" Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113104", "X", "X", "X")					
					Exit sub
				End IF											
			Next				
			
			FOR i = 1 TO  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				IF Trim(frm1.vspddata.value) <> "1"  Then					
					frm1.vspdData.value	= "1"							
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.value	= "1"							
				END IF
				frm1.vspddata.col = C_DocCur
				frm1.vspddata.text = parent.gCurrency				
			Next
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
		CASE "02"
			'출금전표로 바꾸면 대변이 입력되거나 현금계정이 입력되었는지 check한다.	
			FOR i = 1 TO  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				IF  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")					
					Exit sub
				End IF								
				
				frm1.vspddata.col = C_DrCrFg
				IF  Trim(frm1.vspddata.value) = "1" Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113105", "X", "X", "X")					
					Exit sub				
				End IF											
			Next
				
			FOR i = 1 TO  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				IF Trim(frm1.vspddata.value) <> "2"  Then					
					frm1.vspdData.value	= "2"							
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.value	= "2"							
				END IF
				frm1.vspddata.col = C_DocCur
				frm1.vspddata.text = parent.gCurrency				
			Next
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
		CASE "03"
		'대체로 바꾸면 Protect를 풀어준다.		
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
		
	END SELECT	
	
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
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If
 	
    If Not chkField(Document, "1") Then	'⊙: This function check indispensable field
       Exit Function
    End If
    
    Call ggoOper.ClearField(Document, "2")      '⊙: Condition field clear
    Call InitVariables							'⊙: Initializes local global variables

    'Query function call area
    IF  DbQuery = False Then														'☜: Query db data
		Exit Function
	END IF
		
    if frm1.vspddata.maxrows = 0 then	
       frm1.txtTempGlNo.value = ""
    end if
   
    FncQuery = True																'⊙: Processing is OK
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
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	lgBlnExecDelete = False

    Call ggoOper.ClearField(Document, "1")                                  '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '⊙: Clear Contents  Field
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
     
    Call InitComboBoxGrid

    frm1.txtTempGlNo.focus        
    Call SetToolbar(MENU_NEW)
   			 		
	Call ggoOper.SetReqAttr(frm1.txtDeptCd,   "N")
	Call ggoOper.SetReqAttr(frm1.txtTempGlDt, "N")
	Call ggoOper.SetReqAttr(frm1.txtdesc,   "D")
    
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
	
    lgBlnFlgChgValue = False

    FncNew = True                              '⊙: Processing is OK
    lgFormLoad = True							' tempgldt read
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

    ' Update 상태인지를 확인한다.
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = False Then									'변경된 부분이 없을경우 
		intRetCd = DisplayMsgBox("990008", parent.VB_YES_NO, "X", "X")				'삭제하시겠습니까?
		If intRetCd = VBNO Then
			Exit Function
		End IF
    Else
		IntRetCD = DisplayMsgBox("900038", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then    		
      		Exit Function
    	End If
    End If

    '-----------------------
    'Delete function call area
    '-----------------------
    IF  DbDelete = False Then														'☜: Delete db data
    	Exit Function
    End If
    FncDelete = True 
    
End Function


'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                   '☜: Protect system from crashing
    
	'-----------------------
    'Precheck area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False  AND ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          'No data changed!!
        Exit Function
    End If    

    If CheckSpread3 = False then
		IntRetCD = DisplayMsgBox("110420", "X", "X", "X")							'필수입력 check!!
        Exit Function
    End If
    
	If frm1.vspdData.MaxRows < 1 Then												'회계전표존재하지 않음 
		IntRetCD = DisplayMsgBox("114100", "X", "X", "X")
		Exit Function
	End If
  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") Then                                  '⊙: Check contents area
		Exit Function
    End If
    
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Exit Function
    End If
    
	'-----------------------
    'Save function call area
    '----------------------- 	
    IF  DbSave	= False Then			                                                '☜: Save db data
		Exit Function
    End If
    
    FncSave = True                                                          
    
End Function
'========================================================================================
Function FncCopy() 

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    Dim iItemSeq
    Dim RowDocCur

	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	if  frm1.vspdData.MaxRows = 1 Then  Call ggoOper.SetReqAttr(frm1.cboGlType,   "N")

    With frm1.vspdData
        .Row = .ActiveRow
        .Col = 0
        
        If .Text = ggoSpread.InsertFlag Then
			.Col = C_AcctCd
			IF len(Trim(.text)) > 0 Then 
				.Col = C_ItemSeq
				DeleteHSheet(.Text)
			end if	
        End if
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
        
		if .row = 0 then 			
			Exit Function
		end if

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
            frm1.vspdData2.MaxRows = 0
		    Call DbQuery2(.ActiveRow)
        End if
    End With
    Call SetSumItem()
End Function

'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)

End Function

'========================================================================================
Function FncDeleteRow() 

End Function

'========================================================================================
Function FncPrint() 
    On Error Resume Next     
    parent.FncPrint()
End Function

'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
Function FncExcel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function

'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
	        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, 1, -1 ,C_DocCur,C_ItemAmt ,"A" ,"I","X","X")
	        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, 1, -1 ,C_DocCur,C_ExchRate,"D" ,"I","X","X")	        
			
			If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE then
			    Call SetSpreadLock("Q", 1, 1, "")			    
			    Call SetSpread2Lock("",1,1,"")			    
			Else
                Call SetSpreadColor("Q", 0,1, frm1.vspdData.MaxRows)                 
                Call SetSpread2Color()           
			End if
            
		Case "VSPDDATA2"			
			Call PrevspdData2Restore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value <> C_GLINPUTTYPE then
			    Call SetSpread2Lock("",1,1,"")
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
									Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtTempGlNo.Value)
								End If
						End Select
					Next
					'ggoSpread.Source = frm1.vspdData					
					'ggoSpread.EditUndo
					
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtTempGLNo.Value)
					'ggoSpread.Source = frm1.vspdData
					'ggoSpread.EditUndo
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
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.htxtTempGLNo.Value)
						End If
					Next
				Case ggoSpread.DeleteFlag
			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

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
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , LTrim(ISNULL(C.CTRL_VAL,'')), '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & strItemSeq & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_TEMP_GL_DTL C (NOLOCK), A_TEMP_GL_ITEM D (NOLOCK) "
		
		strWhere =			  " D.TEMP_GL_NO = " & FilterVar(UCase(pInvalue1), "''", "S")
		strWhere = strWhere & " AND D.ITEM_SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.TEMP_GL_NO  =  C.TEMP_GL_NO  "
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
	Dim var1,var2

	FncExit = False

	ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True OR ggoSpread.SSCheckChange = True Then  
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"	
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
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
	Dim ObjName

    If Not chkField(Document, "1") Then
		Exit Function
    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
End Function

'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId
    Dim StrEbrFile
    Dim intRetCd
	Dim ObjName

	If Not chkField(Document, "1") Then
       Exit Function
    End If

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

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
End Function

'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)
	Dim intRetCd

	StrEbrFile = "a5118ma1"
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txttempGlDt.Text, parent.gDateFormat, parent.gServerDateType)	
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txttempGlDt.Text, parent.gDateFormat, parent.gServerDateType)

	' 회계전표의 key는 GL_NO이기 때문에 GL_NO만 넘긴다.	
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	VarTempGlNoFr = Trim(frm1.txttempGlNo.value)
	VarTempGlNoTo = Trim(frm1.txttempGlNo.value)
	varOrgChangeId = Trim(frm1.hOrgChangeId.value)
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
		strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
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
					.vspdData.text	=	UNIConvNumPCToCompanyByCurrency(tempAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				End If
			Next		
		End If
	End With
	Call SetSumItem	
End Function

'========================================================================================
' Function Name : gfRealRound
' Function Desc : Arithmetic Rounding Function
'========================================================================================
Function gfRealRound(ByVal x, ByVal Factor )
    Dim lcSwitch, iCurResult
    If x < 0 Then lcSwitch = -1 Else lcSwitch = 1
    x = x * lcSwitch
    iCurResult = Int(x * 10 ^ Factor + 0.5) / 10 ^ Factor
    gfRealRound = iCurResult * lcSwitch
End Function
'========================================================================================
' Function Name : ExchRateCheck
' Function Desc : 
'========================================================================================
Function ExchRateCheck()
	Call FncBtnCalc()
End Function 

'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim RetFlag

    DbQuery = False
    Call LayerShowHide(1)
    frm1.vspdData3.MaxRows = 0 

    Err.Clear                '☜: Protect system from crashing
    
    With frm1    				
		
	    If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtTempGlNo=" & UCase(Trim(.htxtTempGlNo.value))	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 
			strVal = strVal & "&txtTempGlNo=" & UCase(Trim(.txtTempGlNo.value))	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
		End If

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
   
		Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 
		    
    End With
    
    DbQuery = True

End Function

'=======================================================================================================
Function DbQueryOk()
	Dim intI

	With frm1
	   .vspdData.Col = 1:    intItemCnt = .vspddata.MaxRows

		If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value = C_GLINPUTTYPE then
			SetSpreadLock "Q", 0, 1, ""
			SetSpreadColor "Q", 0,1, intItemCnt
		Else
			SetSpreadLock "I", 0, 1, ""
			SetSpreadColor "I", 0,1, intItemCnt
		End if

        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
		'Call ggoOper.LockField(Document, "Q")                                   '⊙: Lock  Suitable  Field

		If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value = C_GLINPUTTYPE Then
			Call SetToolbar(MENU_PRT) 
		Else
			Call SetToolbar(MENU_UPD)									'버튼 툴바 제어 
		End if

		.txtCommandMode.value = "UPDATE"
		'.txtTempGlNo.disabled  = True
    
		InitData
		'Call SetSumItem

		If .vspdData.MaxRows > 0 Then
			.vspdData.Row = 1
			.vspdData.Col = 1
			.hItemSeq.Value = .vspdData.Text				
			
'			Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")
			'Call ggoOper.SetReqAttr(frm1.cboGlType,   "Q")
			Call ggoOper.SetReqAttr(frm1.txtTempGlDt, "Q")
			
			Call DbQuery2(1)
		End If
    End With
    
   	SetGridFocus()
    SetGridFocus2()
'	Call txtDocCur_OnChange()
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
	
	'Err.Clear
	
	With frm1
		'.htxtTempGlNo.value = frm1.txtTempGlNo.value
	    .vspdData.row = Row
	    .vspdData.col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	    
	    frm1.vspdData2.ReDraw = false	
	    
        If CopyFromData(.hItemSeq.Value) = True Then
			If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value = C_GLINPUTTYPE then
				Call SetSpread2Lock("",1,1,"")
			Else
				Call SetSpread2Color()
			End  If 	
			frm1.vspdData2.ReDraw = True
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , LTrim(ISNULL(C.CTRL_VAL,'')), '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & .hItemSeq.Value & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_TEMP_GL_DTL C (NOLOCK), A_TEMP_GL_ITEM D (NOLOCK) "
		
		strWhere =			  " D.TEMP_GL_NO = " & FilterVar(UCase(.htxtTempGlNo.value), "''", "S")
		strWhere = strWhere & " AND D.ITEM_SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.TEMP_GL_NO  =  C.TEMP_GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "	
		
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
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") 
					End If
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
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
'				.vspdData2.Text = lngRows
				strVal = strVal & Chr(11) & lngRows

				strVal = strVal & Chr(11) & Chr(12)
			Next					

			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		intItemCnt = .vspddata.MaxRows
        
		'Call CopyFromData (.hItemSeq.value)
		If frm1.hCongFg.value = "C" Or frm1.cboGlInputType.Value = C_GLINPUTTYPE then
			Call SetSpread2Lock("",1,1,"")
		Else
			Call SetSpread2Color()
		End  If
	End With
	
	frm1.vspdData2.ReDraw = True
	
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

    'Call SetSumItem

	With frm1
		.txtFlgMode.value     = lgIntFlgMode
		.txtUpdtUserId.value  = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtMode.value        = parent.UID_M0002
		.txtAuthorityFlag.value     = lgAuthorityFlag               '권한관리 추가 
				
		'//.hOrgChangeId.value = parent.gChangeOrgId
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
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
				
				'Local 금액을 사용자 입력시 입력금액을 전달 
				'If .Text = "Y" Then
   					.Col = C_ItemLocAmt	'6
					strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep
				'Else
				'	strVal = strVal & "0" & parent.gColSep
				'End If

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

	IF frm1.txtSpread.value = "" Then	
		intRetCd = DisplayMsgBox("990008", parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
		If intRetCd = VBNO Then
			Exit Function
		End IF	
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

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)	

    DbSave = True                                                           
End Function

'========================================================================================
Function DbSaveOk(ByVal TempGlNo)					'☆: 저장 성공후 실행 로직 
	lgBlnFlgChgValue = false
	
	frm1.txtTempGlNo.value = UCase(Trim(TempGlNo))
    frm1.txtCommandMode.value = "UPDATE"
    
	Call ggoOper.ClearField(Document, "2")      '⊙: Condition field clear    
    Call InitVariables							'⊙: Initializes local global variables

	DbQuery
End Function

'========================================================================================
Function DbDelete()
	Dim strVal
	
    Err.Clear
    Call LayerShowHide(1)    
	DbDelete = False														'⊙: Processing is NG

	frm1.hOrgChangeId.value = parent.gChangeOrgId

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003						'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtTempGlNo=" & UCase(Trim(frm1.txtTempGlNo.value))	'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtDeptCd=" & UCase(Trim(frm1.txtDeptCd.value))
	strVal = strVal & "&txtOrgChangeId=" & Trim(frm1.hOrgChangeId.value)
	strVal = strVal & "&txtTempGlDt=" & Trim(frm1.txttempgldt.text)

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True                                                         '⊙: Processing is NG	
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

		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_DocCur ,C_ItemAmt ,"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_DocCur ,C_ExchRate,"D" ,"I","X","X")
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

	Call AutoInputDetail(strAcctCd, Trim(frm1.vspdData.text), frm1.txttempGLDt.text, Row)

	For ii = 1 To frm1.vspdData2.MaxRows
		frm1.vspddata2.col = C_CtrlVal
		frm1.vspddata2.row = ii

		If Trim(frm1.vspddata2.text) <> "" Then
			Call CopyToHSheet2(frm1.vspdData.ActiveRow,ii)			 			
		End If
	Next
End Sub	

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
-->
</SCRIPT>
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
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
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"> </TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%"> 
					    <FIELDSET CLASS="CLSFLD">
						  <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의번호</TD>
								<TD CLASS=TD656 NOWRAP><INPUT NAME="txtTempGlNo" ALT="결의번호" MAXLENGTH="18" SIZE=20 STYLE="TEXT-ALIGN: left" tag  ="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenReftempgl()"></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP >
						<TABLE <%=LR_SPACE_TYPE_60%>>					
							<TR>								
								<TD CLASS=TD5 NOWRAP>결의일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txttempGLDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="결의일자" tag="24" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>전표형태</TD>								
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlType" tag="24" STYLE="WIDTH:82px:" ALT="전표형태"><OPTION VALUE="" selected></OPTION></SELECT></TD> 
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)" tag="22">&nbsp;
													 <INPUT NAME="txtDeptNm" ALT="부서명"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="24X"></TD>
													 <INPUT NAME="txtInternalCd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"  TABINDEX="-1">
<!---- eWare Inf Begin -->
								<TD CLASS=TD5 NOWRAP>승인상태</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboConfFg" tag="24" STYLE="WIDTH:82px:" ALT="승인상태"><OPTION VALUE="" selected></OPTION></SELECT></TD>

<!-- --eWare Inf End -->
						   </TR>
						   <TR>									
								<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="24" STYLE="WIDTH:82px:" ALT="전표입력경로"><OPTION VALUE="" selected></OPTION></SELECT></TD>								
								<TD CLASS=TD5 NOWRAP>
								<TD CLASS=TD6 NOWRAP>
							</TR>						
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="70" tag="22N" ></TD>
							</TR>	
							<TR>
								<TD HEIGHT="60%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
<!--								<TD CLASS=TD5 NOWRAP>차대합계(거래)</TD>
								<TD >
									&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계(거래)" id=OBJECT1></OBJECT>');</SCRIPT>
									&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="대변합계(거래)" id=OBJECT2></OBJECT>');</SCRIPT>
								</TD>
-->								
								<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTTON NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>자국금액계산</BUTTON>&nbsp;
								<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
								<TD CLASS=TD6 NOWRAP>
									&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계(자국)" id=OBJECT3></OBJECT>');</SCRIPT>
									&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="대변합계(자국)" tag="24X2" id=OBJECT4></OBJECT>');</SCRIPT>
								</TD>
							</TR>
			                <TR>
								<TD HEIGHT="40%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
	<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>&nbsp;
					</TD>

					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
<!--
	<TR HEIGHT="20">
  		<TD WIDTH="100%" CLASS="Tab11">
      		<TABLE WIDTH="100%">
    			<TR>
    				<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=50 tag="23" TITLE="SPREAD" id=vaSpread3><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>

    				<TD><BUTTON NAME="button1" CLASS="CLSMBTN" ONCLICK="vbscript:ShowHidden()">SHOW HIDDEN</BUTTON></TD>
    			</TR>
      		</TABLE> 
  		</TD>
    </TR>
-->
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
		<!--<TD WIDTH="100%" HEIGHT=30%><IFRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>-->
	</TR>
</TABLE>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TABINDEX="-1" CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=0 HEIGHT=0 tag="23" TITLE="SPREAD" id=vaSpread3><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
<TEXTAREA class=hidden name=txtSpread		tag="24" Tabindex="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3		tag="24" Tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtTempGlNo"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtCommandMode"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="24" Tabindex="-1">

<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtAuthorityFlag"  tag="24" Tabindex="-1"><!--권한관리추가 -->
<INPUT TYPE=HIDDEN NAME="hCongFg"			tag="24" Tabindex="-1">

<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>
