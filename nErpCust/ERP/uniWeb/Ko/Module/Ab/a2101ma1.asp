<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Chart of Account(E)
'*  3. Program ID           : A2101MA1
'*  4. Program Name         : 계정코드 등록 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2003/11/25
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :  2002/11/25 : ASP Standard for Include improvement
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incUni2KTV.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'==========================================================================================================
Const  C_CMD_TOP_LEVEL	= "LISTTOP"
Const  C_CMD_GP_LEVEL	= "LISTGP"
Const  C_CMD_ACCT_LEVEL = "LISTACCT"

Const  C_Root			= "Root"
Const  C_Folder_Ch		= "folder_ch.gif"
Const  C_URL_Ch			= "URL_Ch"

Const  C_IMG_Folder_Ch	= "../../../CShared/image/folder_ch.gif"
Const  C_IMG_URL_Ch		= "../../../CShared/image/Account_Ch.gif"


Const  C_USER_MENU			= "UNIERP"
Const  C_USER_MENU_KEY		= "$"
Const  C_USER_MENU_STR		= "UM_"
Const  C_UNDERBAR			= "_"

Const  BIZ_SAVE_ACCT_ID		= "a2101mb2.asp"												'☆: 비지니스 로직 ASP명 
Const  BIZ_LOOKUP_ACCT_ID	= "a2101mb3.asp"												'☆: 비지니스 로직 ASP명 
Const  BIZ_MOVE_ACCT_ID		= "a2101mb4.asp"												'☆: 비지니스 로직 ASP명 

Const  C_Sep  = "/"

Const  C_IMG_Root	= "../../../CShared/image/unierp.gif"
Const  C_IMG_Folder	= "../../../CShared/image/Group.gif"
Const  C_IMG_Open	= "../../../CShared/image/Group_op.gif"
Const  C_IMG_URL	= "../../../CShared/image/Account.gif"
Const  C_IMG_None	= "../../../CShared/image/c_none.gif"
Const  C_IMG_Const	= "../../../CShared/image/c_const.gif"

Const  C_MNU_SEP		= "::"
Const  C_MNU_ID		= 0
Const  C_MNU_UPPER	= 1
Const  C_MNU_LVL	= 2
Const  C_MNU_TYPE	= 3
Const  C_MNU_NM		= 4
Const  C_MNU_AUTH	= 5

Const  C_NEW_FOLDER	= "새 폴더"

Const  TAB1 = 1																				'☜: Tab의 위치 
Const  TAB2 = 2

Dim C_CTRLITEM
Dim C_CTRLITEMPB
Dim C_CTRLNM	
Dim C_CTRLITEMSEQ
Dim C_DRFG	
Dim C_CRFG
Dim C_DEFAULT_VALUE
Dim C_GL_ITEM
Dim C_GL_ITEMPB
Dim C_SYSTEM_FG				
Dim C_MAND_FG
Dim C_CHG_DEL

<!-- #Include file="../../inc/lgvariables.inc" -->
 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim  gDragNode , gDropNode, gNewNode , gPrevNode
Dim  lgBlnBizLoadMenu, lgBlnUserLoadMenu, gMenuDat, lgBlnNewNode
Dim  lgBlnLoadMenu

Dim  lgBlnFlgConChg				'☜: Condition 변경 Flag

Dim  lgStrPrevKey1
Dim  lgStrPrevKey2

Dim  lgQueryFlag
Dim  lgRetFlag
Dim  IsOpenPop						 'Popup

Dim  strMode

Dim  lgSaveModFg
Dim  gSelframeFlg
Dim  TempRootNode

Dim	 lglsClicked
Dim  lgUSER_MENU

'========================================================================================================
Sub InitSpreadPosVariables()
	C_CTRLITEM		= 1
	C_CTRLITEMPB	= 2
	C_CTRLNM		= 3
	C_CTRLITEMSEQ	= 4
	C_DRFG			= 5
	C_CRFG			= 6
    C_DEFAULT_VALUE = 7
    C_GL_ITEM       = 8
    C_GL_ITEMPB     = 9
    C_SYSTEM_FG	    = 10
    C_MAND_FG       = 11
    C_CHG_DEL       = 12
End Sub

'========================================================================================================= 
Sub  InitVariables()
	lgBlnBizLoadMenu = False
    lgBlnLoadMenu = False
    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgSortKey = 1
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count

	Call CommonQueryRs("Upper(Co_Cd), Co_NM", "B_Company","",lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	lgUSER_MENU = "[" & Left(lgF0, Len(lgF0)-1) & "]" & Left(lgF1, Len(lgF1)-1)
End Sub

'========================================================================================================= 
Sub  SetDefaultVal()
	frm1.cboBDG_CTRL_FG.Value		= "N"
	frm1.cboFX_EVAL_FG.Value		= "N"
	frm1.cboBAL_FG.value			= "CR" '차대구분 
	frm1.cboHQ_BRCH_FG.Value		= "N"
	frm1.cboTEMP_ACCT_FG.Value		= "N"
	frm1.cboGP_BDG_CTRL_FG.value	= "N"
	frm1.cboDEL_FG.value			= "N"
	frm1.cboSubSystemType.Value	    = ""
	frm1.txtOpenAcctFg.Value		= "N"
	frm1.cboMgntType.Value		    = ""
	lgBlnFlgChgValue = False
End Sub

'======================================================================================== 
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

'==================================================================================================
Sub  InitSpreadSheet()
    Dim sList
    Dim ii

	Call initSpreadPosVariables()
    sList = "Y" & vbTab  & "N"

    With frm1.vspdData
		.MaxCols	= C_CHG_DEL + 1
		.Col		= .MaxCols
		.ColHidden	= True

		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_CTRLITEM     ,"관리항목코드", 21,,,3,2
		ggoSpread.SSSetButton   C_CTRLITEMPB
		ggoSpread.SSSetEdit		C_CTRLNM       ,"관리항목명", 30

'		Call AppendNumberPlace("6","3","0")

		ggoSpread.SSSetFloat    C_CTRLITEMSEQ  ,"NO" ,3,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"         
		ggoSpread.SSSetCombo	C_DRFG         ,"차변입력필수", 13,True
		ggoSpread.SSSetCombo	C_CRFG         ,"대변입력필수", 13,True
		ggoSpread.SetCombo      sList          ,C_DRFG
		ggoSpread.SetCombo      sList          ,C_CRFG
		ggoSpread.SSSetEdit		C_DEFAULT_VALUE,"기본값"      , 20
		ggoSpread.SSSetEdit		C_GL_ITEM      ,"전표항목"    , 20		
		ggoSpread.SSSetButton   C_GL_ITEMPB		
		ggoSpread.SSSetEdit		C_SYSTEM_FG    ,"SYS"         , 1				'is SYS_FG?
		ggoSpread.SSSetEdit		C_MAND_FG      ,"MAND"        , 5				'is subsystem item and mandatory ?
		ggoSpread.SSSetEdit		C_CHG_DEL      ,"DEL"         , 5				'is subsystem change ?

		Call ggoSpread.MakePairsColumn(C_CTRLITEM,C_CTRLITEMPB,"1")
		Call ggoSpread.MakePairsColumn(C_GL_ITEM,C_GL_ITEMPB,"1")
		Call ggoSpread.SSSetColHidden(C_CTRLITEMSEQ,C_CTRLITEMSEQ,True)
		Call ggoSpread.SSSetColHidden(C_SYSTEM_FG,C_SYSTEM_FG,True)
		Call ggoSpread.SSSetColHidden(C_MAND_FG,C_MAND_FG,True)
		Call ggoSpread.SSSetColHidden(C_CHG_DEL,C_CHG_DEL,True)		
		
		Call ggoSpread.SSSetColHidden(C_DEFAULT_VALUE,C_DEFAULT_VALUE,True)				
		Call ggoSpread.SSSetColHidden(C_GL_ITEM,C_GL_ITEM,True)
		Call ggoSpread.SSSetColHidden(C_GL_ITEMPB,C_GL_ITEMPB,True)										
		
		.ReDraw = True
		Call SetSpreadLock("Q", 0, 1, "")
    End With
End Sub

'========================================================================================
Sub  SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2 )
    Dim objSpread

    With frm1
		Select Case Index
			Case 0
				ggoSpread.Source = .vspdData
				Set objSpread = .vspdData
		End Select

		If lRow2 = "" Then lRow2 = objSpread.MaxRows

		objSpread.Redraw = False

		Select Case stsFg
		    Case "Q"
		        Select Case Index
		            Case 0
		                ggoSpread.SpreadLock C_CTRLITEM		, -1, C_CTRLITEM
		                ggoSpread.SpreadLock C_CTRLITEMPB	, -1, C_CTRLITEMPB
		                ggoSpread.SpreadLock C_CTRLNM		, -1, C_CTRLNM
		                ggoSpread.SpreadLock C_CTRLITEMSEQ	, -1, C_CTRLITEMSEQ
		         End Select
		    Case "I"
		        Select Case Index
		            Case 0
		                ggoSpread.SpreadLock C_CTRLITEM		, -1, C_CTRLITEM
		                ggoSpread.SpreadLock C_CTRLITEMPB	, -1, C_CTRLITEMPB
		                ggoSpread.SpreadLock C_CTRLITEMSEQ	, -1, C_CTRLITEMSEQ
		                ggoSpread.SpreadLock C_CTRLNM		, -1, C_CTRLNM
		        End Select
		End Select

		ggoSpread.SSSetRequired C_DRFG, -1, C_DRFG
		ggoSpread.SSSetRequired C_CRFG, -1, C_CRFG

		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData 
		.Redraw = False

		' 아래 상수와 숫자가 섞인 이유는 해당 컬럼명을 몰라서 빠진부분임, 각 개발자는 모두 상수로 처리요망 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetRequired  C_CTRLITEM	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CTRLNM	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DRFG		, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_CRFG		, pvStartRow, pvEndRow

		.Col = 1
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True

		.Redraw = True
    End With
End Sub

'========================================================================================================= 
Sub InitCombo()
	Dim IntRetCD1
	Dim IntRetCD2
	Dim IntRetCD3
	Dim IntRetCD4

	On Error Resume Next	
	Err.Clear 

	'차대변 구분 
	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1012", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	Call SetCombo2(frm1.cboBAL_FG          ,lgF0 ,lgF1  ,Chr(11))	'YN

	IntRetCD2= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1020", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	Call SetCombo2(frm1.cboGP_BDG_CTRL_FG ,lgF0  ,lgF1  ,Chr(11))  '예산관리구분(계정코드용)
	Call SetCombo2(frm1.cboBDG_CTRL_FG    ,lgF0  ,lgF1  ,Chr(11))  '예산관리구분(계정그룹용)
	Call SetCombo2(frm1.cboFX_EVAL_FG     ,lgF0  ,lgF1  ,Chr(11))  '환평가구분 
	Call SetCombo2(frm1.cboTEMP_ACCT_FG   ,lgF0  ,lgF1  ,Chr(11))  '임시계정구분 
	Call SetCombo2(frm1.cboHQ_BRCH_FG     ,lgF0  ,lgF1  ,Chr(11))  '본지점구분 
	Call SetCombo2(frm1.cboDEL_FG         ,lgF0  ,lgF1  ,Chr(11))  '계정폐쇄 

	IntRetCD3= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1046", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	Call SetCombo2(frm1.cboSubSystemType ,lgF0  ,lgF1  ,Chr(11))  'sub_system Type
	
	IntRetCD4= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1038", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	Call SetCombo2(frm1.cboMgntType        ,lgF0  ,lgF1  ,Chr(11))	
End Sub 

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_CTRLITEM		= iCurColumnPos(1)
			C_CTRLITEMPB	= iCurColumnPos(2)
			C_CTRLNM		= iCurColumnPos(3)
			C_CTRLITEMSEQ	= iCurColumnPos(4)
			C_DRFG			= iCurColumnPos(5)
			C_CRFG			= iCurColumnPos(6)
			C_DEFAULT_VALUE = iCurColumnPos(7)
			C_GL_ITEM       = iCurColumnPos(8)
			C_GL_ITEMPB     = iCurColumnPos(9)
			C_SYSTEM_FG	    = iCurColumnPos(10)
			C_MAND_FG       = iCurColumnPos(11)
			C_CHG_DEL       = iCurColumnPos(12)
    End Select
End Sub

'========================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function

	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab
	gSelframeFlg = TAB1
End Function

'========================================================================================================= 
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
	gSelframeFlg = TAB2
End Function

'========================================================================================
Function OpenPopUp(Byval txtValue, Byval IntIndex)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IntIndex = 2 And frm1.cboHQ_BRCH_FG.value = "N" Then Exit Function

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case IntIndex
		Case 0,1
			If frm1.txtSUBLEDGER1.readOnly = True Or  frm1.txtSUBLEDGER2.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If
			
			arrParam(0) = "관리항목 팝업"

			If frm1.vspdData.MaxRows <= 0 Then
				arrParam(1) = "A_CTRL_ITEM"
				arrParam(2) = Trim(txtValue)
				arrParam(3) = ""
				arrParam(4) = ""
				arrParam(5) = "관리항목"

				arrField(0) = "CTRL_CD"
				arrField(1) = "CTRL_NM"
			Else
				arrParam(1) = " A_ACCT A , A_ACCT_CTRL_ASSN B, A_CTRL_ITEM C "
				arrParam(2) = Trim(txtValue)
				arrParam(3) = ""
				arrParam(4) = "A.ACCT_CD =B.ACCT_CD AND B.CTRL_CD = C.CTRL_CD AND A.ACCT_CD = " & FilterVar(frm1.txtACCT_CD.value, "''", "S") & "  "
				arrParam(5) = "관리항목"

				arrField(0) = "UPPER(C.CTRL_CD)"
				arrField(1) = "C.CTRL_NM"
			End If

			arrField(2) = ""
			arrField(3) = ""
			arrField(4) = ""
			arrField(5) = ""

			arrHeader(0) = "관리항목코드"
			arrHeader(1) = "관리항목명"
			arrHeader(2) = ""
			arrHeader(3) = ""
			arrHeader(4) = ""
			arrHeader(5) = "."
		Case 3
			arrParam(0) = "관리항목 팝업"
			arrParam(1) = "A_CTRL_ITEM a, (select distinct(ctrl_cd),mandatory_fg  from a_subsys_item) b "
			arrParam(2) = Trim(txtValue)
			arrParam(3) = ""
			arrParam(4) = "a.ctrl_cd*=b.ctrl_cd"
			arrParam(5) = "관리항목"

			arrField(0) = "a.CTRL_CD"
			arrField(1) = "a.CTRL_NM"
			arrField(2) = "isnull(b.mandatory_fg," & FilterVar("N", "''", "S") & " ) "
			arrField(3) = ""
			arrField(4) = ""
			arrField(5) = ""
			'arrField(6) = ""

			arrHeader(0) = "관리항목코드"
			arrHeader(1) = "관리항목명"
			arrHeader(2) = ""
			arrHeader(3) = ""
			arrHeader(4) = ""
			arrHeader(5) = "."
			'arrHeader(6) = ""
		Case 2
			If frm1.txtREL_BIZ_AREA_CD.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If
			arrParam(0) = "사업장 팝업"						' 팝업 명칭 
			arrParam(1) = "B_Biz_AREA"							' TABLE 명칭 
			arrParam(2) = Trim(txtValue)				 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "사업장"

			arrField(0) = "Biz_AREA_CD"							' Field명(0)
			arrField(1) = "Biz_AREA_NM"							' Field명(1)   

			arrHeader(0) = "사업장코드"						' Header명(0)
			arrHeader(1) = "사업장명"							' Header명(1)
		Case 4
			arrParam(0) = "재무제표코드팝업"					' 팝업 명칭 
			arrParam(1) = "B_MINOR"								' TABLE 명칭 
			arrParam(2) = Trim(txtValue)				 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("A1019", "''", "S") & " "					' Where Condition
			arrParam(5) = "재무제표코드"

			arrField(0) = "MINOR_CD"							' Field명(0)
			arrField(1) = "MINOR_NM"							' Field명(1)

			arrHeader(0) = "재무제표코드"						' Header명(0)
			arrHeader(1) = "재무제표명"						' Header명(1)
		Case 5
			If frm1.txtACCT_TYPE.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If
			arrParam(0) = "계정특성팝업"						' 팝업 명칭 
			arrParam(1) = "B_MINOR"								' TABLE 명칭 
			arrParam(2) = Trim(txtValue)				 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("A1000", "''", "S") & " "					' Where Condition
			arrParam(5) = "계정그룹특성코드"

			arrField(0) = "MINOR_CD"							' Field명(0)
			arrField(1) = "MINOR_NM"							' Field명(1)

			arrHeader(0) = "계정특성코드"						' Header명(0)
			arrHeader(1) = "계정특성명"						' Header명(1)
		Case 6
			arrParam(0) = "계정그룹특성팝업"					' 팝업 명칭 
			arrParam(1) = "B_MINOR"								' TABLE 명칭 
			arrParam(2) = Trim(txtValue)				 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("C2001", "''", "S") & " "					' Where Condition
			arrParam(5) = "계정그룹특성코드"

			arrField(0) = "MINOR_CD"							' Field명(0)
			arrField(1) = "MINOR_NM"							' Field명(1)

			arrHeader(0) = "계정그룹특성코드"					' Header명(0)
			arrHeader(1) = "계정그룹특성명"					' Header명(1)
		Case 7, 8
			If frm1.txtMgntCd1.readOnly = True Or  frm1.txtMgntCd2.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(0) = "관리항목 팝업"
			arrParam(1) = " A_ACCT A , A_ACCT_CTRL_ASSN B, A_CTRL_ITEM C "
			arrParam(2) = Trim(txtValue)
			arrParam(3) = ""
			arrParam(4) = "A.ACCT_CD =B.ACCT_CD AND B.CTRL_CD = C.CTRL_CD AND A.ACCT_CD = " & FilterVar(frm1.txtACCT_CD.value, "''", "S") & "  "
			arrParam(5) = "관리항목"

			arrField(0) = "UPPER(C.CTRL_CD)"
			arrField(1) = "C.CTRL_NM"
			arrField(2) = ""
			arrField(3) = ""
			arrField(4) = ""
			arrField(5) = ""
			'arrField(6) = ""

			arrHeader(0) = "관리항목코드"
			arrHeader(1) = "관리항목명"
			arrHeader(2) = ""
			arrHeader(3) = ""
			arrHeader(4) = ""
			arrHeader(5) = "."
			'arrHeader(6) = ""
		Case 9
			arrParam(0) = "전표항목 팝업"
			arrParam(1) = "A_SUBLEDGER_CTRL "
			arrParam(2) = Trim(txtValue)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "전표항목"

			arrField(0) = "GL_CTRL_FLD"
			arrField(1) = "GL_CTRL_NM"
			arrField(2) = ""
			arrField(3) = ""
			arrField(4) = ""
			arrField(5) = ""
			'arrField(6) = ""

			arrHeader(0) = "전표항목코드"
			arrHeader(1) = "전표항목명"
			arrHeader(2) = ""
			arrHeader(3) = ""
			arrHeader(4) = ""
			arrHeader(5) = "."
			'arrHeader(6) = ""					
    End select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case IntIndex
			Case 0
				frm1.txtSUBLEDGER1.focus
			Case 1
				frm1.txtSUBLEDGER2.focus
			Case 2
				frm1.txtREL_BIZ_AREA_CD.focus
			Case 3
				Call SetActiveCell(frm1.vspdData,C_CTRLITEM,frm1.vspdData.ActiveRow ,"M","X","X")		
			Case 4
				frm1.txtBS_PL_FG.focus
			Case 5
				frm1.txtACCT_TYPE.focus
			Case 6
				frm1.txtGP_TYPE.focus
			Case 7
				frm1.txtMgntCd1.focus
			Case 8
				frm1.txtMgntCd2.focus
			Case 9
				Call SetActiveCell(frm1.vspdData,C_GL_ITEM,frm1.vspdData.ActiveRow ,"M","X","X")						
			Case Else
				
		End Select
	Else
		Call SetPopUp(arrRet, IntIndex)
	End If
End Function

'========================================================================================================= 
Function SetPopUp(Byval arrRet, Byval IntIndex)
	Dim intRtnCnt, strData,IntRetCD,iExistfg,itempRow

	With frm1
		Select Case IntIndex
			Case 0
				.txtSUBLEDGER1.focus
				.txtSUBLEDGER1.value = Trim(arrRet(0))
				.txtSUBLEDGER1_Nm.value = arrRet(1)
			Case 1
				.txtSUBLEDGER2.focus
				.txtSUBLEDGER2.value = Trim(arrRet(0))
				.txtSUBLEDGER2_Nm.value = arrRet(1)
			Case 2
				.txtREL_BIZ_AREA_CD.focus
				.txtREL_BIZ_AREA_CD.value = Trim(arrRet(0))
				.txtREL_BIZ_AREA_nM.value = arrRet(1)
			Case 3
				iTempRow = 	.vspddata.row			
				iExistfg = CheckAddCtrlCd(.vspddata,arrRet(0))
				If iExistfg = "Y" Then
					IntRetCD = DisplayMsgBox("110115","X","X","X") 
					' %1  관리항목이 중복되게 입력될 수 없습니다.
					Exit Function
				Else
					.vspdData.Row = iTempRow
					.vspdData.Col  = C_CTRLITEM
					.vspdData.Text = arrRet(0)
					.vspdData.Col  = C_CTRLNM
					.vspdData.Text = arrRet(1)
'					.vspdData.Col  = C_MAND_FG
'					.vspdData.Text = arrRet(2)
					Call SetDrCrFg(.vspddata,"N",.vspdData.Row)
'					Call ChkGlItemValue(.vspddata,.vspddata.row,"N")
					Call vspdData_Change(.vspdData.Col, .vspdData.Row)									' 변경이 일어났다고 알려줌 
					Call SetActiveCell(.vspdData,C_CTRLITEM,.vspdData.ActiveRow ,"M","X","X")				
				End If					
			Case 4
				.txtBS_PL_FG.focus
				.txtBS_PL_FG.value    = Trim(arrRet(0))
				.txtBS_PL_FG_Nm.value = arrRet(1)
			Case 5
				.txtACCT_TYPE.focus
				.txtACCT_TYPE.value    = Trim(arrRet(0))
				.txtACCT_TYPE_Nm.value = arrRet(1)
			Case 6
				.txtGP_TYPE.focus
				.txtGP_TYPE.value  = Trim(arrRet(0))
				.txtGP_TYPE_Nm.value  = arrRet(1)
			Case 7
				.txtMgntCd1.focus
				.txtMgntCd1.value = Trim(arrRet(0))
				.txtMgntCd1_Nm.value = arrRet(1)
				Call txtMgntCd1_onChange()
			Case 8
				.txtMgntCd2.focus
				.txtMgntCd2.value = Trim(arrRet(0))
				.txtMgntCd2_Nm.value = arrRet(1)
				Call txtMgntCd2_onChange()
			Case 9
				.vspdData.Col  = C_GL_ITEM
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_SYSTEM_FG
				If .vspddata.value = "Y" Then
					.vspdData.Col  = C_MAND_FG
					Call SetDrCrFg(.vspddata,.vspdData.Text,.vspdData.Row)
					Call ChkGlItemValue(.vspddata,.vspddata.row)
				End If
				
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)										' 변경이 일어났다고 알려줌 
				Call SetActiveCell(.vspdData,C_GL_ITEM,.vspdData.ActiveRow ,"M","X","X")			
			Case Else
		End Select

	    lgBlnFlgChgValue = True
    End With
End Function

'==========================================================================================
'   Function Name :ChkDragState
'   Function Desc :Drag 가 어디에 있는지 Drag되는 항목인지 체크 
'==========================================================================================
Function  ChkDragState(ByVal x , ByVal y )
	Dim NewNode
    dim ChildNode
    Dim i

    On Error Resume Next

    ChkDragState = False

    With frm1
		If gDragNode Is Nothing Then Exit Function

		If gDragNode.parent Is Nothing Then Exit Function	' 자신이 Root인 경우 

		Set NewNode = .uniTree1.HitTest(x, y)

		' 폴더가 지정되지 않고 여백이나 기타에 Drop했을 경우 
		If NewNode Is Nothing Then Exit Function

		' 유저메뉴가 아닌곳을 갈때 
		If ChkUserMenu(NewNode, C_USER_MENU_KEY) = False Then Exit Function

		' 자신의 자식에게 갈때 
		If InStr(1, NewNode.Key, gDragNode.Key, vbTextCompare) > 0 Then
		    Set NewNode = Nothing
		    Exit Function
		End If

		'자신의 자리에 있을때 
		If NewNode.Text = gDragNode.Text Then
		    Set NewNode = Nothing
		    Exit Function
		End If

		' URL에 Drop하면 , 즉 폴더가 아닌 최하단일 경우 
		If NewNode.Image = C_URL Then
		    Set NewNode = Nothing
		    Exit Function
		End If

		' 자신의 부모에게 갈때 
		If NewNode.Key = gDragNode.Parent.Key Then
		    Set NewNode = Nothing
		    Exit Function
		End If

		If NewNode.Children > 0 Then 
			Set ChildNode = NewNode.Child
		End If

		For i = 1 To NewNode.Children
			If ChildNode.Key = gDragNode.Key Then
		  		Set NewNode = Nothing
				Exit Function
			End If
			Set ChildNode = ChildNode.Next
		Next

		Set ChildNode = Nothing
		Set NewNode = Nothing
    End With

    ChkDragState = True
    Exit Function
End Function

' UserMenu를 찾는 재귀함수 
Function ChkUserMenu(ParentNode, strFind)
	Dim blnFind

	blnFind = False
	ChkUserMenu = blnFind

	If ParentNode is Nothing Then Exit Function

	If ParentNode.Key <> strFind Then
		blnFind = ChkUserMenu(ParentNode.Parent, strFind)
	Else
		blnFind = True
	End If

	ChkUserMenu = blnFind
End Function

'==========================================================================================
'   Function Name : GetNodeLvl
'   Function Desc : 현재 노드의 Level을 찾는다.
'==========================================================================================
Function  GetNodeLvl(Node)
    Dim tempNode

    Set tempNode = Node
    GetNodeLvl = 0

    If tempNode.Key <> "$" Then
	    Do    	
    		GetNodeLvl = GetNodeLvl + 1
    		Set tempNode = tempNode.Parent
    	Loop Until tempNode.Key = "$"
	End If

	Set tempNode = Nothing
End Function

'==========================================================================================
'   Function Name :GetIndex
'   Function Desc :Node가 부모의 몇번째 위치인가 되돌려준다.
'==========================================================================================
Function GetIndex(Node)
	Dim i, myIndx,  ChildNode, ParentNode

	Set ParentNode = Node.Parent

	If ParentNode Is Nothing Then	' Root일 경우 
		GetIndex = 1
		Exit Function
	End If

	Set ChildNode = ParentNode.Child
	myIndx = 1

	For i = 1 To ParentNode.Children
		If ChildNode.Key = Node.Key Then
			GetIndex = myIndx
			Exit Function
		End If

		If Node.Image = ChildNode.Image Then
			myIndx = myIndx + 1
		End If

		Set ChildNode = ChildNode.Next
	Next
End Function

'==========================================================================================
'   Function Name :GetInsSeq
'   Function Desc : 현재 Insert 되는 Node의 순서를 리턴한다.
'==========================================================================================
Function GetInsSeq(Node)
	Dim i, myIndx,  ChildNode, ParentNode

	Set ChildNode = Node.Child

	myIndx = 1

	For i = 1 To Node.Children
		If gDragNode.Image = ChildNode.Image Then
			myIndx = myIndx + 1
		End If
		
		Set ChildNode = ChildNode.Next
	Next

	GetInsSeq = myIndx
End Function

'==========================================================================================
'   Function Name :GetTotalCnt
'   Function Desc :Add에 관련되 자식수를 되돌려준다.
'==========================================================================================
Function GetTotalCnt(Node)
	If Node.children = 0 Then	' Root일 경우 
		GetTotalCnt = 1
	Else
		GetTotalCnt = Node.children + 1
	End If
End Function

'======================================================================================================
'	화면 사이즈 변경 
'======================================================================================================
Sub DispDivConf(pVal) 
	If pVal = 2 Then
		divconf.style.display = "none"
		tdConf.height = 1
	Else
		divconf.style.display = ""
		tdConf.height = 22
	End If
End Sub

'======================================================================================================
'	메뉴 
'======================================================================================================
Sub MenuRefresh()
	If lgBlnBizLoadMenu = False Then
		Call DisplayAcct()
	End If
End Sub

'======================================================================================================
'	메뉴를 읽어 TreeView에 넣음 
'======================================================================================================
Sub  DisplayAcct()
	Dim NodX

	frm1.uniTree1.Nodes.Clear 

	Set NodX = frm1.uniTree1.Nodes.Add(, tvwChild, C_USER_MENU_KEY, lgUSER_MENU, C_Root, C_Root)

	Call SetDefaultVal()

	frm1.uniTree1.MousePointer = 11

	Call AddNodes(C_CMD_TOP_LEVEL)
End Sub

'========================================================================================
Function DisplayAcctOK()
	Dim NodX

	Set NodX = frm1.uniTree1.Nodes(C_USER_MENU_KEY)

	If Not (nodX.child Is Nothing) Then
		Call uniTree1_NodeClick(nodX.child)
	End If
End Function

'========================================================================================
' Function Name : GetImage
' Function Desc : 이미지 정보 
'========================================================================================
Function GetImage(Byval arrLine)
	Dim strImg
	Select Case arrLine(C_MNU_AUTH)
		Case "A"
			If arrLine(C_MNU_TYPE) = "M" Then
				strImg = C_Folder
			Else
				strImg = C_URL
			End If
		Case "I"
			strImg = C_Const
		Case "N"
			strImg = C_None
	End Select
	
	GetImage = strImg
End Function

'========================================================================================
' Function Name : MakeFolderNodeDataForInsert
' Function Desc : 업무메뉴에서 유저메뉴로 등록시 폴더정보를 읽을때 
'========================================================================================
Function MakeFolderNodeDataForInsert(lDragNode, strKey)
	Dim CNode, strVal, i, strUpKey

	With frm1
		Set CNode = lDragNode.child		' 자식 노드를 할당 

		If CNode is Nothing Then Exit Function

		For i = 1 To lDragNode.children
			If CNode.Image = C_Folder Then	' 자식노드가 서브폴더일때 
				strVal = strVal & MakeNodeDataForIU(CNode, strKey, i)
				strUpKey = strKey & C_UNDERBAR & CNode.key
				strVal = strVal & MakeFolderNodeDataForInsert(CNode, strUpKey)
			Else		' 자식 노드가 프로그램일때 
				strVal = strVal & MakeNodeDataForIU(CNode, strKey, i)
			End If
				
			Set CNode = CNode.Next

			If CNode Is Nothing Then 
				MakeFolderNodeDataForInsert = strVal
				Exit Function
			End If
		Next

		MakeFolderNodeDataForInsert = strVal
	End With
End Function

'========================================================================================
' Function Name : RemoveUpperString
' Function Desc : 
'========================================================================================
Function RemoveUpperString(Byval Node)
	If Node.parent Is Nothing Then 
		RemoveUpperString = Node.Key
		Exit Function
	End If
	
	RemoveUpperString = Replace(Node.key, Node.parent.key & C_UNDERBAR , "")
End Function

'========================================================================================
' Function Name : MakeNodeDataForIU
' Function Desc : 유저메뉴로 등록/이동시 Node 단위 정보를 읽을때 
'========================================================================================
Function MakeNodeDataForIU(lDragNode, strUpKey, Index)
	Dim strVal

	' 0: 신규/수정 상태 
	strVal = strVal & lgIntFlgMode & parent.gColSep		' 신규/수정 상태 

	' 1: Menu ID
	If lgIntFlgMode = parent.OPMD_CMODE Then
		strVal = strVal & strUpKey & C_UNDERBAR & lDragNode.key & parent.gColSep			'☆: Drag 한 폴더/파일의 키 
	Else
		strVal = strVal & lDragNode.key & parent.gColSep			'☆: Drag 한 폴더/파일의 키 
	End If

	' 2: Upper Menu ID
	strVal = strVal & strUpKey & parent.gColSep								'☆: Drop 한 폴더의 키 

	' 3: Menu Name
	strVal = strVal & lDragNode.Text & parent.gColSep								'☆: Drag 한 폴더/파일의 이름 

	' 4: Menu Type
    If lDragNode.image = C_Folder Then
		strVal = strVal & "M" & parent.gColSep
	Else
		strVal = strVal & "P" & parent.gColSep
	End If

	' 5: Menu Seq
	strVal = strVal & Index & parent.gColSep							'☆: Drop 한 폴더/파일의 키 

	' 6: PrevID, PrevUppderID
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = strVal & lDragNode.key	& parent.gColSep							'☆: Drag 한 폴더/파일의 키 
		strVal = strVal & lDragNode.Parent.key & parent.gColSep					'☆: Drag 한 폴더/파일의 키 
	Else
		strVal = strVal & parent.gColSep
		strVal = strVal & parent.gColSep
    End If

	strVal = strVal & parent.gRowSep

	MakeNodeDataForIU = strVal
End Function

'========================================================================================
' Function Name : MakeFolderNodeDataForDelete
' Function Desc : 유저메뉴에서 삭제시 폴더정보를 읽을때 
'========================================================================================
Function MakeFolderNodeDataForDelete(Node)
	Dim CNode, strVal, i

	With frm1
		Set CNode = Node.child		' 자식 노드를 할당 

		If CNode Is Nothing Then Exit Function

		For i = 1 to Node.children
			If CNode.Image = C_Folder Then	' 자식노드가 서브폴더일때 
				strVal = strVal & MakeNodeDataForDelete(CNode)

				strVal = strVal & MakeFolderNodeDataForDelete(CNode)
			Else		' 자식 노드가 프로그램일때 
				strVal = strVal & MakeNodeDataForDelete(CNode)

			End If

			Set CNode = CNode.Next

			If CNode Is Nothing Then 
				MakeFolderNodeDataForDelete = strVal
				Exit Function
			End If
		Next
		
		MakeFolderNodeDataForDelete = strVal
	End With
End Function

'========================================================================================
' Function Name : MakeNodeData
' Function Desc : 유저메뉴에서 유저메뉴로 이동/삭제시 Node 단위 정보를 읽을때 
'========================================================================================
Function MakeNodeDataForDelete(Node)
	Dim strVal

	' 0: 신규/수정 상태 
	strVal = strVal & lgIntFlgMode & parent.gColSep		' 신규/수정 상태 

	' 1: Menu ID
	strVal = strVal & Node.key & parent.gColSep							'☆: Drag 한 폴더/파일의 키 

	' 2: Upper Menu ID
	strVal = strVal & Node.parent.key & parent.gColSep						'☆: Drop 한 폴더의 키 

	' 3: Menu Name
	strVal = strVal & Node.Text & parent.gColSep							'☆: Drag 한 폴더/파일의 이름 

	' 4: Menu Type
    If Node.image = C_Folder Then
		strVal = strVal & "M" & parent.gColSep
	Else
		strVal = strVal & "P" & parent.gColSep
	End If

	' 5: Menu Seq
	strVal = strVal & GetIndex(Node) & parent.gColSep						'☆: Drop 한 폴더/파일의 키 

	' 6: PrevID
	If lgIntFlgMode = parent.OPMD_UMODE Or lgIntFlgMode = parent.UID_M0003 Then
		strVal = strVal & Node.key	& parent.gColSep							'☆: Drag 한 폴더/파일의 키 
		strVal = strVal & Node.Parent.key & parent.gColSep					'☆: Drag 한 폴더/파일의 키 
	Else
		strVal = strVal & parent.gColSep
		strVal = strVal & parent.gColSep
    End If

	strVal = strVal & parent.gRowSep

	MakeNodeDataForDelete = strVal
End Function

'========================================================================================================= 
Sub  Form_Load()
	Dim intColCnt

    Call InitVariables
    Call LoadInfTB19029

    Call ggoOper.LockField(Document, "N")
	Call AppendNumberPlace("7","3","0")
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "3", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call InitSpreadSheet
    Call InitCombo
    Call SetToolbar("1100100000001111")
    
    With frm1
		.uniTree1.SetAddImageCount = 6
		.uniTree1.Indentation = "200"	' 줄 간격 
						' 파일위치,	키명, 위치 
		.uniTree1.AddImage C_IMG_Root,		C_Root,		0
		.uniTree1.AddImage C_IMG_Folder,	C_Folder,	0
		.uniTree1.AddImage C_IMG_Open,		C_Open,		0
		.uniTree1.AddImage C_IMG_URL,		C_URL,		0
		.uniTree1.AddImage C_IMG_None,		C_None,		0
		.uniTree1.AddImage C_IMG_Const,		C_Const,	0

		.uniTree1.OLEDragMode = 1														'⊙: Drag & Drop 을 가능하게 할 것인가 정의 
		.uniTree1.OLEDropMode = 1

		.uniTree1.OpenTitle = "계정그룹입력"
		.uniTree1.AddTitle = "계정입력"
		.uniTree1.RenameTitle = ""
		.uniTree1.DeleteTitle = "삭제"
	End With

	Set gDragNOde = Nothing
	lglsClicked = False
End Sub

'==========================================================================================
Sub Mgnt_QueryOk()
	If frm1.txtOpenAcctFg.value = "Y" Then
'			Call ElementVisible(frm1.cboMgntFg,"1")
		Call ElementVisible(frm1.cboMgntType,"1")
		Call ElementVisible(frm1.txtMgntCd1,"1")
		Call ElementVisible(frm1.txtMgntCd1_Nm,"1")
		Call ElementVisible(frm1.txtMgntCd2,"1")
		Call ElementVisible(frm1.txtMgntCd2_Nm,"1")
		Call ElementVisible(frm1.btnCalType3,"1")
		Call ElementVisible(frm1.btnCalType4,"1")
'			spnMgntFg.innerHTML		= "미결관리여부"
		spnMgntType.innerHTML	= "미결관리구분"
		spnMgntCd1.innerHTML	= "미결관리항목1"
		spnMgntCd2.innerHTML	= "미결관리항목2"
		
		Call ggoOper.SetReqAttr(frm1.cboMgntType,"N")
		Call ggoOper.SetReqAttr(frm1.txtMgntCd1,"N")						
	Else
'			Call ElementVisible(frm1.cboMgntFg,"0")
		Call ElementVisible(frm1.cboMgntType,"0")
		Call ElementVisible(frm1.txtMgntCd1,"0")
		Call ElementVisible(frm1.txtMgntCd1_Nm,"0")
		Call ElementVisible(frm1.txtMgntCd2,"0")
		Call ElementVisible(frm1.txtMgntCd2_Nm,"0")
		Call ElementVisible(frm1.btnCalType3,"0")
		Call ElementVisible(frm1.btnCalType4,"0")
'		spnMgntFg.innerHTML = ""
		spnMgntType.innerHTML = ""
		spnMgntCd1.innerHTML = ""
		spnMgntCd2.innerHTML = ""
		
		frm1.cboMgntType.Value  = ""		
		Call ggoOper.SetReqAttr(frm1.cboMgntType,"D")
		Call ggoOper.SetReqAttr(frm1.txtMgntCd1,"D")			
	End If
End Sub

'========================================================================================
' Function Name : InsMgntItem(2003-8-25 BY JYK)
' Function Desc : Sub System 콤보박스에서 하나의 subsustem_item을 선택했을때 
'========================================================================================
Sub InsMgntItem(ByVal SubSystemTypeCd)
	Dim strSelect,strFrom,strWhere
	Dim arrVal1,arrVal2,ii,iMaxRow,iMandfg
	Dim iExistfg,arrVal3,arrVal4

	On Error Resume Next
	Err.Clear

	With frm1
		If SubSystemTypeCd = "" Then
			ggoSpread.Source = .vspddata
			For ii = 1 To .vspddata.Maxrows
				.vspddata.row = ii
				.vspddata.col = 0
				If Trim(.vspddata.value) = ggoSpread.InsertFlag Then
					ggoSpread.EditUndo(ii)
					ii = ii - 1
				Else
					.vspddata.col = C_SYSTEM_FG
					If .vspddata.value = "Y" Then
						.vspddata.value = "N"
						ggoSpread.UpdateRow(ii)
						.vspddata.col = C_MAND_FG
						.vspddata.value = "N"
						Call SetDrCrFg(.vspddata,.vspddata.value,ii)
						ggoSpread.SSSetRequired  C_CRFG	 , ii, ii
						ggoSpread.SSSetRequired  C_DRFG	 , ii, ii
					End If						
				End If
			Next 
			Exit Sub
		End If	

		Call UndoSubSystem(.vspddata)
	
		strSelect   = " isnull(subsys_type,'') "
		strFrom		= " a_acct "
		strWhere	= " acct_cd = " & FilterVar(.txtACCT_CD.value, "''", "S")
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrVal3 = Split(lgF2By2, Chr(11) & Chr(12))
			arrVal4	= Split(arrVal3(0), chr(11))
			If Trim(arrVal(4)) <> "" Then	
				Call DelSubSysItem(.vspddata,.txtACCT_CD.value,arrVal4(1),SubSystemTypeCd)
			End If				
		End If	

		strSelect	= " a.ctrl_cd,a.mandatory_fg,b.ctrl_nm "
		strFrom		= " a_subsys_item a, a_ctrl_item b "
		strWhere	= " subsys_type = " & FilterVar(SubSystemTypeCd, "''", "S")
		strWhere	= strWhere & " and a.ctrl_cd=b.ctrl_cd "

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
			iMaxRow = .vspddata.MaxRows
			
			For ii = 0 To Ubound(arrVal1,1) - 1
				arrVal2	= Split(arrVal1(ii), chr(11))

				iExistfg = CheckAddCtrlCd(frm1.vspddata, Trim(arrVal2(1)))

				If iExistfg = "N" Then
					.vspddata.Redraw = False
					.vspddata.row = .vspddata.MaxRows
					ggoSpread.insertRow,1
					.vspddata.row = .ActiveRow

					.vspddata.col   = C_CTRLITEM
					.vspddata.value	= Trim(arrVal2(1))
					.vspddata.col   = C_CTRLNM
					.vspddata.value	= Trim(arrVal2(3))
					.vspddata.col   = C_DEFAULT_VALUE
					.vspddata.value = ""
					.vspddata.col   = C_GL_ITEM
					.vspddata.value = ""
					.vspddata.col   = C_SYSTEM_FG
					.vspddata.value = "Y"
					.vspddata.Redraw = True
				End If

				iMandfg = Trim(arrVal2(2))

				Call SetDrCrFg(.vspddata,iMandfg,.vspddata.row)
				Call ChkGlItemValue(.vspddata,.vspddata.row)
				Call SetSpdAddColor(.vspddata,.vspddata.row,"Q","N")  '입력모드이지만 서브시스템관련 입력이므로 관리항목코드 프로텍트 
			Next
		End If	

		Call ChkCount(.vspddata,"Y")
	End With			
End Sub

'========================================================================================
' Function Name : CheckAddCtrlCd(2003-8-25 BY JYK)
' Function Desc : 서브 시스템에서 입력시 기존에 관리항목값이 있는지 체크 
'========================================================================================
Function CheckAddCtrlCd(ByVal CurSpd,ByVal CtrlCd) 
	Dim iObjSpread 
	Dim jj

	On Error Resume Next
	Err.Clear

	Set iObjSpread = CurSpd
	ggoSpread.Source = CurSpd
	
	With iObjSpread
    	.Redraw = False	
		For jj = 1 To .MaxRows
			.row = jj
			.col = 0
			If Trim(.value) <> ggoSpread.DeleteFlag Then
				.col = C_CTRLITEM
				If UCase(Trim(CtrlCd)) = UCase(Trim(.value)) Then
					CheckAddCtrlCd = "Y"
					Exit Function
				End If
			End If				
		Next
    	.Redraw = True
	End With
	
	CheckAddCtrlCd = "N"
End Function

'========================================================================================
' Function Name : SetDrCrFg(2003-8-25 BY JYK)
' Function Desc : 관리항목에 추가된 로우의 컬럼에 디폴트 값을 입력 
'========================================================================================
Sub SetDrCrFg(ByVal CurSpd,ByVal MandatoryFg,ByVal Row)
	Dim iObjSpread 

	On Error Resume Next
	Err.Clear

	Set iObjSpread = CurSpd
	
	With iObjSpread
    	.Redraw = False	
		.row   = Row

		.col   = C_GL_ITEM		
		If .value = "" Then
			If UCase(Trim(frm1.cboBAL_FG.value)) = "DR" Then
				.col = C_DRFG
				.value = "0"
				.col = C_CRFG
				.value = "1"
			End If	
	
			If UCase(Trim(frm1.cboBAL_FG.value)) = "CR" Then
				.col = C_CRFG
				.value = "0"
				.col = C_DRFG
				.value = "1"					
			End If	

			.col   = C_MAND_FG
			.value = MandatoryFg

			If UCase(Trim(MandatoryFg)) = "Y" Then
				.col = C_CRFG
				.value = "0"
				.col = C_DRFG
				.value = "0"
			End If
			
			.col = C_CTRLITEM
			
			If UCase(Trim(frm1.txtSUBLEDGER1.value)) = UCase(Trim(.vlaue))  Or  _
			   UCase(Trim(frm1.txtSUBLEDGER2.value)) = UCase(Trim(.vlaue)) Then
				.col = C_CRFG
				.value = "0"
				.col = C_DRFG
				.value = "0"			
			End If
			
		End If
    	.Redraw = True
	End With			
End Sub

'========================================================================================
' Function Name : ChkGlItemValue(2003-8-25 BY JYK)
' Function Desc : 전표항목값의 변경에 따른 다른 값들의 세팅 
'========================================================================================
Sub ChkGlItemValue(ByVal CurSpd,ByVal Row)
	Dim iObjSpread

	On Error Resume Next
	Err.Clear

	Set iObjSpread = CurSpd

	With iObjSpread
    	.Redraw = False	
		.col = C_GL_ITEM
		If Trim(.value) <> "" Then
			.col = C_DRFG
			.value = "1"
			.col = C_CRFG
			.value = "1"
			.col = C_MAND_FG			
			Call SetDrCrFg(CurSpd,.value,Row)
			.col = C_SYSTEM_FG
			If .value = "Y" Then
				Call SetSpdAddColor(CurSpd,Row,"Q","Y")							
			End If	
		Else
			.col = C_MAND_FG
			Call SetDrCrFg(CurSpd,.value,Row)
			.col = C_SYSTEM_FG
			If .value = "Y" Then			
				Call SetSpdAddColor(CurSpd,Row,"Q","N")							
			End If				
		End If
    	.Redraw = True
	End With
End Sub

'========================================================================================
' Function Name : SetSpdAddColor(2003-8-25 BY JYK)
' Function Desc : 관리항목을 인서트후 해당 컬럼의 입력필수 여부 및 락 적용 
'========================================================================================
Sub SetSpdAddColor(ByVal CurSpd,ByVal Row,ByVal Meth,ByVal GlValue)
	Dim iObjSpread 
	
	On Error Resume Next
	Err.Clear

	Set iObjSpread = CurSpd
	ggoSpread.Source = CurSpd

    With iObjSpread
    	.Redraw = False

		If Meth = "Q" Then
			ggoSpread.SSSetProtected  C_CTRLITEM , Row, Row
		Else
			ggoSpread.SSSetRequired  C_CTRLITEM	 , Row, Row
		End If
		ggoSpread.SSSetProtected  C_CTRLNM	     , Row, Row
		.col = C_MAND_FG
		.row = Row
		If 	UCase(Trim(.value)) = "Y" Or GlValue = "Y" Then
			ggoSpread.SSSetProtected C_DRFG	     , Row, Row
			ggoSpread.SSSetProtected C_CRFG	     , Row, Row
		Else
			.col = C_GL_ITEM
			If Trim(.value) = "" Then
				If UCase(Trim(frm1.cboBAL_FG.value)) = "DR" Then
					ggoSpread.SSSetRequired  C_DRFG	 , Row, Row
					ggoSpread.SSSetProtected C_CRFG	 , Row, Row
				Else			
					ggoSpread.SSSetRequired  C_CRFG	 , Row, Row
					ggoSpread.SSSetProtected C_DRFG	 , Row, Row				
				End If				
			Else
				ggoSpread.SSSetProtected C_DRFG	     , Row, Row
				ggoSpread.SSSetProtected C_CRFG	     , Row, Row			
			End If				
		End If
		.Col = 1
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True

		.Redraw = True
    End With	
End Sub

'========================================================================================
' Function Name : ChkCount(2003-8-25 BY JYK)
' Function Desc : 관리항목의 총갯수가 8개를 초과하는지 여부 판단 
'========================================================================================
Sub ChkCount(ByVal CurSpd,ByVal IsSubSystem)
	Dim iObjSpread
	Dim ii 
	Dim iChkCount
	Dim IntRetCD	
	
	Set iObjSpread = CurSpd	
	ggoSpread.Source = CurSpd	
	
	With iObjSpread
    	.Redraw = False	
		iChkCount = .MaxRows
		For ii = 1 To .MaxRows
			.col = 0
			.row = ii
			If .value = ggoSpread.DeleteFlag Then
				iChkCount = iChkCount - 1
			End If
		Next

		If iChkCount > 8 Then
			IntRetCD = DisplayMsgBox("110304","X","X","X")  				
			If IsSubSystem = "Y" Then		
				Call UndoSubSystem(CurSpd)
			Else
				ggoSpread.EditUndo
			End If	
		End If
    	.Redraw = True
	End With		
End Sub

'========================================================================================
' Function Name : UndoSubSystem(2003-8-25 BY JYK)
' Function Desc : 서브시스템 변경시(저장안된것중 서브시스템에서 추가된것 롤백)
'========================================================================================
Sub  UndoSubSystem(ByVal CurSpd)
	Dim iObjSpread	
	Dim ii	

	Set iObjSpread = CurSpd

	With iObjSpread
    	.Redraw = False	
		For ii = 1 To .MaxRows
			.row = ii
			.col = 0
			If .value <> "" Then
'				.col = C_SYSTEM_FG
'				If .value = "Y" Then
					ggoSpread.Source = CurSpd
					ggoSpread.EditUndo(ii)									'행 취소 실행 
					ii = ii - 1												'Row가 하나 줄었으므로 
'				End If
			End If								
		Next
    	.Redraw = True
	End With
End Sub
	
'====================================================================================================
' Function Name : DelSubSysItem(2003-8-25 BY JYK)
' Function Desc : 서브시스템 변경시(기존에 저장된것중 서브시스템변경으로 인한 이전관리항목 삭제)
'====================================================================================================
Sub DelSubSysItem(ByVal CurSpd,ByVal acct_cd,ByVal old_subsys,Byval new_subsys)
	Dim iObjSpread	
	Dim ii,jj,arrVal1,arrVal2,lDelRows
	Dim strSelect,strFrom,strWhere

	Set iObjSpread = CurSpd	
	ggoSpread.Source = CurSpd

	strSelect	= " distinct(a.ctrl_cd)  "
	strFrom		= " a_acct_ctrl_assn a , a_subsys_item b "
	strWhere	= " a.acct_cd = " & FilterVar(acct_cd, "''", "S")
	strWhere	= strWhere & " and b.subsys_type in ( " & FilterVar(old_subsys, "''", "S") & ","
	strWhere	= strWhere & FilterVar(new_subsys, "''", "S") & ")"
	strWhere	= strWhere & " and a.ctrl_cd=b.ctrl_cd "

	With iObjSpread
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
			For ii = 0 To Ubound(arrVal1,1) - 1

				arrVal2	= Split(arrVal1(ii), chr(11))						'기존에 저장되어 있는 관리항목중 subsystem관련된것들				
				For jj = 1 To .MaxRows
					.row = jj												'Flag가 없는 것들중 
					.col = 0
					If .value = "" Then
'						.col = C_SYSTEM_FG									'sus system 이 "Y" 인 것 
'						If .value = "Y" Then
							.col = C_CTRLITEM
							If Trim(.value) = Trim(arrVal2(1)) Then			'기존관리항목과 일치하는 것이 있으면 삭제 
								.col = C_CHG_DEL
								.value = "Y"
								lDelRows = ggoSpread.DeleteRow(jj,jj) 
 								Exit for
							End If
'						End If	
					End If
				Next	
			Next				
		End If
	End With		
End Sub

'==========================================================================================
Sub  cboSubSystemType_onChange()
	Dim isOpenAcct 

	If frm1.cboSubSystemType.value = "OD" Or frm1.cboSubSystemType.value = "OC" Then
		frm1.txtOpenAcctFg.value = "Y"		
	Else
		frm1.txtOpenAcctFg.value = "N"
	End If		

	Call InsMgntItem(frm1.cboSubSystemType.value)
	Call Mgnt_QueryOk()
	
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  cboBDG_CTRL_FG_onchange()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  cboFX_EVAL_FG_onchange()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  cboDEL_FG_onchange()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  txtACCT_TYPE_onchange()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strGpCd
	Dim strGpNm
	Dim ii, jj
	Dim arrVal1, arrVal2

	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	= " MINOR_NM"
	strFrom		= " B_MINOR(NOLOCK) "
	strWhere	= "MAJOR_CD = " & FilterVar("A1000", "''", "S") & " "
	strWhere	= strWhere & " AND MINOR_CD = " & FilterVar(frm1.txtACCT_TYPE.value, "''", "S")

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			frm1.txtACCT_TYPE_nm.value		= Trim(arrVal2(1))
		Next
	End If
	
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  cboBAL_FG_onchange()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  cboTEMP_ACCT_FG_onchange()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  txtBS_PL_FG_onchange()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strGpCd
	Dim strGpNm
	Dim ii, jj
	Dim arrVal1, arrVal2
	
	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	= " MINOR_NM"
	strFrom		= " B_MINOR(NOLOCK) "
	strWhere	= "MAJOR_CD = " & FilterVar("A1019", "''", "S") & " "
	strWhere	= strWhere & " AND MINOR_CD = " & FilterVar(frm1.txtBS_PL_FG.value, "''", "S")

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			frm1.txtBS_PL_FG_nm.value		= Trim(arrVal2(1))
		Next
	End If
	
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  txtACCT_SEQ_change()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  cboHQ_BRCH_FG_onchange()
	lgBlnFlgChgValue = True

	Select Case Trim(frm1.cboHQ_BRCH_FG.value)
		Case "Y"
			'ggoOper.SetReqAttr(Object, Option)		'N:Required, Q:Protected, D:Default
			Call ggoOper.SetReqAttr(frm1.txtREL_BIZ_AREA_CD, "N")
		Case Else
			frm1.txtREL_BIZ_AREA_CD.value	= ""
			frm1.txtREL_BIZ_AREA_NM.value	= ""

			Call ggoOper.SetReqAttr(frm1.txtREL_BIZ_AREA_CD, "Q")
	End Select
End Sub

'==========================================================================================
'Sub  cboMgntFg_onchange()
'	lgBlnFlgChgValue = True
'
'	Select Case Trim(frm1.cboMgntFg.value)
'	Case "Y"
'		Call ggoOper.SetReqAttr(frm1.cboMgntType, "N")
'	Case Else
'		frm1.cboMgntType.value	= ""
'		Call ggoOper.SetReqAttr(frm1.cboMgntType, "Q")
'
'	End Select
'
'End Sub

Sub  cboMgntType_onchange()
'	Select Case Trim(frm1.cboMgntFg.value)
'	Case "Y"
		lgBlnFlgChgValue = True
'	End Select
End Sub
'==========================================================================================
Sub  subledger_change()
	lgBlnFlgChgValue = True

	Select Case Trim(frm1.txtsubledger_modigy_fg.value)
		Case "Y"
			Call ggoOper.SetReqAttr(frm1.txtSUBLEDGER1, "Q")
			Call ggoOper.SetReqAttr(frm1.txtSUBLEDGER2, "Q")
		Case Else
			Call ggoOper.SetReqAttr(frm1.txtSUBLEDGER1, "D")
			Call ggoOper.SetReqAttr(frm1.txtSUBLEDGER2, "D")
	End Select
End Sub

'==========================================================================================
Sub  mgnt_change()
'	lgBlnFlgChgValue = True
'	Select Case Trim(frm1.txtmgnt_modigy_fg.value)
'		Case "Y"
'			Call ggoOper.SetReqAttr(frm1.txtMgntCd1, "Q")
'			Call ggoOper.SetReqAttr(frm1.txtMgntCd2, "Q")
'		Case Else
'			Call ggoOper.SetReqAttr(frm1.txtMgntCd1, "D")
'			Call ggoOper.SetReqAttr(frm1.txtMgntCd2, "D")
'	End Select
End Sub

'==========================================================================================
Sub  accttype_change()
	lgBlnFlgChgValue = True

	Select Case Trim(frm1.txtaccttype_modigy_fg.value)
		Case "Y"
			Call ggoOper.SetReqAttr(frm1.txtACCT_TYPE, "Q")
		Case Else
			Call ggoOper.SetReqAttr(frm1.txtACCT_TYPE, "D")
	End Select
End Sub

'==========================================================================================
Sub  cboGP_BDG_CTRL_FG_onchange()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  txtGP_TYPE_onchange()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strGpCd
	Dim strGpNm
	Dim ii, jj
	Dim arrVal1, arrVal2

	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	= " MINOR_NM"
	strFrom		= " B_MINOR(NOLOCK) "
	strWhere	= " MAJOR_CD = " & FilterVar("C2001", "''", "S") & " "
	strWhere	= strWhere & " AND MINOR_CD = " & FilterVar(frm1.txtGP_TYPE.value, "''", "S")

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			frm1.txtGP_TYPE_nm.value	= Trim(arrVal2(1))
		Next
	End If
	
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  txtMgntCd1_onChange()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strGpCd
	Dim strGpNm
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim IntRetCD
	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------

	strSelect	= " C.CTRL_NM "
	strFrom		= " A_ACCT A , A_ACCT_CTRL_ASSN B, A_CTRL_ITEM C "
	strWhere	= "A.ACCT_CD =B.ACCT_CD AND B.CTRL_CD = C.CTRL_CD AND A.ACCT_CD = " & FilterVar(frm1.txtACCT_CD.value, "''", "S") & "  "
	strWhere	= strWhere & " AND C.CTRL_CD = " & FilterVar(frm1.txtMgntCd1.value, "''", "S")
	'msgbox " select " & strSelect & " From " & strFrom & " where " & strWhere
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1		
			arrVal2			= Split(arrVal1(ii), chr(11))
			frm1.txtMgntCd1_Nm.value	= Trim(arrVal2(1))
		Next
	Else
'			IntRetCD = DisplayMsgBox("119353","X","X","X")  
'		frm1.txtMgntCd1.value = ""
		frm1.txtMgntCd1_Nm.value = ""
	End If

	lgBlnFlgChgValue = True
End Sub

Sub  txtMgntCd2_onChange()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strGpCd
	Dim strGpNm
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim IntRetCD

	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	= " C.CTRL_NM "
	strFrom		= " A_ACCT A , A_ACCT_CTRL_ASSN B, A_CTRL_ITEM C "
	strWhere	= "A.ACCT_CD =B.ACCT_CD AND B.CTRL_CD = C.CTRL_CD AND A.ACCT_CD = " & FilterVar(frm1.txtACCT_CD.value, "''", "S") & "  "
	strWhere	= strWhere & " AND C.CTRL_CD = " & FilterVar(frm1.txtMgntCd2.value, "''", "S")
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			frm1.txtMgntCd2_Nm.value	= Trim(arrVal2(1))
		Next
	Else
'			IntRetCD = DisplayMsgBox("119353","X","X","X")
'		frm1.txtMgntCd2.value = ""
		frm1.txtMgntCd2_Nm.value = ""
	End If

	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub  vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
				Case C_CTRLITEMPB
					.Col = Col - 1
					.Row = Row
					Call OpenPopup(.text, 3)
				Case C_GL_ITEMPB
					.Col = Col - 1
					.Row = Row
					Call OpenPopup(.text, 9)				
			End Select
		End If
    End With
End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col , ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"   
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
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
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
Sub  vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    Select Case Col
		Case C_DEFAULT_VALUE,C_GL_ITEM
			Call ChkGlItemValue(frm1.vspdData,Row)
		Case Else
		
	End Select	
End Sub

'======================================================================================================
'	창닫기 이벤트 
'======================================================================================================
Function button1_onclick()
End Function

'======================================================================================================
'	추가 
'======================================================================================================
Function Add_onclick()
	Dim strRetValue
	strRetValue = window.showModalDialog("FolderAdd.asp", "", "dialogWidth=400px; dialogHeight=300px; center:Yes; help:No; resizable:No; status:No;")
End Function

'======================================================================================================
'	구성 
'======================================================================================================
Function Form_onclick()
	Dim strRetValue
	strRetValue = window.showModalDialog("FolderConfig.asp", "", "dialogWidth=400px; dialogHeight=300px; center:Yes; help:No; resizable:No; status:No;")
End Function

'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node를 클릭하면 발생 이벤트 
'==========================================================================================
'Sub uniTree1_NodeClick(Node)
'	If Node.Image = C_Root Then Exit Sub
'	
'	Call SetCheck(Node, Not(IsChecked(Node)))
'	Call CheckParent(Node, False)
'	Call CheckChilds(Node)
	
'End Sub


Sub uniTree1_NodeClick(Node)
	Dim Response
	' 트리 조회시에 클릭을 하면 조회가 되지 않도록 조치 

'	frm1.cboSubSystemType.value = ""

	If CheckRunningBizProcess = True Then
	   If lgSaveModFg <> "G" And lgSaveModFg <> "A" Then
		'	Exit Sub 
	   End If
	End If

	If lgBlnNewNode = True Then
		If Node.Key <> gNewNode.Key Then
			Response = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
			If Response = vbYes Then
				frm1.uniTree1.Nodes.Remove gNewNode.Index
				frm1.uniTree1.SetFocus
				lgBlnFlgChgValue = False
				lgBlnNewNode = False
				lgSaveModFg = ""
				Set gNewNode = Nothing
				Call FncNew()
			Else
				frm1.uniTree1.SetFocus
				gNewNode.Selected = True
				Exit Sub
			End If
		Else
			Exit Sub
		End If
	End If

	ggoSpread.Source = frm1.vspdData
	If Node.Key <> gPrevNode And (ggoSpread.SSCheckChange = True Or lgBlnFlgChgValue = True) Then
		Response = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If Response = vbNo Then
			frm1.uniTree1.SetFocus				'3월 22일 추가 
			Exit Sub
		End If
	End If

	gPrevNode = Node.Key
	If Node.Key = C_USER_MENU_KEY Then
		'if Node.children > 0  then  
		'Set Node = Node.child
		'Node.Selected = True 3월 28일 삭제 
		'else
		Exit Sub
		'end if
	End If

	Dim strVal
		
	Call LayerShowHide(1)

	If Node.Image = C_URL Then
		strVal = BIZ_LOOKUP_ACCT_ID & "?txtMode=" & parent.UID_M0001							'☆: 조회 조건 데이타	
		strVal = strVal & "&strCmd=" & "LOOKUPAC"
		strVal = strVal & "&strKey=" & Mid (Node.key,2)
		ClickTab2()
		Call SetToolbar("1100111100001111")														'⊙: 버튼 툴바 제어						 

		frm1.lgstrCmd.value = "ACCT"
	Else
		strVal = BIZ_LOOKUP_ACCT_ID & "?txtMode=" & parent.UID_M0001							'☆: 조회 조건 데이타	
		strVal = strVal & "&strCmd=" & "LOOKUPGP"
		strVal = strVal & "&strKey=" & Mid (Node.key,2)
		ClickTab1()
		Call SetToolbar("1100100000001111")														'⊙: 버튼 툴바 제어					 
		frm1.lgstrCmd.value = "GP"
	End If

	frm1.txtParentGP_CD.value = Mid(Node.parent.key,2)
	frm1.txtParentGP_LVL.value = GetNodeLvl(Node.Parent)
	frm1.txtParentGP_SEQ.value = GetIndex(Node.Parent)

	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 
		
	If Node.Image = C_URL Then
		frm1.lgstrCmd.value = "ACCT"
	Else
		frm1.lgstrCmd.value = "GP"
	End If

	frm1.txtParentGP_CD.value = Mid(Node.parent.key,2)
	frm1.txtParentGP_LVL.value = GetNodeLvl(Node.Parent)
	frm1.txtParentGP_SEQ.value = GetIndex(Node.Parent)

	lgBlnFlgChgValue = False
End Sub

'==========================================================================================
'   Event Name : uniTree1_DblClick
'   Event Desc : Node를 클릭하면 발생 이벤트 
'==========================================================================================
Sub uniTree1_DblClick()
	Dim Node

	With frm1
		Set Node = .uniTree1.SelectedItem

		If Node.Image = C_URL Then
			If ChkUserMenu(Node, C_USER_MENU_KEY) = False Then	' 유저메뉴가 아닐때 
				'Call parent.frToolbar.DBGo(Node.Key)
			Else	' 유저메뉴일때 
				'Call parent.frToolbar.DBGo(Replace(Node.Key, C_USER_MENU_STR, ""))
			End If
		End If
	End With
End Sub

 '******************************  Treeview Operation  *********************************************
'	Function For Treeview Operation
'********************************************************************************************************* 
Sub CheckParent(ByVal Node, ByVal blnEntrprsMnu)
'    If Node.Parent Is Nothing Then Exit Sub
    If Node.Parent.Image = C_Root Then Exit Sub

    If IsChecked(Node) = blnEntrprsMnu Then
	    If IsChecked(Node.Parent) = blnEntrprsMnu Then Exit Sub

		Call SetCheck(Node.Parent, blnEntrprsMnu)
		Call CheckParent(Node.Parent, blnEntrprsMnu)
    End If
End Sub

Sub CheckChilds(ByVal Node)
    Dim ndNode
    Set ndNode = Node.Child

    Do Until ndNode Is Nothing
		Call SetCheck(ndNode, IsChecked(Node))
        Call CheckChilds(ndNode)

        Set ndNode = ndNode.Next
    Loop
End Sub

Function IsChecked(ByVal Node)
	IsChecked = False

	If Node.Image = C_Folder_Ch Or Node.Image = C_URL_Ch Then
		IsChecked = True
	End If
End Function

Sub SetCheck(Node, blnCheck)
	If blnCheck  = True Then
		If Node.Image = C_Folder Then
			Node.Image = C_Folder_Ch
		ElseIf Node.Image = C_URL Then
			Node.Image = C_URL_Ch
		End If
	Else
		If Node.Image = C_Folder_Ch Then
			Node.Image = C_Folder
		ElseIf Node.Image = C_URL_Ch Then
			Node.Image = C_URL
		End If
	End If
End Sub

 '******************************  Treeview Operation  *********************************************
'	Function For Create
'********************************************************************************************************* 
Sub CreateMenu()
	Dim StrKey
	Dim StrText

	frm1.txtMaxRows.value = "0"
	frm1.txtSpread.value = ""

	If frm1.uniTree2.SelectedItem Is Nothing Then	Exit Sub
	If frm1.uniTree2.SelectedItem.Key = False Then	Exit Sub
	If frm1.uniTree2.SelectedItem.image = C_URL Then Exit Sub
	If frm1.uniTree2.SelectedItem.image = C_ROOT Then Exit Sub
	If mid(frm1.uniTree2.SelectedItem.tag,1,1) = "N" Then Exit Sub

	StrKey = frm1.uniTree2.SelectedItem.Key
	StrText = frm1.uniTree2.SelectedItem.Text
	
	Call CreateCoMenu(frm1.uniTree1.Nodes("*"), StrKey, StrText)
End Sub

'========================================================================================================= 
Sub CreateCoMenu(Node, StrKey, StrText)
	Dim ndNode
	Dim errNum

'	On Error Resume Next
	Err.Clear
	
	If Node.Image <> C_Root Then
		Set ndNode = frm1.uniTree2.Nodes(Node.Key)
		errNum = Err.number
		On Error Goto 0

		If IsChecked(Node) = True Then  
			If errNum <> 0 And Node.Image <> C_Folder_Ch Then
				If SetSaveVal(Node, "C", StrKey, StrText) = False Then
					Exit Sub
				End If
			End If
		End If
	End If

	Set ndNode = Node.Child
	
	Do Until (ndNode Is Nothing)
		Call CreateCoMenu(ndNode, StrKey, StrText)
		Set ndNode = ndNode.Next
	Loop

	Call SetCheck(Node, False)
	Set ndNode = Nothing
End Sub

'==========================================================================================
'   Event Name : uniTree1_OLEDragDrop
'   Event Desc : Node를 Drag & Drop 이벤트 
'==========================================================================================
Sub  uniTree1_OLEDragDrop(Data , Effect , Button , Shift , x , y )
	Dim NewNode, IntRetCD
    Dim strVal, strUpKey, Index
    Dim iObjNewNode

	'클릭시 이동할수 없습니다. 메세지 뜨는 버그 수정 
'	On Error Resume Next

    Set iObjNewNode = frm1.uniTree1.HitTest(x, y)
    
    If iObjNewNode Is Nothing Then Exit Sub
	If iObjNewNode.key = gDragNode.key Then Exit Sub

	Set iObjNewNode = Nothing

	If gDragNode Is Nothing Then Exit Sub

	If ChkDragState(x, y) = False Then
        Effect = vbDropEffectNone
		IntRetCD = DisplayMsgBox("990017","X","X","X")	' 해당 위치로는 이동할 수 없습니다!
		frm1.uniTree1.MousePointer = 0
        Exit Sub
	End If

	Call LayerShowHide(1)

	frm1.uniTree1.MousePointer = 11

    Set NewNode = frm1.uniTree1.HitTest(x, y)
    Set gDropNode = NewNode					' 이동해야될 노드를 기억시킴 

	frm1.txtToParentGP_CD.value = Mid(gDropNode.Key, 2)
	frm1.txtToParentGP_LVL.value = GetNodeLvl(gDropNode)
	frm1.txtToParentGP_SEQ.value = GetIndex(gDropNode)

	frm1.txtParentGP_CD.value = Mid(gDragNode.parent.key, 2)
	frm1.txtParentGP_LVL.value = GetNodeLvl(gDragNode.Parent)
	frm1.txtParentGP_SEQ.value = GetIndex(gDragNode.Parent)

	If gDragNode.Image = C_URL Then
		frm1.lgstrCmd.value  = "ACCT"
		frm1.txtToACCT_CD.value = Mid(gDragNode.key, 2)
		frm1.txtToACCT_SEQ.value = GetInsSeq(gDropNode)
		frm1.txtACCT_CD.value = Mid(gDragNode.key, 2)
		frm1.txtACCT_SEQ.value = GetIndex(gDragNode)
	Else
		frm1.lgstrCmd.value = "GP"
		frm1.txtToGP_CD.value = Mid(gDragNode.Key, 2)
		frm1.txtToGP_LVL.value = GetNodeLvl(gDropNode) + 1
		frm1.txtToGP_Seq.value = GetInsSeq(gDropNode)
		frm1.txtGP_CD.value = Mid(gDragNode.Key, 2)
		frm1.txtGP_LVL.value = GetNodeLvl(gDragNode)
		frm1.txtGP_Seq.value = GetInsSeq(gDragNode)
	End If

	lgSaveModFg = "R"
	Call ExecMyBizASP(frm1, BIZ_MOVE_ACCT_ID)										'☜: 비지니스 ASP 를 가동 
End Sub

'========================================================================================================= 
Sub uniTree1_MouseDown(Button, Shift, X, Y)
	If frm1.uniTree1.IsNodeClicked = "Y" Then
		lglsClicked = True
	Else
		lglsClicked = False
	End If
End Sub

'==========================================================================================
'   Event Name : uniTree1_OLEStartDrag
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  uniTree1_OLEStartDrag(Data, AllowedEffects)
	If lglsClicked = True Then
		Set gDragNode = frm1.uniTree1.SelectedItem
		gDragNode.Selected = True
	Else
		Set gDragNode = Nothing
	End If

	lglsClicked = False
End Sub

'==========================================================================================
'   Event Name : uniTree1_MouseUp
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  uniTree1_MouseUp(Node, Button , Shift, X, Y )
	With frm1
		If Button = 2 Or Button = 3 Then
			If Node.Key = C_USER_MENU_KEY Then
				.uniTree1.MenuEnabled C_MNU_OPEN, True
				.uniTree1.MenuEnabled C_MNU_ADD, FALSE
				.uniTree1.MenuEnabled C_MNU_DELETE, False
				.uniTree1.MenuEnabled C_MNU_RENAME, False
				frm1.uniTree1.PopupMenu 
				Exit Sub
			End If

			If ChkUserMenu(Node, C_USER_MENU_KEY) = False Then	' 유저메뉴가 아닌곳에서의 팝업 
				Select Case Node.Image
					Case C_URL, C_Folder, C_Const
						.uniTree1.MenuEnabled C_MNU_OPEN, False
					Case Else
						.uniTree1.MenuEnabled C_MNU_OPEN, False
				End Select

				.uniTree1.MenuEnabled C_MNU_ADD, False
				.uniTree1.MenuEnabled C_MNU_DELETE, False
				.uniTree1.MenuEnabled C_MNU_RENAME, False
			Else
				' 유저메뉴에서의 팝업 
				.uniTree1.MenuEnabled C_MNU_DELETE, True

				Select Case Node.Image
					Case C_URL
						.uniTree1.MenuEnabled C_MNU_OPEN, False
						.uniTree1.MenuEnabled C_MNU_ADD, False
						.uniTree1.MenuEnabled C_MNU_RENAME, False
					Case C_None
						.uniTree1.MenuEnabled C_MNU_RENAME, False
						.uniTree1.MenuEnabled C_MNU_OPEN, False
						.uniTree1.MenuEnabled C_MNU_ADD, False
					Case C_Const
						.uniTree1.MenuEnabled C_MNU_RENAME, False
						.uniTree1.MenuEnabled C_MNU_OPEN, True
						.uniTree1.MenuEnabled C_MNU_ADD, False
					Case C_Folder
						.uniTree1.MenuEnabled C_MNU_OPEN,True
						.uniTree1.MenuEnabled C_MNU_ADD, True
						.uniTree1.MenuEnabled C_MNU_RENAME, False
				End Select

				' 만약 새로운 입력할 노드에서 popup 할 때는 입력메뉴들이 보이면 안된다.
				If lgBlnNewNode = True Then
					If Node.Key = gNewNode.key Then
						.uniTree1.MenuEnabled C_MNU_OPEN,False
						.uniTree1.MenuEnabled C_MNU_ADD, False
						.uniTree1.MenuEnabled C_MNU_RENAME, False
					End If
				End If
			End If

			frm1.uniTree1.PopupMenu
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : uniTree1_MenuAdd - 계정입력 
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  uniTree1_MenuAdd(Node)
	Dim NodX

	'If ChkUserMenu(Node, C_USER_MENU_KEY) = TRUE Then Exit Sub
	Call FncNew

	If Node.Expanded = False Then
		Node.Expanded = True
	End If

	If Node.Key = C_USER_MENU_KEY Then	' 유저메뉴 Root일 경우 
		Set NodX = frm1.uniTree1.Nodes.Add(Node.Key, tvwChild, C_USER_MENU_STR & GetTotalCnt(Node), C_NEW_FOLDER, C_URL, C_URL)
	Else
		Set NodX = frm1.uniTree1.Nodes.Add(Node.Key, tvwChild, Node.Key & C_UNDERBAR & GetTotalCnt(Node), C_NEW_FOLDER, C_URL, C_URL)
	End If

	NodX.Selected = True
	Set gNewNode = NodX
	set gdragnode = NodX

	Call SetToolbar("1100111100001111")									'⊙: 버튼 툴바 제어		
	Call ClickTab2()

	lgIntFlgMode = parent.OPMD_CMODE	' 신규로 등록 
	frm1.lgstrCmd.value  = "ACCT"

	frm1.txtParentGP_CD.value = UCase(Mid(Node.key,2))
	frm1.txtParentGP_LVL.value = GetNodeLvl(node)
	frm1.txtParentGP_SEQ.value = GetIndex(node)
	frm1.txtACCT_SEQ.value = GetIndex(nodX)
	'Call frm1.uniTree1.StartLabelEdit 
	lgBlnFlgChgValue = TRUE
	lgBlnNewNode = TRUE
	lgSaveModFg	= "A"
End Sub

'==========================================================================================
'   Event Name : AddNodes
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  AddNodes( ByVal strCmd )
	Dim strVal

	Call LayerShowHide(1)
	Call AdoQueryTree1()
End sub

'========================================================================================================= 
Sub AdoQueryTree1()
	Dim strSelect
	Dim strFrom
	Dim strWhere

	Dim NodX
	Dim strGpType
	Dim strParGpCd
	Dim strGpCd
	Dim strGpNm
	Dim strGpLvl
	Dim strGpSeq
	Dim strAcctCd
	Dim strAcctNm
	Dim strAcctSeq
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim intColCnt
	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=			 " gp_cd, gp_nm, gp_lvl, gp_seq   "
	strFrom		=			 " a_acct_gp(NOLOCK) "
	strWhere	=			 " gp_lvl = 1  "
	strWhere	= strWhere & " order by gp_lvl, gp_seq , gp_cd "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			strGpCd			= UCase(Trim(arrVal2(1)))
			strGpNm			= Trim(arrVal2(2))
			strGpLvl		= Cstr(Trim(arrVal2(3)))
			strGpSeq		= Cstr(Trim(arrVal2(4)))

			Set NodX = frm1.uniTree1.Nodes.Add (C_USER_MENU_KEY, tvwChild, "G" & strGpCd, "[" & strGpCd & "]" & strGpNm, C_Folder )
			frm1.uniTree1.Nodes("G" & strGpCd).Tag = cstr(strGpLvl) & "|" & cstr(strGpSeq)
		Next
	End If 
'	For intColCnt = 1 To frm1.uniTree1.Nodes.count
'			  frm1.uniTree1.Nodes(intColCnt).Expanded = True
'		Next

	'Level 1이상에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=				" par_gp_cd ,gp_cd, gp_nm,  gp_lvl, gp_seq   "
	strFrom		=				"  a_acct_gp(NOLOCK)  "
	strWhere	=				"  gp_lvl > 1 "
	strWhere	= strWhere	&	" order by  gp_lvl,  gp_seq , gp_cd     "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParGpCd	= UCase(Trim(arrVal2(1)))
			strGpCd		= UCase(Trim(arrVal2(2)))
			strGpNm		= Trim(arrVal2(3))
			strGpLvl	= Trim(arrVal2(4))
			strGpSeq	= Trim(arrVal2(5))

			Set NodX = frm1.uniTree1.Nodes.Add ("G" & strParGpCd , tvwChild, "G" & strGpCd ,  "[" & strGpCd & "]" & strGpNm ,  C_Folder )
			frm1.uniTree1.Nodes("G" & strGpCd ).Tag = cstr( strGpLvl ) & "|" & cstr( strGpSeq )
		Next
	End if

	'계정코드에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=			  " a_acct_gp.par_gp_cd,   a_acct_gp.gp_cd,  a_acct.acct_cd, a_acct.acct_nm, a_acct.acct_seq  "
	strFrom		=			  " a_acct(nolock), a_acct_gp(nolock) "
	strWhere	=			  " a_acct.gp_cd = a_acct_gp.gp_cd "
	strWhere	= strWhere  & " ORDER BY a_acct_gp.gp_cd asc, a_acct.acct_seq asc"

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParGpCd	= UCase(Trim(arrVal2(1)))
			strGpCd		= UCase(Trim(arrVal2(2)))
			strAcctCd	= UCase(Trim(arrVal2(3)))
			strAcctNm	= Trim(arrVal2(4))
			strAcctSeq	= Trim(arrVal2(5))

			Set NodX = frm1.uniTree1.Nodes.Add ("G" & strGpCd , tvwChild, "A" & strAcctCd ,  "[" & strAcctCd & "]" & strAcctNm,  C_URL  )
			frm1.uniTree1.Nodes("A" & strAcctCd ).Tag =  cstr( strAcctSeq )
		Next
	End If

	Call LoadTopGp()
	frm1.uniTree1.Nodes(1).Expanded = True
	frm1.uniTree1.MousePointer = 0
	Call LayerShowHide(0)
End sub

'========================================================================================================= 
Sub LoadTopGp()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strVal

	Dim strGpCd
	Dim strGpNm
	Dim strGpLvl
	Dim strGpSeq

	Dim ArrVal1
	Dim ArrVal2
	Dim ii , jj

	strSelect	=			 " top 1 gp_cd, gp_nm, gp_lvl, gp_seq   "
	strFrom		=			 " a_acct_gp(NOLOCK) "
	strWhere	=			 " gp_lvl = 1  "
	strWhere	= strWhere & " order by gp_lvl, gp_seq , gp_cd "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)
		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			strGpCd			= Trim(arrVal2(1))
			strGpNm			= Trim(arrVal2(2))
			strGpLvl		= Cstr(Trim(arrVal2(3)))
			strGpSeq		= Cstr(Trim(arrVal2(4)))
		Next
	End If

	strVal = BIZ_LOOKUP_ACCT_ID & "?txtMode=" & parent.UID_M0001							'☆: 조회 조건 데이타	
	strVal = strVal & "&strCmd=" & "LOOKUPGP"
	strVal = strVal & "&strKey=" & strGpCd
	ClickTab1()

	Call SetToolbar("1100100000001111")												'⊙: 버튼 툴바 제어					 

	frm1.lgstrCmd.value = "GP"

	frm1.txtParentGP_CD.value = ""
	frm1.txtParentGP_LVL.value = ""
	frm1.txtParentGP_SEQ.value = ""

	Call RunMyBizASP(MyBizASP, strVal)
End Sub

'========================================================================================================= 
Sub InitTreeView()
	lgBlnLoadTreeImage = False

	With frm1
		.uniTree1.SetAddImageCount = 5

		.uniTree1.Indentation = "200"
		.uniTree1.AddImage C_IMG_Root, C_Root, 0
		.uniTree1.AddImage C_IMG_Folder, C_Folder, 0
		.uniTree1.AddImage C_IMG_Folder_Ch, C_Folder_Ch, 0
		.uniTree1.AddImage C_IMG_URL, C_URL, 0
		.uniTree1.AddImage C_IMG_URL_Ch, C_URL_Ch, 0
	End With
	'Call FncQuery()
End Sub

'==========================================================================================
'   Event Name : uniTree1_AfterLabelEdit
'   Event Desc : Add하고 Label을 수정한후 DB등록을 호출할 이벤트 
'==========================================================================================
Sub  uniTree1_AfterLabelEdit(Cancel , NewString )
	Dim Node, strVal

	Set Node = frm1.uniTree1.SelectedItem 

    frm1.uniTree1.MousePointer = 11

    '순서 
	' 0: 신규/수정 상태 
	strVal = strVal & lgIntFlgMode & parent.gColSep		' 신규/수정 상태 

	' 1: Menu ID
	strVal = strVal & Node.key & parent.gColSep			'☆: Drag 한 폴더/파일의 키 

	' 2: Upper Menu ID
	strVal = strVal & Node.parent.key & parent.gColSep		'☆: Drop 한 폴더의 키 

	' 3: Menu Name
	strVal = strVal & NewString & parent.gColSep								'☆: Drag 한 폴더/파일의 이름 

	' 4: Menu Type
    If Node.image = C_Folder Then
		strVal = strVal & "M" & parent.gColSep
	Else
		strVal = strVal & "P" & parent.gColSep
	End If

	' 5: Menu Seq
	strVal = strVal & GetIndex(Node) & parent.gColSep							'☆: Drop 한 폴더/파일의 키 

	' 6: PrevID
	strVal = strVal & parent.gColSep
	strVal = strVal & parent.gColSep

	strVal = strVal & parent.gRowSep

	frm1.txtlgMode.value = parent.UID_M0002
	frm1.txtMulti.value = strVal
	frm1.txtAdd.value = "A"

	'Call ExecMyBizASP(frm1, BIZ_SAVE_ACCT_ID)										'☜: 비지니스 ASP 를 가동 
	'frm1.action = BIZ_SAVE_ACCT_ID
	'frm1.submit 
End Sub

'==========================================================================================
'   Event Name : uniTree1_MenuOpen - 계정그룹입력 
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  uniTree1_MenuOpen(Node)
	Dim NodX

	'If ChkUserMenu(Node, C_USER_MENU_KEY) = True Then Exit Sub

	call FncNew

	If Node.Expanded = False Then
		Node.Expanded = True
	End If

	If Node.Key = C_USER_MENU_KEY Then	' 유저메뉴 Root일 경우 
		Set NodX = frm1.uniTree1.Nodes.Add(Node.Key, tvwChild, C_USER_MENU_STR & GetTotalCnt(Node), C_NEW_FOLDER, C_FOLDER, C_FOLDER)
	Else
		Set NodX = frm1.uniTree1.Nodes.Add(Node.Key, tvwChild, Node.Key & C_UNDERBAR & GetTotalCnt(Node), C_NEW_FOLDER, C_FOLDER, C_FOLDER)
	End If

	NodX.Selected = True
	Set gNewNode = NodX
	set gdragnode = NodX

	Call ClickTab1()
	Call SetToolbar("1100100000001111")									'⊙: 버튼 툴바 제어		

	lgIntFlgMode = parent.OPMD_CMODE	' 신규로 등록 

	frm1.txtParentGP_CD.value = Mid(Node.key,2)
	frm1.txtParentGP_LVL.value = GetNodeLvl(Node)
	frm1.txtParentGP_SEQ.value = GetIndex(Node)
	frm1.txtGP_LVL.value = GetNodeLvl(NodX)
	frm1.txtGP_SEQ.value = GetIndex(NodX)
	frm1.lgstrCmd.value  = "GP"

	lgBlnFlgChgValue = TRUE
	lgBlnNewNode = TRUE
	lgSaveModFg	= "G"
End Sub

'==========================================================================================
'   Event Name : uniTree1_MenuRename
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  uniTree1_MenuRename(Node)
	If ChkUserMenu(Node, C_USER_MENU_KEY) = False Then Exit Sub

	lgIntFlgMode = parent.OPMD_UMODE	' 신규로 등록 

	Call frm1.uniTree1.StartLabelEdit
End Sub

'==========================================================================================
'   Event Name : uniTree1_MenuDelete
'   Event Desc : 삭제메뉴클릭시 
'==========================================================================================
Sub  uniTree1_MenuDelete(Node)
	Dim  OldNode
	Dim IntRetCD

	If Node.Key = C_USER_MENU_KEY Then  Exit Sub

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            'Will you destory previous data"
	
	If IntRetCD = vbNo Then
		frm1.uniTree1.MousePointer = 0
		Exit Sub
	End If

	Call LayerShowHide(1)

	frm1.uniTree1.MousePointer = 11
	Set gdragnode = Node
	TempRootNode = Node.Key

	If lgBlnNewNode = True Then
		If Node.Key = gNewNode.key Then
			Set gdragnode = Node.Next
			frm1.uniTree1.Nodes.Remove gNewNode.Index
			'frm1.uniTree1.SetFocus
			lgBlnFlgChgValue = False
			lgBlnNewNode = False
			lgSaveModFg = ""
			Set gNewNode = Nothing	
			'Set OldNode = frm1.uniTree1.selecteditem
			'call uniTree1_NodeClick(OldNode)      '3월 22일 추가 
			Call LayerShowHide(0)
			frm1.uniTree1.MousePointer = 0
			Exit Sub
		End If
	End If

	If Node.Image = C_URL Then
		frm1.lgstrCmd.value = "ACCT"
		If frm1.txtACCT_CD.value = "" Then
			frm1.txtACCT_CD.value = Mid(Node.key,2)
			frm1.txtACCT_SEQ.value = GetIndex(Node)
		End If
	Else
		frm1.lgstrCmd.value = "GP"
		If frm1.txtGP_CD.value = "" Then
			frm1.txtGP_CD.value = Mid(Node.key,2)
			frm1.txtGP_LVL.value = GetNodeLvl(Node)
			frm1.txtGP_SEQ.value = GetIndex(Node)
		End If
	End If

	frm1.txtParentGP_CD.value = Mid(Node.parent.key,2)
	frm1.txtParentGP_LVL.value = GetNodeLvl(Node.Parent)
	frm1.txtParentGP_SEQ.value = GetIndex(Node.Parent)
	'arrIndx = 0
	'InsCnt = 1
	'StrVal = ""
	'OpMode = "D"

	'Call DelTVParentNodeStr(Node)
	'Call GetDelUpdataNode(Node)
	'lgIntFlgMode = parent.OPMD_CMODE

	frm1.txtlgMode.value = parent.UID_M0003

	lgSaveModFg	= "D"

	Call ExecMyBizASP(frm1, BIZ_SAVE_ACCT_ID)										'☜: 비지니스 ASP 를 가동 
End Sub

'========================================================================================
' Function Name : DelTVParentNodeStr
' Function Desc : 비지니스 로직에 전해줄 스트링 문자열을 만듬 
'========================================================================================
Sub  DelTVParentNodeStr(ByVal nodeParent)
    Dim nodeDummy
    Dim nodeChild

    LastInsNode = DelNodeStr(nodeParent)

    If arrIndx = 0 Then
        ReDim Preserve arrParent(arrIndx)
        arrParent(arrIndx) = LastInsNode
        arrIndx = arrIndx + 1
    End If

    Set nodeChild = nodeParent.Child

    If Not nodeChild Is Nothing Then
        ReDim Preserve arrParent(arrIndx)
        arrParent(arrIndx) = LastInsNode
        arrIndx = arrIndx + 1
    End If

    Do While Not (nodeChild Is Nothing)
        If nodeChild.Children Then
            Call DelTVParentNodeStr(nodeChild)
        Else
            LastInsNode = DelNodeStr(nodeChild)
        End If
        Set nodeChild = nodeChild.Next
    Loop
    
    arrIndx = arrIndx - 1
End Sub

'========================================================================================
' Function Name : DelNodeStr
' Function Desc : 각 노드별 String
'========================================================================================
Function  DelNodeStr(nodeSrc)
    Dim UpperKey

    With nodeSrc
        UpperKey = nodeSrc.Parent.Key
    	DelNodeStr = .Key

	    If getMenuType(.Image) = "P" Then 
	    	If InStr(1, DelNodeStr, "_", 1) > 0 then DelNodeStr = Left(DelNodeStr, InStr(1, DelNodeStr, "_", 1) - 1)
		End If

		StrVal = StrVal & OpMode & parent.gColSep & DelNodeStr & parent.gColSep & UpperKey & parent.gColSep & "" & parent.gColSep & _
                 "" & parent.gColSep & "" & parent.gColSep & "" & parent.gColSep & "" & parent.gColSep & "" & parent.gRowSep
    End With
End Function

'==========================================================================================
'   Function Name : GetDelUpdateNode
'   Function Desc : 현재 삭제 또는 이동되는 Node이후 노드이 Seq를 수정하는 String
'==========================================================================================
Function GetDelUpdataNode(Node)
    Dim i, ChildNode, ParentNode, iNodeCnt, blnFound, NodeKey

    Set ParentNode = Node.Parent
    Set ChildNode = ParentNode.Child

    iNodeCnt = ParentNode.Children
    blnFound = False

    For i = 1 To iNodeCnt
    	If blnFound = False Then
        	If ChildNode.Key = Node.Key Then blnFound = True
		Else
			NodeKey = ChildNode.Key

			If getMenuType(ChildNode.Image) = "P" Then
	    		If InStr(1, NodeKey, "_", 1) > 0 Then NodeKey = Left(NodeKey, InStr(1, NodeKey, "_", 1) - 1)
			End If

			StrVal = StrVal & "U" & parent.gColSep & NodeKey & parent.gColSep & ParentNode.Key & parent.gColSep & ChildNode.Text & parent.gColSep & _
                     getMenuType(ChildNode.Image) & parent.gColSep & GetDelNodeLvl(ChildNode) & parent.gColSep & i - 1 & parent.gColSep & _
                 	 "" & parent.gColSep & "" & parent.gRowSep
		End If
		
        Set ChildNode = ChildNode.Next
    Next
End Function

 '=========================  uniTree1_onAddImgReady()  ====================================
'	Event  Name : uniTree1_onAddImgReady()
'	Description : SetAddImageCount수의 Image가 다운로드 완료되고 TreeView의 ImageList에 
'                 추가되면 발생하는 이벤트 
'========================================================================================= 
Sub uniTree1_onAddImgReady()
	If lgBlnBizLoadMenu = False Then	' 왜 체크를 하는지?
		Call DisplayAcct()
	End If
End Sub

'========================================================================================
Function  FncQuery()
    Dim IntRetCD 

    FncQuery = False
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		'IntRetCD = DisplayMsgBox("900004", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
			frm1.uniTree1.MousePointer = 0
		End If
    End If

    If lgBlnNewNode = TRUE Then
			lgBlnNewNode = FALSE
			Set gNewNode = Nothing
	End if

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    Call InitVariables
    CALL InitSpreadSheet

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery
    
    FncQuery = True
End Function

'========================================================================================
Function  FncNew()
    Dim IntRetCD

    FncNew = False

    Err.Clear
'    On Error Resume Next

    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call SetDefaultVal
    Call InitSpreadSheet

	Call  cboHQ_BRCH_FG_onchange()
'	Call  cboMgntFg_onchange()
	lgBlnFlgChgValue = False

    FncNew = True
End Function


'========================================================================================
Function  FncDelete()
'    On Error Resume Next
End Function

'========================================================================================
Function  FncSave()
	Dim IntRetCD

	FncSave = False

	Err.Clear

	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001","X","X","X")                          'No data changed!!
		Exit Function
	End If

	'-----------------------
	'Check content area
	'-----------------------
	If frm1.lgstrCmd.value = "ACCT" Then
		ggoSpread.Source = frm1.vspdData
		If Not chkField(Document, "2") OR ggoSpread.SSDefaultCheck = False Then                                  '⊙: Check contents area
			Exit Function
		End If 
	ELse
		If Not chkField(Document, "3")  Then                                  '⊙: Check contents area
			Exit Function
		End If
	End If

	if PreCheck = False then
		IntRetCD = DisplayMsgBox("110118","X","X","X") 												' 보조부 항목1 과 보조부 항목2가 동일합니다..
		Exit Function
	end if

	'-----------------------
	'Save function call area
	'-----------------------
	IF DbSave = False Then
		Exit Function
	End IF

	FncSave = True
End Function

'========================================================================================
Function  FncCopy()
	If frm1.vspdData.Maxrows < 1 Then Exit Function
	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	frm1.vspdData.ReDraw = True
End Function

'========================================================================================
Function  FncCancel() 
	if frm1.vspdData.Maxrows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo
End Function

'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow

    On Error Resume Next
    Err.Clear

    FncInsertRow = False

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

		If Err.number = 0 Then
		   FncInsertRow = True
		End If

		.vspddata.col   = C_SYSTEM_FG
		.vspddata.value = "N"
'		Call SetDrCrFg(.vspddata,"N",.vspdData.ActiveRow)
'		Call SetSpdAddColor(.vspddata,.vspdData.ActiveRow,"I","N")
		Call ChkCount(.vspddata,"N")
    End With

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function  FncDeleteRow()
	Dim lDelRows
	
	With Frm1
		If .vspdData.Maxrows < 1 Then Exit Function
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspddata.col =  C_CHG_DEL
		.vspddata.value = "N"
		lDelRows = ggoSpread.DeleteRow
	End With		
End Function

'========================================================================================
Function  FncPrint()
    parent.FncPrint()
End Function

'========================================================================================
Function  FncPrev()
'    On Error Resume Next
End Function

'========================================================================================
Function  FncNext()
'    On Error Resume Next
End Function

'========================================================================================
Function  FncExcel()
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'=======================================================================================================
Function  FncFind()
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
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
Function  FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")
    		If IntRetCD = vbNo Then
      			Exit Function
    		End If
    End If

    FncExit = True
End Function

'========================================================================================
Function  DbQuery()
	DbQuery = False

	Err.Clear

	Call DisplayAcct()

	frm1.uniTree1.SetFocus
	frm1.uniTree1.selecteditem.Selected = True
	Call uniTree1_NodeClick(frm1.uniTree1.selecteditem)

	DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()
	Dim ii
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

    Call SetSpreadLock("Q", 0, 1, "")

    With frm1
		For ii = 1 To .vspddata.maxrows
			.vspddata.row = ii
			.vspddata.col = C_SYSTEM_FG
			If Trim(.vspddata.value) = "Y" Then
				Call SetSpdAddColor(frm1.vspddata,ii,"Q","N")
			End If	
		Next			    
	End With
	
    Call ggoOper.LockField(Document, "Q")

	If gSelframeFlg = TAB2 Then
		Call cboHQ_BRCH_FG_onchange()
		Call Mgnt_QueryOk()
'		Call cboMgntFg_onchange()
		Call subledger_change()
		Call accttype_change()
		Call mgnt_change()
		lgBlnFlgChgValue = False
	End If
End Function

'========================================================================================
Function  DbSave()
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
	Dim strVal,strDel

	Call LayerShowHide(1)

    DbSave = False

    On Error Resume Next

	lgRetFlag = False
	With frm1
		.txtlgMode.value = parent.UID_M0002											'☜: 저장 상태 
		.txtlgMode.value = lgIntFlgMode								         	'☜: 신규입력/수정 상태 

		strMode = .txtlgMode.value
    	'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""	
		'-----------------------
		'Data manipulate area
		'-----------------------
		If frm1.lgstrCmd.value  = "ACCT" Then
'			strCmd = "CREATEAC"

			'계정에 대한 값을 넘겨준다.			

			For lRow = 1 To .vspdData.MaxRows
			    .vspdData.Row = lRow
			    .vspdData.Col = 0

			    Select Case .vspdData.Text
			        Case ggoSpread.InsertFlag													'☜: 신규								
						strVal = strVal & "C" & parent.gColSep  								'☜: C=Create, Row위치 정보 
			            .vspdData.Col = C_CTRLITEM
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CTRLITEMSEQ
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_DRFG
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CRFG
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_DEFAULT_VALUE
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep			            
			            .vspdData.Col = C_GL_ITEM
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep			            
			            .vspdData.Col = C_SYSTEM_FG
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CHG_DEL
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep			            
			            strVal = strVal & lRow & parent.gRowSep

			            lGrpCnt = lGrpCnt + 1
			        Case ggoSpread.UpdateFlag													'☜: 수정 

						strVal = strVal & "U" & parent.gColSep  								'☜: C=Create, Row위치 정보 
			            .vspdData.Col = C_CTRLITEM
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CTRLITEMSEQ
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_DRFG
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CRFG
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_DEFAULT_VALUE
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_GL_ITEM
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_SYSTEM_FG
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CHG_DEL
			            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			            strVal = strVal & lRow & parent.gRowSep

				        lGrpCnt = lGrpCnt + 1
			        Case ggoSpread.DeleteFlag													'☜: 삭제 
						
						strDel = strDel & "D" & parent.gColSep  								'☜: C=Create, Row위치 정보 
			            .vspdData.Col = C_CTRLITEM
			            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CTRLITEMSEQ
			            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_DRFG
			            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CRFG
			            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_DEFAULT_VALUE
			            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_GL_ITEM
			            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_SYSTEM_FG
			            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            .vspdData.Col = C_CHG_DEL
			            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
			            strDel = strDel & lRow & parent.gRowSep

			            lGrpCnt = lGrpCnt + 1
			            lgRetFlag = True
			        Case Else

			    End Select
			Next
		Else

		End If

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_SAVE_ACCT_ID)
	End With

    DbSave = True
    lgRetFlag = True
End Function

'========================================================================================
Function DbSaveOk()
	lgBlnFlgChgValue = False
	 ggoSpread.ssdeleteflag 1
	'Call InitVariables
	Dim addTempNode, ChildNode, strMoveRootKey, I

	If lgSaveModFg	= "R" Then
		strMoveRootKey = gDragNode.Key

		Set addTempNode = frm1.uniTree1.Nodes.Add(gDropNode.Key, tvwChild, "uniTEMPKEY", gDragNode.Text, gDragNode.Image, gDragNode.SelectedImage)
		If addTempNode.Image <> C_URL Then

		    Do
		        If gDragNode.Children = 0 Then Exit Do
			    Set ChildNode = gDragNode.Child.LastSibling
		    	Set ChildNode.Parent = addTempNode
			Loop
			addTempNode.Expanded = False 
		End If
		frm1.uniTree1.Nodes.Remove gDragNode.Index
		addTempNode.Key = strMoveRootKey
	End If

	If lgSaveModFg	= "G" Then

		gDragNode.Key = "G" & UCase(Trim(frm1.txtGp_Cd.value))
		gDragNode.text = "[" & UCase(Trim(frm1.txtGp_Cd.value)) & "]" & frm1.txtGP_SH_NM.value
	End If

	If lgSaveModFg	= "A" Then
		gDragNode.Key = "A" & UCase(Trim(frm1.txtACCT_Cd.value))
		gDragNode.text =  "[" & UCase(Trim(frm1.txtACCT_Cd.value)) & "]" & frm1.txtACCT_Sh_Nm.value
	End If

	If lgSaveModFg	= "D"  Then
		frm1.unitree1.nodes.remove gDragNode.Key
		Call FncNew()
	End If

	Set gDragNOde = Nothing

	If lgBlnNewNode = True Then
		lgBlnNewNode = False
		Set gNewNode = Nothing
	End If

	frm1.uniTree1.Setfocus
	lgSaveModFg = ""

' 저장후 변경된 명을 가져온다.
'	Dim NodX
'	Set NodX = frm1.uniTree1.selecteditem
'	If NodX.Image = C_URL Then
'		NodX.Text = "[" & Trim(frm1.txtACCT_Cd.value) & "]" & frm1.txtACCT_Sh_Nm.value
'	Else
'		NodX.Text = "[" & Trim(frm1.txtGp_Cd.value) & "]" & frm1.txtGP_SH_NM.value
'	End If
'	set NodX = Nothing
	Call uniTree1_NodeClick(frm1.uniTree1.selecteditem)
	Call LayerShowHide(0)
	lgBlnNewNode = False

	frm1.uniTree1.MousePointer = 0
End Function

'========================================================================================
Function  DbDelete()
'    On Error Resume Next
End Function

'========================================================================================
Function DbDeleteOk()
'    On Error Resume Next
End Function

Function PreCheck()
	Precheck = False
	
	If Trim(frm1.txtSUBLEDGER1.value) <> "" Then
		If UCase(Trim(frm1.txtSUBLEDGER1.value)) = UCase(Trim(frm1.txtSUBLEDGER2.value))  Then
			Exit Function
		End If
	End If
	
	Precheck = True
End Function

Sub uniTree1_NodeClick2(Node)
		Dim NodX

		frm1.uniTree1.Nodes.Clear 

		Set NodX = frm1.uniTree1.Nodes.Add(, tvwChild, C_USER_MENU_KEY, lgUSER_MENU, C_Root, C_Root)
	

		'frm1.uniTree1.MousePointer = 11
		Call AdoQueryTree2(Node)

		DIm StrNm,strVal
		
		Call CommonQueryRs("acct_nm","a_acct","(acct_CD = " & FilterVar(Node, "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
		StrNm = replace(lgF0,chr(11),"")
		
		
		Call LayerShowHide(1)
		
		strVal = BIZ_LOOKUP_ACCT_ID & "?txtMode=" & parent.UID_M0001							'☆: 조회 조건 데이타	
		strVal = strVal & "&strCmd=" & "LOOKUPAC"
		strVal = strVal & "&strKey=" & Node
		ClickTab2()
		Call SetToolbar("1100111100001111")														'⊙: 버튼 툴바 제어						 

		frm1.lgstrCmd.value = "ACCT"
		
		Call RunMyBizASP(MyBizASP, strVal)
		
End sub 


Sub AdoQueryTree2(Node)
	Dim strSelect
	Dim strFrom
	Dim strWhere

	Dim NodX
	Dim strGpType
	Dim strParGpCd
	Dim strGpCd
	Dim strGpNm
	Dim strGpLvl
	Dim strGpSeq
	Dim strAcctCd
	Dim strAcctNm
	Dim strAcctSeq,StrL,StrN,StrM
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim intColCnt
	'Level 1에 대한 Node가져오기 

	'----------------------------------------------------------------------------------------
	strSelect	=			 " gp_cd, gp_nm, gp_lvl, gp_seq   "
	strFrom		=			 " a_acct_gp(NOLOCK) "
	strWhere	=			 " gp_lvl = 1  "
	strWhere	= strWhere & " order by gp_lvl, gp_seq , gp_cd "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)
		For ii = 0 To jj - 1
		
			arrVal2			= Split(arrVal1(ii), chr(11))
			strGpCd			= UCase(Trim(arrVal2(1)))
			strGpNm			= Trim(arrVal2(2))
			strGpLvl		= Cstr(Trim(arrVal2(3)))
			strGpSeq		= Cstr(Trim(arrVal2(4)))
			Set NodX = frm1.uniTree1.Nodes.Add (C_USER_MENU_KEY, tvwChild, "G" & strGpCd, "[" & strGpCd & "]" & strGpNm, C_Folder )
			frm1.uniTree1.Nodes("G" & strGpCd).Tag = cstr(strGpLvl) & "|" & cstr(strGpSeq)
			if strGpCd = left(node,1) then
			NodX.Expanded = True
			End if
		Next
		
	End If 
'	For intColCnt = 1 To frm1.uniTree1.Nodes.count
'			  frm1.uniTree1.Nodes(intColCnt).Expanded = True
'		Next
	
	Call CommonQueryRs("gp_cd","a_acct","(acct_cd = " & FilterVar(node, "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	StrL = replace(lgF0,chr(11),"")

	Call CommonQueryRs("par_gp_cd","a_acct_gp","(gp_cd = " & FilterVar(StrL, "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	StrN = replace(lgF0,chr(11),"")

	Call CommonQueryRs("par_gp_cd","a_acct_gp","(gp_cd = " & FilterVar(StrN, "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	StrM = replace(lgF0,chr(11),"")

	
	'Level 1이상에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=				" par_gp_cd ,gp_cd, gp_nm,  gp_lvl, gp_seq   "
	strFrom		=				"  a_acct_gp(NOLOCK)  "
	strWhere	=				"  gp_lvl > 1 "
	strWhere	= strWhere	&	" order by  gp_lvl,  gp_seq , gp_cd     "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParGpCd	= UCase(Trim(arrVal2(1)))
			strGpCd		= UCase(Trim(arrVal2(2)))
			strGpNm		= Trim(arrVal2(3))
			strGpLvl	= Trim(arrVal2(4))
			strGpSeq	= Trim(arrVal2(5))

			Set NodX = frm1.uniTree1.Nodes.Add ("G" & strParGpCd , tvwChild, "G" & strGpCd ,  "[" & strGpCd & "]" & strGpNm ,  C_Folder )
			frm1.uniTree1.Nodes("G" & strGpCd ).Tag = cstr( strGpLvl ) & "|" & cstr( strGpSeq )
			
			if strGpCd = StrM  then
			NodX.Expanded = True
			End if
			
			if strGpCd = StrN  then
			NodX.Expanded = True
			End if
			
			if strGpCd = StrL  then
			'msgbox "1"
			NodX.Expanded = True
			End if
			
			
			 
		Next

	End if

	'계정코드에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=			  " a_acct_gp.par_gp_cd,   a_acct_gp.gp_cd,  a_acct.acct_cd, a_acct.acct_nm, a_acct.acct_seq  "
	strFrom		=			  " a_acct(nolock), a_acct_gp(nolock) "
	strWhere	=			  " a_acct.gp_cd = a_acct_gp.gp_cd "
	strWhere	= strWhere  & " ORDER BY a_acct_gp.gp_cd asc, a_acct.acct_seq asc"

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParGpCd	= UCase(Trim(arrVal2(1)))
			strGpCd		= UCase(Trim(arrVal2(2)))
			strAcctCd	= UCase(Trim(arrVal2(3)))
			strAcctNm	= Trim(arrVal2(4))
			strAcctSeq	= Trim(arrVal2(5))

			Set NodX = frm1.uniTree1.Nodes.Add ("G" & strGpCd , tvwChild, "A" & strAcctCd ,  "[" & strAcctCd & "]" & strAcctNm,  C_URL  )

			frm1.uniTree1.Nodes("A" & strAcctCd ).Tag =  cstr( strAcctSeq )

			if strAcctCd = node then
				NodX.Expanded = True
			End if
			
		Next
	End If

	Call LoadTopGp2(Node)
	frm1.uniTree1.Nodes(1).Expanded = True
	
	frm1.uniTree1.MousePointer = 0
	Call LayerShowHide(0)
End sub

Sub LoadTopGp2(Node)
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strVal

	Dim strGpCd
	Dim strGpNm
	Dim strGpLvl
	Dim strGpSeq

	Dim ArrVal1
	Dim ArrVal2
	Dim ii , jj

	strSelect	=			 " top 1 gp_cd, gp_nm, gp_lvl, gp_seq   "
	strFrom		=			 " a_acct_gp(NOLOCK) "
	strWhere	=			 " gp_lvl = 1  "
	strWhere	= strWhere & " order by gp_lvl, gp_seq , gp_cd "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)
		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			strGpCd			= Trim(arrVal2(1))
			strGpNm			= Trim(arrVal2(2))
			strGpLvl		= Cstr(Trim(arrVal2(3)))
			strGpSeq		= Cstr(Trim(arrVal2(4)))
		Next
	End If

	strVal = BIZ_LOOKUP_ACCT_ID & "?txtMode=" & parent.UID_M0001							'☆: 조회 조건 데이타	
	strVal = strVal & "&strCmd=" & "LOOKUPGP"
	strVal = strVal & "&strKey=" & strGpCd
	ClickTab1()

	Call SetToolbar("1100100000001111")												'⊙: 버튼 툴바 제어					 

	frm1.lgstrCmd.value = "GP"

	frm1.txtParentGP_CD.value = ""
	frm1.txtParentGP_LVL.value = ""
	frm1.txtParentGP_SEQ.value = ""

	Call RunMyBizASP(MyBizASP, strVal)
End Sub


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR WIDTH=100%>
					<!-- TreeView AREA  -->
					<TD HEIGHT=* WIDTH=30%>
						<script language =javascript src='./js/a2101ma1_uniTree1_N612524169.js'></script>
					</TD>

					<!-- DATA AREA  -->
					<TD HEIGHT=* WIDTH=70%>
						<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
							<TR HEIGHT=23>
								<TD WIDTH="100%">
									<TABLE <%=LR_SPACE_TYPE_10%>>
										<TR>
											<TD WIDTH=10>&nbsp;</TD>
											<TD CLASS="CLSMTABP">
												<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>	
													<TR>
														<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>계정그룹</font></td>
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
													</TR>
												</TABLE>
											</TD>
											<TD CLASS="CLSMTABP">
												<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23" ></td>
														<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>계정코드</font></td>
														<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23" ></td>
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
											<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
										</TR>
										<TR>
											<TD WIDTH=100% HEIGHT=* valign=top>
												<!-- 첫번째 탭 내용  -->
												<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=no>
													<TABLE <%=LR_SPACE_TYPE_60%>>
														<TR>
															<TD CLASS=TD5 HEIGHT=5 WIDTH="100%"></TD>
															<TD CLASS=TD6 HEIGHT=5 WIDTH="100%"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>계정그룹코드</TD>
															<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGP_CD" ALT="계정그룹코드" MAXLENGTH="20" tag  ="33XXXU"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>계정그룹명(단문)</TD>
															<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGP_SH_NM" ALT="계정그룹명(단문)" MAXLENGTH="30" SIZE=30 tag  ="32"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>계정그룹명(장문)</TD>
															<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGP_FULL_NM" ALT="계정그룹명(장문)" MAXLENGTH="50" SIZE=50 tag  ="31"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>계정그룹명(영문)</TD>
															<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGP_ENG_NM" ALT="계정그룹명(영문)" MAXLENGTH="50" SIZE=50 tag  ="31" style="ime-mode:disabled"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>레벨</TD>
															<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGP_LVL" ALT="레벨" MAXLENGTH="3" SIZE=3 STYLE="TEXT-ALIGN: center" tag  ="34"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>순서</TD>
															<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGP_SEQ" ALT="순서" MAXLENGTH="3" SIZE=3 STYLE="TEXT-ALIGN: center" tag  ="34"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>예산관리구분</TD>
															<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGP_BDG_CTRL_FG" ALT="예산관리구분" STYLE="WIDTH: 150px" tag="32"></SELECT></TD>
														</TR>
													</TABLE>
												</DIV> 
												<!-- 두번째 탭 내용  -->
												<DIV ID="TabDiv" SCROLL=no>
													<TABLE <%=LR_SPACE_TYPE_60%>>
														<TR>
															<TD CLASS=TD5 NOWRAP>계정코드</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtACCT_CD" ALT="계정코드" MAXLENGTH="20" tag ="23XXXU"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>계정코드명(단문)</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtACCT_SH_NM" ALT="계정코드명(단문)" MAXLENGTH="30" SIZE=30 tag  ="22"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>계정코드명(장문)</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtACCT_FULL_NM" ALT="계정코드명(장문)" MAXLENGTH="50" SIZE=50 tag  ="21"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>순서</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><script language =javascript src='./js/a2101ma1_OBJECT1_txtACCT_SEQ.js'></script></TD>
															<!--	<INPUT NAME="txtACCT_SEQ" ALT="순서" MAXLENGTH="3" SIZE=3 tag  ="22"></TD>  -->
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>예산관리구분</TD>
															<TD CLASS=TD6 NOWRAP><SELECT NAME="cboBDG_CTRL_FG" ALT="예산관리구분" STYLE="WIDTH: 150px" tag="21"><OPTION VALUE="" selected></OPTION></SELECT>
															<TD CLASS=TD5 NOWRAP>예산통제그룹</TD>
															<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBDG_CTRL_GP_LVL" ALT="예산통제그룹" MAXLENGTH="3" SIZE=3 STYLE="TEXT-ALIGN: center" tag  ="24"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>환평가구분</TD>
															<TD CLASS=TD6 NOWRAP><SELECT NAME="cboFX_EVAL_FG" ALT="환평가구분" STYLE="WIDTH: 150px" tag="22"></TD>
															<TD CLASS=TD5 NOWRAP>차대구분</TD>
															<TD CLASS=TD6 NOWRAP><SELECT NAME="cboBAL_FG" ALT="차대구분" STYLE="WIDTH: 150px" tag="22"></SELECT></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>임시계정구분</TD>
															<TD CLASS=TD6 NOWRAP><SELECT NAME="cboTEMP_ACCT_FG" ALT="임시계정구분" STYLE="WIDTH: 150px" tag="21"><OPTION VALUE="" selected></OPTION></SELECT></TD>
															<TD CLASS=TD5 NOWRAP>계정폐쇄</TD>
															<TD CLASS=TD6 NOWRAP><SELECT NAME="cboDEL_FG" ALT="계정폐쇄" STYLE="WIDTH: 150px" tag="22"></SELECT></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>재무제표구분</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtBS_PL_FG" ALT="재무제표구분" maxlength=2 SIZE=10  tag ="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtBS_PL_FG.value, 4)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT NAME="txtBS_PL_FG_Nm" ALT="재무제표명" MAXLENGTH="40" size=40 STYLE="TEXT-ALIGN: LEFT" tag="24"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>계정특성</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtACCT_TYPE" ALT="계정특성" maxlength=2 SIZE=10  tag ="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtACCT_TYPE.value, 5)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT NAME="txtACCT_TYPE_Nm" ALT="계정특성명" MAXLENGTH="40" size=40  STYLE="TEXT-ALIGN: LEFT" tag="24"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>원가&경영손익구분</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtGP_TYPE" ALT="원가&경영손익구분"maxlength=2 SIZE=10  tag ="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtGP_TYPE.value, 6)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT NAME="txtGP_TYPE_Nm" ALT="계정그룹특성명" MAXLENGTH="40" size=40  STYLE="TEXT-ALIGN: LEFT" tag="24"></TD>
														</TR>	
														<TR>
															<TD CLASS=TD5 NOWRAP>Subsystem 유형</TD>
															<TD CLASS=TD6 NOWRAP><SELECT NAME="cboSubSystemType" ALT="Subsystem 유형" STYLE="WIDTH: 150px" tag="21"><OPTION VALUE="" selected></OPTION></SELECT></TD>
															<TD CLASS=TD5 NOWRAP>본지점구분</TD>
															<TD CLASS=TD6 NOWRAP><SELECT NAME="cboHQ_BRCH_FG" ALT="본지점구분" STYLE="WIDTH: 150px" tag="21"><OPTION VALUE="" selected></OPTION></SELECT></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>본지점관계사업장</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtREL_BIZ_AREA_CD" ALT="본지점관계사업장" MAXLENGTH="10" SIZE=10 tag ="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtREL_BIZ_AREA_CD.value, 2)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT NAME="txtREL_BIZ_AREA_NM" ALT="관계사업장명" MAXLENGTH="30" STYLE="TEXT-ALIGN: LEFT" tag  ="24"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>보조부항목1</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtSUBLEDGER1" ALT="보조부항목1" MAXLENGTH="3" SIZE=10 tag ="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtSUBLEDGER1.value, 0)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT NAME="txtSUBLEDGER1_Nm" ALT="보조부항목명" MAXLENGTH="30" STYLE="TEXT-ALIGN: LEFT" tag  ="24"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP>보조부항목2</TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtSUBLEDGER2" ALT="보조부항목2" MAXLENGTH="3" SIZE=10 tag ="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtSUBLEDGER2.value, 1)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT NAME="txtSUBLEDGER2_nm" ALT="보조부항목명" MAXLENGTH="30" STYLE="TEXT-ALIGN: LEFT" tag  ="24"></TD>
														</TR>
														<TR>															
															<TD CLASS=TD5 NOWRAP><span id="spnMgntType">미결관리구분</span></TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><SELECT NAME="cboMgntType" ALT="미결관리구분" STYLE="WIDTH: 150px" tag="21"></TD>															
														</TR>														
														<TR>
															<TD CLASS=TD5 NOWRAP><span id="spnMgntCd1">미결관리항목1</span></TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtMgntCd1" ALT="미결관리항목1" MAXLENGTH="3" SIZE=10 tag ="21NXXU"><IMG align=top name=btnCalType3 onclick="vbscript:CALL OpenPopUp(frm1.txtMgntCd1.value, 7)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT NAME="txtMgntCd1_Nm" ALT="미결관리코드1" MAXLENGTH="30" STYLE="TEXT-ALIGN: LEFT" tag  ="24"></TD>
														</TR>
														<TR>
															<TD CLASS=TD5 NOWRAP><span id="spnMgntCd2">미결관리항목2</span></TD>
															<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtMgntCd2" ALT="미결관리항목2" MAXLENGTH="3" SIZE=10 tag ="21NXXU"><IMG align=top name=btnCalType4 onclick="vbscript:CALL OpenPopUp(frm1.txtMgntCd2.value, 8)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT NAME="txtMgntCd2_Nm" ALT="미결관리코드2" MAXLENGTH="30" STYLE="TEXT-ALIGN: LEFT" tag  ="24"></TD>
														</TR>
														<TR>
															<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
																<script language =javascript src='./js/a2101ma1_OBJECT1_vspdData.js'></script>
															</TD>
														</TR>
													</TABLE>
												</DIV>
											</TD>
										</TR>
									</TABLE>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  src="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="lgstrCmd" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtParentGP_CD" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtParentGP_LVL" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtParentGP_SEQ" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToParentGP_CD" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToParentGP_LVL" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToParentGP_SEQ" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToGP_CD" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToGP_LVL" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToGP_SEQ" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToACCT_CD" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToACCT_SEQ" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtOpenAcctFg" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtsubledger_modigy_fg" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtmgnt_modigy_fg" tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtaccttype_modigy_fg" tag="21" tabindex="-1">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>
