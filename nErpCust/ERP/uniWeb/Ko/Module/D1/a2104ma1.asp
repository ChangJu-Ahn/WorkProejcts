<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Account Configuration
'*  3. Program ID           : A2104MA1
'*  4. Program Name         : 유형별 계정분류 등록 
'*  5. Program Desc         :
'*  6. Component LIst       : +B19029LookupNumericFormat
'*  7. ModIfied date(First) : 1999/09/10
'*  8. ModIfied date(Last)  : 1999/09/10
'*  9. ModIfier (First)     : Mr  Kim
'* 10. ModIfier (Last)      : Mrs Kim / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : thIs mark(☜) means that "Do not change"
'*                            thIs mark(⊙) Means that "may  change"
'*                            thIs mark(☆) Means that "must change"
'* 13. HIstory              :
'*                            -1999/09/12 : ..........
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<Script Language="VBScript"		SRC="../../inc/incUni2KTV.vbs">          </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit	

'==========================================================================================
'	1. Constant는 반드시 대문자 표기.
'==========================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_SAVE_CLASS_ID			= "a2104mb2.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_LOAD_GRID_CLASS_ID	= "a2104mb6.asp"

'==========================================================================================
Const C_IMG_Root        = "../../../CShared/image/unierp.gIf"
Const C_IMG_Folder      = "../../../CShared/image/folder.gIf"
Const C_IMG_Folder_Ch   = "../../../CShared/image/folder_ch.gIf"
Const C_IMG_URL         = "../../../CShared/image/Account.gIf"
Const C_IMG_URL_Ch      = "../../../CShared/image/Account_Ch.gIf"

Const C_CMD_TOP_LEVEL   = "LIsTTOP"
Const C_CMD_LIsT_LEVEL  = "LIsT"
Const C_CMD_LIsT_DIsT   = "ACCTDIsT"
Const C_CMD_ACCT_LEVEL  = "LIsTACCT"
Const C_CMD_GP_LEVEL    = "LIsTGP"

Const C_USER_MENU       = "UNIERP"
Const C_USER_MENU_KEY   = "*"
Const C_USER_MENU_STR   = "UM_"
Const C_UNDERBAR        = "_"

Const C_NEW_FOLDER      = "새 폴더"

Dim C_ClassCd
Dim C_ClassNm
Dim C_ClassLvl
Dim C_ClassSeq
Dim C_LeftRightFg
Dim C_IndexFg
Dim C_IndexFgNm
Dim C_BalFg
Dim C_ClassFg
Dim C_ClassFgNm
Dim C_PropertyFg
Dim C_PropertyFgNm
Dim C_OccurType
Dim C_OccurTypeNm
Dim C_PmFg
Dim C_AcctCd
Dim C_AcctNm
Dim C_ParClassCd
Dim C_ParClassCd2

Const  C_Root      = "Root"
Const  C_Folder_Ch = "folder_ch.gIf"
Const  C_URL_Ch    = "URL_Ch"

'==========================================================================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'==========================================================================================

Dim  lgStrPrevKey1
Dim  lgStrPrevKey2
Dim  lgQueryFlag
Dim	 lgUSER_MENU

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim  gDragNode , gDropNode, gNewNode

Dim  IsOpenPop
Dim  lgBlnLoadTreeImage
Dim  lgBlnNewNode
'==========================================================================================
Dim  lgSpreadNo   '노드 클릭시 현재 로우값 
Dim  TempRootNode
Dim  lgRetFlag
Dim  lgDelCnt
Dim  orderseq


Dim strFirst

'==========================================================================================
Sub initSpreadPosVariables()
     C_ClassCd         = 1
     C_ClassNm         = 2
     C_ClassLvl        = 3
     C_ClassSeq        = 4
     C_LeftRightFg     = 5
     C_IndexFg         = 6
     C_IndexFgNm       = 7
     C_BalFg           = 8
     C_ClassFg         = 9
     C_ClassFgNm       = 10
     C_PropertyFg      = 11
     C_PropertyFgNm    = 12
     C_OccurType       = 13
     C_OccurTypeNm     = 14
     C_PmFg            = 15
     C_AcctCd          = 16
     C_AcctNm          = 17
     C_ParClassCd      = 18
     C_ParClassCd2     = 19
End Sub

'==========================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgQueryFlag = "1"
	Call CommonQueryRs("Co_Cd", "B_Company","",lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	lgUSER_MENU = Left(lgF0, Len(lgF0)-1)
	strFirst = 1
End Sub

'==========================================================================================
Sub  SetDefaultVal()
	Dim NodX
	frm1.uniTree1.Nodes.Clear 
	Set NodX = frm1.uniTree1.Nodes.Add(, tvwChild, C_USER_MENU_KEY, lgUSER_MENU, C_Root, C_Root)
	strFirst = 1
End Sub

'==========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================================================================
Sub  InitSpreadSheet()
    Dim  sLIst
    Dim  sLIst1
    Dim  sLIst2

    sLIst = "DR" & vbTab  & "CR" & vbTab  & "OD" & vbTab  & "OC"
    sLIst1 = "+" & vbTab  & "-"
    sLIst2 = "R" & vbTab  & "L"

	Call initSpreadPosVariables()

    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread

    With frm1.vspdData
		.MaxCols = C_ParClassCd2 + 1   '' 마지막 상수명 사용 
		.MaxRows = 0
		.ReDraw = False

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit     C_ClassCd,      "계정분류코드",  20, ,   ,   20, 2	'20
		ggoSpread.SSSetEdit     C_ClassNm,      "계정분류명",    30, ,   ,   50	'30
		ggoSpread.SSSetEdit     C_ClassLvl,     "레벨",         6,  2    
		Call AppEndNumberPlace("6","3","0")
		ggoSpread.SSSetFloat    C_ClassSeq,     "순서" ,        6,  "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,  2,  ,   ,   "1", "9999"
		ggoSpread.SSSetCombo    C_LeftRightFg,  "좌우구분",      12, True
		ggoSpread.SSSetCombo    C_IndexFg,      "INDEX TYPE",   12, True
		ggoSpread.SSSetCombo    C_IndexFgNm,    "INDEX TYPE",   20, True	'20
		ggoSpread.SSSetCombo    C_BalFg,        "차대구분",      12, True
		ggoSpread.SSSetCombo    C_ClassFg,      "손익분류",      12, True
		ggoSpread.SSSetCombo    C_ClassFgNm,    "손익분류",      20, True		'20
		ggoSpread.SSSetCombo    C_PropertyFg,   "계정특성",      12, True
		ggoSpread.SSSetCombo    C_PropertyFgNm, "계정특성",      20, True		'20
		ggoSpread.SSSetCombo    C_OccurType,    "기초기말구분",  12
		ggoSpread.SSSetCombo    C_OccurTypeNm,  "기초기말구분",  12
		ggoSpread.SSSetCombo    C_PmFg,         "계정가감구분",  12, True
		ggoSpread.SSSetEdit     C_AcctCd,       "계정코드",      20
		ggoSpread.SSSetEdit     C_AcctNm,       "계정코드명",    30
		ggoSpread.SSSetEdit     C_ParClassCd,   "", 30,2  '상위코드 
		ggoSpread.SSSetEdit     C_ParClassCd2,   "", 30,2
		'call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctNm)

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_IndexFg,C_IndexFg,True)
		Call ggoSpread.SSSetColHidden(C_ClassFg,C_ClassFg,True)
		Call ggoSpread.SSSetColHidden(C_PropertyFg,C_PropertyFg,True)
		Call ggoSpread.SSSetColHidden(C_OccurType,C_OccurType,True)
		Call ggoSpread.SSSetColHidden(C_OccurTypeNm,C_OccurTypeNm,True)
		Call ggoSpread.SSSetColHidden(C_PmFg,C_PmFg,True)
		Call ggoSpread.SSSetColHidden(C_ParClassCd,C_ParClassCd,True)
		Call ggoSpread.SSSetColHidden(C_ParClassCd2,C_ParClassCd2,True)
		Call ggoSpread.SSSetColHidden(C_ParClassCd2 + 1,C_ParClassCd2 + 1,True)

		.ReDraw = True

		Call SetSpreadLock("I", 0, 1, "")
		' ggoSpread.SetCombo sLIst, C_IndexFg
		ggoSpread.SetCombo sLIst , C_BalFg
		ggoSpread.SetCombo sLIst1, C_PmFg
		ggoSpread.SetCombo sLIst2, C_LeftRightFg
    End With
End Sub

'==========================================================================================
Sub  SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2 )
    With frm1
        ggoSpread.Source = .vspdData
        If lRow2 = "" Then lRow2 = .vspdData.MaxRows
            .vspdData.Redraw = False
            ggoSpread.SpreadLock C_ClassCd      ,  -1 , C_ClassCd
            ggoSpread.SpreadLock C_ClassLvl     ,  -1 , C_ClassLvl
            ggoSpread.SpreadLock C_AcctCd       ,  -1 , C_AcctCd
            ggoSpread.SpreadLock C_AcctNm       ,  -1 , C_AcctNm
            ggoSpread.SpreadLock C_OccurTypeNm  ,  -1 , C_OccurTypeNm 
			ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
            .vspdData.Redraw = True
    End With
End Sub


'==========================================================================================
Sub  SetSpreadColor(ByVal lRow , byval classfg)

	With frm1.vspdData

	.Redraw = False
		ggoSpread.Source = frm1.vspdData
		If classfg = "1" Then
		'class 필수 
			ggoSpread.SSSetRequired  C_ClassCd     , lRow, lRow		'계정분류 
			ggoSpread.SSSetRequired  C_ClassNm     , lRow, lRow		'계정분류 
			ggoSpread.SSSetProtected C_ClassLvl    , lRow, lRow
			ggoSpread.SSSetRequired  C_BalFg       , lRow, lRow		'계정분류 
			'  ggoSpread.SSSetRequired  C_LeftRightFg, lRow, lRow	'좌우구분 

			'class Protect
			ggoSpread.SSSetProtected C_OccurType   , lRow, lRow		'기초기말구분 
			ggoSpread.SSSetProtected C_OccurTypeNm , lRow, lRow		'기초기말구분 
			ggoSpread.SSSetProtected C_AcctCd      , lRow, lRow
			ggoSpread.SSSetProtected C_AcctNm      , lRow, lRow
		Else
			'계정 필수 
			ggoSpread.SSSetProtected C_ClassCd     , lRow, lRow		'계정분류 
			ggoSpread.SSSetProtected C_ClassNm     , lRow, lRow		'계정분류명 
			ggoSpread.SSSetProtected C_ClassLvl    , lRow, lRow
			ggoSpread.SSSetProtected C_ClassSeq    , lRow, lRow
			ggoSpread.SSSetProtected C_LeftRightFg , lRow, lRow		'좌우구분 
			ggoSpread.SSSetProtected C_IndexFgNm   , lRow, lRow
			ggoSpread.SSSetProtected C_BalFg       , lRow, lRow		'차대구분 20030623 jsk
			ggoSpread.SSSetProtected C_ClassFgNm   , lRow, lRow
			ggoSpread.SSSetProtected C_PropertyFgNm, lRow, lRow

			'ggoSpread.SSSetRequired C_OccurType   , lRow, lRow		'기초기말구분 
			ggoSpread.SSSetProtected C_OccurTypeNm , lRow, lRow		'기초기말구분 
			ggoSpread.SSSetProtected C_AcctCd      , lRow, lRow
			ggoSpread.SSSetProtected C_AcctNm      , lRow, lRow
			'계정 Protect
		End If
		.Col = 1
		.Row = .ActiveRow
		.Action = 0                         'Parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True
		.Redraw = True

	End With
End Sub

'==========================================================================================
Sub InitComboBox()

    Dim iCodeArr,iNameArr


    'INDEX_FG
    Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1024", "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    iCodeArr = vbtab & lgF0
    iNameArr = vbtab & lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_IndexFg
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_IndexFgNm


    '손익분류 
    Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1023", "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    iCodeArr = vbtab & lgF0
    iNameArr = vbtab & lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_ClassFg
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ClassFgNm



    '계정특성 
    Call CommonQueryRs("minor_cd, minor_nm", "b_minor", " major_cd=" & FilterVar("A1000", "''", "S") & "   And  minor_cd  in  (" & FilterVar("A", "''", "S") & " ," & FilterVar("Q1", "''", "S") & "," & FilterVar("P", "''", "S") & " ," & FilterVar("G", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    iCodeArr = vbtab & lgF0
    iNameArr = vbtab & lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PropertyFg
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PropertyFgNm


    '기초기말구분 
    Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1022", "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    iCodeArr = vbtab & lgF0
    iNameArr = vbtab & lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_OccurType
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_OccurTypeNm

End Sub

'==========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
    Case "A"
        ggoSpread.Source = frm1.vspdData

        Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
        C_ClassCd         = iCurColumnPos(1)
        C_ClassNm         = iCurColumnPos(2)
        C_ClassLvl        = iCurColumnPos(3)
        C_ClassSeq        = iCurColumnPos(4)
        C_LeftRightFg     = iCurColumnPos(5)
        C_IndexFg         = iCurColumnPos(6)
        C_IndexFgNm       = iCurColumnPos(7)
        C_BalFg           = iCurColumnPos(8)
        C_ClassFg         = iCurColumnPos(9)
        C_ClassFgNm       = iCurColumnPos(10)
        C_PropertyFg      = iCurColumnPos(11)
        C_PropertyFgNm    = iCurColumnPos(12)
        C_OccurType       = iCurColumnPos(13)
        C_OccurTypeNm     = iCurColumnPos(14)
        C_PmFg            = iCurColumnPos(15)
        C_AcctCd          = iCurColumnPos(16)
        C_AcctNm          = iCurColumnPos(17)
        C_ParClassCd      = iCurColumnPos(18)
        C_ParClassCd2     = iCurColumnPos(19)
    End select
End Sub

'==========================================================================================
'	Name : OpenTransType()
'	Description : Plant PopUp
'==========================================================================================
Function OpenClassType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정분류형태 팝업"
	arrParam(1) = "A_ACCT_CLASS_TYPE"
	arrParam(2) = Trim(frm1.txtClassType.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "계정분류형태"

    arrField(0) = "CLASS_TYPE"	
	arrField(1) = "CLASS_TYPE_NM"

    arrHeader(0) = "계정분류형태코드"
	arrHeader(1) = "계정분류형태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtClassType.focus
		Exit Function
	Else
		Call SetClass(arrRet)
	End If

End Function

'==========================================================================================
'	Name : SetClass()
'	Description : Item Popup에서 Return되는 값 setting
'==========================================================================================
Function SetClass(Byval arrRet)
	With frm1
		.txtClassType.focus
		.txtClassType.value = arrRet(0)
		.txtClassTypeNm.value = arrRet(1)
	     'lgBlnFlgChgValue = True
	End With
End Function

'======================================================================================================
'	메뉴를 읽어 TreeView에 넣음 
'======================================================================================================
Sub  DIsplayAcct()
	frm1.uniTree1.MousePointer = 11
	Call SetDefaultVal
	Call AddNodes(C_CMD_TOP_LEVEL)
End Sub

'==========================================================================================
'   Event Name : AddNodes
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  AddNodes(ByVal strCmd)

	Call LayerShowHide(1)
	Call AdoQueryTree1()
End Sub

'==========================================================================================
Sub AdoQueryTree1()
	'on error resume Next
	'err.clear

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


	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=			 " gp_cd, gp_nm, gp_lvl, gp_seq   "
	strFrom		=			 " a_acct_gp(NOLOCK) "
	strWhere	=			 " gp_lvl = 1  "
	strWhere	= strWhere & " order by gp_lvl, gp_seq "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			strGpCd			= UCase(Trim(arrVal2(1)))
			strGpNm			= Trim(arrVal2(2))
			strGpLvl		= Cstr(Trim(arrVal2(3)))
			strGpSeq		= Cstr(Trim(arrVal2(4)))
			Set NodX = frm1.uniTree1.Nodes.Add (C_USER_MENU_KEY, tvwChild, "G" & strGpCd, strGpNm, C_Folder )
			frm1.uniTree1.Nodes("G" & strGpCd).Tag = cstr(strGpLvl) & "|" & cstr(strGpSeq)

		Next
	End If


	'Level 1이상에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=				" par_gp_cd ,gp_cd, gp_nm,  gp_lvl, gp_seq   "
	strFrom		=				"  a_acct_gp(NOLOCK)  "
	strWhere	=				"  gp_lvl > 1 "
	strWhere	= strWhere	&	" order by  gp_lvl,  gp_cd  ,  gp_seq "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParGpCd	= UCase(Trim(arrVal2(1)))
			strGpCd		= UCase(Trim(arrVal2(2)))
			strGpNm		= Trim(arrVal2(3))
			strGpLvl	= Trim(arrVal2(4))
			strGpSeq	= Trim(arrVal2(5))

			Set NodX = frm1.uniTree1.Nodes.Add ("G" & strParGpCd , tvwChild, "G" & strGpCd ,  strGpNm ,  C_Folder )
			frm1.uniTree1.Nodes("G" & strGpCd ).Tag = cstr( strGpLvl ) & "|" & cstr( strGpSeq )

		Next
	End If

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
			strAcctCd	= Trim(arrVal2(3))
			strAcctNm	= Trim(arrVal2(4))
			strAcctSeq	= Trim(arrVal2(5))


			Set NodX = frm1.uniTree1.Nodes.Add ("G" & strGpCd , tvwChild, "A" & strAcctCd ,  strAcctNm,  C_URL  )
			frm1.uniTree1.Nodes("A" & strAcctCd ).Tag =  cstr( strAcctSeq )

		Next
	End If
	frm1.uniTree1.MousePointer = 0

	If Not(frm1.uniTree1.Nodes("*").Child Is Nothing) Then
		frm1.uniTree1.Nodes("*").Child.EnsureVIsible
		frm1.uniTree1.Nodes("*").Child.Selected = True
	End If

	Call LayerShowHide(0)
End Sub

'==========================================================================================
'   Event Name : AddClassNodes
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  AddClassNodes(ByVal strCmd)
	Call LayerShowHide(1)
	Call AdoQueryTree2()
End Sub

'==========================================================================================
'   Event Name : AdoQueryTree
'   Event Desc : 계정분류형태를 입력후 조회시 두번째 TreeView를 Setting한다.'@@
'==========================================================================================
Sub AdoQueryTree2()
	'on error resume Next
	'err.clear
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim NodX
	Dim strClassType
	Dim strParClassCd
	Dim strClassCd
	Dim strClassNm
	Dim strClassLvl
	Dim strClassSeq
	Dim strAcctCd
	Dim strAcctNm
	Dim ii, jj
	Dim arrVal1, arrVal2

	strClassType = Trim(frm1.txtClassType.value)

	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=			 " class_cd, class_nm, class_lvl, class_seq "
	strFrom		=			 " a_acct_class(NOLOCK) "
	strWhere	=			 " class_type = " & FilterVar(strClassType, "''", "S")
	strWhere	= strWhere & " And class_lvl = 1 "
	strWhere	= strWhere & " order by class_seq, class_cd "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			strClassCd		= UCase(Trim(arrVal2(1)))
			strClassNm		= Trim(arrVal2(2))
			strClassLvl		= Cstr(Trim(arrVal2(3)))
			strClassSeq		= Cstr(Trim(arrVal2(4)))

			Set NodX = frm1.uniTree2.Nodes.Add (C_USER_MENU_KEY, tvwChild, "K" & strClassCd , strClassNm, C_Folder)
			frm1.uniTree2.Nodes("K" & strClassCd).Tag = strClassLvl & "|" & strClassSeq

		Next
	End If 

	'Level 1이상에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=				" par_class_cd , class_cd, class_nm, class_lvl, class_seq  "
	strFrom		=				" a_acct_class(NOLOCK) "
	strWhere	=				" class_type = " & FilterVar(strClassType, "''", "S")
	strWhere	= strWhere	&	" And class_lvl > 1 "
	strWhere	= strWhere	&	" order by class_lvl, class_seq, class_cd  "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)


		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParClassCd	= UCase(Trim(arrVal2(1)))
			strClassCd		= UCase(Trim(arrVal2(2)))
			strClassNm		= Trim(arrVal2(3))
			strClassLvl	= Trim(arrVal2(4))
			strClassSeq	= Trim(arrVal2(5))

			Set NodX = frm1.uniTree2.Nodes.Add ("K" & strParClassCd, tvwChild, "K" & strClassCd , strClassNm, C_Folder)

			frm1.uniTree2.Nodes("K" & strClassCd).Tag = cstr(strClassLvl) & "|" & cstr(strClassSeq)

		Next
	End If
	'계정코드에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=			  " dIstinct a.par_class_cd, a.class_cd, a.class_nm, c.acct_cd,	c.acct_nm  "
	strFrom		=			  " a_acct_class  a(nolock), a_clssfc_of_acct  b(nolock), a_acct  c(nolock) "
	strWhere	=			  " a.class_cd = b.class_cd "
	strWhere	= strWhere	& " And	a.class_type= b.class_type "
	strWhere	= strWhere  & " And	a.class_type = " & FilterVar(strClassType, "''", "S")  
	strWhere	= strWhere	& " And	b.acct_cd = c.acct_cd "
	strWhere	= strWhere  & " And	b.query_type_fg <> " & FilterVar("N", "''", "S") & "  "
	strWhere	= strWhere  & " order by a.class_cd, c.acct_cd"

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParClassCd	= UCase(Trim(arrVal2(1)))
			strClassCd		= UCase(Trim(arrVal2(2)))
			strClassNm		= Trim(arrVal2(3))
			strAcctCd	= Trim(arrVal2(4))
			strAcctNm	= Trim(arrVal2(5))

			Set NodX = frm1.uniTree2.Nodes.Add ("K" & strClassCd, tvwChild, "K" & strClassCd & "#" & strAcctCd , strAcctNm, C_URL)
		Next
	End If

	Call GridQuery()

End Sub




'==========================================================================================
'   Event Name : GridQuery
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Function GridQuery()
	Dim strVal

	Call LayerShowHide(1)
    strVal = BIZ_LOAD_GRID_CLASS_ID & "?txtClassType=" & Trim(Frm1.txtClassType.value)
	Call RunMyBizASP(MyBizASP, strVal)
End Function

'==========================================================================================
'   Function Name :ChkDragState
'   Function Desc :Drag 가 어디에 있는지 Drag되는 항목인지 체크 
'==========================================================================================
Function  ChkDragState(ByVal x , ByVal y)

    Dim NewNode
    Dim ChildNode

    On Error Resume Next

    ChkDragState = False

    With frm1

	    If gDragNode Is Nothing Then Exit Function
	    If gDragNode.parent Is Nothing Then Exit Function				' 자신이 Root인 경우 

	    Set NewNode = .uniTree2.HitTest(x, y)

	    ' 폴더가 지정되지 않고 여백이나 그런데 Drop했을 경우 
	    If NewNode Is Nothing Then
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

	    If NewNode.Children > 0 Then Set ChildNode = NewNode.Child

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

'==========================================================================================
' UserMenu를 찾는 재귀함수 
'==========================================================================================
Function ChkUserMenu(ParentNode, strFind)
	Dim blnFind

	blnFind = False

	ChkUserMenu = blnFind

	If ParentNode Is Nothing Then Exit Function

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

    If tempNode.Key <> "*" Then
	    Do
    		GetNodeLvl = GetNodeLvl + 1
    		Set tempNode = tempNode.Parent
    	Loop Until tempNode.Key = "*"
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

	For i = 1 to ParentNode.Children

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
	For i = 1 to Node.Children
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

'==========================================================================================
Sub  Form_Load()
	Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitSpreadSheet
	call InitComboBox
	Call SetToolbar("1100100100001111")
	Call InitTreeView()	
	frm1.txtClassType.focus
End Sub

'==========================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
'	Window에 발생 하는 모든 Even 처리 
'==========================================================================================
Sub btnCreate_onmousedown()
	frm1.btnCreate.src = "../../../CShared/image/btnCreate_dn.gIf"
End Sub

Sub btnCreate_onmouseup()
	frm1.btnCreate.src = "../../../CShared/image/btnCreate.gIf"
End Sub

Sub btnCreate_onmouseout()
	Call btnCreate_onmouseup()
End Sub

Sub btnCreate_onclick()
	SetTimeOut "CreateMenu", 10
End Sub

'==========================================================================================
'	Treeview Operation
'==========================================================================================
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

		.uniTree2.SetAddImageCount = 5

		.uniTree2.Indentation = "200"
		.uniTree2.AddImage C_IMG_Root, C_Root, 0
		.uniTree2.AddImage C_IMG_Folder, C_Folder, 0
		.uniTree2.AddImage C_IMG_Folder_Ch, C_Folder_Ch, 0
		.uniTree2.AddImage C_IMG_URL, C_URL, 0
		.uniTree2.AddImage C_IMG_URL_Ch, C_URL_Ch, 0

		.uniTree2.OpenTitle = "입력"
		.uniTree2.AddTitle = ""
		.uniTree2.RenameTitle = ""
		.uniTree2.DeleteTitle = "삭제"
	End With
End Sub

'==========================================================================================
Sub uniTree1_onAddImgReady()
	Call DIsplayAcct()
End Sub

'==========================================================================================
Sub uniTree2_onAddImgReady()
End Sub

'==========================================================================================
Sub uniTree1_NodeClick(Node)
	If Node.Image = C_Root Then Exit Sub
	Call SetCheck(Node, Not(IsChecked(Node)))
	Call CheckParent(Node, False)
	Call CheckChilds(Node,True)
End Sub

'==========================================================================================
Sub uniTree2_NodeClick(Node)

	Dim Response

	If Node.Image = C_Root Then Exit Sub

	Select Case Node.Image
		Case C_Folder, C_Folder_Ch
			lgSpreadNo = findCrrRow("GRP" , node.key )
		Case Else
			lgSpreadNo = findCrrRow("ACCT" , node.key )
	End Select

	If lgSpreadNo > 0  Then
		frm1.vspdData.focus
		frm1.vspdData.Row = lgSpreadNo
		frm1.vspdData.Col = 1
		frm1.vspdData.Action = 0
	End If

End Sub


'==========================================================================================
'   Event Name : uniTree2_MouseUp
'   Event 'Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  uniTree2_MouseUp(Node, Button , ShIft, X, Y)

With frm1
	If Button = 2 Or Button = 3 Then

		If ChkUserMenu(Node, C_USER_MENU_KEY) = False Then	' 유저메뉴가 아닌곳에서의 팝업 

'			Select Case Node.Image									'이 로직이 필요한가??????
'				Case C_URL, C_Folder, C_URL_Ch, C_Folder_Ch
'					.uniTree2.MenuEnabled C_MNU_OPEN, False
'				Case Else
'					.uniTree2.MenuEnabled C_MNU_OPEN, False
'					.uniTree2.MenuEnabled C_MNU_ADD, False
'					.uniTree2.MenuEnabled C_MNU_DELETE, False
'					.uniTree2.MenuEnabled C_MNU_RENAME, False
'			End Select
			Exit Sub												'2002-06-12 update line

		Else

			' 유저메뉴에서의 팝업 
			.uniTree2.MenuEnabled C_MNU_DELETE, True

			Select Case Node.Image
				Case C_URL, C_URL_Ch
					.uniTree2.MenuEnabled C_MNU_OPEN, False
					.uniTree2.MenuEnabled C_MNU_ADD, False
					.uniTree2.MenuEnabled C_MNU_RENAME, False
				Case C_None
					.uniTree2.MenuEnabled C_MNU_RENAME, False
					.uniTree2.MenuEnabled C_MNU_OPEN, False
					.uniTree2.MenuEnabled C_MNU_ADD, False
				Case C_Folder, C_Folder_Ch
		  			.uniTree2.MenuEnabled C_MNU_OPEN,True
					.uniTree2.MenuEnabled C_MNU_ADD, True
					.uniTree2.MenuEnabled C_MNU_RENAME, True
			End Select
		End If

		If Node.Key = C_USER_MENU_KEY Then
			.uniTree2.MenuEnabled C_MNU_OPEN, True
			.uniTree2.MenuEnabled C_MNU_ADD, False
			.uniTree2.MenuEnabled C_MNU_DELETE, False
			.uniTree2.MenuEnabled C_MNU_RENAME, False
			frm1.uniTree2.PopupMenu
			Exit Sub
		End If

		If mid(node.tag ,1,1) = "N" Then
			frm1.uniTree2.MenuEnabled C_MNU_OPEN,False
			frm1.uniTree2.MenuEnabled C_MNU_ADD, False
			frm1.uniTree2.MenuEnabled C_MNU_DELETE, False
		End If

		frm1.uniTree2.PopupMenu

	End If
End With

End Sub

'==========================================================================================
'   Event Name : uniTree1_MenuAdd
'   Event Desc : 사용안함(2002-06-17.확인)
'==========================================================================================
Sub  uniTree2_MenuAdd(Node)
	Dim NodX

	'CALL FNCNEW

	If Node.ExpAnded = False Then
		Node.ExpAnded = True
	End If

	If Node.Key = C_USER_MENU_KEY Then	' 유저메뉴 Root일 경우 
		Set NodX = frm1.uniTree2.Nodes.Add (Node.Key, tvwChild, C_USER_MENU_STR & GetTotalCnt(Node), C_NEW_FOLDER, C_Folder, C_Folder)
	Else
		Set NodX = frm1.uniTree2.Nodes.Add (Node.Key, tvwChild, Node.Key & C_UNDERBAR & GetTotalCnt(Node), C_NEW_FOLDER, C_Folder, C_Folder)
	End If

	NodX.Selected = True
	Set gNewNode = NodX	

	With frm1

		ggoSpread.Source = .vspdData
		.vspdData.Row = .vspdData.MaxRows
		ggoSpread.InsertRow

		SetSpreadColor .vspdData.ActiveRow , 1

		lgSpreadNo = .vspdData.ActiveRow
		node.tag = "N" & lgSpreadNo

		.vspdData.Col = C_ClassLvl:				.vspdData.Text = CStr(GetNodeLvl(NodX))
		'.vspdData.Col = C_ClassSeq:				.vspdData.Text = Cstr(GetIndex(NodX))
		If Node.Key = C_USER_MENU_KEY Then
			.vspdData.Col = C_ParClassCd:			.vspdData.Text = ""	
		ELSE
			.vspdData.Col = C_ParClassCd:			.vspdData.Text = MID(Node.key,2)
		End If
	End With
	'Call ClickTab2()

	lgIntFlgMode = parent.OPMD_CMODE	' 신규로 등록 

	lgBlnFlgChgValue = True
	lgBlnNewNode = True

End Sub

'==========================================================================================
'   Event Name : uniTree2_Menuopen
'   Event Desc : 
'==========================================================================================
Sub  uniTree2_Menuopen(Node)
	Dim NodX

	'CALL FNCNEW

	If Node.ExpAnded = False Then
		Node.ExpAnded = True
	End If

	If Node.Key = C_USER_MENU_KEY Then	' 유저메뉴 Root일 경우 
		Set NodX = frm1.uniTree2.Nodes.Add (Node.Key, tvwChild, C_USER_MENU_STR & GetTotalCnt(Node), C_NEW_FOLDER, C_Folder, C_Folder)
	Else
		Set NodX = frm1.uniTree2.Nodes.Add (Node.Key, tvwChild, Node.Key & C_UNDERBAR & GetTotalCnt(Node), C_NEW_FOLDER, C_Folder, C_Folder)
	End If

	NodX.Selected = True
	Set gNewNode = NodX	

	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.Row = .vspdData.MaxRows
		ggoSpread.InsertRow
		SetSpreadColor .vspdData.ActiveRow , 1

		lgSpreadNo = .vspdData.ActiveRow

		gNewNode.tag = "N" & .vspdData.ActiveRow

		.vspdData.Col = C_ClassLvl:						.vspdData.Text = CStr(GetNodeLvl(NodX))
		.vspdData.Col = C_ParClassCd + 2:				.vspdData.Text = NodX.key

		'.vspdData.Col = C_ClassSeq:				.vspdData.Text = Cstr(GetIndex(NodX))
		If Node.Key = C_USER_MENU_KEY Then
			.vspdData.Col = C_ParClassCd:			.vspdData.Text = ""	
		ELSE
			.vspdData.Col = C_ParClassCd:			.vspdData.Text = MID(Node.key,2)
		End If
	End With

	lgIntFlgMode = parent.OPMD_CMODE	' 신규로 등록 

	lgBlnFlgChgValue = True
	lgBlnNewNode = True

End Sub

'==========================================================================================
'   Event Name : Newnodedelete(byval Node)
'   Event Desc : 
'==========================================================================================
Function Newnodedelete(byval Node)
	Dim ndNode
    Set ndNode = Node.Child

    Do Until ndNode Is Nothing
		If mid(ndNode.tag,1,1) = "N" Then
			frm1.vspdData.Row = SearchVspdKey(ndnode.key)

			If frm1.vspdData.Row = 0 Then Exit function

			frm1.vspdData.Action = 0
			FncCancel

		Else
			Exit function
		End If

        Call Newnodedelete(ndNode)

        Set ndNode = ndNode.Next
    Loop    
End function


'==========================================================================================
'   Event Name : ReadNodeDelete(byval Node)
'   Event Desc : 
'==========================================================================================
Function ReadNodeDelete(byval Node)
	Dim ndNode
	Dim tempNo
	Dim lDelRow

    Set ndNode = Node.Child

	Do Until ndNode Is Nothing

		If mid(ndNode.tag,1,1) = "N" Then Exit function

		Select Case ndNode.Image
		Case C_Folder, C_Folder_Ch
			tempNo = findCrrRow("GRP" , ndNode.key )
		Case Else
			tempNo = findCrrRow("ACCT" , ndNode.key )
		End Select

	    frm1.vspddata.row = tempNo
	    frm1.vspddata.col = 1
	    frm1.vspddata.action = 0

	    ggoSpread.Source = frm1.vspdData
		lDelRow = ggoSpread.DeleteRow

        Call Readnodedelete(ndNode )

        Set ndNode = ndNode.Next
    Loop
End function


'==========================================================================================
'   Event Name : uniTree2_MenuDelete(Node)
'   Event Desc : uniTree2의 트리 메뉴 삭제시 처리되는 로직.
'==========================================================================================
Sub uniTree2_MenuDelete(Node)											'
	Dim lRow
	Dim lDelRow
    Dim lGrpCnt
    Dim strVal
    Dim strAcct
    Dim strClass
	Dim strDelAcct
    Dim strDelClass

    TempRootNode = Node.Key
	set gdragnode = Node

	lgDelCnt = 0

	If mid(Node.tag,1,1) = "N" Then

		set gdragnode = Node.Next

		frm1.vspdData.Row = mid(Node.Tag, InStr(Node.Tag, "N") + 1) 

		If frm1.vspdData.Row = 0 Then Exit Sub

		frm1.vspdData.Action = 0

		FncCancel

		call Newnodedelete(node)

		lgBlnFlgChgValue = False
		lgBlnNewNode = False

		frm1.uniTree2.SetFocus

		Exit Sub

	End If

	frm1.vspdData.focus
	frm1.vspdData.Row = lgSpreadNo
	frm1.vspdData.Col = 1
	frm1.vspdData.Action = 0

	ggoSpread.Source = frm1.vspdData

    lDelRow = ggoSpread.DeleteRow
End Sub


'==========================================================================================
'	Function For Treeview Operation
'   Description: UniTree1에서 Tree2로 노드(계정-child) 삽입시 체크할 사항 로직.
'==========================================================================================
Sub CheckParent(ByVal Node, ByVal blnEntrprsMnu)
'   If Node.Parent Is Nothing Then Exit Sub

    If Node.Parent.Image = C_Root Then Exit Sub

    If IsChecked(Node) = blnEntrprsMnu Then
	    If IsChecked(Node.Parent) = blnEntrprsMnu Then Exit Sub
      	Call SetCheck(Node.Parent, blnEntrprsMnu)
		Call CheckParent(Node.Parent, blnEntrprsMnu)
    End If
End Sub

'==========================================================================================
Sub CheckChilds(ByVal Node,ByVal blnEntrprsMnu)
    Dim ndNode
    Set ndNode = Node.Child
    Do Until ndNode Is Nothing
		If blnEntrprsMnu = True Then
			Call SetCheck(ndNode, IsChecked(Node))
		Else
			Call SetCheck(ndNode, False)
		End If
		Call CheckChilds(ndNode,blnEntrprsMnu)
		Set ndNode = ndNode.Next
    Loop
End Sub

'==========================================================================================
Sub CheckBrothers(ByVal Node)
    Dim ndNode,cdNode
    Set ndNode = Node.Parent.Child
    Do Until ndNode Is Nothing
		If Node <> ndNode Then
			Call SetCheck(ndNode, False)
			Call CheckChilds(ndNode,False)
	    End If
	    Set ndNode = ndNode.Next
    Loop
End Sub

'==========================================================================================
Function IsChecked(ByVal Node)
	IsChecked = False
	If Node.Image = C_Folder_Ch Or Node.Image = C_URL_Ch Then
		IsChecked = True
	End If
End Function

'==========================================================================================
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

'==========================================================================================
'	Function For Create
'==========================================================================================
Sub CreateMenu()
	Dim StrKey
	Dim StrText
	Dim intRetCD

	frm1.txtMaxRows.value = "0"
	frm1.txtSpread.value = ""
	
	If frm1.uniTree2.SelectedItem Is Nothing Then		Exit Sub
	If frm1.uniTree2.SelectedItem.Key = False Then		Exit Sub
	If frm1.uniTree2.SelectedItem.image = C_URL Then	Exit Sub
	
	If mid(frm1.uniTree2.SelectedItem.tag,1,1) = "N" Then 
		intRetCD =  DIsplayMsgBox("110604", vbOKOnly, "x", "x")
		Exit Sub
	End If
	
	If frm1.uniTree2.SelectedItem.image = C_ROOT  And frm1.group.checked = False And frm1.account.checked = False Then
		intRetCD =  DIsplayMsgBox("110620", vbOKOnly, "x", "x")
		Exit Sub
	End If

	StrKey = frm1.uniTree2.SelectedItem.Key
	StrText = frm1.uniTree2.SelectedItem.Text
	strFirst = 1

	Call CreateCoMenu(frm1.uniTree1.Nodes("*"), StrKey, StrText)
	strFirst = 1
	Call OrderAllocmain() '순서자동생성 
End Sub

'==========================================================================================
Sub CreateCoMenu(Node, StrKey, StrText)
	Dim ndNode
	Dim errNum

	If Node.Image <> C_Root Then
		'If IsChecked(Node) = False Then
		'	Exit Sub
		'End If

		On Error Resume Next
		Err.Clear

		Set ndNode = frm1.uniTree2.Nodes(Node.Key)
		errNum = Err.number
		
		On Error Goto 0

		If IsChecked(Node) = True And errNum <> 0 Then
			If Node.Image <> C_Folder_Ch Then
				If frm1.group.checked  = True And frm1.account.checked  = True Then
					If SetSaveVal(Node, "CD1", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				ElseIf frm1.group.checked  = False And frm1.account.checked  = True Then
					If SetSaveVal(Node, "CD2", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				ElseIf frm1.group.checked  = False And frm1.account.checked  = False Then
					If SetSaveVal(Node, "C1", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				ElseIf frm1.group.checked  = True And frm1.account.checked  = False Then
					If SetSaveVal(Node, "C2", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				End If
			Else
				If frm1.group.checked  = True Then
					If SetSaveVal(Node, "D", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				End If
			End If
			strFirst = 2
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
'	Set Save Value(Grid Insert)
'==========================================================================================
Function SetSaveVal(Node, strMode, StrKey, StrText, strFirst)
	Dim NodX
	Dim NodP
	Dim intRetCD
	Dim ndNode
	Dim strParentKey
	Dim strDIff
	Dim strLvl, i, strSpace, strAcctText, strBalFg

	SetSaveVal = False

	On Error Resume Next
	Err.Clear

	With frm1
		If Trim(frm1.txtClassIndex.value) <> "" Then
			strAcctText = frm1.txtClassIndex.value
		Else
			strAcctText = frm1.hClassType.value
		End If

		If strFirst = 1  or strMode = "C1" or strMode = "C3" or strMode = "CD2" Then
			strParentKey = StrKey
		Else
			strParentKey = "K" & strAcctText & MID(Node.parent.Key,2)
		End If

		If strMode = "C1" or strMode = "C2"  or strMode = "C3" Then
			'msgbox "strParentKey=" & strParentKey & "::" & "strMode=" & strMode & "::" & "tvwChild=" & tvwChild & "::" & "strAcctText=" & strAcctText & "::" & "MID(Node.Key,2)=" & MID(Node.Key,2) & "::" & "Node.Text=" & Node.Text 
			Set NodX = frm1.uniTree2.Nodes.Add (strParentKey, tvwChild, StrKey & "#" & MID(Node.Key,2), Node.Text, C_Url)
		ElseIf strMode = "D" or strMode = "CD1" or strMode = "CD2" Then
			'msgbox "strParentKey=" & strParentKey & "::" & "strMode=" & strMode & "::" & "tvwChild=" & tvwChild & "::" & "strAcctText=" & strAcctText & "::" & "MID(Node.Key,2)=" & MID(Node.Key,2) & "::" & "Node.Text=" & Node.Text 
			Set NodX = frm1.uniTree2.Nodes.Add (strParentKey, tvwChild, "K" & strAcctText & MID(Node.Key,2), Node.Text, C_Folder)
		End If

		If Err.Number <> 0 Then
			intRetCD =  DIsplayMsgBox("110102", vbOKOnly, "x", "x")
			Exit Function
		End If

		Set NodP = NodX.parent
		ggoSpread.Source = .vspdData
		'.vspdData.MaxRows = .vspdData.MaxRows + 1
		.vspdData.Row = .vspdData.MaxRows
		ggoSpread.InsertRow

		strBalFg = ""
		For i=1 to .vspddata.maxrows
			.vspdData.Row = i
			.vspddata.Col = C_ClassCd
			If Trim(.vspddata.text) = UCase(Trim(MID(NodX.parent.Key,2))) Then
				.vspddata.Col = C_BalFg
				strBalFg = .vspddata.text
				Exit For
			End If
		Next

		.vspdData.Col = C_BalFg
		.vspddata.row = .vspddata.activeRow

		If NodX.parent.Image = C_ROOT or NodX.Image= C_URL Then
			.vspdData.Text = ""
		ELSE
			If strBalFg <> "" And (strMode = "D" or strMode = "CD1"or strMode = "CD2") Then
				.vspddata.text = strBalFg
			End If
		End If

		nodX.tag = "N" & .vspdData.ActiveRow 

		.vspdData.Col = C_ClassLvl
		.vspdData.Text = CStr(GetNodeLvl(NodX.Parent)+1)

		strLvl = .vspdData.Text


'////자동생성시 계정분류코드의 생성 규칙 /////////////////////////////////////////////

'////계정그룹에 의한 계정분류코드 자동생성 ///////////////////////////////////////////

'/////1) 계정분류 index 존재시 : txtClassIndex
'///// 계정분류코드 : 계정분류index + ltrim(a_acct_gp.gp_Cd)
'///// 계정분류명 : 2space *(level-1) + ltrim(a_acct_gp.gp_nm)

'/////2)계정분류index존재하지 않을때 
'///// 계정분류코드 : 계정분류형태 + ltrim(a_acct_gp.gp_Cd)
'///// 계정분류명 : 2space *(level-1) + ltrim(a_acct_gp.gp_nm)

'////계정코드에 의한 계정분류코드 자동생성 ///////////////////////////////////////////

'/////1) 계정분류 index 존재시 
'///// 계정분류코드 : 계정분류index + ltrim(a_acct.acct_Cd)
'///// 계정분류명 : 2space *(level-1) + ltrim(a_acct.acct_nm)

'/////2)계정분류index존재하지 않을때 
'///// 계정분류코드 : 계정분류형태 + ltrim(a_acct.acct_Cd)
'///// 계정분류명 : 2space *(level-1) + ltrim(a_acct.acct_nm)
		If strMode = "D" or left(strMode,2) = "CD" or strMode = "C3" Then
			.vspdData.Col = C_ClassCd
			.vspdData.Text = strAcctText & MID(Node.Key,2)
			.vspdData.Col = C_ClassNm
			strSpace = ""
			for i=1 to strlvl-1
				strSpace= strSpace & "  "
			Next
			.vspdData.Text = strSpace & Node.Text
		Else '//계정코드를 spread sheet에 추가시............
			If strFirst = 1  or strMode = "C1" Then
				.vspdData.Col = C_BalFg
				.vspdData.Text = strBalFg	'20030623 Jsk

				.vspdData.Col = C_ClassCd
				.vspdData.Text = MID(strKey,2)
				.vspdData.Col = C_ClassNm
				strSpace = ""
				for i=1 to strlvl-1
					strSpace= strSpace & "  "
				Next
				.vspdData.Text = strSpace & strText
			Else
				.vspdData.Col = C_ClassCd
				.vspdData.Text = strAcctText & MID(Node.parent.Key,2)
				.vspdData.Col = C_ClassNm
				strSpace = ""
				for i=1 to strlvl-1
					strSpace= strSpace & "  "
				Next
				.vspdData.Text = strSpace & Node.parent.Text
			End If
		End If

		.vspdData.Col = C_OccurType
		.vspdData.Text = "N"

		If strMode = "C1" or strMode = "C2"  or strMode = "C3" Then
			.vspdData.Col = C_AcctCd
			.vspdData.Text = MID(Node.Key,2)
			.vspdData.Col = C_AcctNm
			.vspdData.Text = Node.Text
			.vspdData.Col = C_ClassSeq
			.vspdData.Text = "1"

			SetSpreadColor .vspddata.activerow , 2
		Else
			.vspdData.Col = C_AcctCd
			.vspdData.Text = ""
			.vspdData.Col = C_AcctNm
			.vspdData.Text = ""
			.vspdData.Col = C_ClassSeq
			.vspdData.Text = Cstr(GetIndex(NodX))

			SetSpreadColor .vspddata.activerow , 1
		End If

		.vspdData.Col = C_ParClassCd

		If NodX.parent.Image = C_ROOT Then
			.vspdData.Text = ""
		Else
			.vspdData.Text = MID(NodX.parent.Key,2)
		End If

		.vspdData.Col = .vspdData.MaxCols
		.vspdData.Text = NodX.key

		lgBlnFlgChgValue = True
		lgBlnNewNode = True
	End With

	If left(strMode,2) = "CD" Then
		If SetSaveVal(Node, "C3", NodX.Key, NodX.Text, strFirst) = False Then
			Exit Function
		End If
	End If

	SetSaveVal = True
End Function

'==========================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0001111111")
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

'==========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , ShIft , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
    Dim nodX
    Dim nodekeyval

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True

    If col = C_ClassNm Then
        frm1.vspdData.col = frm1.vspdData.maxcols
        nodekeyval = frm1.vspdData.text
        If nodekeyval = False  Then          '-->update part.error message:'키가 잘못되었습니다."
        set nodX = frm1.uniTree2.Nodes(nodekeyval)     '-->반영사항:not(IsNumeric(nodekeyval)) -> nodekeyval = False 
            frm1.vspdData.col = C_ClassNm
            nodX.text = LTrim(frm1.vspdData.text)
        End If
    End If

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 
End Sub

'==========================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
			Exit Sub
		End If
    End With
End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------   
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case  C_IndexFgNm
				.Col = Col
				intIndex = .Value
				.Col = C_IndexFg
				.Value = intIndex
			Case  C_ClassFgNm
				.Col = Col
				intIndex = .Value
				.Col = C_ClassFg
				.Value = intIndex
			Case  C_PropertyFgNm
				.Col = Col
				intIndex = .Value
				.Col = C_PropertyFg
				.Value = intIndex
			Case  C_OccurTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_OccurType
				.Value = intIndex

		End Select
	End With
End Sub

'==========================================================================================
Sub  InitData()
	Dim intRow
	Dim intIndex
	
	With frm1.vspdData
	    .Redraw = False

		For intRow = 1 To .MaxRows
			.Row = intRow

			.Col = C_IndexFg
			intIndex = .value
			.col = C_IndexFgNm
			.value = intindex

			.Col = C_ClassFg
			intIndex = .value
			.col = C_ClassFgNm
			.value = intindex

			.Col = C_PropertyFg
			intIndex = .value
			.col = C_PropertyFgNm
			.value = intindex

			.Col = C_OccurType
			intIndex = .value
			.col = C_OccurTypeNm
			.value = intindex

			.Col = C_AcctCd
			.Row = intRow
			' Grid Column Color 
			If Trim(.Text) <> "" Then
				ggoSpread.SpreadLock	C_ClassNm,		intRow,	C_ClassNm,		intRow
				ggoSpread.SpreadLock	C_ClassSeq,		intRow,	C_ClassSeq,		intRow
				ggoSpread.SpreadLock	C_LeftRightFg,	intRow,	C_LeftRightFg,	intRow
				ggoSpread.SpreadLock	C_IndexFgNm,	intRow,	C_IndexFgNm,	intRow
				ggoSpread.SpreadLock	C_BalFg,		intRow,	C_BalFg,		intRow	'20030623 jsk
				ggoSpread.SpreadLock	C_ClassFgNm,	intRow,	C_ClassFgNm,	intRow
				ggoSpread.SpreadLock	C_PropertyFgNm,	intRow,	C_PropertyFgNm,	intRow
			Else
				ggoSpread.SpreadLock	C_OccurTypeNm,	intRow,	C_OccurTypeNm,	intRow
			End If
		Next

	    .Redraw = True
	End With
End Sub

'==========================================================================================
Function FncQuery()
    Dim IntRetCD
    DIM NodX

    FncQuery = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DIsplayMsgBox("900013", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables
	call InitComboBox

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
		Exit Function
    End If

	frm1.uniTree2.Nodes.Clear

	If ChkClsType = True Then
		Set NodX = frm1.uniTree2.Nodes.Add(, tvwChild, C_USER_MENU_KEY, UCase(frm1.txtClassType.value), C_ROOT, C_ROOT)
	Else
		frm1.txtClassTypeNm.value = ""
		Exit Function 
	End If

	frm1.hClassType.value = Trim(frm1.txtClassType.value)
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery

    FncQuery = True
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : ChkClsType()
' Function Desc : ThIs function Is used in checking account class type .
' ModIfied date(First) : 2002-06-11
' ModIfied date(Last)  : 
' ModIfier (First)     : Heo Chung Ku
' ModIfier (Last)      : 
'========================================================================================
Function ChkClsType()																'User Defined Function 
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strClassType
	Dim IntRetCD

	strClassType = Trim(frm1.txtClassType.value)

	'class type yes/no check
	'------------------------
	strSelect = " class_type, class_type_nm "
	strFrom = " a_acct_class_type(NOLOCK)"
	strWhere = " class_type = " & FilterVar(strClassType, "''", "S")

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
		ChkClsType= True
	Else
		ChkClsType = False
        IntRetCD = DIsplayMsgBox("110500",vbOkOnly,"X","X")
	End If
End Function

'==========================================================================================
Function FncSave()
	Dim IntRetCD

    FncSave = False

    Err.Clear

    '-----------------------
    'Precheck area
    '-----------------------
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DIsplayMsgBox("900001","X","X","X")                            '⊙: No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
	ggoSpread.Source = frm1.vspdData
	If Not chkField(Document, "2") OR ggoSpread.SSDefaultCheck = False Then        '⊙: Check contents area
		Exit Function
	End If

    If Not chkField(Document, "1") Then
		Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave

    FncSave = True
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function FncCopy()
	Dim IntRetCD

	If frm1.vspdData.maxrows < 1 Then Exit function

	frm1.vspdData.row = frm1.vspdData.activerow
	frm1.vspdData.col = C_AcctCd

	If Trim(frm1.vspdData.text) = "" Then
		'IntRetCD = DIsplayMsgBox("900001","X","X","X")
		Exit function
	End If

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow , 2

    frm1.vspdData.Col = C_OccurType
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_OccurTypeNm
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function FncCancel()
	Dim nodX
	On error resume Next
	Err.clear

	If frm1.vspddata.maxrows < 1  Then Exit function
    If frm1.vspddata.row = 0 Then Exit function

    frm1.vspddata.row =frm1.vspddata.activerow
    frm1.vspddata.col = frm1.vspddata.maxcols

    If not(IsNumeric(frm1.vspddata.text)) And Trim(frm1.vspddata.text) <> "" Then
    	set nodX = frm1.unitree2.nodes(frm1.vspddata.Text)
        frm1.uniTree2.Nodes.Remove NodX.Index
    End If


	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo

	Call InitData
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function FncInsertRow()
	Dim StrKey
	Dim StrText
	Dim NodX
	Dim NodP

	If frm1.uniTree2.SelectedItem.Key = False Then	Exit Function
	If frm1.uniTree2.SelectedItem.image = C_Folder Then Exit Function
	If frm1.uniTree2.SelectedItem.image = C_ROOT Then Exit Function

	Set NodP = frm1.uniTree2.SelectedItem.Parent
	Set NodX = frm1.uniTree2.SelectedItem
	StrKey = frm1.uniTree2.SelectedItem.Key
	StrText = frm1.uniTree2.SelectedItem.Text

	ggoSpread.Source = frm1.vspdData

	frm1.vspdData.ReDraw = False

	ggoSpread.InsertRow

	With frm1		
	.vspdData.Col = C_ClassLvl:				.vspdData.Text = CStr(GetNodeLvl(NodP))
	.vspdData.Col = C_ClassSeq:				.vspdData.Text = Cstr(GetIndex(NodP))
	.vspdData.Col = C_ClassCd:				.vspdData.Text = MID(NodP.key,2)
	.vspdData.Col = C_ClassNm:				.vspdData.Text = NodP.text
	.vspdData.Col = C_AcctCd:				.vspdData.Text = MID(NodX.Key,2)
	.vspdData.Col = C_AcctNm:				.vspdData.Text = NodX.Text

	End With
	frm1.vspdData.ReDraw = True
	lgBlnFlgChgValue = True
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function FncDeleteRow()
	Dim lDelRows
    Dim iDelRowCnt, i

    If frm1.vspdData.Maxrows < 1 Then Exit Function
		With frm1.vspdData
			.focus
			ggoSpread.Source = frm1.vspdData
			lDelRows = ggoSpread.DeleteRow
			lgBlnFlgChgValue = True
		End With
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function FncPrint()
    On Error Resume Next
    parent.FncPrint()
End Function

'==========================================================================================
Function FncPrev()
    On Error Resume Next
End Function

'==========================================================================================
Function FncNext()
    On Error Resume Next
End Function


'==========================================================================================
Function FncPrint()
    On Error Resume Next
End Function


'==========================================================================================
Function  FncExcel()
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function


'==========================================================================================
Function  FncPrint()
    On Error Resume Next
    parent.FncPrint()
End Function

'==========================================================================================
Function  FncFind()
    Call parent.FncFind(parent.C_SINGLEMULTI , True)
End Function

'==========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'==========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'==========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
    'Call SetSpreadLock( "I", 0, 1, "")
	Call InitData()
End Sub

'==========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DIsplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function

'==========================================================================================
Function DbQuery()
    Dim LngLastRow
    Dim LngMaxRow
    Dim StrNextKey

    DbQuery = False
    Err.Clear

    Call AddClassNodes(C_CMD_TOP_LEVEL)

    DbQuery = True    
End Function

'==========================================================================================
Function DbQueryOk()

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

    SetSpreadLock "Q", 0, 1, ""
    Call SetToolbar("1100100100011111")	
'    Call ggoOper.LockField(Document, "Q")	
	If Not(frm1.uniTree1.Nodes("*").Child Is Nothing) Then
		frm1.uniTree1.Nodes("*").Child.EnsureVIsible
		frm1.uniTree1.Nodes("*").Child.Selected = True
	End If
	If Not(frm1.uniTree2.Nodes("*").Child Is Nothing) Then
		frm1.uniTree2.Nodes("*").Child.EnsureVIsible
		frm1.uniTree2.Nodes("*").Child.Selected = True
	End If
'	Call LayerShowHide(0)
	Call InitData
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function  DbSave() 
	Dim lRow
    Dim lGrpCnt
    Dim strVal
    Dim strDel

    DbSave = False

    'On Error Resume Next


	Call LayerShowHide(1)


	With frm1
		.txtSpread.value = ""
		.txtMode.value = parent.UID_M0002
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows

		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text

			Case ggoSpread.InsertFlag											'☜: 신규 
			    strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep  					'☜: C=Create, Row위치 정보 
				.vspdData.Col = C_ClassCd
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassNm
				strVal = strVal & .vspdData.Text & parent.gColSep
				.vspdData.Col = C_ClassLvl
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassSeq
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_LeftRightFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_IndexFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_BalFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PropertyFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_OccurType
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PmFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_AcctCd
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ParClassCd
				strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

				lGrpCnt = lGrpCnt + 1

			Case ggoSpread.UpdateFlag											'☜: 수정 

				strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep  					'☜: U=Update, Row위치 정보 
				.vspdData.Col = C_ClassCd
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassNm
				strVal = strVal & .vspdData.Text & parent.gColSep
				.vspdData.Col = C_ClassLvl
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassSeq
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_LeftRightFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_IndexFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_BalFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PropertyFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_OccurType
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PmFg
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_AcctCd
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ParClassCd
				strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
				   
				lGrpCnt = lGrpCnt + 1

			Case ggoSpread.DeleteFlag											'☜: 삭제

				strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep  					'☜: D=Delete, Row위치 정보 
				.vspdData.Col = C_ClassCd
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassNm
				strDel = strDel & .vspdData.Text & parent.gColSep
				.vspdData.Col = C_ClassLvl
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassSeq
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_LeftRightFg
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_IndexFg
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_BalFg
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ClassFg
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PropertyFg
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_OccurType
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_PmFg
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_AcctCd
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_ParClassCd
				strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

				lGrpCnt = lGrpCnt + 1

				lgRetFlag = True
			End Select
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value =  strVal & strDel 
		Call ExecMyBizASP(frm1, BIZ_SAVE_CLASS_ID)
	End With	
    DbSave = True
End Function

'==========================================================================================
Function DbSaveOk()
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	Call FncQuery()
	lgBlnNewNode = False
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function  DbDelete() 
End Function


'==========================================================================================
Function  OrderAllocmain()
	orderseq = 0
	If frm1.unitree2.nodes.count =0 Then Exit function
	
	call OrderAlloc(frm1.unitree2.nodes("*"))
End Function

'==========================================================================================
Function  OrderAlloc(ByVal Node)
	Dim ndNode
	Dim tempNo
	Dim iRow

    Set ndNode = Node.Child

	For iRow = 1 to frm1.vspddata.maxRows
		frm1.vspddata.row = iRow
		frm1.vspdData.col = C_AcctCd
		If frm1.vspdData.text = "" Then
			orderseq = orderseq + 1
			frm1.vspddata.col = C_ClassSeq
			frm1.vspddata.text = orderseq 
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow frm1.vspddata.row
		End If
	Next
End Function

'==========================================================================================
' 현재의 노드 키값으로 스프레드의 row 값을 반환 (즉 새로운 입력상태의 노드만 찾을 수 있음)
'==========================================================================================
Function SearchVspdKey(byval nodekey)
	Dim iRow

	For iRow = 1 to frm1.vspddata.maxRows
		frm1.vspddata.row = iRow
		frm1.vspddata.col = frm1.vspddata.maxcols
		
		If frm1.vspddata.text = nodekey Then
			SearchVspdKey = iRow
			Exit Function
		End If
	Next
End function

'==========================================================================================
' 현재 노드의 그룹상태(그룹이냐 계정이냐) 와 키값으로 스프레드의 row 값을 반환 
'==========================================================================================
Function  findCrrRow(byVal Flag , byval FindVal)
	Dim iRow
	Dim FindValOfCol, FindClassCode
	Dim SharpFlag

	If Flag = "GRP" Then
		FindValOfCol = mid(FindVal,2)

		With frm1.vspdData
			For iRow = 1 to .Maxrows
				.Row = iRow
				.Col = C_ClassCd

				If UCase(.Text) = UCase(FindValOfCol) Then
					findCrrRow = iRow
					Exit Function
				End If
			Next
		End With
	Else
		SharpFlag		= inStr(1,FindVal,"#")
		FindValOfCol	= mid(FindVal,SharpFlag+1)
		FindClassCode	= mid(FindVal,2,SharpFlag-2 )

		With frm1.vspdData
			For iRow = 1 to .Maxrows
				.Row = iRow
				.Col = C_ClassCd

				If UCase(.Text) = UCase(FindClassCode) Then
					.Col = C_AcctCd

					If UCase(.Text) = UCase(FindValOfCol) Then
						findCrrRow = iRow
						Exit Function
					End If
				End If
			Next
		End With
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gIf"><IMG src="../../../CShared/image/table/seltab_up_left.gIf" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="center" CLASS="CLSMTABP"><font color=white>유형별계정분류등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gIf" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:OrderAllocmain()" >순서자동생성</a></TD>
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
									<TD CLASS="TD5">계정분류형태</TD>
									<TD CLASS="TD656">
										<INPUT NAME="txtClassType" MAXLENGTH="4" SIZE=10 ALT ="계정분류형태" tag="13XXXU"><IMG SRC="../../../CShared/image/btnPopup.gIf" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenClassType()">&nbsp;
										<INPUT NAME="txtClassTypeNm" MAXLENGTH="50" SIZE=20 ALT ="" tag="14X">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">계정분류INDEX</TD>
									<TD CLASS="TD656">
										<INPUT NAME="txtClassIndex" MAXLENGTH="4" SIZE=10 ALT ="계정분류INDEX" tag="11XXXU">&nbsp;&nbsp;&nbsp;
										<input type="checkbox" class = "check" name="chkGroup" value="Y" id="group"><label for="group">계정그룹</label>&nbsp;&nbsp;
										<input type="checkbox" class = "check" name="chkAccount" value="N" id="account"><label for="account">계정코드</label>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100% WIDTH=20% ROWSPAN=4>
									<script language =javascript src='./js/a2104ma1_uniTree1_N393563538.js'></script>
								</TD>
								<TD HEIGHT=* WIDTH=10>&nbsp;</TD>
								<TD HEIGHT=100% WIDTH=20% ROWSPAN=4>
									<script language =javascript src='./js/a2104ma1_uniTree2_N120885396.js'></script>
								</TD>
								<TD HEIGHT="100%" WIDTH=* ROWSPAN=4>
									<script language =javascript src='./js/a2104ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=10 WIDTH=10>
									<IMG SRC="../../../CShared/image/btnCreate.gIf" NAME="btnCreate">
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=* WIDTH=10>&nbsp;</TD>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IfRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IfRAME></TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="hClassType" tag="24">
<INPUT TYPE=hidden NAME="hClassIndex" tag="24">


</FORM>
<DIV ID="MousePT" NAME="MousePT">
<Iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></Iframe>
</DIV>
</BODY>
</HTML>

