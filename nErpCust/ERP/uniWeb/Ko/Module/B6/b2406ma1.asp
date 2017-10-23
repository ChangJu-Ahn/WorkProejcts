<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Basic Architecture
'*  2. Function Name        : 
'*  3. Program ID           : B2406ma1
'*  4. Program Name         : 부서정보등록 
'*  5. Program Desc         :
'*  6. Component LIST       : 
'*  7. ModIfied date(First) : 2005/10/19
'*  8. ModIfied date(Last)  : 
'*  9. ModIfier (First)     : Jeong Yong Kyun
'* 10. ModIfier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : thIs mark(☜) means that "Do not change"
'*                            thIs mark(⊙) Means that "may  change"
'*                            thIs mark(☆) Means that "must change"
'* 13. HIstory              :
'*                            
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

Const BIZ_SAVE_CLASS_ID			= "b2406mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_LOAD_GRID_CLASS_ID	= "b2406mb2.asp"
Const BIZ_BATCH_PGM_ID          = "b2406mb3.asp"
Const BIZ_ORDER_PGM_ID          = "b2406mb4.asp"  
Const BIZ_DEPT_MOVE_ID          = "b2406mb5.ASP"

'==========================================================================================
Const C_IMG_Root        = "../../../CShared/image/unierp.gIf"
Const C_IMG_Folder      = "../../../CShared/image/folder.gIf"
Const C_IMG_Folder_Ch   = "../../../CShared/image/folder_ch.gIf"
Const C_IMG_URL         = "../../../CShared/image/Account.gIf"
Const C_IMG_URL_Ch      = "../../../CShared/image/Account_Ch.gIf"

Const C_CMD_TOP_LEVEL   = "LISTTOP"
Const C_CMD_LIST_LEVEL  = "LIST"
Const C_CMD_LIST_DIsT   = "ACCTDIST"
Const C_CMD_ACCT_LEVEL  = "LISTACCT"
Const C_CMD_GP_LEVEL    = "LISTGP"

Const C_USER_MENU       = "UNIERP"
Const C_USER_MENU_KEY   = "*"
Const C_USER_MENU_STR   = "UM_"
Const C_UNDERBAR        = "_"

Const C_NEW_FOLDER      = "새 부서"

Dim C_ORGID
Dim C_DEPT
Dim C_PDEPT
Dim C_LDEPTNM
Dim C_BUILDID
Dim C_LVL
Dim C_SEQ
Dim C_SDEPTNM
Dim C_EDEPTNM
Dim C_ENDDEPTYN
Dim C_ENTRY_FG

Const  C_Root      = "Root"
Const  C_Folder_Ch = "folder_ch.gIf"
Const  C_URL_Ch    = "URL_Ch"

'==========================================================================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'==========================================================================================

Dim  lgStrPrevKey1
Dim  lgStrPrevKey2
Dim  lgQueryFlag
Dim	 lgCur_Orgid
Dim  lgGroup
Dim  lgCode
Dim	 lglsClicked

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
	C_ORGID		= 1
	C_DEPT		= 2
	C_PDEPT		= 3
	C_LDEPTNM	= 4
	C_BUILDID	= 5
	C_LVL		= 6
	C_SEQ		= 7
	C_SDEPTNM	= 8
	C_EDEPTNM	= 9
	C_ENDDEPTYN	= 10
	C_ENTRY_FG  = 11
End Sub

'==========================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgQueryFlag = "1"

	Call CommonQueryRs("orgid", "horg_abs"," currentyn='Y'",lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	lgCur_Orgid = Left(lgF0, Len(lgF0)-1)
	strFirst = 1
End Sub

'==========================================================================================
Sub  SetDefaultVal()
	Dim NodX
	frm1.uniTree1.Nodes.Clear 
	Set NodX = frm1.uniTree1.Nodes.Add(, tvwChild, C_USER_MENU_KEY, lgCur_Orgid, C_Root, C_Root)
	strFirst = 1
	
	lgGroup = True
	lgCode = False

End Sub

'==========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================================================================
Sub  InitSpreadSheet()
	Call initSpreadPosVariables()

    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread

    With frm1.vspdData
		.MaxCols = C_ENTRY_FG + 1   '' 마지막 상수명 사용 
		.MaxRows = 0
		.ReDraw = False

		Call AppendNumberPlace("6","2","0")
		Call AppendNumberPlace("7","3","0")
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit  C_ORGID,     "부서개편ID"  , 12,,,5,2
		ggoSpread.SSSetEdit  C_DEPT,      "부서코드"    , 10,,,10,2  
		ggoSpread.SSSetEdit  C_PDEPT,     "모부서"      , 10,,,10,2
		ggoSpread.SSSetEdit  C_LDEPTNM,   "부서명"      , 20,,,200,1
		ggoSpread.SSSetEdit  C_BUILDID,   "내부부서코드", 10,,,10
		ggoSpread.SSSetFloat C_LVL,       "레벨"        ,  6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","10"
		ggoSpread.SSSetFloat C_SEQ,       "순서"        ,  6,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","999"
		ggoSpread.SSSetEdit  C_SDEPTNM,   "약칭부서명"  , 15,,,24,1
		ggoSpread.SSSetEdit  C_EDEPTNM,   "영문부서명"  , 15,,,100,1
		ggoSpread.SSSetCheck C_ENDDEPTYN, "말단부서여부", 14, 2, "말단부서", False    
	    ggoSpread.SSSetEdit  C_ENTRY_FG  , "", 4
    
		Call ggoSpread.SSSetColHidden(C_ENTRY_FG,C_ENTRY_FG,True)				        
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = True

		Call SetSpreadLock
    End With
End Sub

Sub SetSpreadLock()
	Dim ii
		
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_ORGID, -1, C_ORGID
		ggoSpread.SpreadLock C_DEPT, -1, C_PDEPT
		ggoSpread.SSSetRequired C_LDEPTNM, -1, -1
		ggoSpread.SpreadLock C_BUILDID, -1, C_BUILDID
		ggoSpread.SpreadLock C_LVL, -1, C_LVL
		ggoSpread.SSSetRequired C_SEQ, -1, -1
		ggoSpread.SSSetRequired C_SDEPTNM, -1, -1
		'ggoSpread.SpreadLock C_ENDDEPTYN, -1, C_ENDDEPTYN
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
		.vspdData.ReDraw = True

		For ii = 1 To .vspdData.MaxRows
			.vspddata.col = C_ENTRY_FG
			.vspddata.row = ii
			
			If Trim(.vspddata.text) = "E" Then			
				ggoSpread.SpreadLock C_ORGID, ii, C_ENTRY_FG ,ii
			End If
		Next
	End With	
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_ORGID, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_DEPT, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_LDEPTNM, pvStartRow, pvEndRow
		ggoSpread.SpreadUnlock C_PDEPT, pvStartRow, C_PDEPT,pvEndRow		
		ggoSpread.SSSetProtected C_BUILDID, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LVL, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_SEQ, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_SDEPTNM, pvStartRow, pvEndRow  
		ggoSpread.SSSetProtected C_ENDDEPTYN, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub

'==========================================================================================
Sub InitComboBox()

End Sub

'==========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
    Case "A"
        ggoSpread.Source = frm1.vspdData

        Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

        C_ORGID     = iCurColumnPos(1)
        C_DEPT      = iCurColumnPos(2)
        C_PDEPT     = iCurColumnPos(3)
        C_LDEPTNM   = iCurColumnPos(4)
        C_BUILDID   = iCurColumnPos(5)
        C_LVL       = iCurColumnPos(6)
        C_SEQ       = iCurColumnPos(7)
        C_SDEPTNM   = iCurColumnPos(8)
        C_EDEPTNM   = iCurColumnPos(9)
        C_ENDDEPTYN = iCurColumnPos(10)
    End select
End Sub

'==========================================================================================
'	Name : OpenTransType()
'	Description : Plant PopUp
'==========================================================================================
Function OpenOrgId()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "부서개편ID팝업"
	arrParam(1) = "  horg_abs "
	arrParam(2) = Trim(frm1.txtOrgId.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "부서개편ID"

    arrField(0) = "Orgid"	
	arrField(1) = "orgdt"

    arrHeader(0) = "부서개편ID"
	arrHeader(1) = "부서개편일자"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtOrgId.focus
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
		.txtOrgId.focus
		.txtOrgId.value = arrRet(0)
		.txtOrgDt.value = arrRet(1)
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
	Dim strParDeptCd
	Dim strDeptCd
	Dim strDeptNm
	Dim strDeptLvl
	Dim strDeptSeq

	Dim ii, jj
	Dim arrVal1, arrVal2

	'----------------------------------------------------------------------------------------
	'Level 1이상에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=				" dept, ldeptnm,  lvl, seq,sdeptnm,edeptnm   "
	strFrom		=				" horg_mas  "
	strWhere	=               " lvl = 1 and orgid = " & lgCur_Orgid
	strWhere	= strWhere	&	" order by buildid, lvl, seq ,dept "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2		= Split(arrVal1(ii), chr(11))
			strDeptCd	= UCase(Trim(arrVal2(1)))
			strDeptNm	= Trim(arrVal2(2)) 
			strDeptLvl	= Cstr(Trim(arrVal2(3)))
			strDeptSeq	= Cstr(Trim(arrVal2(4)))

			Set NodX = frm1.uniTree1.Nodes.Add (C_USER_MENU_KEY, tvwChild, "G" & strDeptCd, strDeptNm, C_Folder )
'			frm1.uniTree1.Nodes("G" & strDeptCd).Tag = cstr(strDeptLvl) & "|" & cstr(strDeptSeq)
			frm1.uniTree1.Nodes("G" & strDeptCd).Tag = Trim(arrVal2(5)) & parent.gcolSep & Trim(arrVal2(6))
		Next
	End If


	'Level 1이상에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=				" pdept ,dept, ldeptnm,  lvl, seq ,sdeptnm,edeptnm   "
	strFrom		=				"  horg_mas  "
	strWhere	=               " lvl > 1 and orgid = " & lgCur_Orgid
	strWhere	= strWhere	&	" order by  buildid, lvl, seq, dept "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParDeptCd = UCase(Trim(arrVal2(1)))
			strDeptCd	 = UCase(Trim(arrVal2(2)))
			strDeptNm	 = Trim(arrVal2(3))
			strDeptLvl	 = Trim(arrVal2(4))
			strDeptSeq	 = Trim(arrVal2(5))

			Set NodX = frm1.uniTree1.Nodes.Add ("G" & strParDeptCd , tvwChild, "G" & strDeptCd ,  strDeptNm ,  C_Folder )
'			frm1.uniTree1.Nodes("G" & strDeptCd ).Tag = cstr( strDeptLvl ) & "|" & cstr( strDeptSeq )
			frm1.uniTree1.Nodes("G" & strDeptCd ).Tag = Trim(arrVal2(6)) & parent.gcolSep & Trim(arrVal2(7))

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
'	on error resume Next
'	err.clear
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim NodX
	Dim strOrgId
	Dim strParDeptCd
	Dim strDeptCd
	Dim strDeptNm
	Dim strDeptLvl
	Dim strDeptSeq

	Dim ii, jj
	Dim arrVal1, arrVal2

	strOrgId = Trim(frm1.txtOrgId.value)

	'Level 1에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=			 " dept, ldeptnm,  lvl, seq ,sdeptnm, edeptnm  "
	strFrom		=			 " horg_mas  "
	strWhere	=			 " orgid = " & FilterVar(strOrgId, "''", "S")
	strWhere	= strWhere & " And lvl = 1 "
	strWhere	= strWhere & " order by buildid,lvl ,seq ,dept"

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			strDeptCd		= UCase(Trim(arrVal2(1)))
			strDeptNm		= Trim(arrVal2(2)) 
			strDeptLvl		= Cstr(Trim(arrVal2(3)))
			strDeptSeq		= Cstr(Trim(arrVal2(4)))

			Set NodX = frm1.uniTree2.Nodes.Add (C_USER_MENU_KEY, tvwChild, "K" & strDeptCd , strDeptNm, C_Folder)
'			frm1.uniTree2.Nodes("K" & strDeptCd).Tag = strDeptLvl & "|" & strDeptSeq
			frm1.uniTree2.Nodes("K" & strDeptCd).Tag = Trim(arrVal2(5)) & parent.gColSep & Trim(arrVal2(6))
		Next
	End If 

	'Level 1이상에 대한 Node가져오기 
	'----------------------------------------------------------------------------------------
	strSelect	=				" pdept ,dept, ldeptnm,  lvl, seq ,sdeptnm, edeptnm  "
	strFrom		=				"  horg_mas  "
	strWhere	=				" orgid = " & FilterVar(strOrgId, "''", "S")
	strWhere	= strWhere	&	" And lvl > 1 "
	strWhere	= strWhere	&	" order by buildid,lvl, seq ,dept "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1)

		For ii = 0 To jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			strParDeptCd	= UCase(Trim(arrVal2(1)))
			strDeptCd		= UCase(Trim(arrVal2(2))) 
			strDeptNm		= Trim(arrVal2(3)) 
			strDeptLvl	    = Trim(arrVal2(4))
			strDeptSeq	    = Trim(arrVal2(5))

			Set NodX = frm1.uniTree2.Nodes.Add ("K" & strParDeptCd, tvwChild, "K" & strDeptCd , strDeptNm, C_Folder)

'			frm1.uniTree2.Nodes("K" & strDeptCd).Tag = Cstr(strDeptLvl) & "|" & Cstr(strDeptSeq)
			frm1.uniTree2.Nodes("K" & strDeptCd).Tag = Trim(arrVal2(6)) & parent.gColSep & Trim(arrVal2(7))
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

    strVal = BIZ_LOAD_GRID_CLASS_ID & "?txtOrgId=" & Trim(Frm1.txtOrgId.value)
	Call RunMyBizASP(MyBizASP, strVal)
End Function

'==========================================================================================
'   Function Name :ChkDragState
'   Function Desc :Drag 가 어디에 있는지 Drag되는 항목인지 체크 
'==========================================================================================
Function  ChkDragState(ByVal x , ByVal y)

    Dim NewNode
    Dim ChildNode
    Dim i

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

	    If NewNode.Children > 0 Then 
	       Set ChildNode = NewNode.Child

			For i = 1 To NewNode.Children
				If ChildNode.Key = gDragNode.Key Then
			  		Set NewNode = Nothing
					Exit Function
				End If
				Set ChildNode = ChildNode.Next
			Next
		End if	

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

	frm1.uniTree2.OLEDragMode = 1														'⊙: Drag & Drop 을 가능하게 할 것인가 정의 
	frm1.uniTree2.OLEDropMode = 1

	frm1.txtOrgId.focus
	lglsClicked = False 
End Sub

'==========================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
'	Window에 발생 하는 모든 Event 처리 
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

	lgSpreadNo = findCrrRow("GRP" , node.key )

	If lgSpreadNo > 0  Then
		frm1.vspdData.focus
		frm1.vspdData.Row = lgSpreadNo
		frm1.vspdData.Col = 1
		frm1.vspdData.Action = 0
	End If

End Sub

'==========================================================================================
'   Event Name : uniTree2_OLEDragDrop
'   Event Desc : Node를 Drag & Drop 이벤트 
'==========================================================================================
Sub  uniTree2_OLEDragDrop(Data , Effect , Button , Shift , x , y )
	Dim NewNode, IntRetCD
    Dim strVal, strUpKey, Index
    Dim iObjNewNode

	'클릭시 이동할수 없습니다. 메세지 뜨는 버그 수정 
'	On Error Resume Next

    Set iObjNewNode = frm1.uniTree2.HitTest(x, y)
    
    If iObjNewNode Is Nothing Then Exit Sub
	If iObjNewNode.key = gDragNode.key Then Exit Sub

	Set iObjNewNode = Nothing

	If gDragNode Is Nothing Then Exit Sub

	If ChkDragState(x, y) = False Then
        Effect = vbDropEffectNone
		IntRetCD = DisplayMsgBox("990017","X","X","X")	' 해당 위치로는 이동할 수 없습니다!
		frm1.uniTree2.MousePointer = 0
        Exit Sub
	End If

	Call LayerShowHide(1)

    Set NewNode = frm1.uniTree2.HitTest(x, y)
    Set gDropNode = NewNode					' 이동해야될 노드를 기억시킴 

'msgbox "af_cd::" & Mid(gDropNode.Key, 2)
'msgbox "af_lvl::" & GetNodeLvl(gDropNode)
'msgbox "af_seq::" & GetIndex(gDropNode)
'msgbox "be_cd::" &  Mid(gDragNode.parent.key, 2)
'msgbox "be_lvl::" & GetNodeLvl(gDragNode.Parent)
'msgbox "be_seq::" & GetIndex(gDragNode.Parent)

	frm1.txtToParentDeptCd.value = Mid(gDropNode.Key, 2)
	frm1.txtToParentDeptLvl.value = GetNodeLvl(gDropNode)
	frm1.txtToParentDeptSeq.value = GetIndex(gDropNode)

	frm1.txtParentDeptCd.value = Mid(gDragNode.parent.key, 2)
	frm1.txtParentDeptLvl.value = GetNodeLvl(gDragNode.Parent)
	frm1.txtParentDeptSeq.value = GetIndex(gDragNode.Parent)

'	frm1.lgstrCmd.value = "GP"
	frm1.txtMoveDeptCd.value = Mid(gDragNode.Key, 2)
	frm1.txtMoveDeptLvl.value = GetNodeLvl(gDropNode)
	frm1.txtMoveDeptSeq.value = GetInsSeq(gDropNode)

	Call ExecMyBizASP(frm1, BIZ_DEPT_MOVE_ID)										'☜: 비지니스 ASP 를 가동 
End Sub

'========================================================================================================= 
Sub uniTree2_MouseDown(Button, Shift, X, Y)
	If frm1.uniTree2.IsNodeClicked = "Y" Then
		lglsClicked = True
	Else
		lglsClicked = False
	End If
End Sub

'==========================================================================================
'   Event Name : uniTree_OLEStartDrag
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  uniTree2_OLEStartDrag(Data, AllowedEffects)
	If lglsClicked = True Then
		Set gDragNode = frm1.uniTree2.SelectedItem
		gDragNode.Selected = True
	Else
		Set gDragNode = Nothing
	End If

	lglsClicked = False
End Sub

'==========================================================================================
'   Event Name : uniTree2_MouseUp
'   Event 'Desc : Node를 Drag할때 이벤트 
'==========================================================================================
Sub  uniTree2_MouseUp(Node, Button , ShIft, X, Y)

	With frm1
		If Button = 2 Or Button = 3 Then
			If ChkUserMenu(Node, C_USER_MENU_KEY) = False Then	' 유저메뉴가 아닌곳에서의 팝업 
				Exit Sub												'2002-06-12 update line
			Else
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
'   Event Name : uniTree2_MenuAdd
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

		'.vspdData.Col = C_LVL:				.vspdData.Text = CStr(GetNodeLvl(NodX))
		'.vspdData.Col = C_SEQ:				.vspdData.Text = Cstr(GetIndex(NodX))
'		msgbox Node.key
		If Node.Key = C_USER_MENU_KEY Then
			.vspdData.Col = C_PDEPT:			.vspdData.Text = ""	
		ELSE
			.vspdData.Col = C_PDEPT:			.vspdData.Text = MID(Node.key,2)
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
		SetSpreadColor .vspdData.ActiveRow ,.vspdData.ActiveRow

		lgSpreadNo = .vspdData.ActiveRow

		gNewNode.tag = "N" & .vspdData.ActiveRow

		.vspdData.Col = C_LVL:						.vspdData.Text = CStr(GetNodeLvl(NodX))
'		.vspdData.Col = C_PDEPT + 2:				.vspdData.Text = NodX.key

		.vspddata.col = C_ORGID	:			.vspdData.Text = Trim(frm1.hOrgId.value)

		'.vspdData.Col = C_SEQ:				.vspdData.Text = Cstr(GetIndex(NodX))

		If Node.Key = C_USER_MENU_KEY Then
			.vspdData.Col = C_PDEPT:			.vspdData.Text = ""	
		Else
			.vspdData.Col = C_PDEPT:			.vspdData.Text = MID(Node.key,2)
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

		tempNo = findCrrRow("GRP" , ndNode.key )

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
'	 msgbox frm1.uniTree2.SelectedItem.tag
	If mid(frm1.uniTree2.SelectedItem.tag,1,1) = "N" Then 
		intRetCD =  DIsplayMsgBox("124543", vbOKOnly, "x", "x")
		Exit Sub
	End If
	
'	If frm1.uniTree2.SelectedItem.image = C_ROOT  Then
'		intRetCD =  DIsplayMsgBox("110620", vbOKOnly, "x", "x")
'		Exit Sub
'	End If

	StrKey = frm1.uniTree2.SelectedItem.Key
	StrText = frm1.uniTree2.SelectedItem.Text
	strFirst = 1

	Call CreateCoMenu(frm1.uniTree1.Nodes("*"), StrKey, StrText)
	strFirst = 1
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
				If lgGroup  = True And lgCode  = True Then
					If SetSaveVal(Node, "CD1", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				ElseIf lgGroup  = False And lgCode  = True Then
					If SetSaveVal(Node, "CD2", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				ElseIf lgGroup  = False And lgCode  = False Then
					If SetSaveVal(Node, "C1", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				ElseIf lgGroup  = True And lgCode  = False Then
					If SetSaveVal(Node, "C2", StrKey, StrText, strFirst) = False Then
						Exit Sub
					End If
				End If
			Else
				If lgGroup  = True Then
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
	Dim strParentKey,strTemp
	Dim strDIff
	Dim strLvl, i, strSpace, strAcctText, strBalFg

	SetSaveVal = False

	On Error Resume Next
	Err.Clear

	With frm1
		If strFirst = 1  or strMode = "C1" or strMode = "C3" or strMode = "CD2" Then
			strParentKey = StrKey
		Else
			strParentKey = "K" & strAcctText & MID(Node.parent.Key,2)
		End If
'msgbox strParentKey
		If strMode = "C1" or strMode = "C2"  or strMode = "C3" Then
			'msgbox "strParentKey=" & strParentKey & "::" & "strMode=" & strMode & "::" & "tvwChild=" & tvwChild & "::" & "strAcctText=" & strAcctText & "::" & "MID(Node.Key,2)=" & MID(Node.Key,2) & "::" & "Node.Text=" & Node.Text 
			Set NodX = frm1.uniTree2.Nodes.Add (strParentKey, tvwChild, StrKey & "#" & MID(Node.Key,2), Node.Text, C_Url)
		ElseIf strMode = "D" or strMode = "CD1" or strMode = "CD2" Then
'			msgbox "strParentKey=" & strParentKey & "::" & "strMode=" & strMode & "::" & "tvwChild=" & tvwChild & "::" & "strAcctText=" & strAcctText & "::" & "MID(Node.Key,2)=" & MID(Node.Key,2) & "::" & "Node.Text=" & Node.Text 
			Set NodX = frm1.uniTree2.Nodes.Add (strParentKey, tvwChild, "K" & strAcctText & MID(Node.Key,2), Node.Text, C_Folder)
		End If

		If Err.Number <> 0 Then
			intRetCD =  DIsplayMsgBox("127802", vbOKOnly, "x", "x")
			Exit Function
		End If

		Set NodP = NodX.parent
		ggoSpread.Source = .vspdData
		'.vspdData.MaxRows = .vspdData.MaxRows + 1
		.vspdData.Row = .vspdData.MaxRows
		ggoSpread.InsertRow

		If strMode = "C1" or strMode = "C2"  or strMode = "C3" Then
			SetSpreadColor .vspddata.activerow , 2
		Else
			SetSpreadColor .vspddata.activerow , .vspddata.activerow
		End If

		nodX.tag = "N" & .vspdData.ActiveRow 

		.vspdData.Col = C_LVL
		.vspdData.Text = CStr(GetNodeLvl(NodX.Parent)+1)

		strLvl = .vspdData.Text

		If strMode = "D" or left(strMode,2) = "CD" or strMode = "C3" Then
			.vspdData.Col = C_DEPT
			.vspdData.Text = strAcctText & MID(Node.Key,2)
			.vspdData.Col = C_LDEPTNM
			.vspdData.Text = Node.Text
			strTemp = split(node.tag,parent.gColSep)
			.vspdData.Col = C_SDEPTNM
			.vspdData.Text = strTemp(0)
			.vspdData.Col = C_EDEPTNM
			.vspdData.Text = strTemp(1)
			.vspdData.Col = C_ORGID
			.vspdData.Text = frm1.txtOrgId.value 
			
		End If

		.vspdData.Col = C_PDEPT

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

    If col = C_LDEPTNM Then
        frm1.vspdData.col = frm1.vspdData.maxcols
        nodekeyval = frm1.vspdData.text
        If nodekeyval = False  Then          '-->update part.error message:'키가 잘못되었습니다."
        set nodX = frm1.uniTree2.Nodes(nodekeyval)     '-->반영사항:not(IsNumeric(nodekeyval)) -> nodekeyval = False 
            frm1.vspdData.col = C_LDEPTNM
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

End Sub

'==========================================================================================
Sub  InitData()

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
		Set NodX = frm1.uniTree2.Nodes.Add(, tvwChild, C_USER_MENU_KEY, UCase(frm1.txtOrgId.value), C_ROOT, C_ROOT)
	Else
		frm1.txtOrgDt.value = ""
		Exit Function 
	End If

	frm1.hOrgId.value = Trim(frm1.txtOrgId.value)
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
'========================================================================================
Function ChkClsType()																'User Defined Function 
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strOrgId
	Dim IntRetCD

	strOrgId = Trim(frm1.txtOrgId.value)

	'class type yes/no check
	'------------------------
	strSelect = " orgid "
	strFrom   = " horg_abs"
	strWhere  = " orgid = " & FilterVar(strOrgId, "''", "S")

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
		ChkClsType= True
	Else
		ChkClsType = False
        IntRetCD = DIsplayMsgBox("124700",vbOkOnly,"X","X")
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

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow , 2

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
    	Set nodX = frm1.unitree2.nodes(frm1.vspddata.Text)
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

	With frm1		
		If .uniTree2.SelectedItem.Key = False Then	Exit Function
		If .uniTree2.SelectedItem.image = C_Folder Then Exit Function
		If .uniTree2.SelectedItem.image = C_ROOT Then Exit Function

		Set NodP = .uniTree2.SelectedItem.Parent
		Set NodX = .uniTree2.SelectedItem
		StrKey = .uniTree2.SelectedItem.Key
		StrText = .uniTree2.SelectedItem.Text

		ggoSpread.Source = .vspdData

		.vspdData.ReDraw = False

		ggoSpread.InsertRow

		.vspdData.Col = C_LVL:				.vspdData.Text = CStr(GetNodeLvl(NodP))
		.vspdData.Col = C_SEQ:				.vspdData.Text = 0 'Cstr(GetIndex(NodP))
		.vspdData.Col = C_DEPT:				.vspdData.Text = MID(NodP.key,2)
		.vspdData.Col = C_LDEPTNM:			.vspdData.Text = NodP.text

		.vspdData.ReDraw = True
	End With

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
	Dim Nodx 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

    Call SetSpreadLock()
    Call SetToolbar("1100100100011111")	
'    Call ggoOper.LockField(Document, "Q")	
	
	If Not (frm1.uniTree1.Nodes("*").Child Is Nothing) Then
		frm1.uniTree1.Nodes("*").Child.EnsureVIsible
		frm1.uniTree1.Nodes("*").Child.Selected = True
	End If

	If Not (frm1.uniTree2.Nodes("*").Child Is Nothing) Then
		frm1.uniTree2.Nodes("*").Child.EnsureVIsible
		frm1.uniTree2.Nodes("*").Child.Selected = True
	End If

	Call allCollapse_ButtonClicked

'	Call LayerShowHide(0)
	Call InitData
	frm1.vspdData.focus
	Call SetSpreadLock
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function  DbSave() 
	Dim lRow
    Dim lGrpCnt
    Dim strVal
    Dim strDel
    Dim sPDept

    DbSave = False

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
		        Case ggoSpread.InsertFlag								'☜: 신규 
					strVal = strVal & "C" & parent.gColSep						'☜: C=Create
		        Case ggoSpread.UpdateFlag								'☜: 수정 
					strVal = strVal & "U" & parent.gColSep						'☜: U=Update
			End Select
			
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag			'☜: 수정, 신규 
		            .vspdData.Col = C_ORGID	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_DEPT		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_PDEPT		'4		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            sPDept = Trim(.vspdData.Text)
		            .vspdData.Col = C_LDEPTNM		'5
		            strVal = strVal & .vspdData.Text & parent.gColSep
		            .vspdData.Col = C_BUILDID		'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_LVL		'7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            If Trim(.vspdData.Text) <> "1" and sPDept = "" Then
						Call LayerShowHide(0)
						.vspdData.Col = C_PDEPT
						.vspdData.Row = 0
						Call DisplayMsgBox("970021", "X", .vspdData.Text , "X")
						.vspdData.Focus
						.vspdData.Row = lRow
						.vspdData.Action = 0
						Exit Function
		            End If
		            .vspdData.Col = C_SEQ		'8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_SDEPTNM		'9
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_EDEPTNM		'10
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
					
		            lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag								'☜: 삭제 
					strDel = strDel & "D" & parent.gColSep						'☜: U=Update

		            .vspdData.Col = C_ORGID	'2
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_DEPT		'3
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
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

	frm1.uniTree2.MousePointer = 0

	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function  DbDelete()

End Function


'==========================================================================================
' 현재의 노드 키값으로 스프레드의 row 값을 반환 (즉 새로운 입력상태의 노드만 찾을 수 있음)
'==========================================================================================
Function SearchVspdKey(byval nodekey)
	Dim iRow

	For iRow = 1 To frm1.vspddata.maxRows
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
				.Col = C_DEPT

				If UCase(.Text) = UCase(FindValOfCol) Then
					findCrrRow = iRow
					Exit Function
				End If
			Next
		End With
	End If
End Function

Function  OrderAllocmain()
	orderseq = 0
	If frm1.unitree2.nodes.count =0 Then Exit function
	
	call OrderAlloc(frm1.unitree2.nodes("*"))
End Function

'==========================================================================================
Function  OrderAlloc(ByVal Node)
	Dim lRow
	Dim IntRetCD
	Dim strVal
	Dim sPDept

    IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")

	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call LayerShowHide(1)
	
	With frm1
		For lRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

			strVal = strVal & "U" & parent.gColSep						'☜: U=Update
		    .vspdData.Col = C_ORGID	'2
		    strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		    .vspdData.Col = C_DEPT		'3
		    strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		    .vspdData.Col = C_PDEPT		'4		            
		    strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		    sPDept = Trim(.vspdData.Text)
		    .vspdData.Col = C_LDEPTNM		'5
		    strVal = strVal & .vspdData.Text & parent.gColSep
		    .vspdData.Col = C_BUILDID		'6
		    strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		    .vspdData.Col = C_LVL		'7
		    strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		    If Trim(.vspdData.Text) <> "1" And sPDept = "" Then
				Call LayerShowHide(0)
				.vspdData.Col = C_PDEPT
				.vspdData.Row = 0
				Call DisplayMsgBox("970021", "X", .vspdData.Text , "X")
				.vspdData.Focus
				.vspdData.Row = lRow
				.vspdData.Action = 0
				Exit Function
		    End If
		    .vspdData.Col = C_SEQ		'8
		    strVal = strVal & 0 & parent.gColSep
		    .vspdData.Col = C_SDEPTNM		'9
		    strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		    .vspdData.Col = C_EDEPTNM		'10
		    strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		Next

		.txtSpread.value =  strVal 

		Call ExecMyBizASP(frm1, BIZ_ORDER_PGM_ID)
	End With
End Function

Function  CreateInternal()
	orderseq = 0
	If frm1.unitree2.nodes.count =0 Then Exit function
	
	Call CrtInternalCd()
End Function

Function CrtInternalCd()
	Dim strVal
	Dim IntRetCD
		
	If frm1.txtOrgId.value = "" Then		
		Call DisplayMsgBox("970029", "X","부서개편ID", "X")
		Exit Function
	End If

    IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")

	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call LayerShowHide(1)
	
    strVal = BIZ_BATCH_PGM_ID & "?txtMode=Gen"
	strVal = strVal & "&txtOrgId=" & frm1.txtOrgId.value

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
End Function

Function Batch_OK()
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function

Sub AllExpand(nodx)
	
	With frm1.uniTree1
		.nodes(nodx.key).expanded = True

		If .nodes(nodx.key).children > 0 Then
			allExpand(.nodes(nodx.key).child)
		End If
		
		If .nodes(nodx.key) <> .nodes(nodx.key).LastSibling Then
			allExpand(.nodes(nodx.key).next)
		Else 
			Exit Sub
		End If
	End With
End Sub

Sub AllExpand2(nodx)
	With frm1.uniTree2
		.nodes(nodx.key).expanded = True
		If .nodes(nodx.key).children > 0 Then
			AllExpand2(.nodes(nodx.key).child)
		End If
		
		If .nodes(nodx.key) <> .nodes(nodx.key).LastSibling Then
			allExpand2(.nodes(nodx.key).next)
		Else 
			Exit Sub
		End If
	End With
End Sub


Sub AllCollapse(nodx)
	With frm1.uniTree1
		.nodes(nodx.key).expanded = False
		If .nodes(nodx.key).children > 0 Then
			allCollapse(.nodes(nodx.key).child)
		End If
		
		If .nodes(nodx.key) <> .nodes(nodx.key).LastSibling Then
			allCollapse(.nodes(nodx.key).next)
		Else 
			Exit sub
		End If
	End With
End Sub

Sub allCollapse2(nodx)
	With frm1.uniTree2
		.nodes(nodx.key).expanded = False
		If .nodes(nodx.key).children > 0 Then
			allCollapse2(.nodes(nodx.key).child)
		End If
		
		If .nodes(nodx.key) <> .nodes(nodx.key).LastSibling Then
			allCollapse2(.nodes(nodx.key).next)
		Else 
			Exit Sub
		End If
	End With
End Sub

sub allExpand_ButtonClicked()
	Dim Nodx

	Set	Nodx = frm1.uniTree1.nodes("*").child
	
	Call allExpand(Nodx)

	Set Nodx = Nothing

	If lgIntFlgMode =parent.OPMD_UMODE Then
		Set	Nodx = frm1.uniTree2.nodes("*").child

		Call allExpand2(Nodx)
	End If	

End Sub

sub allCollapse_ButtonClicked()
	Dim Nodx 

	Set	Nodx = frm1.uniTree1.nodes("*").child
	Call allCollapse(NodX)
	NodX.expanded = True

	Set Nodx = Nothing

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Set	Nodx = frm1.uniTree2.nodes("*").child
		Call allCollapse2(NodX)
		NodX.expanded = True		
	End If
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gIf" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:OrderAllocmain()" >순서재생성</a> | <a href = "VBSCRIPT:CreateInternal()" >내부부서코드생성</a></TD>
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
									<TD CLASS="TD5">부서개편ID</TD>
									<TD CLASS="TD656">
										<INPUT NAME="txtOrgId" MAXLENGTH="5" SIZE=10 ALT ="부서개편ID" tag="13XXXU"><IMG SRC="../../../CShared/image/btnPopup.gIf" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrgId()">&nbsp;
										<INPUT NAME="txtOrgDt" MAXLENGTH="8" SIZE=20 ALT ="" tag="14X">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
<!--				<TR>
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
				</TR>  -->
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100% WIDTH=20% ROWSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT <%=UNI2KTV_IDVER%> id=uniTree1 width=100% height=100%> <PARAM NAME="ImageWidth" VALUE="16">  <PARAM NAME="ImageHeight" VALUE="16">  <PARAM NAME="LineStyle" VALUE="1"> <PARAM NAME="Style" VALUE="7">  <PARAM NAME="LabelEdit" VALUE="1">  </OBJECT>');</SCRIPT>
								</TD>
								<TD HEIGHT=* WIDTH=10>&nbsp;</TD>
								<TD HEIGHT=100% WIDTH=20% ROWSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT <%=UNI2KTV_IDVER%> id=uniTree2 width=100% height=100%> <PARAM NAME="ImageWidth" VALUE="16">  <PARAM NAME="ImageHeight" VALUE="16">  <PARAM NAME="LineStyle" VALUE="1"> <PARAM NAME="Style" VALUE="7">  <PARAM NAME="LabelEdit" VALUE="1">  </OBJECT>');</SCRIPT>
								</TD>
								<TD HEIGHT="100%" WIDTH=* ROWSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> height="100%" name=vspdData width="100%" tag="23" title="SPREAD" id=OBJECT1> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT=20 Valign=TOP>
		<TD  colspan=2>
			<TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
		            <TD WIDTH=10>&nbsp;</TD>
				    <TD><BUTTON NAME="btnCb_allExpand" CLASS="CLSMBTN" ONCLICK="VBScript: allExpand_ButtonClicked()">전체확장</BUTTON>&nbsp;&nbsp;
						<BUTTON NAME="btnCb_allCollapse" CLASS="CLSMBTN" ONCLICK="VBScript: allCollapse_ButtonClicked()">전체축소</BUTTON></TD>
					<TD WIDTH=* ALIGN="right"></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
				
	</TR>	
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IfRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING= no noresize framespacing=0></IfRAME></TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread			tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"				tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"			tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hOrgId"				tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hClassIndex"			tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hgroup"				tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hcode"					tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtParentDeptCd"		tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtParentDeptLvl"		tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtParentDeptSeq"		tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToParentDeptCd"		tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToParentDeptLvl"	tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtToParentDeptSeq"	tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMoveDeptCd"			tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMoveDeptLvl"		tag="21" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMoveDeptSeq"		tag="21" tabindex="-1">


</FORM>
<DIV ID="MousePT" NAME="MousePT">
<Iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></Iframe>
</DIV>
</BODY>
</HTML>

