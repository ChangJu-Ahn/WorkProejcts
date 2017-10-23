
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Sale
'*  2. Function Name        : Sales Organization
'*  3. Program ID           : b1256ma1.asp
'*  4. Program Name         : 영업조직/그룹등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/09/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seong Bae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2002/12/24 Include 성능향상 강준구 
'*                            b1254mb1호출시 argument prgramid넘겨줌 b1254ma1과 구분짓기위해...
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ====================================== -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->
<!--'==========================================  1.1.2 공통 Include   ======================================
'============================================================================================================ -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<Script Language="vbscript"	  src="../../inc/incUni2KTV.vbs"></Script>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'*****************<A HREF="\\ferrari\uniWEB\Template\inc\incUni2KTV.vbs">\\ferrari\uniWEB\Template\inc\incUni2KTV.vbs</A>*****************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'Const  tvwChild = 4

Const  C_Root = "Root"
Const  C_ORG = "ORG"
Const  C_GRP = "GRP"

Const  C_ROOT_DESC = "UNIERP"
Const  C_ROOT_KEY = "$"
Const  C_ROOT_KEY_STR = "RT_"
Const  C_UNDERSCORE = "_"

Const BIZ_MOVE_TREE = "b1256mb1.asp"										 '☆: 비지니스 로직 ASP명 
Const BIZ_SALES_GRP = "b1254mb1.asp"										 '☆: 영업그룹등록 
Const BIZ_SALES_ORG = "b1255mb1.asp"										 '☆: 영업조직등록 

Const C_IMG_Root = "../../../CShared/image/unierp.gif"
Const C_IMG_ORG = "../../../CShared/image/Orglvl_2.gif"
Const C_IMG_Open = "../../../CShared/image/Group_op.gif"
Const C_IMG_GRP = "../../../CShared/image/HumanC.gif"
Const C_IMG_None = "../../../CShared/image/c_none.gif"
Const C_IMG_Const = "../../../CShared/image/c_const.gif"

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Const   C_CTRLITEM		= 1
Const   C_CTRLITEMPB	= 2
Const   C_CTRLNM		= 3
Const	C_CTRLITEMSEQ	= 4
Const   C_DRFG			= 5
Const   C_CRFG			= 6

Const	C_CostCD = 1

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim  lgObjDragNode, lgObjDropNode, lgNewNode
Dim  lgStrPrevNodeKey
Dim  lgBlnNewNode

Dim  lgStrPrevKey1
Dim  lgStrPrevKey2

Dim  lgQueryFlag
Dim  IsOpenPop						 'Popup

Dim  lgSaveModFg
Dim  lgSelframeFlg

Dim	lglsClicked

Dim lgStrCmd					' 
Dim lgArrOrgLvl					' 영업조직 Level정보 
Dim lgIntLastOrvLvl
Dim	lgIntLastOrgLvlIndex
Dim	lgBlnRemakeNodes				' 레벨이 변경된 Tag의 레벨값을 변경하기위한 재쿼리 여부(하위 조직이 존재하는 경우 재쿼리)
Dim	lgBlnLvlChanged
Dim lgBlnOpenPopup
Dim lgBlnOrgLvlExists
 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub  InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgSortKey = 1
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count    
End Sub

'========================================================================================================
' Function Name : InitTree
' Function Desc : This method initializes tree object
'========================================================================================================
Sub InitTree()
    With frm1.uniTree1
		.HideSelection = false
		.SetAddImageCount = 6
		.Indentation = "200"	' 줄 간격 
						' 파일위치,	키명, 위치 
		.AddImage C_IMG_Root,		C_Root,		0
		.AddImage C_IMG_ORG,		C_ORG,		0
		.AddImage C_IMG_Open,		C_Open,		0
		.AddImage C_IMG_GRP,		C_GRP,		0
		.AddImage C_IMG_None,		C_None,		0
		.AddImage C_IMG_Const,		C_Const,	0
	
		.PathSeparator = parent.gColSep
		
		.OLEDragMode = 1														'⊙: Drag & Drop 을 가능하게 할 것인가 정의 
		.OLEDropMode = 1	
	
		.OpenTitle = "영업조직입력"											
		.AddTitle = "영업그룹입력"		
		.RenameTitle = ""
		.DeleteTitle = "삭제"
	End With
End Sub		

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ===================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub  SetDefaultVal()
	Call GetSalesOrgLvlInfo()
	lgBlnOpenPopup = False
End Sub

'==========================================  2.2.2 SetDefaultScreen()  ===================================
'	Name : SetDefaultScreen()
'	Description : Default Screen을 설정한다.
'========================================================================================================= 
Sub SetDefaultScreen()
	ClickTab1()
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call ggoOper.ClearField(Document, "1")                                  '⊙: Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  '⊙: Clear Contents  Field
	Call ggoOper.ClearField(Document, "3")
	With frm1
		Call ggoOper.SetReqAttr(.txtSalesOrg, "Q")
		Call ggoOper.SetReqAttr(.txtSalesOrgnm, "Q")
		Call ggoOper.SetReqAttr(.txtSalesOrgFullnm, "Q")
		Call ggoOper.SetReqAttr(.txtSalesOrgEngnm, "Q")
		Call ggoOper.SetReqAttr(.txtHeadusrnm, "Q")
		Call ggoOper.SetReqAttr(.txtSalesOrgnm, "Q")
		Call ggoOper.SetReqAttr(.rdoORgUsageflagN, "Q")
		Call ggoOper.SetReqAttr(.rdoORgUsageflagY, "Q")
		Call ggoOper.SetReqAttr(.rdoEndOrgFlagY, "Q")
		Call ggoOper.SetReqAttr(.rdoEndOrgFlagN, "Q")
	End With
End Sub
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다.
'********************************************************************************************************* 
 '==========================================  2.3.1 Tab Click 처리  =================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=================================================================================================================== 
 '----------------  ClickTab1(): Header Tab처리 부분 (Header Tab이 있는 경우만 사용)  ---------------------------- 
Function ClickTab1()
	
	If lgSelframeFlg = TAB1 Then Exit Function
	 
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 	
	lgSelframeFlg = TAB1

End Function

Function ClickTab2()

	If lgSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab 
	lgSelframeFlg = TAB2

End Function

 '******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
'------------------------------------------  OpenConPopup()  -------------------------------------------------
'	Name : OpenSheetPopup()
'	Description : Sales Org Display PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSheetPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenSheetPopup = False
	
	If lgBlnOpenPopup = True Then Exit Function

	lgBlnOpenPopup = True

	Select Case pvIntWhere

	Case C_CostCd												
		iArrParam(1) = "dbo.b_cost_center"				<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtCostCenter.value)	<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = ""								<%' Where Condition%>
		iArrParam(5) = frm1.txtCostCenter.alt 			<%' TextBox 명칭 %>
			
		iArrField(0) = "ED15" & parent.gColSep & "cost_cd"	<%' Field명(0)%>
		iArrField(1) = "ED30" & parent.gColSep & "cost_nm"	<%' Field명(1)%>
		    
		iArrHeader(0) = "비용집계처"					<%' Header명(0)%>
		iArrHeader(1) = "비용집계처명"					<%' Header명(1)%>
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPopup = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSheetPopup = SetSheetPopup(iArrRet,pvIntWhere)
	End If	
	
End Function

 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetRefOpenAp()  --------------------------------------------------
'	Name : SetSheetPopup()
'	Description : OpenSheetPopup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSheetPopup(Byval pvArrRet, Byval pvIntWhere)

	SetSheetPopup = False
	
	With frm1
		Select Case pvIntWhere
		Case C_CostCd
			.txtCostCenter.value = pvArrRet(0)
			.txtCostCenterNm.value = pvArrRet(1)
		End Select	
    End With
    
    SetSheetPopup = True
    lgBlnFlgChgValue = True
End Function

'==========================================================================================
'   Function Name :ChkDragState
'   Function Desc :Drag 가 어디에 있는지 Drag되는 항목인지 체크 
'==========================================================================================
Function  ChkDragState(ByVal x , ByVal y )
    
    Dim iObjNewNode
    dim ChildNode
    Dim iArrNewNodeTag, iArrDragNodeTag
    Dim iIntIndex, iIntCurOrgIndex
    
    On Error Resume Next
    
    ChkDragState = False

    If lgObjDragNode.parent Is Nothing Then Exit Function	' 자신이 Root인 경우 
    
    Set iObjNewNode = frm1.uniTree1.HitTest(x, y)

    ' 폴더가 지정되지 않은 경우 
    If iObjNewNode Is Nothing Then Exit Function
    
    ' 트리내의 존재여부 Check
	If Not ChkOrgTree(iObjNewNode, C_ROOT_KEY) Then Exit Function

    iArrNewNodeTag = Split(iObjNewNode.Tag, parent.gColSep)
    iArrDragNodeTag = Split(lgObjDragNode.Tag, parent.gColSep)

	' Drag된 Node가 영업 조직인 경우에는 말단조직에는 Drop될 수 없다.
	If lgObjDragNode.Image = C_GRP Then
		' 영업그룹은 말단조직에만 종속될 수 있다.
		If iObjNewNode.Key = C_ROOT_KEY OR iArrNewNodeTag(1) = "N" Then Exit Function
		
	Else
    	If iObjNewNode.Key = C_ROOT_KEY Then
			iIntCurOrgIndex = 0
		Else
			' 하위 레벨로 이동하는 경우 새 조직이 자신의 하위 조직인지 check
			If iArrNewNodeTag(0) > iArrDragNodeTag(0) Then
				' 트리내의 존재여부 Check
				If ChkOrgTree(iObjNewNode, lgObjDragNode.Key) Then Exit Function
			End If
			
			' 말단 조직 아래에는 영업조직이 올 수 없다.
			If iArrNewNodeTag(1) = "Y" Then Exit Function
		
			For iIntIndex = 0 to lgIntLastOrgLvlIndex - 1
				If lgArrOrgLvl(iIntIndex, 0) = iArrNewNodeTag(0) then
					iIntCurOrgIndex = iIntIndex + 1
					Exit For
				End If
			Next
		End If

		' 조직레벨의 최대값 Check
		If iIntCurOrgIndex + GetSubOrgLvlCnt(iArrDragNodeTag(0), Mid(lgObjDragNode.Key,2)) > lgIntLastOrgLvlIndex Then Exit Function
	End If
	
    '자신의 자리에 있을때 
    If iObjNewNode.Text = lgObjDragNode.parent.Text Then Exit Function
    
    ' 자신의 부모에게 갈때 
    If iObjNewNode.Key = lgObjDragNode.Key Then Exit Function
    
    ' 영업그룹에 Drop된 경우 
    If iObjNewNode.Image = C_GRP Then Exit Function
    
    ChkDragState = True
    
End Function

' 특정 트리(pvStrFind)내에 존재하는 check하는 재귀함수 
Function ChkOrgTree(prObjParentNode, prStrFind)
	Dim blnFind
	
	blnFind = False
	
	ChkOrgTree = blnFind
	
	If prObjParentNode is Nothing Then Exit Function
	
	If prObjParentNode.Key <> prStrFind Then
		blnFind = ChkOrgTree(prObjParentNode.Parent, prStrFind)
	Else
		blnFind = True
	End If
	
	ChkOrgTree = blnFind
End Function

'==========================================================================================
'   Function Name : GetSubOrgLvlCnt
'   Function Desc : 현재 노드가 포함하고 있는 영업조직레벨 수를 구한다.
'==========================================================================================

Function  GetSubOrgLvlCnt(ByVal pvIntOrgLvl, ByVal pvStrOrg)
    On Error Resume Next
    
    Dim iStrSelect, iStrFrom, iStrWhere, iStrResult
    Dim iArrRow, iArrCol

	iStrSelect	= " ISNULL(MAX(lvl), 0) + 1 "
	iStrFrom	= " dbo.ufn_s_ListSalesOrgHierarchy(" & pvIntOrgLvl & ",  " & FilterVar(pvStrOrg, "''", "S") & ", Default)"
	iStrWhere	= ""
		
	If CommonQueryRs2by2(iStrSelect, iStrFrom ,  iStrWhere , iStrResult) Then 
	
		iArrRow = Split(iStrResult, parent.gColSep & parent.gRowSep)
		iArrCol	= Split(iArrRow(0), parent.gColSep)
		
		GetSubOrgLvlCnt = CInt(iArrCol(1))
	End If
	
End Function

'==========================================================================================
'   Function Name :GetTotalCnt
'   Function Desc :Add에 관련되 자식수를 되돌려준다.
'==========================================================================================

Function GetTotalCnt(prObjNode)
	
	If prObjNode.children = 0 Then	' Root일 경우 
		GetTotalCnt = 1
	Else
		GetTotalCnt = prObjNode.children + 1
	End If
	
End Function


'======================================================================================================
'	화면 사이즈 변경 
'======================================================================================================

Sub DispDivConf(pVal) 
	if pVal = 2 then
		divconf.style.display = "none"
		tdConf.height = 1
	else
		divconf.style.display = ""
		tdConf.height = 22
	end if
End Sub

 '#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub  Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call AppendNumberPlace("7","3","0")
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "3", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
                             
    Call InitVariables                                                      '⊙: Initializes local global variables
	Call SetDefaultVal()
		
    '----------  Coding part  -------------------------------------------------------------
    Call InitTree()

	Set lgObjDragNode = Nothing

	lglsClicked = False
End Sub


'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

 '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

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

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Sub rdoEndOrgFlagN_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoEndOrgFlagY_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoORgUsageflagN_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoORgUsageflagY_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoGrpUsageflagN_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoGrpUsageflagY_OnClick()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node를 클릭하면 발생 이벤트 
'==========================================================================================

Sub uniTree1_NodeClick(pvObjNode)
	On Error Resume Next
	Dim Response
	Dim iBlnProtect
	
	' 트리 조회시에 클릭을 하면 조회가 되지 않도록 조치 
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If lgBlnNewNode = TRUE Then
		If pvObjNode.Key <> lgNewNode.Key Then
			Response = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")

			If Response = vbYes Then
				frm1.uniTree1.Nodes.Remove lgNewNode.Index
				frm1.uniTree1.SetFocus
				lgBlnFlgChgValue = False
				lgBlnNewNode = False
				lgSaveModFg = ""
				Set lgNewNode = Nothing
				Call FncNew()
			Else
				frm1.uniTree1.SetFocus
				lgNewNode.Selected = True
				Exit Sub
			End If
		Else
			Exit Sub			
		End If			
	End If

	If pvObjNode.Key <> lgStrPrevNodeKey And lgBlnFlgChgValue Then
		Response = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If Response = vbNo Then
			frm1.uniTree1.Nodes(lgStrPrevNodeKey).Selected = True
			frm1.uniTree1.SetFocus
			Exit Sub
		End if
	End if

	lgStrPrevNodeKey = pvObjNode.Key
	  
	IF pvObjNode.Key = C_ROOT_KEY Then
		Call SetDefaultScreen()
		Exit Sub
	End If
	  
	Dim iStrSelect, iStrFrom, iStrWhere
	Dim iArrCol, iArrRow

	iBlnProtect = False
		
	Call LayerShowHide(1)
	Call SetToolbar("1100100000001111")									'⊙: 버튼 툴바 제어					 
	
	' Get the Sales Org. Info.
	If pvObjNode.Image = C_ORG Then
		iStrSelect	= " SO1.sales_org, SO1.sales_org_nm, SO1.sales_org_full_nm, SO1.sales_org_eng_nm, SO1.head_usr_nm, SO1.lvl, SO1.end_org_flag, SO1.usage_flag, SO1.upper_sales_org, SO2.sales_org_nm "
		iStrFrom	= " dbo.b_sales_org SO1 LEFT OUTER JOIN dbo.b_sales_org SO2 ON (SO2.sales_org = SO1.upper_sales_org) "
		iStrWhere	= " SO1.sales_org =  " & FilterVar(Mid (pvObjNode.key,2), "''", "S") & " "
		
		ClickTab1()
		lgStrCmd = "ORG"		
	Else
		iStrSelect	= " SG.sales_grp, SG.sales_grp_nm, SG.sales_grp_full_nm, SG.sales_grp_eng_nm, SG.usage_flag, SG.sales_org, SO.sales_org_nm, SG.cost_cd, CC.cost_nm "
		iStrFrom	= " dbo.b_sales_grp SG INNER JOIN dbo.b_sales_org SO ON (SO.sales_org = SG.sales_org) INNER JOIN dbo.b_cost_center CC ON (CC.cost_cd = SG.cost_cd) "
		iStrWhere	= " SG.sales_grp =  " & FilterVar(Mid (pvObjNode.key,2), "''", "S") & " "
		
		ClickTab2()
		lgStrCmd = "GRP"
	End If
	 
	If CommonQueryRs2by2(iStrSelect, iStrFrom ,  iStrWhere , lgF2By2) Then 
	
		iArrRow = Split(lgF2By2, parent.gColSep & parent.gRowSep)
		iArrCol	= Split(iArrRow(0), parent.gColSep)			

		With frm1
			If pvObjNode.Image = C_ORG Then
				.txtSalesOrg.value = iArrCol(1)
				.txtSalesOrgnm.value = iArrCol(2)
				.txtSalesOrgFullnm.value = iArrCol(3)
				.txtSalesOrgEngnm.value = iArrCol(4)
				.txtHeadusrnm.value = iArrCol(5)
				.txtSalesOrgLvl.value = iArrCol(6)
				If iArrCol(7) = "Y" Then
					.rdoEndOrgFlagY.checked = True
				Else
					.rdoEndOrgFlagN.checked = True
				End If
			
				If iArrCol(8) = "Y" Then
					.rdoOrgUsageflagY.checked = True
				Else
					.rdoOrgUsageflagN.checked = True
				End If
				.txtUpperSalesOrg.value = iArrCol(9)
				.txtUpperSalesOrgNm.value = iArrCol(10)
				
				' if Last level, you cannot edit 'End Org. Flag'
				If lgArrOrgLvl(lgIntLastOrgLvlIndex - 1, 0) = iArrCol(6) Then
					iBlnProtect = True
				Else
					IF pvObjNode.Children > 0 THEN
						' If it has sales group as child node, you cannot edit 'End org. flag'
						If pvObjNode.Child.Image = C_GRP Then
							iBlnProtect = True
						End If
					end if 
				End If
			Else
				.txtSalesGrp.value = iArrCol(1)
				.txtSalesGrpnm.value = IArrCol(2)
				.txtSalesGrpFullnm.value = iArrCol(3)
				.txtSalesGrpEngnm.value = iArrCol(4)
				If iArrCol(5) = "Y" Then
					.rdoGrpUsageflagY.checked = True
				Else
					.rdoGrpUsageflagN.checked = True
				End If
				.txtSalesOrgInGrp.value = iArrCol(6)
				.txtSalesOrgNmInGrp.value = iArrCol(7)
				.txtCostCenter.value = iArrcol(8)
				.txtCostCenterNm.value = iArrCol(9)
			End If
			
		End With
	Else
		If lgStrCmd = "ORG" Then
			IntRetCD = DisplayMsgBox("125500","X","X","X")	' 영업조직정보가 존재하지 않습니다.
		Else
			IntRetCD = DisplayMsgBox("125400","X","X","X")	' 영업그룹정보가 존재하지 않습니다.
		End If
	End if 

    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    ' End Org Flag Protect 처리 
    If pvObjNode.Image = C_ORG And iBlnProtect Then
		Call ggoOper.SetReqAttr(frm1.rdoEndOrgFlagY, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoEndOrgFlagN, "Q")
    End If
	Call LayerShowHide(0)
	lgBlnFlgChgValue = False
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode    
End Sub

'==========================================================================================
'   Event Name : uniTree1_OLEDragDrop
'   Event Desc : Node를 Drag & Drop 이벤트 
'==========================================================================================

Sub  uniTree1_OLEDragDrop(Data , Effect , Button , Shift , x , y )

	Dim IntRetCD
    Dim iStrVal
    Dim iArrIndex, iArrTag

	If lgObjDragNode Is Nothing Then Exit Sub

	' mscomctl.ocx 버전업후 클릭시OLEDragDrop Event가 발생하여 
	' "해당 위치로는 이동할 수 없습니다!" 에러 발생 
	Dim iobjNewNode
	
    Set iObjNewNode = frm1.uniTree1.HitTest(x, y)

    If Not iObjNewNode is nothing then
        If iObjNewNode.key = lgObjDragNode.key Then
            Exit sub
        End If  
    End if

    Set iObjNewNode = Nothing

	If ChkDragState(x, y) = False Then
        Effect = vbDropEffectNone
		IntRetCD = DisplayMsgBox("990017","X","X","X")	' 해당 위치로는 이동할 수 없습니다!
        Exit Sub
	End If

	' 레벨이 변경된 Tag의 레벨값을 변경하기위한 재쿼리 여부 
	lgBlnRemakeNodes = False
	
	Call LayerShowHide(1)

	frm1.uniTree1.MousePointer = 11

    Set lgObjDropNode = frm1.uniTree1.HitTest(x, y)					' 이동해야될 노드를 기억시킴 
 
	' 영업그룹이 이동된 경우는 b_sales_grp.sales_org만 변경하면 된다.	
	If lgObjDragNode.Image = C_GRP Then
		lgStrCmd  = "GRP"
		lgBlnLvlChanged = False

		iStrVal = BIZ_MOVE_TREE & "?txtMode=" & parent.UID_M0002
		iStrVal = iStrVal & "&txtFlag="		& "GRP"							' Sales Group
		iStrVal = iStrVal & "&txtSalesGrp=" & Mid(lgObjDragNode.key, 2)		' Sales Group
		iStrVal = iStrVal & "&txtSalesOrg=" & Mid(lgObjDropNode.key, 2)		' Sales Org.
		iStrVal = iStrVal & "&txtUserId="	& parent.gUsrID
	ELSE
		iStrVal = BIZ_MOVE_TREE & "?txtMode=" & parent.UID_M0002

		' 조직 레벨의 변경 여부 check
		lgBlnLvlChanged = True
		If lgObjDropNode.Key <> C_ROOT_KEY And lgObjDragNode.parent.Key <> C_ROOT_KEY THEN
			If lgObjDropNode.parent.fullpath = lgObjDragNode.parent.parent.fullpath Then
				lgBlnLvlChanged = False
				iStrVal = iStrVal & "&txtFlag="		& "ORG1"								' Sales Org.
			End If
		End If

		If lgBlnLvlChanged Then
			' 새로운 레벨 설정 
			iArrIndex = Split(lgObjDropNode.fullpath, parent.gColSep)
			iStrVal = iStrVal & "&txtSalesOrgNewLvl=" & lgArrOrgLvl(Ubound(iArrIndex, 1), 0)	' Sales Org. New Level
			If Ubound(iArrIndex, 1) = lgIntLastOrgLvlIndex - 1 Then
				iStrVal = iStrVal & "&txtEndOrgFlag=Y"
			Else
				iStrVal = iStrVal & "&txtEndOrgFlag=N"
			End If
			
			iArrTag = Split(lgObjDragNode.Tag, parent.gColSep)
			iStrVal = iStrVal & "&txtSalesOrgCurLvl=" & iArrTag(0)								' Sales Org. Current Level

			' 하위조직의 존재여부 Check
			If lgObjDragNode.Children = 0 Then
				iStrVal = iStrVal & "&txtFlag="	& "ORG2"								' Sales Org.
			Else
				'말단조직여부 
				IF lgObjDragNode.Child.Image = C_GRP Then
					iStrVal = iStrVal & "&txtFlag="	& "ORG2"							' Sales Org.
				Else
					lgBlnRemakeNodes = True
					iStrVal = iStrVal & "&txtFlag="	& "ORG3"							' Sales Org.
				End If
			End If
		End If
		
		iStrVal = iStrVal & "&txtSalesOrg=" & Mid(lgObjDragNode.key, 2)			' Sales Org.
		
		If lgObjDropNode.Key = C_ROOT_KEY Then
			iStrVal = iStrVal & "&txtUpperSalesOrg="								' Upper Sales Org.
		Else
			iStrVal = iStrVal & "&txtUpperSalesOrg=" & Mid(lgObjDropNode.key, 2)	' Upper Sales Org.
		End If
		iStrVal = iStrVal & "&txtUserId="	& parent.gUsrID
		lgStrCmd = "ORG"
	END IF

	Call LayerShowHide(0)
	frm1.uniTree1.MousePointer = 0
	
	lgSaveModFg = "R"
	
	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '☜: 비지니스 ASP 를 가동 

End Sub

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
		Set lgObjDragNode = frm1.uniTree1.SelectedItem
		lgObjDragNode.Selected = True
	Else
		Set lgObjDragNode = Nothing
	End If

	lglsClicked = False		

End Sub

'==========================================================================================
'   Event Name : uniTree1_MouseUp
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================

Sub  uniTree1_MouseUp(pvObjNode, Button , Shift, X, Y )
	Dim iArrTag
	
	With frm1
		If Button = 2 Or Button = 3 Then
			If pvObjNode.Key = C_ROOT_KEY Then
				.uniTree1.MenuEnabled C_MNU_OPEN, True
				.uniTree1.MenuEnabled C_MNU_ADD, FALSE
				.uniTree1.MenuEnabled C_MNU_DELETE, False
				.uniTree1.MenuEnabled C_MNU_RENAME, False
				frm1.uniTree1.PopupMenu 
				Exit Sub
			End If
			
			' Mouse Pointer가 트리에 위치하는지 Check
			If ChkOrgTree(pvObjNode, C_ROOT_KEY) = False Then
				Select Case pvObjNode.Image
					Case C_GRP, C_ORG, C_Const
						.uniTree1.MenuEnabled C_MNU_OPEN, False
					Case Else
						.uniTree1.MenuEnabled C_MNU_OPEN, False
				End Select
				
				.uniTree1.MenuEnabled C_MNU_ADD, False
				.uniTree1.MenuEnabled C_MNU_DELETE, False
				.uniTree1.MenuEnabled C_MNU_RENAME, False
			Else
				.uniTree1.MenuEnabled C_MNU_DELETE, True
				
				' 만약 새로운 입력할 노드에서 popup 할 때는 입력메뉴들이 보이면 안된다.
				If lgBlnNewNode = TRUE Then
					if pvObjNode.Key = lgNewNode.key then
						
						.uniTree1.MenuEnabled C_MNU_OPEN,False					
						.uniTree1.MenuEnabled C_MNU_ADD, False
						.uniTree1.MenuEnabled C_MNU_RENAME, False
					end if
				Else
					Select Case pvObjNode.Image
						Case C_GRP
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
						Case C_ORG
							iArrTag = Split(pvObjNode.tag, parent.gColSep)

							' 말단조직인 경우 
							If iArrTag(1) = "N" Then
								.uniTree1.MenuEnabled C_MNU_OPEN,True
								.uniTree1.MenuEnabled C_MNU_ADD, False
							Else
								.uniTree1.MenuEnabled C_MNU_OPEN,False
								.uniTree1.MenuEnabled C_MNU_ADD, True
							End If
							.uniTree1.MenuEnabled C_MNU_RENAME, False
					End Select
				End if
			End If
			frm1.uniTree1.PopupMenu
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : uniTree1_MenuOpen - 영업조직입력 
'   Event Desc : Node Popup
'==========================================================================================

Sub  uniTree1_MenuOpen(pvObjNode)
	on Error Resume next

	Dim iObjDummyNode		
	Dim iArrTag
	Dim ii
	
	call FncNew
	
	If pvObjNode.Expanded = False Then
		pvObjNode.Expanded = True
	End If

	If pvObjNode.Key = C_ROOT_KEY Then	' Root일 경우 
		Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, C_ROOT_KEY_STR & GetTotalCnt(pvObjNode), "새 영업조직", C_ORG, C_ORG)

		With frm1
		
			.txtSalesOrglvl.value = lgArrOrgLvl(0, 0)
			' 조직레벨이 하나인 경우 말단조직여부를 'Y'로 설정 
			If lgIntLastOrgLvlIndex = 1 Then
				.rdoEndOrgFlagY.checked = True
				Call ggoOper.SetReqAttr(.rdoEndOrgFlagY, "Q")
				Call ggoOper.SetReqAttr(.rdoEndOrgFlagN, "Q")
			Else
				.rdoEndOrgFlagN.checked = True
				Call ggoOper.SetReqAttr(.rdoEndOrgFlagY, "N")
				Call ggoOper.SetReqAttr(.rdoEndOrgFlagN, "N")
			End If
		End With
	Else
		Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, pvObjNode.Key & C_UNDERSCORE & GetTotalCnt(pvObjNode), "새 영업조직", C_ORG, C_ORG)

		With frm1
			.txtUpperSalesOrg.value = Mid(pvObjNode.Key,2)
			iArrTag = Split(pvObjNode.tag, parent.gColSep)
			For ii = 0 to lgIntLastOrgLvlIndex - 1
				If lgArrOrgLvl(ii, 0) = iArrTag(0) then
					.txtSalesOrglvl.value = lgArrOrgLvl(ii + 1, 0)

					If (ii + 1) = (lgIntLastOrgLvlIndex - 1)Then
						.rdoEndOrgFlagY.checked = True
						Call ggoOper.SetReqAttr(.rdoEndOrgFlagY, "Q")
						Call ggoOper.SetReqAttr(.rdoEndOrgFlagN, "Q")
					Else
						.rdoEndOrgFlagN.checked = True
						Call ggoOper.SetReqAttr(.rdoEndOrgFlagY, "N")
						Call ggoOper.SetReqAttr(.rdoEndOrgFlagN, "N")
					End If
					
					Exit For
				End If
			Next
		End With
	End If
	
	iObjDummyNode.Selected = True	
	Set lgNewNode = iObjDummyNode
	set lgObjDragNode = iObjDummyNode
	
	Call ClickTab1()

	Call SetToolbar("1100100000001111")									'⊙: 버튼 툴바 제어		
			
	lgIntFlgMode = parent.OPMD_CMODE	' 신규로 등록 
	
	lgStrCmd  = "ORG"
	
	lgBlnFlgChgValue = TRUE
	lgBlnNewNode = TRUE
	lgSaveModFg	= "O"	
End Sub


'==========================================================================================
'   Event Name : uniTree1_MenuAdd - 영업그룹등입력 
'   Event Desc : Node Popup
'==========================================================================================
Sub  uniTree1_MenuAdd(pvObjNode)

	Dim iObjDummyNode
		
	'If ChkOrgTree(Node, C_ROOT_KEY) = TRUE Then Exit Sub
	CALL FNCNEW
	
	If pvObjNode.Expanded = False Then
		pvObjNode.Expanded = True
	End If
	
	Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, pvObjNode.Key & C_UNDERSCORE & GetTotalCnt(pvObjNode), "새 영업그룹", C_GRP, C_GRP)
	
	iObjDummyNode.Selected = True
	Set lgNewNode = iObjDummyNode
	set lgObjDragNode = iObjDummyNode	
	
	Call SetToolbar("1100100000001111")									'⊙: 버튼 툴바 제어		
	 
	Call ClickTab2()

	frm1.txtSalesOrgInGrp.value = Mid(pvObjNode.Key,2)
		
	lgIntFlgMode = parent.OPMD_CMODE	' 신규로 등록 
	lgStrCmd  = "GRP"
		
	lgBlnFlgChgValue = TRUE
	lgBlnNewNode = TRUE
	lgSaveModFg	= "G"	
End Sub

'==========================================================================================
'   Event Name : DisplayNodes
'   Event Desc : 
'==========================================================================================

Sub DisplayNodes()
		
	Dim iObjDummyNode
	Dim iStrSelect, iStrFrom, iStrWhere 	
	Dim iStrSalesOrgCd, iStrSalesOrgNm
	Dim iStrNode, iStrImg
	Dim ii, jj
	Dim iArrRow, iArrCol

	On Error Resume Next

	frm1.uniTree1.MousePointer = 11
	
	Call LayerShowHide(1)

	frm1.uniTree1.Nodes.Clear 
	
	' Add the top level(uniERP)
	Set iObjDummyNode = frm1.uniTree1.Nodes.Add(, tvwChild, C_ROOT_KEY, C_ROOT_DESC, C_Root, C_Root)

	iStrSelect	= " CASE WHEN upper_sales_org IS NULL THEN  " & FilterVar(C_ROOT_KEY, "''", "S") & " ELSE " & FilterVar("O", "''", "S") & "  + upper_sales_org END , " & FilterVar("O", "''", "S") & "  + sales_org, " & FilterVar("[", "''", "S") & "  + sales_org + " & FilterVar("]", "''", "S") & " + sales_org_nm, lvl, end_org_flag,  " & FilterVar(C_ORG, "''", "S") & ""
	iStrFrom	= " dbo.b_sales_org "
	iStrFrom	= iStrFrom & " UNION ALL "
	iStrFrom	= iStrFrom & " SELECT " & FilterVar("O", "''", "S") & "  + SG.sales_org, " & FilterVar("G", "''", "S") & "  + SG.sales_grp, " & FilterVar("[", "''", "S") & "  + SG.sales_grp + " & FilterVar("]", "''", "S") & " + SG.sales_grp_nm, SO.lvl + 1, " & FilterVar("N", "''", "S") & " ,  " & FilterVar(C_GRP, "''", "S") & ""
	iStrFrom	= iStrFrom & " FROM dbo.b_sales_grp SG INNER JOIN dbo.b_sales_org SO ON (SO.sales_org = SG.sales_org) "
	iStrFrom	= iStrFrom & " ORDER BY lvl, 2 "
	iStrWhere	= ""
	
	If CommonQueryRs2by2(iStrSelect, iStrFrom ,  iStrWhere , lgF2By2) Then 
	
		iArrRow = Split(lgF2By2, parent.gColSep & parent.gRowSep)			
		jj = Ubound(iArrRow,1)
		
		For ii = 0 To jj - 1		
			iArrCol			= Split(iArrRow(ii), parent.gColSep)			
			
			iStrNode		= Trim(iArrCol(1))
			iStrSalesOrgCd	= Trim(iArrCol(2))
			iStrSalesOrgNm	= Trim(iArrCol(3))
			iStrImg			= Trim(iArrCol(6))

			Set iObjDummyNode = frm1.uniTree1.Nodes.Add (iStrNode, tvwChild, iStrSalesOrgCd, iStrSalesOrgNm, iStrImg )
			' Org Level, End org Flag
			frm1.uniTree1.Nodes(iStrSalesOrgCd).Tag = Trim(iArrCol(4)) & parent.gColSep & Trim(iArrCol(5))
		Next
	End if 
	Call LayerShowHide(0)
	frm1.uniTree1.MousePointer = 0

	If Not(frm1.uniTree1.Nodes(C_ROOT_KEY).Child Is Nothing) Then
		frm1.uniTree1.Nodes(C_ROOT_KEY).Child.EnsureVisible						' Expand Tree	
	End If
	frm1.uniTree1.Nodes(C_ROOT_KEY).Selected = True
End sub

'==========================================================================================
'   Event Name : RemakeNodes
'   Event Desc : 레벨이 변경되고 하위 조직이 존재하는 경우 관련 Nodes를 재성성한다.
'==========================================================================================
Sub RemakeNodes()
		
	Dim iObjDummyNode
	Dim iStrSelect, iStrFrom, iStrWhere, iStrResult 	
	Dim ii, jj
	Dim iArrRow, iArrCol, iArrTag

	On Error Resume Next
	iArrTag = Split(lgObjDragNode.Tag, parent.gColSep)

	iStrSelect	= " CASE WHEN SO.sales_org =  " & FilterVar(Mid(lgObjDragNode.Key, 2), "''", "S") & " THEN  " & FilterVar(lgObjDropNode.Key, "''", "S") & " ELSE " & FilterVar("O", "''", "S") & "  + SO.upper_sales_org END , " & FilterVar("O", "''", "S") & "  + SO.sales_org, " & FilterVar("[", "''", "S") & "  + SO.sales_org + " & FilterVar("]", "''", "S") & " + SO.sales_org_nm, SO.lvl, SO.end_org_flag,  " & FilterVar(C_ORG, "''", "S") & " "
	iStrFrom	= " dbo.b_sales_org SO INNER JOIN  "
	iStrFrom	= iStrFrom & " (SELECT	 " & FilterVar(Mid(lgObjDragNode.Key, 2), "''", "S") & " AS sales_org "
	iStrFrom	= iStrFrom & " UNION ALL "
	iStrFrom	= iStrFrom & " SELECT leaf_org "
	iStrFrom	= iStrFrom & " FROM dbo.ufn_s_ListSalesOrgHierarchy(" & iArrTag(0) & ",  " & FilterVar(Mid(lgObjDragNode.Key, 2), "''", "S") & ",  default)) T ON (T.sales_org = SO.sales_org) "
	iStrFrom	= iStrFrom & " UNION ALL "
	iStrFrom	= iStrFrom & " SELECT " & FilterVar("O", "''", "S") & "  + SG.sales_org, " & FilterVar("G", "''", "S") & "  + SG.sales_grp, " & FilterVar("[", "''", "S") & "  + SG.sales_grp + " & FilterVar("]", "''", "S") & " + SG.sales_grp_nm, SO.lvl + 1, " & FilterVar("N", "''", "S") & " ,  " & FilterVar(C_GRP, "''", "S") & ""
	iStrFrom	= iStrFrom & " FROM dbo.b_sales_grp SG INNER JOIN dbo.b_sales_org SO ON (SO.sales_org = SG.sales_org) "
	iStrFrom	= iStrFrom & " INNER JOIN "
	iStrFrom	= iStrFrom & " (SELECT  leaf_org as sales_org "
	iStrFrom	= iStrFrom & " FROM dbo.ufn_s_ListSalesOrgHierarchy(" & iArrTag(0) & ",  " & FilterVar(Mid(lgObjDragNode.Key, 2), "''", "S") & ",  default) "
	iStrFrom	= iStrFrom & " WHERE end_org_flag = " & FilterVar("Y", "''", "S") & " ) T ON (T.sales_org = SG.sales_org) "	
	iStrFrom	= iStrFrom & " ORDER BY lvl, 2 "
	iStrWhere	= ""
	
	If CommonQueryRs2by2(iStrSelect, iStrFrom ,  iStrWhere , iStrResult) Then 
	
		iArrRow = Split(iStrResult, parent.gColSep & parent.gRowSep)			
		jj = Ubound(iArrRow,1)
		
		For ii = 0 To jj - 1		
			iArrCol			= Split(iArrRow(ii), parent.gColSep)			
			
			Set iObjDummyNode = frm1.uniTree1.Nodes.Add (Trim(iArrCol(1)), tvwChild, Trim(iArrCol(2)), Trim(iArrCol(3)), Trim(iArrCol(6)) )
			
			' Org Level, End org Flag
			frm1.uniTree1.Nodes(Trim(iArrCol(2))).Tag = Trim(iArrCol(4)) & parent.gColSep & Trim(iArrCol(5))
		Next

		frm1.uniTree1.Nodes(lgObjDragNode.Key).Selected = True
	Else
		If Err.number <> 0 Then	Msgbox Err.Description
	End if 

End sub

'==========================================================================================
'   Event Name : Get Sales Org. level Info.
'   Event Desc : 
'==========================================================================================

Sub GetSalesOrgLvlInfo()
		
	Dim iStrSelect, iStrFrom, iStrWhere, iStrResult 	
	Dim ii, iIntRows
	Dim iArrRow, iArrCol
	
	iStrSelect	= " MI.minor_cd, IsNull(CF.reference , " & FilterVar("N", "''", "S") & " ) "
	iStrFrom	= " dbo.b_minor MI LEFT OUTER JOIN dbo.b_configuration CF ON (CF.major_cd = MI.major_cd AND CF.minor_cd = MI.minor_cd) "
	iStrWhere	= " MI.major_cd = " & FilterVar("S0016", "''", "S") & " "
	
	If CommonQueryRs2by2(iStrSelect, iStrFrom ,  iStrWhere , iStrResult) Then 
	
		iArrRow = Split(iStrResult, parent.gColSep & parent.gRowSep)			
		iIntRows = Ubound(iArrRow,1)
		
		Redim lgArrOrgLvl(iIntRows, 1)
		
		For ii = 0 To iIntRows - 1		
			iArrCol	= Split(iArrRow(ii), parent.gColSep)			
			
			lgArrOrgLvl(ii, 0) = Trim(iArrCol(1))
			lgArrOrgLvl(ii, 1) = Trim(iArrCol(2))
		Next
		lgIntLastOrvLvl = Trim(iArrCol(1))
		lgIntLastOrgLvlIndex = ii
		lgBlnOrgLvlExists = True
	Else
		lgBlnOrgLvlExists = False
	END if
	
End sub

'==========================================================================================
'   Event Name : uniTree1_MenuRename
'   Event Desc : Node를 Drag할때 이벤트 
'==========================================================================================

Sub  uniTree1_MenuRename(Node)
	If ChkOrgTree(Node, C_ROOT_KEY) = False Then Exit Sub

	lgIntFlgMode = parent.OPMD_UMODE	' 신규로 등록 
	
	Call frm1.uniTree1.StartLabelEdit 
End Sub

'==========================================================================================
'   Event Name : uniTree1_MenuDelete
'   Event Desc : 삭제메뉴클릭시 
'==========================================================================================

Sub  uniTree1_MenuDelete(prObjNode)
	Dim  OldNode
	dIM IntRetCD
	Dim iStrVal
	Dim iArrTag

	On Error Resume Next
	
	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Sub
	End If
		
	frm1.uniTree1.MousePointer = 11
	Call LayerShowHide(1)
	
	set lgObjDragNode = prObjNode
			
	If lgBlnNewNode = TRUE Then
		if prObjNode.Key = lgNewNode.key then
		
			set lgObjDragNode = prObjNode.Next
			
			frm1.uniTree1.Nodes.Remove lgNewNode.Index
			lgBlnFlgChgValue = False
			lgBlnNewNode = False
			lgSaveModFg = ""
			Set lgNewNode = Nothing	
			Call LayerShowHide(0)

			frm1.uniTree1.MousePointer = 0
			frm1.uniTree1.Setfocus
			Call uniTree1_NodeClick(frm1.uniTree1.selecteditem)
			
			Exit sub
		end if
	end if

	frm1.uniTree1.MousePointer = 0

	If prObjNode.Image = C_GRP Then				
		lgStrCmd = "GRP"
					
		iStrVal = BIZ_SALES_GRP & "?txtMode="	& parent.UID_M0003
		iStrVal = iStrVal & "&txtSales_Grp2="	& Mid(prObjNode.Key, 2)
	Else	
		lgStrCmd = "ORG"
		' 하위조직이나 영업그룹이 존재하지 않는 경우 
		If prObjNode.Children = 0 Then
			iStrVal = BIZ_SALES_ORG & "?txtMode="	& parent.UID_M0003
			iStrVal = iStrVal & "&txtSales_Org2="	& Mid(prObjNode.Key, 2)
		Else
			iStrVal = BIZ_MOVE_TREE & "?txtMode=" & parent.UID_M0002
			iStrVal = iStrVal & "&txtFlag="		& "ORG4"						' Delete Sales Org. Tree
			iStrVal = iStrVal & "&txtSalesOrg=" & Mid(lgObjDragNode.key, 2)		' Sales Org.

			iArrTag = Split(lgObjDragNode.Tag, parent.gColSep)
			iStrVal = iStrVal & "&txtSalesOrgCurLvl="	& iArrTag(0)			' Sales Org. Current Level
			iStrVal = iStrVal & "&txtEndOrgFlag="		& iArrTag(1)			' End Org. Flag
		End If
	End If	

	lgSaveModFg	= "D"	 	
	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '☜: 비지니스 ASP 를 가동 
End Sub

 '=========================  uniTree1_onAddImgReady()  ====================================
'	Event  Name : uniTree1_onAddImgReady()
'	Description : SetAddImageCount수의 Image가 다운로드 완료되고 TreeView의 ImageList에 
'                 추가되면 발생하는 이벤트 
'========================================================================================= 
Sub uniTree1_onAddImgReady()
	If lgBlnOrgLvlExists Then
		Call DbQuery()
		Call SetToolbar("1100100000001111")									'⊙: 버튼 툴바 제어 
	Else
		Call SetToolbar("1000000000001111")									'⊙: 버튼 툴바 제어 
		Call SetDefaultScreen()
	End If
End Sub

'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function  FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    If lgBlnNewNode = TRUE Then
		lgBlnNewNode = FALSE		
		Set lgNewNode = Nothing
	end if
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")										'⊙: Clear Contents  Field
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoOper.ClearField(Document, "3")										'⊙: Clear Contents  Field
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function  FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '⊙: Clear Contents  Field
    Call ggoOper.ClearField(Document, "3")
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables

    FncNew = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function  FncDelete() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function  FncSave() 
	Dim IntRetCD    
	    
	FncSave = False                                                         '⊙: Processing is NG
	    
	Err.Clear                                                               '☜: Protect system from crashing
	On Error Resume Next                                                    '☜: Protect system from crashing
	    
	'-----------------------
	'Precheck area
	'-----------------------

	If Not lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900001","X","X","X")                          'No data changed!!
		Exit Function
	End If 
	    
	'-----------------------
	'Check content area
	'-----------------------
	If lgStrCmd = "ORG" Then
		If Not chkField(Document, "2") Then  Exit Function                        '⊙: Check contents area
	Else
		If Not chkField(Document, "3") Then  Exit Function                        '⊙: Check contents area
	End If

	'-----------------------
	'Save function call area
	'-----------------------
	IF DbSave = False Then
		Exit Function
	End IF					                                                  '☜: Save db data
	    
	FncSave = True                                                          '⊙: Processing is OK
	    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function  FncCopy() 
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function  FncCancel() 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function  FncPrint() 
    parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function  FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function  FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function  FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                          
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function  FncExit()
	Dim IntRetCD
	FncExit = False
    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    		If IntRetCD = vbNo Then
      			Exit Function
    		End If
    End If

    FncExit = True
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function  DbQuery() 
	DbQuery = False
	    
	Err.Clear                                                               '☜: Protect system from crashing

	Call DisplayNodes()
	Call DbQueryOk
	DbQuery = True    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode    

    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    
	If lgSelframeFlg = TAB2 Then
		lgBlnFlgChgValue = False
	End If
	Call SetDefaultScreen()

End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function  DbSave()     
	Dim iStrVal

	Call LayerShowHide(1)
	
    DbSave = False                                                          '⊙: Processing is NG
  
    On Error Resume Next                                                   '☜: Protect system from crashing
	With frm1
		' 영업그룹 
		If lgStrCmd  = "GRP" Then
			iStrVal = BIZ_SALES_GRP & "?txtMode="		& parent.UID_M0002
			iStrVal = iStrVal & "&txtFlgMode="			& lgIntFlgMode
			iStrVal = iStrVal & "&txtSales_Grp2="		& UCase(.txtSalesGrp.value)
			iStrVal = iStrVal & "&txtSales_Grp_nm2="	& .txtSalesGrpnm.value
			iStrVal = iStrVal & "&txtSales_Org_Fullnm=" & .txtSalesGrpFullnm.value
			iStrVal = iStrVal & "&txtSales_Org_Engnm="	& .txtSalesGrpEngnm.value
			iStrVal = iStrVal & "&txtCost_center="		& .txtCostCenter.value
			iStrVal = iStrVal & "&txtSales_Org="		& .txtSalesOrgInGrp.value
			If .rdoGrpUsageflagY.checked Then
				iStrVal = iStrVal & "&txtRadio=Y"
			Else
				iStrVal = iStrVal & "&txtRadio=N"
			End If
			iStrVal = iStrVal & "&txtInsrtUserId="		& parent.gUsrID
			iStrVal = iStrVal & "&txtprogramId=b1256ma1"
		ELSE
			iStrVal = BIZ_SALES_ORG & "?txtMode="		& parent.UID_M0002
			iStrVal = iStrVal & "&txtFlgMode="			& lgIntFlgMode
			iStrVal = iStrVal & "&txtSales_Org2="		& UCase(.txtSalesOrg.value)
			iStrVal = iStrVal & "&txtSales_Org_nm2="	& .txtSalesOrgnm.value
			iStrVal = iStrVal & "&txtSales_Org_Fullnm=" & .txtSalesOrgFullnm.value
			iStrVal = iStrVal & "&txtSales_Org_Engnm="	& .txtSalesOrgEngnm.value
			iStrVal = iStrVal & "&txtUpper_Sales_Org="	& .txtUpperSalesOrg.value
			iStrVal = iStrVal & "&txtHead_usr_nm="		& .txtHeadusrnm.value
			iStrVal = iStrVal & "&txtlvl="				& .txtSalesOrglvl.value
			
			'말단조직여부 
			If .rdoEndOrgFlagY.checked Then
				iStrVal = iStrVal & "&txtEndOrgFlag=Y"
			Else
				iStrVal = iStrVal & "&txtEndOrgFlag=N"
			End If
			'사용여부 
			If .rdoOrgUsageflagY.checked Then
				iStrVal = iStrVal & "&txtRadio=Y"
			Else
				iStrVal = iStrVal & "&txtRadio=N"
			End If
			iStrVal = iStrVal & "&txtInsrtUserId="		& parent.gUsrID
		End If
    	'-----------------------
		'Data manipulate area
		'-----------------------
		
	End With	
 	
	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '☜: 비지니스 ASP 를 가동 
	
    DbSave = True                                                           '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()	
	Dim iArrIndex, iArrTag
	Dim iStrText

	On Error Resume Next												'☆: 저장 성공후 실행 로직 

	lgBlnFlgChgValue = False
	
	If lgSaveModFg	= "R" Then
		If Not lgBlnRemakeNodes Then
			' 레벨이 변경된 경우 Node의 Tag값 재설정 
			If lgBlnLvlChanged Then
				' 새로운 레벨 설정 
				iArrIndex = Split(lgObjDropNode.fullpath, parent.gColSep)
				If Ubound(iArrIndex, 1) = lgIntLastOrgLvlIndex - 1 Then
					lgObjDragNode.Tag = lgArrOrgLvl(Ubound(iArrIndex, 1), 0) & parent.gColSep & "Y"
				Else
					lgObjDragNode.Tag = lgArrOrgLvl(Ubound(iArrIndex, 1), 0) & parent.gColSep & "N"
				End If
			End If
			Set lgObjDragNode.parent = lgObjDropNode
		Else
			' Drag된 Node 삭제 
			frm1.uniTree1.Nodes.Remove lgObjDragNode.Index
			' Drag된 Node 재생성 
			Call RemakeNodes()
		End If
	End If
	
	' 영업조직 입력 
	IF lgSaveModFg	= "O" Then	
		With frm1
			lgObjDragNode.Key = "O" & UCase(Trim(.txtSalesOrg.value))
			lgObjDragNode.text = "[" & UCase(Trim(.txtSalesOrg.value)) & "]" & .txtSalesOrgnm.value
			If .rdoEndOrgFlagN.checked Then
				lgObjDragNode.Tag = .txtSalesOrglvl.value & parent.gColSep & "N"
			Else
				lgObjDragNode.Tag = .txtSalesOrglvl.value & parent.gColSep & "Y"
			End If
		End With
	END IF	
	
	' 영업그룹 입력 
	IF lgSaveModFg	= "G" Then
		With frm1
			lgObjDragNode.Key = "G" & UCase(Trim(.txtSalesGrp.value))
			lgObjDragNode.text =  "[" & UCase(Trim(.txtSalesGrp.value)) & "]" & .txtSalesGrpnm.value
			iArrTag = Split(.unitree1.nodes(lgObjDragNode.Key).parent.Tag)
			lgObjDragNode.Tag = iArrTag(0) & parent.gColSep & "N"
		End With
	END IF	

	' 삭제 
	IF lgSaveModFg	= "D"  Then
		frm1.unitree1.nodes.remove lgObjDragNode.Key
		Call FncNew()
	End If
	
	Set lgObjDragNode = Nothing
	
	If lgBlnNewNode = TRUE Then
		lgBlnNewNode = FALSE		
		Set lgNewNode = Nothing
	end if

	' 트리뷰 Tag 재설정 
	If lgSaveModFg = "" Then
		With frm1
			If lgStrCmd = "ORG" Then
				If .rdoEndOrgFlagN.checked Then
					.uniTree1.selecteditem.Tag = .txtSalesOrglvl.value & parent.gColSep & "N"
				Else
					.uniTree1.selecteditem.Tag = .txtSalesOrglvl.value & parent.gColSep & "Y"
				End If
				
				iStrText = "[" & .txtSalesOrg.value & "]" & Trim(.txtSalesOrgnm.value)
				If Trim(.uniTree1.selecteditem.Text) <> iStrText Then
					.uniTree1.selecteditem.Text = iStrText
				End If
			Else
				iStrText = "[" & .txtSalesGrp.value & "]" & Trim(.txtSalesGrpNm.value)
				If Trim(.uniTree1.selecteditem.Text) <> iStrText Then
					.uniTree1.selecteditem.Text = iStrText
				End If
			End If
		End With
	End If	

	lgSaveModFg = ""

	lgBlnNewNode = False
	lgBlnFlgChgValue = False

	frm1.uniTree1.Setfocus
	
	Call LayerShowHide(0)

	frm1.uniTree1.MousePointer = 0
	Call uniTree1_NodeClick(frm1.uniTree1.selecteditem)
	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function  DbDelete() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
    On Error Resume Next
    Call DbSaveOk()
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

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
				<TR>
					<!-- TreeView AREA  -->
					<TD HEIGHT=* WIDTH=30%>
						<script language =javascript src='./js/b1256ma1_uniTree1_N785120457.js'></script>
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
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>영업조직</font></td>
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
													</TR>
												</TABLE>
											</TD>
											<TD CLASS="CLSMTABP">
												<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23" ></td>
														<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>영업그룹</font></td>
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
															<TD CLASS="TD5" NOWRAP>영업조직</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrg" TYPE="Text" MAXLENGTH="4" tag="23XXXU" size="10" ALT="영업조직"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>영업조직명</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrgnm" TYPE="Text" MAXLENGTH="50" tag="22XXX" size="34" ALT="영업조직명"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>조직레벨</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrglvl" TYPE="Text" MAXLENGTH="2" tag="24XXXU" size="10" ALT="조직레벨"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>영업조직총칭</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrgFullnm" TYPE="Text" MAXLENGTH="70" tag="21XXX" size="50"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>영업조직영문명</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrgEngnm" TYPE="Text" MAXLENGTH="50" tag="21XXX" size="50"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>상위영업조직</TD>
															<TD CLASS="TD656">
																<input NAME="txtUpperSalesOrg" TYPE="Text" MAXLENGTH="4" tag="24XXXU" size="10">&nbsp;<input NAME="txtUpperSalesOrgNm" TYPE="Text" MAXLENGTH="30" tag="24" size="30"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>영업조직장명</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtHeadusrnm" TYPE="Text" MAXLENGTH="50" tag="21XXX" size="50"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>말단조직여부</TD>
															<TD CLASS="TD656" NOWRAP>
															<input type=radio CLASS="RADIO" id=rdoEndOrgFlagY name="rdoEndOrgFlag" value="Y" tag = "21XXX" checked>
																<label for="rdoEndOrgFlagY">예</label>&nbsp;&nbsp;&nbsp;&nbsp;
															<input type=radio CLASS = "RADIO" id=rdoEndOrgFlagN name="rdoEndOrgFlag" value="N" tag = "21XXX">
																<label for="rdoEndOrgFlagN">아니오</label></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>사용여부</TD>
															<TD CLASS="TD656" NOWRAP>
																<input type=radio CLASS="RADIO" id=rdoOrgUsageflagY name="rdoOrgUsageflag" value="Y" tag = "21" checked>
																	<label for="rdoOrgUsageflagY">예</label>&nbsp;&nbsp;&nbsp;&nbsp;
																<input type=radio CLASS = "RADIO" id=rdoORgUsageflagN name="rdoOrgUsageflag" value="N" tag = "21">
																	<label for="rdoOrgUsageflagN">아니오</label></TD>
														</TR>
																									
													</TABLE>
												</DIV> 
												<!-- 두번째 탭 내용  -->
												<DIV ID="TabDiv" SCROLL=no>
													<TABLE <%=LR_SPACE_TYPE_60%>>
														<TR>
														  <TD CLASS="TD5" NOWRAP>영업그룹</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtSalesGrp" TYPE="Text" MAXLENGTH="4" tag="33XXXU" size="10" ALT="영업그룹"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>영업그룹명</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtSalesGrpnm" TYPE="Text" MAXLENGTH="50" tag="32XXX" size="50" ALT="영업그룹명"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>영업그룹총칭</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtSalesGrpFullnm" TYPE="Text" MAXLENGTH="120" tag="31XXX" size="50"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>영업그룹영문명</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtSalesGrpEngnm" TYPE="Text" MAXLENGTH="50" tag="31XXX" size="50"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>비용집계처</TD>
														  <TD CLASS="TD656" NOWRAP>
															<input NAME="txtCostCenter" TYPE="Text" MAXLENGTH="10" tag="32XXXU" size="10" ALT="비용집계처"><img SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCenter" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenSheetPopUp C_CostCd">
															<input TYPE=Text NAME="txtCostCenterNm" MAXLENGTH="20" tag="34" size="20"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>영업조직</TD>
														  <TD CLASS="TD656" NOWRAP>
															<input NAME="txtSalesOrgInGrp" TYPE="Text" MAXLENGTH="4" tag="34XXXU" size="10" ALT="영업조직"><img SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesOrgInGrp" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20">
															<input TYPE=Text NAME="txtSalesOrgNmInGrp" MAXLENGTH="50" tag="34" size="20"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>사용여부</TD>
														  <TD CLASS="TD656" NOWRAP>
															<input type=radio CLASS="RADIO" id=rdoGrpUsageflagY name="rdoGrpUsageflag" value="Y" tag = "31XXX" checked>
																<label for="rdoGrpUsageflagY">예</label>&nbsp;&nbsp;&nbsp;&nbsp;
															<input type=radio CLASS = "RADIO" id=rdoGrpUsageflagN name="rdoGrpUsageflag" value="N" tag = "31XXX">
																<label for="rdoGrpUsageflagN">아니오</label></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>

