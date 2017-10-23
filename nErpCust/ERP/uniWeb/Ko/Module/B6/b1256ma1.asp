
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Sale
'*  2. Function Name        : Sales Organization
'*  3. Program ID           : b1256ma1.asp
'*  4. Program Name         : ��������/�׷��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/09/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seong Bae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2002/12/24 Include ������� ���ر� 
'*                            b1254mb1ȣ��� argument prgramid�Ѱ��� b1254ma1�� ������������...
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ====================================== -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->
<!--'==========================================  1.1.2 ���� Include   ======================================
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
Option Explicit                                                             '��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'*****************<A HREF="\\ferrari\uniWEB\Template\inc\incUni2KTV.vbs">\\ferrari\uniWEB\Template\inc\incUni2KTV.vbs</A>*****************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
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

Const BIZ_MOVE_TREE = "b1256mb1.asp"										 '��: �����Ͻ� ���� ASP�� 
Const BIZ_SALES_GRP = "b1254mb1.asp"										 '��: �����׷��� 
Const BIZ_SALES_ORG = "b1255mb1.asp"										 '��: ����������� 

Const C_IMG_Root = "../../../CShared/image/unierp.gif"
Const C_IMG_ORG = "../../../CShared/image/Orglvl_2.gif"
Const C_IMG_Open = "../../../CShared/image/Group_op.gif"
Const C_IMG_GRP = "../../../CShared/image/HumanC.gif"
Const C_IMG_None = "../../../CShared/image/c_none.gif"
Const C_IMG_Const = "../../../CShared/image/c_const.gif"

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

Const   C_CTRLITEM		= 1
Const   C_CTRLITEMPB	= 2
Const   C_CTRLNM		= 3
Const	C_CTRLITEMSEQ	= 4
Const   C_DRFG			= 5
Const   C_CRFG			= 6

Const	C_CostCD = 1

 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
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
Dim lgArrOrgLvl					' �������� Level���� 
Dim lgIntLastOrvLvl
Dim	lgIntLastOrgLvlIndex
Dim	lgBlnRemakeNodes				' ������ ����� Tag�� �������� �����ϱ����� ������ ����(���� ������ �����ϴ� ��� ������)
Dim	lgBlnLvlChanged
Dim lgBlnOpenPopup
Dim lgBlnOrgLvlExists
 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
 '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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
		.Indentation = "200"	' �� ���� 
						' ������ġ,	Ű��, ��ġ 
		.AddImage C_IMG_Root,		C_Root,		0
		.AddImage C_IMG_ORG,		C_ORG,		0
		.AddImage C_IMG_Open,		C_Open,		0
		.AddImage C_IMG_GRP,		C_GRP,		0
		.AddImage C_IMG_None,		C_None,		0
		.AddImage C_IMG_Const,		C_Const,	0
	
		.PathSeparator = parent.gColSep
		
		.OLEDragMode = 1														'��: Drag & Drop �� �����ϰ� �� ���ΰ� ���� 
		.OLEDropMode = 1	
	
		.OpenTitle = "���������Է�"											
		.AddTitle = "�����׷��Է�"		
		.RenameTitle = ""
		.DeleteTitle = "����"
	End With
End Sub		

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ===================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub  SetDefaultVal()
	Call GetSalesOrgLvlInfo()
	lgBlnOpenPopup = False
End Sub

'==========================================  2.2.2 SetDefaultScreen()  ===================================
'	Name : SetDefaultScreen()
'	Description : Default Screen�� �����Ѵ�.
'========================================================================================================= 
Sub SetDefaultScreen()
	ClickTab1()
	Call InitVariables                                                      '��: Initializes local global variables
	Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  '��: Clear Contents  Field
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

'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�.
'********************************************************************************************************* 
 '==========================================  2.3.1 Tab Click ó��  =================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=================================================================================================================== 
 '----------------  ClickTab1(): Header Tabó�� �κ� (Header Tab�� �ִ� ��츸 ���)  ---------------------------- 
Function ClickTab1()
	
	If lgSelframeFlg = TAB1 Then Exit Function
	 
	Call changeTabs(TAB1)	 '~~~ ù��° Tab 	
	lgSelframeFlg = TAB1

End Function

Function ClickTab2()

	If lgSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ ù��° Tab 
	lgSelframeFlg = TAB2

End Function

 '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
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
		iArrParam(1) = "dbo.b_cost_center"				<%' TABLE ��Ī %>
		iArrParam(2) = Trim(frm1.txtCostCenter.value)	<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = ""								<%' Where Condition%>
		iArrParam(5) = frm1.txtCostCenter.alt 			<%' TextBox ��Ī %>
			
		iArrField(0) = "ED15" & parent.gColSep & "cost_cd"	<%' Field��(0)%>
		iArrField(1) = "ED30" & parent.gColSep & "cost_nm"	<%' Field��(1)%>
		    
		iArrHeader(0) = "�������ó"					<%' Header��(0)%>
		iArrHeader(1) = "�������ó��"					<%' Header��(1)%>
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' �˾� ��Ī %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPopup = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSheetPopup = SetSheetPopup(iArrRet,pvIntWhere)
	End If	
	
End Function

 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetRefOpenAp()  --------------------------------------------------
'	Name : SetSheetPopup()
'	Description : OpenSheetPopup���� Return�Ǵ� �� setting
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
'   Function Desc :Drag �� ��� �ִ��� Drag�Ǵ� �׸����� üũ 
'==========================================================================================
Function  ChkDragState(ByVal x , ByVal y )
    
    Dim iObjNewNode
    dim ChildNode
    Dim iArrNewNodeTag, iArrDragNodeTag
    Dim iIntIndex, iIntCurOrgIndex
    
    On Error Resume Next
    
    ChkDragState = False

    If lgObjDragNode.parent Is Nothing Then Exit Function	' �ڽ��� Root�� ��� 
    
    Set iObjNewNode = frm1.uniTree1.HitTest(x, y)

    ' ������ �������� ���� ��� 
    If iObjNewNode Is Nothing Then Exit Function
    
    ' Ʈ������ ���翩�� Check
	If Not ChkOrgTree(iObjNewNode, C_ROOT_KEY) Then Exit Function

    iArrNewNodeTag = Split(iObjNewNode.Tag, parent.gColSep)
    iArrDragNodeTag = Split(lgObjDragNode.Tag, parent.gColSep)

	' Drag�� Node�� ���� ������ ��쿡�� ������������ Drop�� �� ����.
	If lgObjDragNode.Image = C_GRP Then
		' �����׷��� ������������ ���ӵ� �� �ִ�.
		If iObjNewNode.Key = C_ROOT_KEY OR iArrNewNodeTag(1) = "N" Then Exit Function
		
	Else
    	If iObjNewNode.Key = C_ROOT_KEY Then
			iIntCurOrgIndex = 0
		Else
			' ���� ������ �̵��ϴ� ��� �� ������ �ڽ��� ���� �������� check
			If iArrNewNodeTag(0) > iArrDragNodeTag(0) Then
				' Ʈ������ ���翩�� Check
				If ChkOrgTree(iObjNewNode, lgObjDragNode.Key) Then Exit Function
			End If
			
			' ���� ���� �Ʒ����� ���������� �� �� ����.
			If iArrNewNodeTag(1) = "Y" Then Exit Function
		
			For iIntIndex = 0 to lgIntLastOrgLvlIndex - 1
				If lgArrOrgLvl(iIntIndex, 0) = iArrNewNodeTag(0) then
					iIntCurOrgIndex = iIntIndex + 1
					Exit For
				End If
			Next
		End If

		' ���������� �ִ밪 Check
		If iIntCurOrgIndex + GetSubOrgLvlCnt(iArrDragNodeTag(0), Mid(lgObjDragNode.Key,2)) > lgIntLastOrgLvlIndex Then Exit Function
	End If
	
    '�ڽ��� �ڸ��� ������ 
    If iObjNewNode.Text = lgObjDragNode.parent.Text Then Exit Function
    
    ' �ڽ��� �θ𿡰� ���� 
    If iObjNewNode.Key = lgObjDragNode.Key Then Exit Function
    
    ' �����׷쿡 Drop�� ��� 
    If iObjNewNode.Image = C_GRP Then Exit Function
    
    ChkDragState = True
    
End Function

' Ư�� Ʈ��(pvStrFind)���� �����ϴ� check�ϴ� ����Լ� 
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
'   Function Desc : ���� ��尡 �����ϰ� �ִ� ������������ ���� ���Ѵ�.
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
'   Function Desc :Add�� ���õ� �ڽļ��� �ǵ����ش�.
'==========================================================================================

Function GetTotalCnt(prObjNode)
	
	If prObjNode.children = 0 Then	' Root�� ��� 
		GetTotalCnt = 1
	Else
		GetTotalCnt = prObjNode.children + 1
	End If
	
End Function


'======================================================================================================
'	ȭ�� ������ ���� 
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
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub  Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call AppendNumberPlace("7","3","0")
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "3", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
                             
    Call InitVariables                                                      '��: Initializes local global variables
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

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

'======================================================================================================
'	�߰� 
'======================================================================================================
Function Add_onclick()
	Dim strRetValue
	strRetValue = window.showModalDialog("FolderAdd.asp", "", "dialogWidth=400px; dialogHeight=300px; center:Yes; help:No; resizable:No; status:No;")
End Function
  	
'======================================================================================================
'	���� 
'======================================================================================================
Function Form_onclick()
	Dim strRetValue
	strRetValue = window.showModalDialog("FolderConfig.asp", "", "dialogWidth=400px; dialogHeight=300px; center:Yes; help:No; resizable:No; status:No;")
End Function

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
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
'   Event Desc : Node�� Ŭ���ϸ� �߻� �̺�Ʈ 
'==========================================================================================

Sub uniTree1_NodeClick(pvObjNode)
	On Error Resume Next
	Dim Response
	Dim iBlnProtect
	
	' Ʈ�� ��ȸ�ÿ� Ŭ���� �ϸ� ��ȸ�� ���� �ʵ��� ��ġ 
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
	Call SetToolbar("1100100000001111")									'��: ��ư ���� ����					 
	
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
			IntRetCD = DisplayMsgBox("125500","X","X","X")	' �������������� �������� �ʽ��ϴ�.
		Else
			IntRetCD = DisplayMsgBox("125400","X","X","X")	' �����׷������� �������� �ʽ��ϴ�.
		End If
	End if 

    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    ' End Org Flag Protect ó�� 
    If pvObjNode.Image = C_ORG And iBlnProtect Then
		Call ggoOper.SetReqAttr(frm1.rdoEndOrgFlagY, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoEndOrgFlagN, "Q")
    End If
	Call LayerShowHide(0)
	lgBlnFlgChgValue = False
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode    
End Sub

'==========================================================================================
'   Event Name : uniTree1_OLEDragDrop
'   Event Desc : Node�� Drag & Drop �̺�Ʈ 
'==========================================================================================

Sub  uniTree1_OLEDragDrop(Data , Effect , Button , Shift , x , y )

	Dim IntRetCD
    Dim iStrVal
    Dim iArrIndex, iArrTag

	If lgObjDragNode Is Nothing Then Exit Sub

	' mscomctl.ocx �������� Ŭ����OLEDragDrop Event�� �߻��Ͽ� 
	' "�ش� ��ġ�δ� �̵��� �� �����ϴ�!" ���� �߻� 
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
		IntRetCD = DisplayMsgBox("990017","X","X","X")	' �ش� ��ġ�δ� �̵��� �� �����ϴ�!
        Exit Sub
	End If

	' ������ ����� Tag�� �������� �����ϱ����� ������ ���� 
	lgBlnRemakeNodes = False
	
	Call LayerShowHide(1)

	frm1.uniTree1.MousePointer = 11

    Set lgObjDropNode = frm1.uniTree1.HitTest(x, y)					' �̵��ؾߵ� ��带 ����Ŵ 
 
	' �����׷��� �̵��� ���� b_sales_grp.sales_org�� �����ϸ� �ȴ�.	
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

		' ���� ������ ���� ���� check
		lgBlnLvlChanged = True
		If lgObjDropNode.Key <> C_ROOT_KEY And lgObjDragNode.parent.Key <> C_ROOT_KEY THEN
			If lgObjDropNode.parent.fullpath = lgObjDragNode.parent.parent.fullpath Then
				lgBlnLvlChanged = False
				iStrVal = iStrVal & "&txtFlag="		& "ORG1"								' Sales Org.
			End If
		End If

		If lgBlnLvlChanged Then
			' ���ο� ���� ���� 
			iArrIndex = Split(lgObjDropNode.fullpath, parent.gColSep)
			iStrVal = iStrVal & "&txtSalesOrgNewLvl=" & lgArrOrgLvl(Ubound(iArrIndex, 1), 0)	' Sales Org. New Level
			If Ubound(iArrIndex, 1) = lgIntLastOrgLvlIndex - 1 Then
				iStrVal = iStrVal & "&txtEndOrgFlag=Y"
			Else
				iStrVal = iStrVal & "&txtEndOrgFlag=N"
			End If
			
			iArrTag = Split(lgObjDragNode.Tag, parent.gColSep)
			iStrVal = iStrVal & "&txtSalesOrgCurLvl=" & iArrTag(0)								' Sales Org. Current Level

			' ���������� ���翩�� Check
			If lgObjDragNode.Children = 0 Then
				iStrVal = iStrVal & "&txtFlag="	& "ORG2"								' Sales Org.
			Else
				'������������ 
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
	
	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 

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
'   Event Desc : Node�� Drag�Ҷ� �̺�Ʈ 
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
'   Event Desc : Node�� Drag�Ҷ� �̺�Ʈ 
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
			
			' Mouse Pointer�� Ʈ���� ��ġ�ϴ��� Check
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
				
				' ���� ���ο� �Է��� ��忡�� popup �� ���� �Է¸޴����� ���̸� �ȵȴ�.
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

							' ���������� ��� 
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
'   Event Name : uniTree1_MenuOpen - ���������Է� 
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

	If pvObjNode.Key = C_ROOT_KEY Then	' Root�� ��� 
		Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, C_ROOT_KEY_STR & GetTotalCnt(pvObjNode), "�� ��������", C_ORG, C_ORG)

		With frm1
		
			.txtSalesOrglvl.value = lgArrOrgLvl(0, 0)
			' ���������� �ϳ��� ��� �����������θ� 'Y'�� ���� 
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
		Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, pvObjNode.Key & C_UNDERSCORE & GetTotalCnt(pvObjNode), "�� ��������", C_ORG, C_ORG)

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

	Call SetToolbar("1100100000001111")									'��: ��ư ���� ����		
			
	lgIntFlgMode = parent.OPMD_CMODE	' �űԷ� ��� 
	
	lgStrCmd  = "ORG"
	
	lgBlnFlgChgValue = TRUE
	lgBlnNewNode = TRUE
	lgSaveModFg	= "O"	
End Sub


'==========================================================================================
'   Event Name : uniTree1_MenuAdd - �����׷���Է� 
'   Event Desc : Node Popup
'==========================================================================================
Sub  uniTree1_MenuAdd(pvObjNode)

	Dim iObjDummyNode
		
	'If ChkOrgTree(Node, C_ROOT_KEY) = TRUE Then Exit Sub
	CALL FNCNEW
	
	If pvObjNode.Expanded = False Then
		pvObjNode.Expanded = True
	End If
	
	Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, pvObjNode.Key & C_UNDERSCORE & GetTotalCnt(pvObjNode), "�� �����׷�", C_GRP, C_GRP)
	
	iObjDummyNode.Selected = True
	Set lgNewNode = iObjDummyNode
	set lgObjDragNode = iObjDummyNode	
	
	Call SetToolbar("1100100000001111")									'��: ��ư ���� ����		
	 
	Call ClickTab2()

	frm1.txtSalesOrgInGrp.value = Mid(pvObjNode.Key,2)
		
	lgIntFlgMode = parent.OPMD_CMODE	' �űԷ� ��� 
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
'   Event Desc : ������ ����ǰ� ���� ������ �����ϴ� ��� ���� Nodes�� �缺���Ѵ�.
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
'   Event Desc : Node�� Drag�Ҷ� �̺�Ʈ 
'==========================================================================================

Sub  uniTree1_MenuRename(Node)
	If ChkOrgTree(Node, C_ROOT_KEY) = False Then Exit Sub

	lgIntFlgMode = parent.OPMD_UMODE	' �űԷ� ��� 
	
	Call frm1.uniTree1.StartLabelEdit 
End Sub

'==========================================================================================
'   Event Name : uniTree1_MenuDelete
'   Event Desc : �����޴�Ŭ���� 
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
		' ���������̳� �����׷��� �������� �ʴ� ��� 
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
	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 
End Sub

 '=========================  uniTree1_onAddImgReady()  ====================================
'	Event  Name : uniTree1_onAddImgReady()
'	Description : SetAddImageCount���� Image�� �ٿ�ε� �Ϸ�ǰ� TreeView�� ImageList�� 
'                 �߰��Ǹ� �߻��ϴ� �̺�Ʈ 
'========================================================================================= 
Sub uniTree1_onAddImgReady()
	If lgBlnOrgLvlExists Then
		Call DbQuery()
		Call SetToolbar("1100100000001111")									'��: ��ư ���� ���� 
	Else
		Call SetToolbar("1000000000001111")									'��: ��ư ���� ���� 
		Call SetDefaultScreen()
	End If
End Sub

'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function  FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

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
    Call ggoOper.ClearField(Document, "1")										'��: Clear Contents  Field
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call ggoOper.ClearField(Document, "3")										'��: Clear Contents  Field
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function  FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    On Error Resume Next                                                    '��: Protect system from crashing

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
    Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '��: Clear Contents  Field
    Call ggoOper.ClearField(Document, "3")
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitVariables                                                      '��: Initializes local global variables

    FncNew = True                                                           '��: Processing is OK
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
	    
	FncSave = False                                                         '��: Processing is NG
	    
	Err.Clear                                                               '��: Protect system from crashing
	On Error Resume Next                                                    '��: Protect system from crashing
	    
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
		If Not chkField(Document, "2") Then  Exit Function                        '��: Check contents area
	Else
		If Not chkField(Document, "3") Then  Exit Function                        '��: Check contents area
	End If

	'-----------------------
	'Save function call area
	'-----------------------
	IF DbSave = False Then
		Exit Function
	End IF					                                                  '��: Save db data
	    
	FncSave = True                                                          '��: Processing is OK
	    
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function  FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
' Function Desc : ȭ�� �Ӽ�, Tab���� 
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    		If IntRetCD = vbNo Then
      			Exit Function
    		End If
    End If

    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function  DbQuery() 
	DbQuery = False
	    
	Err.Clear                                                               '��: Protect system from crashing

	Call DisplayNodes()
	Call DbQueryOk
	DbQuery = True    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode    

    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    
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
	
    DbSave = False                                                          '��: Processing is NG
  
    On Error Resume Next                                                   '��: Protect system from crashing
	With frm1
		' �����׷� 
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
			
			'������������ 
			If .rdoEndOrgFlagY.checked Then
				iStrVal = iStrVal & "&txtEndOrgFlag=Y"
			Else
				iStrVal = iStrVal & "&txtEndOrgFlag=N"
			End If
			'��뿩�� 
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
 	
	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 
	
    DbSave = True                                                           '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()	
	Dim iArrIndex, iArrTag
	Dim iStrText

	On Error Resume Next												'��: ���� ������ ���� ���� 

	lgBlnFlgChgValue = False
	
	If lgSaveModFg	= "R" Then
		If Not lgBlnRemakeNodes Then
			' ������ ����� ��� Node�� Tag�� �缳�� 
			If lgBlnLvlChanged Then
				' ���ο� ���� ���� 
				iArrIndex = Split(lgObjDropNode.fullpath, parent.gColSep)
				If Ubound(iArrIndex, 1) = lgIntLastOrgLvlIndex - 1 Then
					lgObjDragNode.Tag = lgArrOrgLvl(Ubound(iArrIndex, 1), 0) & parent.gColSep & "Y"
				Else
					lgObjDragNode.Tag = lgArrOrgLvl(Ubound(iArrIndex, 1), 0) & parent.gColSep & "N"
				End If
			End If
			Set lgObjDragNode.parent = lgObjDropNode
		Else
			' Drag�� Node ���� 
			frm1.uniTree1.Nodes.Remove lgObjDragNode.Index
			' Drag�� Node ����� 
			Call RemakeNodes()
		End If
	End If
	
	' �������� �Է� 
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
	
	' �����׷� �Է� 
	IF lgSaveModFg	= "G" Then
		With frm1
			lgObjDragNode.Key = "G" & UCase(Trim(.txtSalesGrp.value))
			lgObjDragNode.text =  "[" & UCase(Trim(.txtSalesGrp.value)) & "]" & .txtSalesGrpnm.value
			iArrTag = Split(.unitree1.nodes(lgObjDragNode.Key).parent.Tag)
			lgObjDragNode.Tag = iArrTag(0) & parent.gColSep & "N"
		End With
	END IF	

	' ���� 
	IF lgSaveModFg	= "D"  Then
		frm1.unitree1.nodes.remove lgObjDragNode.Key
		Call FncNew()
	End If
	
	Set lgObjDragNode = Nothing
	
	If lgBlnNewNode = TRUE Then
		lgBlnNewNode = FALSE		
		Set lgNewNode = Nothing
	end if

	' Ʈ���� Tag �缳�� 
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
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()												'��: ���� ������ ���� ���� 
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

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
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
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>��������</font></td>
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
													</TR>
												</TABLE>
											</TD>
											<TD CLASS="CLSMTABP">
												<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23" ></td>
														<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�����׷�</font></td>
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
												<!-- ù��° �� ����  -->
												<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=no>
													<TABLE <%=LR_SPACE_TYPE_60%>>
														<TR>
															<TD CLASS=TD5 HEIGHT=5 WIDTH="100%"></TD>
															<TD CLASS=TD6 HEIGHT=5 WIDTH="100%"></TD>												
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>��������</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrg" TYPE="Text" MAXLENGTH="4" tag="23XXXU" size="10" ALT="��������"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>����������</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrgnm" TYPE="Text" MAXLENGTH="50" tag="22XXX" size="34" ALT="����������"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>��������</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrglvl" TYPE="Text" MAXLENGTH="2" tag="24XXXU" size="10" ALT="��������"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>����������Ī</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrgFullnm" TYPE="Text" MAXLENGTH="70" tag="21XXX" size="50"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>��������������</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtSalesOrgEngnm" TYPE="Text" MAXLENGTH="50" tag="21XXX" size="50"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>������������</TD>
															<TD CLASS="TD656">
																<input NAME="txtUpperSalesOrg" TYPE="Text" MAXLENGTH="4" tag="24XXXU" size="10">&nbsp;<input NAME="txtUpperSalesOrgNm" TYPE="Text" MAXLENGTH="30" tag="24" size="30"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>�����������</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtHeadusrnm" TYPE="Text" MAXLENGTH="50" tag="21XXX" size="50"></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>������������</TD>
															<TD CLASS="TD656" NOWRAP>
															<input type=radio CLASS="RADIO" id=rdoEndOrgFlagY name="rdoEndOrgFlag" value="Y" tag = "21XXX" checked>
																<label for="rdoEndOrgFlagY">��</label>&nbsp;&nbsp;&nbsp;&nbsp;
															<input type=radio CLASS = "RADIO" id=rdoEndOrgFlagN name="rdoEndOrgFlag" value="N" tag = "21XXX">
																<label for="rdoEndOrgFlagN">�ƴϿ�</label></TD>
														</TR>
														<TR>
															<TD CLASS="TD5" NOWRAP>��뿩��</TD>
															<TD CLASS="TD656" NOWRAP>
																<input type=radio CLASS="RADIO" id=rdoOrgUsageflagY name="rdoOrgUsageflag" value="Y" tag = "21" checked>
																	<label for="rdoOrgUsageflagY">��</label>&nbsp;&nbsp;&nbsp;&nbsp;
																<input type=radio CLASS = "RADIO" id=rdoORgUsageflagN name="rdoOrgUsageflag" value="N" tag = "21">
																	<label for="rdoOrgUsageflagN">�ƴϿ�</label></TD>
														</TR>
																									
													</TABLE>
												</DIV> 
												<!-- �ι�° �� ����  -->
												<DIV ID="TabDiv" SCROLL=no>
													<TABLE <%=LR_SPACE_TYPE_60%>>
														<TR>
														  <TD CLASS="TD5" NOWRAP>�����׷�</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtSalesGrp" TYPE="Text" MAXLENGTH="4" tag="33XXXU" size="10" ALT="�����׷�"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>�����׷��</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtSalesGrpnm" TYPE="Text" MAXLENGTH="50" tag="32XXX" size="50" ALT="�����׷��"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>�����׷���Ī</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtSalesGrpFullnm" TYPE="Text" MAXLENGTH="120" tag="31XXX" size="50"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>�����׷쿵����</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtSalesGrpEngnm" TYPE="Text" MAXLENGTH="50" tag="31XXX" size="50"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>�������ó</TD>
														  <TD CLASS="TD656" NOWRAP>
															<input NAME="txtCostCenter" TYPE="Text" MAXLENGTH="10" tag="32XXXU" size="10" ALT="�������ó"><img SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCenter" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenSheetPopUp C_CostCd">
															<input TYPE=Text NAME="txtCostCenterNm" MAXLENGTH="20" tag="34" size="20"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>��������</TD>
														  <TD CLASS="TD656" NOWRAP>
															<input NAME="txtSalesOrgInGrp" TYPE="Text" MAXLENGTH="4" tag="34XXXU" size="10" ALT="��������"><img SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesOrgInGrp" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20">
															<input TYPE=Text NAME="txtSalesOrgNmInGrp" MAXLENGTH="50" tag="34" size="20"></TD>
														</TR>
														<TR>
														  <TD CLASS="TD5" NOWRAP>��뿩��</TD>
														  <TD CLASS="TD656" NOWRAP>
															<input type=radio CLASS="RADIO" id=rdoGrpUsageflagY name="rdoGrpUsageflag" value="Y" tag = "31XXX" checked>
																<label for="rdoGrpUsageflagY">��</label>&nbsp;&nbsp;&nbsp;&nbsp;
															<input type=radio CLASS = "RADIO" id=rdoGrpUsageflagN name="rdoGrpUsageflag" value="N" tag = "31XXX">
																<label for="rdoGrpUsageflagN">�ƴϿ�</label></TD>
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

