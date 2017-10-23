<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Cost
'*  2. Function Name        : Cost Center
'*  3. Program ID           : b1256ma1.asp
'*  4. Program Name         : ���C/C �׷��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/08/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : choe0tae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            
'*                            
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

Const BIZ_PGM_ID = "c4001mb1.asp"

Const  C_Root	= "Root"
Const  C_CCG	= "CCG"
Const  C_CC1	= "CC1"

Const  C_ROOT_DESC = "[*]uniERP"
Const  C_ROOT_KEY = "$"
Const  C_ROOT_KEY_STR = "RT_"
Const  C_UNDERSCORE = "_"

Const BIZ_MOVE_TREE = "c4001mb1.asp"										 '��: Ʈ�� �� ��ȸ 
Const BIZ_SALES_GRP = "C4001mb2.asp"										 '��: �ڽ�Ʈ���� �׷� ��ȸ/����/���� 
Const BIZ_SALES_ORG = "C4001mb3.asp"										 '��: �ڽ�Ʈ���� ��ȸ/����/���� 

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
Dim lgArrOrgLvl					' ���C/C �׷��� Level���� 
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
		.AddImage C_IMG_ORG,		C_CCG,		0
		.AddImage C_IMG_Open,		C_Open,		0
		.AddImage C_IMG_GRP,		C_CC1,		0
		.AddImage C_IMG_None,		C_None,		0
		.AddImage C_IMG_Const,		C_Const,	0
	
		.PathSeparator = parent.gColSep
		
		.OLEDragMode = 1														'��: Drag & Drop �� �����ϰ� �� ���ΰ� ���� 
		.OLEDropMode = 1	
	
		.OpenTitle = "Cost Center Group �߰�"
		.AddTitle = "Cost Center �߰�"		
		.RenameTitle = "������"
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

	Call GetCostCenterLvlInfo()
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
		Call ggoOper.SetReqAttr(.txtCOST_CD, "Q")
		Call ggoOper.SetReqAttr(.txtCOST_NM, "Q")
		Call ggoOper.SetReqAttr(.txtCOST_CD_2, "Q")
		Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_Y, "Q")
		Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_N, "Q")
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
'	Name : OpenCostCenter()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenCostCenter(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenCostCenter = False
	
	If lgBlnOpenPopup = True Then Exit Function

	lgBlnOpenPopup = True

	Select Case pvIntWhere

	Case 0												
		iArrParam(1) = "dbo.b_cost_center"				<%' TABLE ��Ī %>
		iArrParam(2) = Trim(frm1.txtCOST_CD_2.value)	<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = ""								<%' Where Condition%>
		iArrParam(5) = frm1.txtCOST_CD_2.alt 			<%' TextBox ��Ī %>
			
		iArrField(0) = "ED15" & parent.gColSep & "cost_cd"	<%' Field��(0)%>
		iArrField(1) = "ED30" & parent.gColSep & "cost_nm"	<%' Field��(1)%>
		    
		iArrHeader(0) = "Cost Center"					<%' Header��(0)%>
		iArrHeader(1) = "Cost Center��"					<%' Header��(1)%>
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' �˾� ��Ī %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPopup = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenCostCenter = SetCostCenter(iArrRet,pvIntWhere)
	End If	
	
End Function

 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetRefOpenAp()  --------------------------------------------------
'	Name : SetSheetPopup()
'	Description : OpenSheetPopup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCostCenter(Byval pvArrRet, Byval pvIntWhere)

	SetCostCenter = False
	
	With frm1
		Select Case pvIntWhere
		Case 0
			.txtCOST_CD_2.value = pvArrRet(0)
			.txtCOST_NM_2.value = pvArrRet(1)
		End Select	
    End With
    
    SetCostCenter = True
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
	If lgObjDragNode.Image = C_CC1 Then
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
			
			' ���� ���� �Ʒ����� ���C/C �׷����� �� �� ����.
			If iArrNewNodeTag(1) = "Y" Then Exit Function
		
			For iIntIndex = 0 to lgIntLastOrgLvlIndex - 1
				If lgArrOrgLvl(iIntIndex, 0) = iArrNewNodeTag(0) then
					iIntCurOrgIndex = iIntIndex + 1
					Exit For
				End If
			Next
		End If

		' ���������� �ִ밪 Check
		'If iIntCurOrgIndex + GetSubOrgLvlCnt(iArrDragNodeTag(0), Mid(lgObjDragNode.Key,2)) > lgIntLastOrgLvlIndex Then Exit Function
	End If
	
    '�ڽ��� �ڸ��� ������ 
    If iObjNewNode.Text = lgObjDragNode.parent.Text Then Exit Function
    
    ' �ڽ��� �θ𿡰� ���� 
    If iObjNewNode.Key = lgObjDragNode.Key Then Exit Function
    
    ' �����׷쿡 Drop�� ��� 
    If iObjNewNode.Image = C_CC1 Then Exit Function
    
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

' -- ���� ���� ���� 
Function CheckLeafFlagY()
	CheckLeafFlag = True
	
	If frm1.rdoLEAF_FLAG_Y.checked = True Then	' -- �����̸�  True
		Exit Function
	End If
	
	CheckLeafFlag = False
End Function
'==========================================================================================
'   Function Name : GetSubOrgLvlCnt
'   Function Desc : ���� ��尡 �����ϰ� �ִ� ���C/C �׷��Ϸ��� ���� ���Ѵ�.
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
Sub rdoLEAF_FLAG_N_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoLEAF_FLAG_Y_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtCOST_CD_2_onChange()
	
	If Len(frm1.txtCOST_CD_2.Value) > 0 Then
		If CommonQueryRs("COST_CD, COST_NM", " B_COST_CENTER " , " COST_CD = '" & Trim(frm1.txtCOST_CD_2.Value) &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    	frm1.txtCOST_CD_2.Value	= Replace(lgF0, Chr(11), "")
			frm1.txtCOST_NM_2.Value	= Replace(lgF1, Chr(11), "")
		Else
			Call DisplayMsgBox("970000", "x",frm1.txtCOST_CD_2.alt & " '" & UCase(Me.Value) & "' " ,"x")
			frm1.txtCOST_CD_2.Value	= ""
			frm1.txtCOST_NM_2.Value	= ""
		End If
	Else
		frm1.txtCOST_NM_2.Value = ""
	End If

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
	If pvObjNode.Image = C_CCG Then
		iStrSelect	= "  "
		iStrSelect	= " COST_CD, COST_NM, LEVEL_CD, UPPER_COST_CD, LEAF_FLAG "
		iStrFrom	= " dbo.C_COST_CENTER_HIERARCHY_S "
		iStrWhere	= " COST_CD =  " & FilterVar(Mid (pvObjNode.key,2), "''", "S") & " "
		
		ClickTab1()
		lgStrCmd = "CCG"		
	Else
		iStrSelect	= " COST_CD, COST_NM, LEVEL_CD, UPPER_COST_CD, LEAF_FLAG "
		iStrFrom	= " dbo.C_COST_CENTER_HIERARCHY_S "
		iStrWhere	= " COST_CD =  " & FilterVar(Mid (pvObjNode.key,2), "''", "S") & " "
		
		ClickTab2()
		lgStrCmd = "CC1"
	End If
	 
	If CommonQueryRs2by2(iStrSelect, iStrFrom ,  iStrWhere , lgF2By2) Then 
	
		iArrRow = Split(lgF2By2, parent.gColSep & parent.gRowSep)
		iArrCol	= Split(iArrRow(0), parent.gColSep)			

		With frm1
			If pvObjNode.Image = C_CCG Then
			
				.txtCOST_CD.value		= iArrCol(1)
				.txtCOST_CD_OLD.value	= iArrCol(1)	' -- ���������� ���� Ű�� 
				.txtCOST_NM.value		= iArrCol(2)
				.txtLEVEL_CD.value		= iArrCol(3)
				.txtUPPER_COST_CD.value = iArrCol(4)

				If iArrCol(5) = "Y" Then
					.rdoLEAF_FLAG_Y.checked = True
				Else
					.rdoLEAF_FLAG_N.checked = True
				End If

				.txtCOST_CD_2.value = ""
				.txtCOST_CD_2_OLD.value = ""
				.txtCOST_NM_2.value = ""
				.txtUPPER_COST_CD_2.value = ""
				
				' if Last level, you cannot edit 'End Org. Flag'
				If lgArrOrgLvl(lgIntLastOrgLvlIndex - 1, 0) = iArrCol(6) Then
					iBlnProtect = True
				Else
					IF pvObjNode.Children > 0 THEN
						' If it has sales group as child node, you cannot edit 'End org. flag'
						If pvObjNode.Child.Image = C_CC1 Then
							iBlnProtect = True
						End If
					end if 
				End If
				
			Else	' -- C/C�� ��� 
				.txtCOST_CD.value		= ""
				.txtCOST_CD_OLD.value		= ""
				.txtCOST_NM.value		= ""
				.txtLEVEL_CD.value		= ""
				.txtUPPER_COST_CD.value = ""

				If iArrCol(5) = "Y" Then
					.rdoLEAF_FLAG_Y.checked = True
				Else
					.rdoLEAF_FLAG_N.checked = True
				End If

				.txtCOST_CD_2.value			= iArrCol(1)
				.txtCOST_CD_2_OLD.value		= iArrCol(1)	' -- �������� Ű�� 
				.txtCOST_NM_2.value			= iArrCol(2)
				.txtUPPER_COST_CD_2.value	= iArrCol(4)
			End If
			
		End With
	Else
		If lgStrCmd = "CCG" Then
			IntRetCD = DisplayMsgBox("970000","X","Cost Center Group " & Mid (pvObjNode.key,2) ,"X")	' ���C/C �׷��������� �������� �ʽ��ϴ�.
		Else
			IntRetCD = DisplayMsgBox("970000","X","Cost Center " & Mid (pvObjNode.key,2),"X")	' �����׷������� �������� �ʽ��ϴ�.
		End If
	End if 

    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    ' End Org Flag Protect ó�� 
    If pvObjNode.Image = C_CCG And iBlnProtect Then
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
    Dim iArrIndex, iArrTag, iLevelCd

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
	If lgObjDragNode.Image = C_CC1 Then
		lgStrCmd  = "CC1"
		lgBlnLvlChanged = False

	ELSE
		'iStrVal = BIZ_MOVE_TREE & "?txtMode=" & parent.UID_M0002
		' -- �׷��� �̵��� ��� ������ ���������ش�.
		iLevelCd = CInt(Left(lgObjDropNode.Tag, 1)) + 1

		' ���� ������ ���� ���� check
		lgBlnLvlChanged = True
		If lgObjDropNode.Key <> C_ROOT_KEY And lgObjDragNode.parent.Key <> C_ROOT_KEY THEN
			If lgObjDropNode.parent.fullpath = lgObjDragNode.parent.parent.fullpath Then
				lgBlnLvlChanged = False
				'iStrVal = iStrVal & "&txtFlag="		& "ORG1"								' Sales Org.
			End If
		End If

		If lgBlnLvlChanged Then
			' ���ο� ���� ���� 
			iArrIndex = Split(lgObjDropNode.fullpath, parent.gColSep)
			'iStrVal = iStrVal & "&txtSalesOrgNewLvl=" & lgArrOrgLvl(Ubound(iArrIndex, 1), 0)	' Sales Org. New Level
			If Ubound(iArrIndex, 1) = lgIntLastOrgLvlIndex - 1 Then
				'iStrVal = iStrVal & "&txtEndOrgFlag=Y"
			Else
				'iStrVal = iStrVal & "&txtEndOrgFlag=N"
			End If
			
			iArrTag = Split(lgObjDragNode.Tag, parent.gColSep)
			iStrVal = iStrVal & "&txtSalesOrgCurLvl=" & iArrTag(0)								' Sales Org. Current Level

'			' ���������� ���翩�� Check
'			If lgObjDragNode.Children = 0 Then
'				iStrVal = iStrVal & "&txtFlag="	& "ORG2"								' Sales Org.
'			Else
'				'������������ 
'				IF lgObjDragNode.Child.Image = C_CC1 Then
'					iStrVal = iStrVal & "&txtFlag="	& "ORG2"							' Sales Org.
'				Else
'					lgBlnRemakeNodes = True
'					iStrVal = iStrVal & "&txtFlag="	& "ORG3"							' Sales Org.
'				End If
'			End If
		End If
		
'		iStrVal = iStrVal & "&txtSalesOrg=" & Mid(lgObjDragNode.key, 2)			' Sales Org.
		
'		If lgObjDropNode.Key = C_ROOT_KEY Then
'			iStrVal = iStrVal & "&txtUpperSalesOrg="								' Upper Sales Org.
'		Else
'			iStrVal = iStrVal & "&txtUpperSalesOrg=" & Mid(lgObjDropNode.key, 2)	' Upper Sales Org.
'		End If
'		iStrVal = iStrVal & "&txtUserId="	& parent.gUsrID
		lgStrCmd = "CCG"
	END IF

	Call LayerShowHide(0)
	frm1.uniTree1.MousePointer = 0
	
	lgSaveModFg = "R"
	
	'Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 

	With Frm1
	
		If lgStrCmd  = "CC1" Then
			iStrVal = C_CC1 &  Parent.gColSep
		Else
			iStrVal = C_CCG &  Parent.gColSep
		End If
		
		iStrVal = iStrVal & "U" &  Parent.gColSep
	
		' C/C 
		If lgStrCmd  = "CC1" Then
			
			iStrVal = iStrVal & ""							&  Parent.gColSep	' -- COST_CD
			iStrVal = iStrVal & ""							&  Parent.gColSep	' -- COST_NM
			iStrVal = iStrVal & Mid(lgObjDropNode.key, 2)	&  Parent.gColSep	' -- UPPER_COST_CD
			iStrVal = iStrVal & ""							&  Parent.gColSep	' -- LEVEL_CD
			iStrVal = iStrVal & ""							&  Parent.gColSep	' -- LEAFT_FLAG
			iStrVal = iStrVal & ""							&  Parent.gColSep	' -- TEMP_STR_1
			iStrVal = iStrVal & ""							&  Parent.gColSep	' -- TEMP_STR_2
			iStrVal = iStrVal & .txtCOST_CD_2_OLD.value		&  Parent.gColSep	' -- OLD_KEY

		ELSE

			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & Mid(lgObjDropNode.key, 2)	&  Parent.gColSep	' -- UPPER_COST_CD
			iStrVal = iStrVal & iLevelCd					&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & .txtCOST_CD_OLD.value		&  Parent.gColSep

		End If
    	'-----------------------
		'Data manipulate area
		'-----------------------
		
	End With	
 	
 	Frm1.txtSpread.value      = iStrVal
	'Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 
	Frm1.txtMode.value        =  Parent.UID_M0002									' ��������, �������� ó���� 
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 	

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
			
			' Ʈ���� �����ڽ� üũ 
			If ChkOrgTree(pvObjNode, C_ROOT_KEY) = False Then
				Select Case pvObjNode.Image
					Case C_CC1, C_CCG, C_Const
						.uniTree1.MenuEnabled C_MNU_OPEN, False
					Case Else
						.uniTree1.MenuEnabled C_MNU_OPEN, False
				End Select
				
				.uniTree1.MenuEnabled C_MNU_ADD, False
				.uniTree1.MenuEnabled C_MNU_DELETE, False
				.uniTree1.MenuEnabled C_MNU_RENAME, False
			Else
				' -- ���� �ڽ��� �ִ� ��� 
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
						Case C_CC1
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
						Case C_CCG
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
'   Event Name : uniTree1_MenuOpen - ���C/C �׷����Է� 
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
		Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, C_ROOT_KEY_STR & GetTotalCnt(pvObjNode), "�� Cost Center Group ���", C_CCG, C_CCG)

		With frm1
		
			.txtUPPER_COST_CD.value		= "*"
			.txtLEVEL_CD.value			= 1
			' ���������� �ϳ��� ��� �����������θ� 'Y'�� ���� 
			If lgIntLastOrgLvlIndex = 1 Then
				.rdoLEAF_FLAG_Y.checked = True
				Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_Y, "Q")
				Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_N, "Q")
			Else
				.rdoLEAF_FLAG_N.checked = True
				Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_Y, "N")
				Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_N, "N")
			End If
		End With
	Else
		Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, pvObjNode.Key & C_UNDERSCORE & GetTotalCnt(pvObjNode), "�� Cost Center Group ���", C_CCG, C_CCG)

		With frm1
			.txtUPPER_COST_CD.value = Mid(pvObjNode.Key,2)
			iArrTag = Split(pvObjNode.tag, parent.gColSep)
			.txtLEVEL_CD.value			= CInt(iArrTag(0)) + 1
			.rdoLEAF_FLAG_Y.checked		= False
'			For ii = 0 to lgIntLastOrgLvlIndex - 1
'				If lgArrOrgLvl(ii, 0) = iArrTag(0) then
'					.txtLEVEL_CD.value = lgArrOrgLvl(ii + 1, 0)
'
'					If (ii + 1) = (lgIntLastOrgLvlIndex - 1)Then
'						.rdoLEAF_FLAG_Y.checked = True
'						Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_Y, "Q")
'						Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_N, "Q")
'					Else
'						.rdoLEAF_FLAG_N.checked = True
'						Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_Y, "N")
'						Call ggoOper.SetReqAttr(.rdoLEAF_FLAG_N, "N")
'					End If
'					
'					Exit For
'				End If
'			Next
		End With
	End If
	
	iObjDummyNode.Selected = True	
	Set lgNewNode = iObjDummyNode
	set lgObjDragNode = iObjDummyNode
	
	Call ClickTab1()

	Call SetToolbar("1100100000001111")									'��: ��ư ���� ����		
			
	lgIntFlgMode = parent.OPMD_CMODE	' �űԷ� ��� 
	
	lgStrCmd  = "CCG"
	
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
	
	Set iObjDummyNode = frm1.uniTree1.Nodes.Add(pvObjNode.Key, tvwChild, pvObjNode.Key & C_UNDERSCORE & GetTotalCnt(pvObjNode), "�� Cost Center ���", C_CC1, C_CC1)
	
	iObjDummyNode.Selected = True
	Set lgNewNode = iObjDummyNode
	set lgObjDragNode = iObjDummyNode	
	
	Call SetToolbar("1100100000001111")									'��: ��ư ���� ����		
	 
	Call ClickTab2()

	frm1.txtUPPER_COST_CD_2.value = Mid(pvObjNode.Key,2)
	frm1.txtCOST_CD_2_OLD.value = "" 
		
	lgIntFlgMode = parent.OPMD_CMODE	' �űԷ� ��� 
	lgStrCmd  = "CC1"
		
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

	iStrSelect	= " CASE WHEN upper_cost_cd = '*' THEN  " & FilterVar(C_ROOT_KEY, "''", "S") & " ELSE " & FilterVar("O", "''", "S") & "  + upper_cost_cd END , CASE WHEN LEVEL_CD = '99' THEN " & FilterVar("O", "''", "S") & "  + cost_cd ELSE " & FilterVar("O", "''", "S") & " + cost_cd END, " & FilterVar("[", "''", "S") & "  + cost_cd + " & FilterVar("]", "''", "S") & " + cost_nm, level_cd, LEAF_FLAG,  "
	iStrSelect	= iStrSelect & " CASE WHEN COST_FLAG = 'N' THEN " & FilterVar(C_CCG, "''", "S") & " ELSE " & FilterVar(C_CC1, "''", "S") & " END"
	iStrFrom	= " dbo.ufn_c_getTreeView_C_COST_CENTER_HIERARCHY_S('*') "
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

	iStrSelect	= " CASE WHEN SO.sales_org =  " & FilterVar(Mid(lgObjDragNode.Key, 2), "''", "S") & " THEN  " & FilterVar(lgObjDropNode.Key, "''", "S") & " ELSE " & FilterVar("O", "''", "S") & "  + SO.upper_sales_org END , " & FilterVar("O", "''", "S") & "  + SO.sales_org, " & FilterVar("[", "''", "S") & "  + SO.sales_org + " & FilterVar("]", "''", "S") & " + SO.sales_org_nm, SO.lvl, SO.end_org_flag,  " & FilterVar(C_CCG, "''", "S") & " "
	iStrFrom	= " dbo.b_sales_org SO INNER JOIN  "
	iStrFrom	= iStrFrom & " (SELECT	 " & FilterVar(Mid(lgObjDragNode.Key, 2), "''", "S") & " AS sales_org "
	iStrFrom	= iStrFrom & " UNION ALL "
	iStrFrom	= iStrFrom & " SELECT leaf_org "
	iStrFrom	= iStrFrom & " FROM dbo.ufn_s_ListSalesOrgHierarchy(" & iArrTag(0) & ",  " & FilterVar(Mid(lgObjDragNode.Key, 2), "''", "S") & ",  default)) T ON (T.sales_org = SO.sales_org) "
	iStrFrom	= iStrFrom & " UNION ALL "
	iStrFrom	= iStrFrom & " SELECT " & FilterVar("O", "''", "S") & "  + SG.sales_org, " & FilterVar("G", "''", "S") & "  + SG.sales_grp, " & FilterVar("[", "''", "S") & "  + SG.sales_grp + " & FilterVar("]", "''", "S") & " + SG.sales_grp_nm, SO.lvl + 1, " & FilterVar("N", "''", "S") & " ,  " & FilterVar(C_CC1, "''", "S") & ""
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

Sub GetCostCenterLvlInfo()
		
	Dim iStrSelect, iStrFrom, iStrWhere, iStrResult 	
	Dim ii, iIntRows
	Dim iArrRow, iArrCol
	
	iStrSelect	= " cost_cd, level_cd, leaf_flag "
	iStrFrom	= " dbo.ufn_c_getTreeView_C_COST_CENTER_HIERARCHY_S('*') "
	iStrWhere	= " "
	
	If CommonQueryRs2by2(iStrSelect, iStrFrom ,  iStrWhere , iStrResult) Then 
	
		iArrRow = Split(iStrResult, parent.gColSep & parent.gRowSep)			
		iIntRows = Ubound(iArrRow,1)
		
		Redim lgArrOrgLvl(iIntRows, 1)
		
		For ii = 0 To iIntRows - 1		
			iArrCol	= Split(iArrRow(ii), parent.gColSep)			
			
			lgArrOrgLvl(ii, 0) = Trim(iArrCol(2))
			lgArrOrgLvl(ii, 1) = Trim(iArrCol(3))
		Next
		lgIntLastOrvLvl = "0" 'Trim(iArrCol(2))
		lgIntLastOrgLvlIndex =  0 'ii
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

	With Frm1
	
		If lgStrCmd  = "CC1" Then
			iStrVal = C_CC1 &  Parent.gColSep
		Else
			iStrVal = C_CCG &  Parent.gColSep
		End If
		
		iStrVal = iStrVal & "D" &  Parent.gColSep
	
		' C/C 
		If lgStrCmd  = "CC1" Then
			
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & .txtCOST_CD_2_OLD.value		&  Parent.gColSep

		ELSE

			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & .txtCOST_CD_OLD.value		&  Parent.gColSep

		End If
    	'-----------------------
		'Data manipulate area
		'-----------------------
		
	End With	
 	
 	Frm1.txtSpread.value      = iStrVal
	'Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 
	Frm1.txtMode.value        =  Parent.UID_M0002									' ��������, �������� ó���� 
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 	
	lgSaveModFg	= "D"	 	
	'Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 
End Sub

 '=========================  uniTree1_onAddImgReady()  ====================================
'	Event  Name : uniTree1_onAddImgReady()
'	Description : SetAddImageCount���� Image�� �ٿ�ε� �Ϸ�ǰ� TreeView�� ImageList�� 
'                 �߰��Ǹ� �߻��ϴ� �̺�Ʈ 
'========================================================================================= 
Sub uniTree1_onAddImgReady()
	Call DbQuery()
	'If lgBlnOrgLvlExists Then
		Call SetToolbar("1100100000001111")									'��: ��ư ���� ���� 
	'Else
		Call SetToolbar("1000000000001111")									'��: ��ư ���� ���� 
	'	Call SetDefaultScreen()
	'End If
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
	If lgStrCmd = "CCG" Then
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
	
		If lgStrCmd  = "CC1" Then
			iStrVal = C_CC1 &  Parent.gColSep
		Else
			iStrVal = C_CCG &  Parent.gColSep
		End If
		
		' C/C 
		If lgStrCmd  = "CC1" Then

			If .txtCOST_CD_2_OLD.value <> "" Then
				iStrVal = iStrVal & "U" &  Parent.gColSep
			Else
				iStrVal = iStrVal & "C" &  Parent.gColSep
			End If
	
			
			iStrVal = iStrVal & .txtCOST_CD_2.value			&  Parent.gColSep
			iStrVal = iStrVal & .txtCOST_NM_2.value			&  Parent.gColSep
			iStrVal = iStrVal & .txtUPPER_COST_CD_2.value	&  Parent.gColSep
			iStrVal = iStrVal & "99"						&  Parent.gColSep
			iStrVal = iStrVal & "Y"							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & .txtCOST_CD_2_OLD.value		&  Parent.gColSep

		ELSE

			If .txtCOST_CD_OLD.value <> "" Then
				iStrVal = iStrVal & "U" &  Parent.gColSep
			Else
				iStrVal = iStrVal & "C" &  Parent.gColSep
			End If

			iStrVal = iStrVal & .txtCOST_CD.value			&  Parent.gColSep
			iStrVal = iStrVal & .txtCOST_NM.value			&  Parent.gColSep
			iStrVal = iStrVal & .txtUPPER_COST_CD.value		&  Parent.gColSep
			iStrVal = iStrVal & .txtLEVEL_CD.value 			&  Parent.gColSep

			If .rdoLEAF_FLAG_Y.checked Then
				iStrVal = iStrVal & "Y"						&  Parent.gColSep
			Else
				iStrVal = iStrVal & "N"						&  Parent.gColSep
			End If

			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & ""							&  Parent.gColSep
			iStrVal = iStrVal & .txtCOST_CD_OLD.value		&  Parent.gColSep

		End If
    	'-----------------------
		'Data manipulate area
		'-----------------------
		
	End With	
 	
 	Frm1.txtSpread.value      = iStrVal
	'Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 
	Frm1.txtMode.value        =  Parent.UID_M0002
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 	
	
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
	
	' ���C/C �׷��� �Է� 
	IF lgSaveModFg	= "O" Then	
		With frm1
			lgObjDragNode.Key = "O" & UCase(Trim(.txtCOST_CD.value))
			lgObjDragNode.text = "[" & UCase(Trim(.txtCOST_CD.value)) & "]" & .txtCOST_NM.value
			If .rdoLEAF_FLAG_N.checked Then
				lgObjDragNode.Tag = .txtLEVEL_CD.value & parent.gColSep & "N"
			Else
				lgObjDragNode.Tag = .txtLEVEL_CD.value & parent.gColSep & "Y"
			End If
		End With
	END IF	
	
	' C/C �Է� 
	IF lgSaveModFg	= "G" Then
		With frm1
			lgObjDragNode.Key = "G" & UCase(Trim(.txtCOST_CD_2.value))
			lgObjDragNode.text =  "[" & UCase(Trim(.txtCOST_CD_2.value)) & "]" & .txtCOST_NM_2.value
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
			If lgStrCmd = "CCG" Then
				If .rdoLEAF_FLAG_N.checked Then
					.uniTree1.selecteditem.Tag = .txtLEVEL_CD.value & parent.gColSep & "N"
				Else
					.uniTree1.selecteditem.Tag = .txtLEVEL_CD.value & parent.gColSep & "Y"
				End If
				
				iStrText = "[" & .txtCOST_CD.value & "]" & Trim(.txtCOST_NM.value)
				If Trim(.uniTree1.selecteditem.Text) <> iStrText Then
					.uniTree1.selecteditem.Text = iStrText
				End If

				' -- Ű�� ����� ��� Ʈ����Ű�� �����Ѵ�.
				If Mid(.uniTree1.selecteditem.Key,2) <> Trim(.txtCOST_CD.value) Then
					.uniTree1.selecteditem.Key = "O" & Trim(.txtCOST_CD.value)
				End If

			Else
				iStrText = "[" & .txtCOST_CD_2.value & "]" & Trim(.txtCOST_NM_2.value)
				If Trim(.uniTree1.selecteditem.Text) <> iStrText Then
					.uniTree1.selecteditem.Text = iStrText
				End If

				' -- Ű�� ����� ��� Ʈ����Ű�� �����Ѵ�.
				If Mid(.uniTree1.selecteditem.Key,2) <> Trim(.txtCOST_CD_2.value) Then
					.uniTree1.selecteditem.Key = "O" & Trim(.txtCOST_CD_2.value)
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=uniTree1 width=100% height=100% <%=UNI2KTV_IDVER%>> <PARAM NAME="ImageWidth" VALUE="16">  <PARAM NAME="ImageHeight" VALUE="16">  <PARAM NAME="LineStyle" VALUE="1"> <PARAM NAME="Style" VALUE="7">  <PARAM NAME="LabelEdit" VALUE="1">  </OBJECT>');</SCRIPT>
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
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>Cost Center Group</font></td>
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
													</TR>
												</TABLE>
											</TD>
											<TD CLASS="CLSMTABP">
												<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23" ></td>
														<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>Cost Center</font></td>
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
														<TR HEIGHT="20">
															<TD CLASS="TD5" NOWRAP>Cost Center Group</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtCOST_CD" TYPE="Text" MAXLENGTH="10" tag="23XXXU" size="20" ALT="Cost Center Group">
															<input NAME="txtCOST_CD_OLD" TYPE="hidden" MAXLENGTH="10" tag="21XXX" size="20" ALT="Cost Center Group"></TD>
														</TR>
														<TR HEIGHT="20">
															<TD CLASS="TD5" NOWRAP>Cost Center Group��</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtCOST_NM" TYPE="Text" MAXLENGTH="20" tag="22XXX" size="30" ALT="Cost Center Group��"></TD>
														</TR>
														<TR HEIGHT="20">
															<TD CLASS="TD5" NOWRAP>LEVEL</TD>
															<TD CLASS="TD656" NOWRAP><input NAME="txtLEVEL_CD" TYPE="Text" MAXLENGTH="2" tag="24XXXU" size="10"></TD>
														</TR>
														<TR HEIGHT="20">
															<TD CLASS="TD5" NOWRAP>�����׷�</TD>
															<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtUPPER_COST_CD" MAXLENGTH=10 TAG="24XXXU" size="30"></TD>
														</TR>
														<TR HEIGHT="20">
															<TD CLASS="TD5" NOWRAP>���� ����</TD>
															<TD CLASS="TD656" NOWRAP>
																<input type=radio CLASS="RADIO" id=rdoLEAF_FLAG_Y name="rdoLEAF_FLAG" value="Y" tag = "21">
																	<label for="rdoLEAF_FLAG_Y">��</label>&nbsp;&nbsp;&nbsp;&nbsp;
																<input type=radio CLASS = "RADIO" id=rdoLEAF_FLAG_N name="rdoLEAF_FLAG" value="N" tag = "21" checked>
																	<label for="rdoLEAF_FLAG_N">�ƴϿ�</label></TD>
															</TD>
														</TR>
														<TR HEIGHT="*">
														  <TD CLASS="TD5" NOWRAP> </TD>
														  <TD CLASS="TD656" NOWRAP> </TD>
														</TR>
													</TABLE>
												</DIV> 
												<!-- �ι�° �� ����  -->
												<DIV ID="TabDiv" SCROLL=no>
													<TABLE <%=LR_SPACE_TYPE_60%>>
														<TR HEIGHT="20">
														  <TD CLASS="TD5" NOWRAP>Cost Center</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtCOST_CD_2" TYPE="Text" MAXLENGTH="10" tag="32XXX" size="20" ALT="Cost Center">
														  <IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCostCenter(0)">
														  <input NAME="txtCOST_CD_2_OLD" TYPE="hidden" MAXLENGTH="10" tag="31XXX" size="20" ALT="Cost Center"></TD>
														</TR>
														<TR HEIGHT="20">
														  <TD CLASS="TD5" NOWRAP>Cost Center��</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtCOST_NM_2" TYPE="Text" MAXLENGTH="20" tag="34XXX" size="30"></TD>
														</TR>
														<TR HEIGHT="20">
														  <TD CLASS="TD5" NOWRAP>Cost Center Group</TD>
														  <TD CLASS="TD656" NOWRAP><input NAME="txtUPPER_COST_CD_2" TYPE="Text" MAXLENGTH="20" tag="34XXX" size="30"></TD>
														</TR>
														<TR HEIGHT="*">
														  <TD CLASS="TD5" NOWRAP></TD>
														  <TD CLASS="TD656" NOWRAP></TD>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX=-1>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:none" TABINDEX=-1></TEXTAREA>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>

