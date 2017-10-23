<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : b1b03ma2.asp
'*  4. Program Name         : ǰ��׷���ȸ 
'*  5. Program Desc         :
'*  6. Component List       : ADO
'*  7. Modified date(First) : 1999/12/12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit   

'========================================================================================================
'=                       1.2.1 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
                                                       

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Const BIZ_PGM_QRY_ID				= "b1b03mb4.asp"			
Const BIZ_PGM_LOOKUPITEMGROUP_ID	= "b1b03mb5.asp"			
Const BIZ_PGM_LOOKUPITEM_ID			= "b1b03mb6.asp"			

Const TAB1 = 1
Const TAB2 = 2

Dim IsOpenPop						 'Popup
Dim gSelframeFlg

Dim lgCurNode

Const C_Sep  = "/"

Const C_GROUP  = "GROUP"
Const C_OPEN = "OPEN"
Const C_PROD  = "PROD"
Const C_MATL  = "MATL"
Const C_PHANTOM = "PHANTOM"
Const C_ASSEMBLY = "ASSEMBLY"

Const C_IMG_GROUP = "../../../CShared/image/Group.gif"
Const C_IMG_OPEN = "../../../CShared/image/Group_op.gif"
Const C_IMG_PROD = "../../../CShared/image/product.gif"
Const C_IMG_MATL = "../../../CShared/image/material.gif"
Const C_IMG_PHANTOM = "../../../CShared/image/phantom.gif"
Const C_IMG_ASSEMBLY = "../../../CShared/image/subcon.gif"

Const tvwChild = 4


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
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgCurNode = 0
    
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ===================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()

End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================

Sub LoadInfTB19029()
	
	<!--#Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<%Call LoadInfTB19029A("Q", "*","NOCOOKIE", "MA")%>
	
End Sub


'========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5, lgF6)
    call SetCombo2(frm1.cboItemAcct ,lgF0, lgF1, chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1002", "''", "S") & "  " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5, lgF6)
    call SetCombo2(frm1.cboItemClass  ,lgF0, lgF1, chr(11))
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
	
	If gSelframeFlg = TAB1 Then Exit Function
	 
	Call changeTabs(TAB1)	
	gSelframeFlg = TAB1

	'++++++++++++  Insert Your Code  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'Call SetToolBar()
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)	
	gSelframeFlg = TAB2
	'++++++++++++  Insert Your Code  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ %>
	'Call SetToolBar()
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ %>
    	
End Function

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenItemGroup()  --------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  " 			
	arrParam(5) = "ǰ��׷�"			
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
    
    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array( arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
	
End Function

'------------------------------------------  SetItemGroup()  ---------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)		
	frm1.txtItemGroupNm.Value    = arrRet(1)		
End Function

'========================================================================================
' Function Name : InitTreeImage
' Function Desc : �̹��� �ʱ�ȭ 
'========================================================================================
Function InitTreeImage()
	Dim NodX, lHwnd
	
	With frm1

	.uniTree1.SetAddImageCount = 6
	.uniTree1.Indentation = "200"
	
    .uniTree1.AddImage C_IMG_GROUP, C_GROUP, 0												'��: TreeView�� ���� �̹��� ���� 
	.uniTree1.AddImage C_IMG_OPEN, C_OPEN, 0
	.uniTree1.AddImage C_IMG_PROD, C_PROD, 0												'��: TreeView�� ���� �̹��� ���� 
	.uniTree1.AddImage C_IMG_MATL, C_MATL, 0
	.uniTree1.AddImage C_IMG_ASSEMBLY, C_ASSEMBLY, 0												'��: TreeView�� ���� �̹��� ���� 
	.uniTree1.AddImage C_IMG_PHANTOM, C_PHANTOM, 0
	
	.uniTree1.OLEDragMode = 0														'��: Drag & Drop �� �����ϰ� �� ���ΰ� ���� 
	.uniTree1.OLEDropMode = 0
	
	End With

End Function

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
Sub Form_Load()
	
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitComboBox
    Call InitVariables                                                      '��: Initializes local global variables
       
    '----------  Coding part  -------------------------------------------------------------
        
    Call SetToolbar("11000000000011")									'��: ��ư ���� ���� 
    
    gTabMaxCnt = 2
    gIsTab = "Y"
   
    Call InitTreeImage
    
    frm1.txtItemGroupCd.focus
    Set gActiveElement = document.activeElement 
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'==========================================================================================
'   Event Name : cboBDG_CTRL_FG_onchange()
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================

Sub txtItemGroupCd_onChange()
	If frm1.txtItemGroupCd.value = "" Then
		frm1.txtItemGroupNm.value = ""
	End If	
End Sub

Sub LookUpItemGroup(ByVal txtItemGroup)

    Err.Clear                                                               
    
    Call ggoOper.ClearField(Document, "2")									
    
    Call LayerShowHide(1)
    
    Dim strVal
      
    strVal = BIZ_PGM_LOOKUPITEMGROUP_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtItemGroupCd=" & Trim(txtItemGroup)
    strVal = strVal & "&PrevNextFlg=" & ""
        
	Call RunMyBizASP(MyBizASP, strVal)										

	Call ClickTab1()
End Sub

Sub LookUpItemGroupOk()
End Sub

Sub LookUpItem(ByVal txtItem, ByVal intLevel)
    Err.Clear                                                               
    Call ggoOper.ClearField(Document, "2")									
    
    Call LayerShowHide(1)													
        
    Dim strVal
    strVal = BIZ_PGM_LOOKUPITEM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtItemCd=" & Trim(txtItem)
    strVal = strVal & "&txtLevelCd=" & Trim(intLevel)
    strVal = strVal & "&txtRootLevel=" & Trim(frm1.hRootLevel.value)
	strVal = strVal & "&PrevNextFlg=" & ""
	    
	Call RunMyBizASP(MyBizASP, strVal)										
	
	Call ClickTab2()

End Sub

Sub LookUpItemOk()
	
End Sub
'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*****************************************************************************************************

'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node�� Ŭ���ϸ� �߻� �̺�Ʈ 
'==========================================================================================

Sub uniTree1_NodeClick(Node)
    Dim NodX
    
	Dim iPos1
	Dim iPos2
	Dim iPos3
	
	Dim ItemFlg
	Dim intLevel
	Dim txtKey
	
	Dim prntNode
	
	Err.Clear                                                              
		
	With frm1
	
    Set NodX = .uniTree1.SelectedItem
        
	If lgCurNode = NodX.Index Then Exit Sub
	
	lgCurNode = NodX.Index
        
    If Not NodX Is Nothing Then ' ���õ� ������ ������ 

		'-------------------------------------
		'Hidden Value Init
		'---------------------------------------
		
		Set PrntNode = NodX.Parent
		
		If PrntNode is Nothing Then	' Root�� ��� 
			
			'--------------------------------------
			'Item Group Key
			'--------------------------------------				
			
			txtKey  = Trim(NodX.Text)

			Call LookUpItemGroup(txtKey) 
			
		Else

			'--------------------------------------
			'Item/Item Group Flag
			'--------------------------------------				
			iPos1 = InStr(1,NodX.Key, "|^|^|")       
			ItemFlg = Mid(NodX.Key, 1, iPos1 - 1)

			'--------------------------------------
			'Level
			'--------------------------------------				
			iPos2 = InStr(iPos1 + 5, NodX.Key, "|^|^|")								'Child Item Seq
			intLevel = Trim(Mid(NodX.Key, iPos1 + 5, iPos2 - (iPos1 + 5)))
		    
			'--------------------------------------
			'Item Group Key
			'--------------------------------------				
		    
		    txtKey = Trim(NodX.Text)
		    
		    If CInt(ItemFlg) = 0 Then
				Call LookUpItemGroup(txtKey)
		    Else
				Call LookUpItem(txtKey, intLevel)
		    End If
		End IF
	End If
    
    Set NodX = Nothing
    Set PrntNode = Nothing
    
    End With

	
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

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
    '-----------------------
    'Erase contents area
    '-----------------------
   	If frm1.txtItemGroupCd.value = "" Then
		frm1.txtItemGroupNm.value = ""
	End If	
	
	frm1.uniTree1.Nodes.Clear
    							
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    
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
    Call ClickTab1
    
    If DbQuery = False Then   
		Exit Function           
    End If 
    														'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    
End Function
'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next 	
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next                                                 '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
    On Error Resume Next 	
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow()
    On Error Resume Next    
End Function
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.fncExport(parent.C_SINGLE)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                  
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'*********************************************************************************************************
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Err.Clear                                                              
    
    DbQuery = False                                                        
	
	Call LayerShowHide(1)													
	    
    Dim strVal
    
    frm1.txtUpdtUserId.value= parent.gUsrID
      
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtPlantCd=" & "****"
    strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
    strVal = strVal & "&txtUpdtUserId=" & Trim(frm1.txtUpdtUserId.value)
    strVal = strVal & "&txtSrchType=" & "2"
        
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbQuery = True                                                          

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												
	
    '-----------------------
    'Reset variables area
    '-----------------------
	Dim NodX
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Set NodX = frm1.uniTree1
		NodX.SetFocus
		Set NodX = Nothing
		Set gActiveElement = document.activeElement
    End If
    
    lgIntFlgMode = parent.OPMD_UMODE										
        
    Call ggoOper.LockField(Document, "Q")							
	Call SetToolbar("11000000000111")								
	
	
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()     

End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													
	
End Function
'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function
'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()											
End Function


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
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
						<TABLE CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ��׷���ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="12XXXU"  ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()" >&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=40 tag="14"></TD>									
								</TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>								
								<!-- TreeView AREA -->
								<TD HEIGHT=100% WIDTH=30%>
									<script language =javascript src='./js/b1b03ma2_uniTree1_N487839591.js'></script>
								</TD>
								<!-- DATA AREA -->
								<TD WIDTH="70%" HEIGHT="100%">
									<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
										<TR HEIGHT=23>
											<TD WIDTH="100%">
												<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH="100%" border=0>
													<TR>
														<TD WIDTH=10>&nbsp;</TD>
														<TD CLASS="CLSMTABP">
															<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
																<TR>
																	<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
																	<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>ǰ��׷�����</font></td>
																	<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
															    </TR>
															</TABLE>
														</TD>
														<TD CLASS="CLSMTABP">
															<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
																<TR>
																	<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>ǰ������</font></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
															    </TR>
															</TABLE>
														</TD>
														<TD WIDTH=300>&nbsp;</TD>
														<TD WIDTH=*>&nbsp;</TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD WIDTH="100%" CLASS="TB3">
												<!-- ù��° �� ���� -->
												<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
													<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD WIDTH=100% HEIGHT=* valign=top>
																<TABLE CLASS="TB2" CELLSPACING=0 CELLPADDING=0 HEIGHT=100%>
																	<TR>
																		<TD CLASS=TD5 HEIGHT=5 WIDTH="100%"></TD>
																		<TD CLASS=TD6 HEIGHT=5 WIDTH="100%"></TD>												
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtlevel1" SIZE=5 tag="24" ALT="����"></TD>												
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemGroupCd1" SIZE=20 tag="24" ALT="ǰ��׷�"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ��׷��</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemGroupNm1" SIZE=40 tag="24" ALT="ǰ��׷��"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>����ǰ��׷�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtUpperItemGroupCd" SIZE=20 tag="24" ALT="����ǰ��׷�"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>����ǰ��׷��</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtUpperItemGroupNm" SIZE=40 tag="24" ALT="����ǰ��׷��"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>������ǰ��׷쿩��</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLowItemGroupFlg" tag="24" ID="rdoLowItemGroupFlg1" VALUE="Y"><LABEL FOR="rdoLowItemGroupFlg1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLowItemGroupFlg" tag="24" ID="rdoLowItemGroupFlg2" VALUE="N"><LABEL FOR="rdoLowItemGroupFlg2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtValidFromDt1" SIZE=10 STYLE="TEXT-ALIGN: center" tag="24" ALT="��ȿ�Ⱓ������"> ~ <INPUT TYPE=TEXT NAME="txtValidToDt1" SIZE=10 STYLE="TEXT-ALIGN: center" tag="24" ALT="��ȿ�Ⱓ������"></TD>
																	</TR>											
																	<TR>
																		<TD CLASS=TD5 HEIGHT=200 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD6 HEIGHT=200  NOWRAP>&nbsp;</TD>
																	</TR>	
																</TABLE>
															</TD>
														</TR>										
													</TABLE>
												</DIV> 
												<!-- �ι�° �� ���� -->
												<DIV ID="TabDiv"  SCROLL="no" style="display:none">
													<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD WIDTH=100% HEIGHT=* valign=top>
																<TABLE CLASS="TB2" CELLSPACING=0 CELLPADDING=0 HEIGHT=100%>
																	<TR>
																		<TD CLASS=TD5 HEIGHT=5 WIDTH="100%" NOWRAP></TD>
																		<TD CLASS=TD6 HEIGHT=5 WIDTH="100%" NOWRAP></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT NAME="txtlevel2" SIZE=5 tag="24" ALT="����"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ��</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 tag="24" ALT="ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 MAXLENGTH=40 tag="24" ALT="ǰ���"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ�����ĸ�Ī</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemDesc" SIZE=60 tag="24" ALT="ǰ�����ĸ�Ī"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
																		<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" ALT="ǰ�����" STYLE="Width: 98px;" tag="24"></SELECT></TD>
																	</TR>
																	<TR>	
																		<TD CLASS=TD5 NOWRAP>���ش���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasicUnit" SIZE=5 tag="24"  ALT="���ش���"></TD>
																	</TR>											
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd2" SIZE=20 tag="24" ALT="ǰ��׷�">&nbsp;<INPUT NAME="txtItemGroupNm2" MAXLENGTH=40 SIZE=40 tag=24" ALT="ǰ��׷��"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>Phantom����</TD>
																		<TD CLASS=TD6 NOWRAP>
																			<INPUT TYPE="RADIO" NAME="rdoPhantomType" ID="rdoPhantomType1" Value="Y" CLASS="RADIO" tag="24"><LABEL FOR="rdoPhantomType1">��</LABEL>
																			<INPUT TYPE="RADIO" NAME="rdoPhantomType" ID="rdoPhantomType2" Value="N" CLASS="RADIO" tag="24"><LABEL FOR="rdoPhantomType2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>���ձ��ű���</TD>
																		<TD CLASS=TD6 NOWRAP>
																			<INPUT TYPE="RADIO" NAME="rdoUnifyPurFlg" ID="rdoUnifyPurFlg1" Value="Y" CLASS="RADIO" tag="24"><LABEL FOR="rdoUnifyPurFlg1">��</LABEL>
																			<INPUT TYPE="RADIO" NAME="rdoUnifyPurFlg" ID="rdoUnifyPurFlg2" Value="N" CLASS="RADIO" tag="24"><LABEL FOR="rdoUnifyPurFlg2">�ƴϿ�</LABEL></TD>
																	</TR>											
																	<TR>
																		<TD CLASS=TD5 NOWRAP>����ǰ��</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasisItemCd" SIZE=20 tag="24" ALT="����ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtBasisItemNm" SIZE=40 MAXLENGTH=40 tag="24" ALT="����ǰ���"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�����ǰ��Ŭ����</TD>
																		<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemClass" ALT="�����ǰ��Ŭ����" STYLE="Width: 98px;" tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ��԰�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=50 tag="24" ALT="ǰ��԰�"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>Net�߷�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWeight" SIZE=20 STYLE="TEXT-ALIGN: right" tag="24X3" ALT="Net�߷�">&nbsp;<INPUT TYPE=TEXT NAME="txtWeightUnit" SIZE=5 MAXLENGTH=3 tag="24"  ALT="����"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>Gross�߷�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtGrossWeight" SIZE=20  tag="24X3" ALT="Gross�߷�" STYLE="TEXT-ALIGN: right">&nbsp;<INPUT TYPE=TEXT NAME="txtGrossWeightUnit" align=top SIZE=5 MAXLENGTH=3  tag="24" ALT="����"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>CBM(����)</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCBM" SIZE=20  tag="24X3" ALT="CBM(����)" STYLE="TEXT-ALIGN: right">&nbsp;<INPUT TYPE=TEXT NAME="txtCBMInfo" align=top SIZE=40 MAXLENGTH=50  tag="24" ALT="CBM����"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�����ȣ</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDrawNo" SIZE=30 tag="24" ALT="�����ȣ"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>HS�ڵ�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHSCd" SIZE=20 tag="24" ALT="HS�ڵ�">&nbsp;<INPUT TYPE=TEXT NAME="txtHSUnit" SIZE=5 MAXLENGTH=3 tag="24"  ALT="HS����"></TD>								
																	</TR>
																	<TR>	
																		<TD CLASS=TD5 NOWRAP>ǰ���������</TD>
																		<TD CLASS=TD6 NOWRAP>
																			<INPUT TYPE="RADIO" NAME="rdoPhoto" ID="rdoPhoto1" Value="Y" CLASS="RADIO" tag="24"><LABEL FOR="rdoPhoto1">��</LABEL>
															 				<INPUT TYPE="RADIO" NAME="rdoPhoto" ID="rdoPhoto2" Value="N" CLASS="RADIO" tag="24"><LABEL FOR="rdoPhoto2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>��ȿ����</TD>
																		<TD CLASS=TD6 NOWRAP>
																			<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg1" Value="Y" CLASS="RADIO" tag="24"><LABEL FOR="rdoValidFlg1">��</LABEL>
																			<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg2" Value="N" CLASS="RADIO" tag="24"><LABEL FOR="rdoValidFlg2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtValidFromDt2" SIZE=10 STYLE="TEXT-ALIGN: center" tag="24" ALT="��ȿ�Ⱓ������"> ~ <INPUT TYPE=TEXT NAME="txtValidToDt2" SIZE=10 STYLE="TEXT-ALIGN: center" tag="24" ALT="��ȿ�Ⱓ������"></TD>
																	</TR>											
																</TABLE>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=20 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME></TD>
	</TR>
</TABLE><TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtlgMode" tag="24">
<INPUT TYPE=hidden NAME="hRootLevel" tag="14">
<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24">
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>
