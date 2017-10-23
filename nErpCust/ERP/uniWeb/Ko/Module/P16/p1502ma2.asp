
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name		: Production
'*  2. Function Name	: 
'*  3. Program ID		: p1502ma2.asp
'*  4. Program Name		: �ڿ��׷� ��ȸ 
'*  5. Program Desc		:
'*  6. Comproxy List	: +B19029LookupNumericFormat
'*  7. Modified date(First)	: 2001/11/29
'*  8. Modified date(Last)	: 2002/11/18
'*  9. Modifier (First)		: Jung Yu Kyung
'* 10. Modifier (Last)		: Ryu Sung Won
'* 11. Comment		:
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'<Script LANGUAGE="vbscript"	  SRC="../../inc/incUni2KTV.vbs"></Script>
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<!--'==========================================  1.1.2 ���� Include   ======================================
'============================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID				= "p1502mb9.asp"			

Const C_Sep  = "/"

Const C_GROUP  = "GROUP"
Const C_OPEN = "OPEN"
Const C_PROD  = "PROD"
Const C_MATL  = "MATL"
Const C_PHANTOM ="PHANTOM"
Const C_ASSEMBLY = "ASSEMBLY"
Const C_SUBCON  = "SUBCON"

Const C_IMG_GROUP = "../../../CShared/image/Group.gif"
Const C_IMG_OPEN = "../../../CShared/image/Group_op.gif"
Const C_IMG_PROD = "../../../CShared/image/product.gif"
Const C_IMG_MATL = "../../../CShared/image/material.gif"
Const C_IMG_PHANTOM = "../../../CShared/image/phantom.gif"
Const C_IMG_ASSEMBLY = "../../../CShared/image/subcon.gif"
Const C_IMG_SUBCON = "../../../CShared/image/product.gif"

Const tvwChild = 4
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop						 'Popup
Dim gSelframeFlg
Dim lgCurNode
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
    
    '---- Coding part--------------------------------------------------------------------
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
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub

'========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

End Sub

'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�.
'*********************************************************************************************************

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    arrField(2) = "CUR_CD"
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    arrHeader(2) = "��ȭ�ڵ�"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenResourceGroup()  -------------------------------------------------
'	Name : OpenResourceGroup()
'	Description : ResourceGroup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResourceGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "�ڿ��׷��˾�"	
	arrParam(1) = "P_RESOURCE_GROUP"				
	arrParam(2) = Trim(frm1.txtResourceGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " 
				  			
	arrParam(5) = "�ڿ��׷�"			
	    
    arrField(0) = "RESOURCE_GROUP_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ��׷�"		
    arrHeader(1) = "�ڿ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetResourceGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtResourceGroupCd.focus
	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
End Function

'------------------------------------------  SetResourceGroup()  --------------------------------------------------
'	Name : SetResourceGroup()
'	Description : ResourceGroup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResourceGroup(byval arrRet)
	frm1.txtResourceGroupCd.Value    = arrRet(0)		
	frm1.txtResourceGroupNm.Value    = arrRet(1)		
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
	.uniTree1.AddImage C_IMG_SUBCON, C_SUBCON, 0
	
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

	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtResourceGroupCd.focus()
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If
       
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11000000000011")									'��: ��ư ���� ���� 
    
    gTabMaxCnt = 2
    gIsTab = "Y"
   
    Call InitTreeImage
    
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
Sub LookUpResource(ByVal txtResource, ByVal intLevel)
    Err.Clear                                                               
    
    Call ggoOper.ClearField(Document, "2")									
    
    Call LayerShowHide(1)													
        
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0003				
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				
    strVal = strVal & "&txtResourceCd=" & Trim(txtResource)						
    strVal = strVal & "&txtLevelCd=" & intLevel								
	strVal = strVal & "&PrevNextFlg=" & ""
	    
	Call RunMyBizASP(MyBizASP, strVal)										
	
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
	
	Dim ResourceFlg
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
		
		If NOT PrntNode is Nothing Then	' Root�� �ƴ� ��� 
			'--------------------------------------
			'Resource Group Key
			'--------------------------------------				
		    
		    txtKey = Trim(NodX.Text)
		    
		  	Call LookUpResource(txtKey,intLevel)
		 
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
   	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtResourceGroupCd.value = "" Then
		frm1.txtResourceGroupNm.value = ""
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
    
    If DbQuery = False Then   
		Exit Function           
    End If 
    														'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    
End Function
'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
End Function
'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
	Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
	
	iColumnLimit = frm1.vspdData.MaxCols-1
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit : Frm1.vspdData.Row = 0  :	iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			 '��: Data is changed.  Do you want to exit? 
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
Function DbQuery() 
    Err.Clear                                                              
    
    DbQuery = False                                                        
	
	Call LayerShowHide(1)													
	    
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value )					
    strVal = strVal & "&txtResourceGroupCd=" & Trim(frm1.txtResourceGroupCd.value)	
    strVal = strVal & "&txtSrchType=" & "2"
        
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbQuery = True                                                          

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												
    Dim NodX
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Set NodX = frm1.uniTree1
		NodX.SetFocus
		Set NodX = Nothing
		Set gActiveElement = document.activeElement
    End If
    
    lgIntFlgMode = parent.OPMD_UMODE										
        
    Call ggoOper.LockField(Document, "Q")							
	Call SetToolbar("11000000000111")								
	
	lgCurNode = 1	
	 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ڿ��׷���ȸ</font></td>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU"  ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()" >&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP>�ڿ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceGroupCd" SIZE=15 MAXLENGTH=10 tag="12XXXU"  ALT="�ڿ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenResourceGroup()" >&nbsp;<INPUT TYPE=TEXT NAME="txtResourceGroupNm" SIZE=30 tag="14"></TD>									
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
									<script language =javascript src='./js/p1502ma2_uniTree1_N572672176.js'></script>
								</TD>
								<!-- DATA AREA -->
								<TD WIDTH=* HEIGHT="100%">
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ڿ�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=20 MAXLENGTH=10 tag="24" ALT="�ڿ�">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=30 MAXLENGTH=40 tag="24" ALT="�ڿ���"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ڿ��׷�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceGroupCd2" SIZE=20 MAXLENGTH=10 tag="24" ALT="�ڿ��׷�">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceGroupNm2" SIZE=30 tag="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ڿ�����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceType" SIZE=20 MAXLENGTH=10 tag="24" ALT="�ڿ�����"></TD>																								
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ڿ���</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p1502ma2_I535131281_txtNoOfResource.js'></script>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>ȿ��</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p1502ma2_I856668007_txtEfficiency.js'></script>&nbsp;%
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p1502ma2_I310801726_txtUtilization.js'></script>&nbsp;%
											</TD>
										</TR>
										<TR ID=Q1>
											<TD CLASS=TD5 NOWRAP>RCCP���ϰ����</TD>
											<TD CLASS=TD6 NOWRAP>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="24" ID="rdoRunRccp1"><LABEL FOR="rdoRunRccp1">��</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="24" ID="rdoRunRccp2" checked><LABEL FOR="rdoRunRccp2">�ƴϿ�</LABEL></TD>
										</TR>		
										<TR ID=Q3>
											<TD CLASS=TD5 NOWRAP>CRP���ϰ����</TD>
											<TD CLASS=TD6 NOWRAP>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="24" ID="rdoRunCrp1"><LABEL FOR="rdoRunCrp1">��</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="24" ID="rdoRunCrp2" checked><LABEL FOR="rdoRunCrp2">�ƴϿ�</LABEL></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����������</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p1502ma2_I185918879_txtOverloadTol.js'></script>&nbsp;%
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ڿ����ؼ���</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p1502ma2_I166762489_txtResourceEa.js'></script>												
											</TD>
										</TR>																																	
										<TR>
											<TD CLASS=TD5 NOWRAP>�ڿ����ش���</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceUnitCd" SIZE=5 MAXLENGTH=3 tag="24" ALT="�ڿ����ش���"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���ش����� �����������</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLPADDING=0 CELLSPACING=0>
													<TR>
														<TD>															
															<script language =javascript src='./js/p1502ma2_I449396455_txtMfgCost.js'></script>
														</TD>
														<TD>											
															&nbsp;<INPUT TYPE=TEXT NAME="txtCurCd" tag=24 SIZE=5 MAXLENGTH=3 ALT="��ȭ�ڵ�">&nbsp;/&nbsp;
														</TD>
														<TD>
															<script language =javascript src='./js/p1502ma2_I894750171_txtResourceEa1.js'></script>												
														</TD>
														<TD>
															&nbsp;<INPUT TYPE=TEXT NAME="txtResourceUnitCd1" SIZE=5 MAXLENGTH=3 tag="24XXXU" ALT="�ڿ����ش���">
														</TD>
													</TR>
												</TABLE>												
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE=TEXT NAME="txtValidFromDt" CLASS=FPDTYYYYMMDD tag="24" ALT="��ȿ�Ⱓ������">&nbsp;~&nbsp;
												<INPUT TYPE=TEXT NAME="txtValidToDt" CLASS=FPDTYYYYMMDD tag="24" ALT="��ȿ�Ⱓ������">
											</TD>
										</TR>																		
										<TR>
											<TD CLASS=TD5 HEIGHT=100 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 HEIGHT=100 NOWRAP>&nbsp;</TD>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=20 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE><TEXTAREA class=hidden name=txtSpread tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TABINDEX="-1">
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
