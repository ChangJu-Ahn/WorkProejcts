<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1413MA5
'*  4. Program Name         : �������ý��� �˻��� �跮������ ����ȭ�� 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/07/30
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Koh Jae Woo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>T�跮 ������ ���ø� �˻��� ����</TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit  

Const BIZ_PGM_QRY_ID = "q1413Mb5.asp"							'��: Query �����Ͻ� ���� ASP�� 
Const PGM_JUMP_ID1 = "q1411ma1"
Const PGM_JUMP_ID3 = "Q1442MA1.asp"

Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 

Dim lgInspClassCd
Dim lgPlantCd
Dim lgItemCd
Dim lgInspReqNo
Dim lgBpCd

Dim hPlantCd
Dim hItemCd
Dim hInspReqNo
Dim hBpCd

Dim arrParam					 '--- First Parameter Group 
Dim arrReturn					 '--- Return Parameter Group 

Dim IsOpenPop          

<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================

Function InitVariables()
	
End Function

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'========================================== 2.2.1 SetDefaultVal() ======================================== 
' Name : SetDefaultVal() 
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting) 
'=========================================================================================================  
Sub SetDefaultVal() 
	With frm1
		.txtAlpha.Text = UNIFormatNumber(5, 2, -2, 0, 3, 0)
		.txtBeta.Text = UNIFormatNumber(10, 2, -2, 0, 3, 0)
	
		'/* Issue: �ʱⰪ�� �ȵ����� ���� - START */
		.txtP1.AllowNull = True
		.txtP1.Text = ""
		.txtP2.AllowNull = True
		.txtP2.Text = ""
		.txtSD.AllowNull = True
		.txtSD.Text = ""
		'/* Issue: �ʱⰪ�� �ȵ����� ���� - END */
	End With
End Sub 

'=============================================  2.3.3()  ======================================
'=	Event Name : ReturnClick
'=	Event Desc :
'========================================================================================================
Function ReturnClick()
	PgmJump(PGM_JUMP_ID1)
End Function
 
'=========================================== 2.3.1 ResultClick() ========================================== 
'= Name : ResultClick()  
'= Description : Return Array to Opener Window when OK button click = 
'======================================================================================================== 
Function ResultClick() 

	Dim strVal 
	If Not chkField(Document, "2") Then  '��: Check contents area
    	Exit Function
    End If
    	
	IF frm1.rdoSTDack.rdoSTDack1.checked = true then
		frm1.txtSTDack.Value = "0"		'ǥ�������� �˰� �ִ�.		
	End IF
	If frm1.rdoSTDack.rdoSTDack2.checked = true then
		frm1.txtSTDack.Value = "1" 		'ǥ�������� ���� ���Ѵ�.
	End IF
		
	IF frm1.rdoInsCri.rdoInsCri1.checked = true then
		frm1.txtInsCri.Value = "0"		'���ѱ԰� ���� 
	ElseIF frm1.rdoInsCri.rdoInsCri2.checked = true then
		frm1.txtInsCri.Value = "1"		'���ѱ԰� ���� 
	ElseIF frm1.rdoInsCri.rdoInsCri3.checked = true then
		frm1.txtInsCri.Value = "2" 		'���ʱ԰� ���� 
	Else
	
	End IF
	
	IF frm1.txtSTDack.Value = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'���û����� üũ�Ͻʽÿ� 
		Exit Function	
		
	ElseIF  frm1.txtInsCri.Value = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'���û����� üũ�Ͻʽÿ� 
		Exit Function	
		
	ElseIF  frm1.txtP1.Text = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'���û����� üũ�Ͻʽÿ� 
		Exit Function	
	End IF
	
	IF frm1.rdoSTDack.rdoSTDack1.checked = true and frm1.txtSD.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'���û��װ� �Է»����� ��ġ���� �ʽ��ϴ� 
		Exit Function	
	End IF
	
	IF frm1.rdoInsCri.rdoInsCri1.checked = true and frm1.txtLowerBound.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'���û��װ� �Է»����� ��ġ���� �ʽ��ϴ� 
		Exit Function	
	End IF
	IF frm1.rdoInsCri.rdoInsCri2.checked = true and frm1.txtUpperBound.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'���û��װ� �Է»����� ��ġ���� �ʽ��ϴ� 
		Exit Function	
	End IF
	IF frm1.rdoInsCri.rdoInsCri3.checked = true and frm1.txtUpperBound.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'���û��װ� �Է»����� ��ġ���� �ʽ��ϴ� 
		Exit Function	
	End IF
	IF frm1.rdoInsCri.rdoInsCri3.checked = true and frm1.txtLowerBound.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'���û��װ� �Է»����� ��ġ���� �ʽ��ϴ� 
		Exit Function	
	End IF
	
	Call ggoOper.ClearField(Document, "1")										'��: Clear Contents  Field
	
	Call LayerShowHide(1)

	strVal = BIZ_PGM_QRY_ID & "?txtSD=" & frm1.txtSD.Text '��: '��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtAlpha=" & frm1.txtAlpha.Text 
	strVal = strVal & "&txtBeta=" & frm1.txtBeta.Text
	strVal = strVal & "&txtP1=" & frm1.txtP1.Text 
	strVal = strVal & "&txtP2=" & frm1.txtP2.Text 
	
	strVal = strVal & "&txtSTDack=" & Trim(frm1.txtSTDack.Value) 	'CheckBox���� �Ѱ��ݴϴ�.
	strVal = strVal & "&txtInsCri=" & Trim(frm1.txtInsCri.Value) 		'CheckBox���� �Ѱ��ݴϴ�.

	strVal = strVal & "&txtLowerBound=" & frm1.txtLowerBound.Text
	strVal = strVal & "&txtUpperBound=" & frm1.txtUpperBound.Text
	
	Call RunMyBizASP(MyBizASP, strVal) 
			
End Function 
 
'========================================= 2.3.2 ShowGraphClick() ======================================== 
'= Name : ShowGraphClick() = 
'= Description : Return Array to Opener Window for Cancel button click = 
'======================================================================================================== 
 
Function ShowGraphClick() 

	Dim strVal	
	
	IF frm1.txtSampleSize.Text = "" then
		Call DisplayMsgBox("229920", "X", "X", "X") 		'����׸��� �����ϴ� 
		Exit Function	
	End IF
			
	strVal = PGM_JUMP_ID3 & "?txtSampleSize=" & frm1.txtSampleSize.Text
	If frm1.rdoSTDack.rdoSTDack1.checked = true then
		strVal = strVal & "&txtSD=" & frm1.txtSD.Text
	Else
		Call DisplayMsgBox("229922", "X", "X", "X") 		'����� �������� �ʽ��ϴ� 
		Exit Function	
		'strVal = strVal & "&txtSD=" & ""
	End If
	strVal = strVal & "&txtInsCri=" & Trim(frm1.txtInsCri.Value) 		'CheckBox���� �Ѱ��ݴϴ�.

	If frm1.txtInsCri.value = 0 Then 
		strVal = strVal & "&txtLowerBound=" & frm1.txtLowerBound.Text
	End If 
	
	If frm1.txtInsCri.value = 1 Then 
		strVal = strVal & "&txtUpperBound=" & frm1.txtUpperBound.Text
	End If 
	
	If frm1.txtInsCri.value = 2 Then
		strVal = strVal & "&txtUpperBound=" & frm1.txtUpperBound.Text
		strVal = strVal & "&txtLowerBound=" & frm1.txtLowerBound.Text
	End if
	'/* Issue: �˻��� �������� Return - START */
	strVal = strVal & "&txtPageCode=" & "OV"
	'/* Issue: �˻��� �������� Return - END */
	Location.href = strVal
	
End Function 

'/* Issue: ǥ�������� �ʱⰪ�� �ȵ��� �Ͱ� ���� ���� - START */
'=======================================================================================================
'   Sub Name : LockSD()
'   Sub Desc : 
'=======================================================================================================
Sub LockSD(Byval vSTDack)
	If vSTDack = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtSD, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtSD, "Q")
	End If
End Sub
'/* Issue: ǥ�������� �ʱⰪ�� �ȵ��� �Ͱ� ���� ���� - END */

'/* Issue: �˻�԰� ���ÿ� ���� �ش� �ʵ� Enable/Disable - START */
'=======================================================================================================
'   Sub Name : LockInspSpec()
'   Sub Desc : 
'=======================================================================================================
Sub LockInspSpec(Byval vSTrdoInsCri)
	SELECT CASE vSTrdoInsCri
		CASE "A"
			Call ggoOper.SetReqAttr(frm1.txtUpperBound, "N")
			Call ggoOper.SetReqAttr(frm1.txtLowerBound, "N")
		CASE "U"
			Call ggoOper.SetReqAttr(frm1.txtUpperBound, "N")
			Call ggoOper.SetReqAttr(frm1.txtLowerBound, "Q")
		CASE "L"
			Call ggoOper.SetReqAttr(frm1.txtUpperBound, "Q")
			Call ggoOper.SetReqAttr(frm1.txtLowerBound, "N")
		CASE ELSE
			Call ggoOper.SetReqAttr(frm1.txtUpperBound, "Q")
			Call ggoOper.SetReqAttr(frm1.txtLowerBound, "Q")
	END SELECT
End Sub
'/* Issue: �˻�԰� ���ÿ� ���� �ش� �ʵ� Enable/Disable - END */

'========================================== 3.1.1 Form_Load() ====================================== 
' Name : Form_Load() 
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================  
Sub Form_Load() 
	Call LoadInfTB19029                                                     	'��: Load table , B_numeric_format
	Call AppendNumberPlace("6", "3", "2")
	Call AppendNumberPlace("7", "11", "4")
	Call ggoOper.LockField(Document, "N")                                   	'��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	'/* Issue: ǥ�������� �ʱⰪ�� �ȵ��� �Ͱ� ���� ���� - START */
	Call LockSD("")
	'/* Issue: ǥ�������� �ʱⰪ�� �ȵ��� �Ͱ� ���� ���� - END */
	'/* Issue: �˻�԰� ���ÿ� ���� �ش� �ʵ� Enable/Disable - START */
	Call LockInspSpec("")
	'/* Issue: �˻�԰� ���ÿ� ���� �ش� �ʵ� Enable/Disable - END */
	Call InitVariables																'��: Initializes local global variables          
	Call SetDefaultVal
	Call SetToolbar("10000000000111")
End Sub 
 
'========================================================================================== 
' Event Name : Form_QueryUnload 
' Event Desc : 
'========================================================================================== 
Sub Form_QueryUnload(Cancel , UnloadMode ) 

End Sub 
 
'/* Issue: ǥ�������� �ʱⰪ�� �ȵ��� �Ͱ� ���� ���� - START */
'========================================================================================== 
' Event Name : rdoSTDack1_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoSTDack1_onclick()
	Call LockSD("Y")
End Sub

'========================================================================================== 
' Event Name : rdoSTDack2_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoSTDack2_onclick()
	Call LockSD("N")
End Sub
'/* Issue: ǥ�������� �ʱⰪ�� �ȵ��� �Ͱ� ���� ���� - END */

'/* Issue: �˻�԰� ���ÿ� ���� �ش� �ʵ� Enable/Disable - START */
'========================================================================================== 
' Event Name : rdoInsCri1_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoInsCri1_onclick()
	Call LockInspSpec("L")
End Sub

'========================================================================================== 
' Event Name : rdoInsCri2_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoInsCri2_onclick()
	Call LockInspSpec("U")
End Sub

'========================================================================================== 
' Event Name : rdoInsCri3_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoInsCri3_onclick()
	Call LockInspSpec("A")
End Sub
'/* Issue: �˻�԰� ���ÿ� ���� �ش� �ʵ� Enable/Disable - END */

'======================================================================================== 
' Function Name : FncQuery 
' Function Desc : This function is related to Query Button of Main ToolBar 
'======================================================================================== 
Function FncQuery() 
	FncQuery = False
End Function 
 
'======================================================================================== 
' Function Name : FncNew 
' Function Desc : This function is related to New Button of Main ToolBar 
'======================================================================================== 
Function FncNew() 
	FncNew = False
End Function 
 
'======================================================================================== 
' Function Name : Fnc 
' Function Desc : This function is related to Delete Button of Main ToolBar 
'======================================================================================== 
Function FncDelete() 
	FncDelete = False
End Function 
 
'======================================================================================== 
' Function Name : FncSave 
' Function Desc : This function is related to Delete Button of Main ToolBar 
'======================================================================================== 
Function FncSave() 
	FncSave = False
End Function 	
 
'======================================================================================== 
' Function Name : FncCopy 
' Function Desc : This function is related to Copy Button of Main ToolBar 
'======================================================================================== 
Function FncCopy() 
	FncCopy = False
End Function 
 
'======================================================================================== 
' Function Name : FncCancel 
' Function Desc : This function is related to Cancel Button of Main ToolBar 
'======================================================================================== 
Function FncCancel() 
	FncCancel = False
End Function 
 
'======================================================================================== 
' Function Name : FncInsertRow 
' Function Desc : This function is related to InsertRow Button of Main ToolBar 
'======================================================================================== 
Function FncInsertRow() 
	FncInsertRow = False
End Function 
 
'======================================================================================== 
' Function Name : FncDeleteRow 
' Function Desc : This function is related to DeleteRow Button of Main ToolBar 
'======================================================================================== 
Function FncDeleteRow() 
	FncDeleteRow = False
End Function 

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call Parent.FncPrint()
	FncPrint = True    
End Function
 
'======================================================================================== 
' Function Name : FncPrev 
' Function Desc : This function is related to Previous Button 
'======================================================================================== 
Function FncPrev() 
	FncPrev = False
End Function 
 
'======================================================================================== 
' Function Name : FncNext 
' Function Desc : This function is related to Next Button 
'======================================================================================== 
Function FncNext() 
	FncNext = False
End Function 

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
	Call parent.FncExport(Parent.C_SINGLE)					'��: ȭ�� ���� 
	FncExcel = True
End Function

Function FncExit() 
	FncExit = True
End Function 

'******************************* 5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function ******************************* 
' ���� : 
'*********************************************************************************************************  
</SCRIPT> 
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD> 
<BODY SCROLL=NO TABINDEX="-1"> 
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> BORDER=0>
				<TR> 
					<TD WIDTH=10>&nbsp;</TD> 
					<TD CLASS="CLSMTABP"> 
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 > 
						<TR> 
							<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD> 
							<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>�跮 ������ ���ø� �˻�</FONT></TD> 
							<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD> 
						</TR> 
						</TABLE> 
					</TD> 
					<TD WIDTH=* align=right>&nbsp;</TD> 
				</TR> 
			</TABLE> 
		</TD> 
	</TR>
	<TR height=*>
		<TD  VALIGN="TOP" WIDTH=100% CLASS="Tab11"> 
			<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%> </TD>
				</TR>
				<TR>
					<TD VALIGN="top"  WIDTH="100%">
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;">
						<LEGEND>���û���</LEGEND>
							<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
								<TR>
									<TD CLASS=TD5  HEIGHT=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>ǥ������ ��������</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSTDack" TAG="2X" ID="rdoSTDack1"><LABEL FOR="rdoSTDack1">ǥ�������� �˰� �ֽ�</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSTDack" TAG="2X" ID="rdoSTDack2"><LABEL FOR="rdoSTDack2">ǥ�������� ���� ����</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>�˻�԰� ����</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInsCri" TAG="2X" ID="rdoInsCri1"><LABEL FOR="rdoInsCri1">����(����)�԰�</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInsCri" TAG="2X" ID="rdoInsCri2"><LABEL FOR="rdoInsCri2">����(����)�԰�</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInsCri" TAG="2X" ID="rdoInsCri3"><LABEL FOR="rdoInsCri3">���ʱ԰�</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5  HEIGHT=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;"> 
						<LEGEND>�Է»���</LEGEND> 
							<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%> 
								<TR>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>������ ����(��)</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma5_txtAlpha_txtAlpha.js'></script>&nbsp;%
									</TD>
									<TD CLASS="TD5" NOWRAP>�Һ��� ����(��)</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma5_txtBeta_txtBeta.js'></script>&nbsp;%
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>P1(AQL)</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma5_txtP1_txtP1.js'></script>&nbsp;%
									</TD>
									<TD CLASS="TD5" NOWRAP>P2(LTPD)</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma5_txtP2_txtP2.js'></script>&nbsp;%
									</TD>
								</TR>
								<!-- '/* Issue: �˻�԰� ���ÿ� ���� �ش� �ʵ� Enable/Disable - START */ -->
								<TR> 
									<TD CLASS="TD5" NOWRAP>���ѱ԰�</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma5_txtUpperBound_txtUpperBound.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>���ѱ԰�</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma5_txtLowerBound_txtLowerBound.js'></script>
									</TD>
								</TR> 
								<!-- '/* Issue: �˻�԰� ���ÿ� ���� �ش� �ʵ� Enable/Disable - END */ -->
								<TR> 
									<TD CLASS="TD5" NOWRAP>ǥ������</TD> 
									<TD CLASS="TD6" NOWRAP>
										<!-- /* 8�� ������ġ: ��ġ ���� ���� Tag ���� */ -->
										<script language =javascript src='./js/q1413ma5_txtSD_txtSD.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD> 
									<TD CLASS="TD6" NOWRAP></TD> 
								</TR> 
								<TR>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
								</TR>
							</TABLE> 
						</FIELDSET> 		
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;"> 
						<LEGEND>���</LEGEND> 
							<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%> 
								<TR>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
								</TR>
								<TR> 
									<TD CLASS="TD5" NOWRAP>����ũ��</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma5_txtSampleSize_txtSampleSize.js'></script>
									</TD> 
									<TD CLASS="TD5" NOWRAP>�հ��������</TD> 
									<TD CLASS="TD6" NOWRAP>
										<!-- /* 8�� ������ġ: ��ġ ���� ���� Tag ���� */ -->
										<script language =javascript src='./js/q1413ma5_txtAcceptSize_txtAcceptSize.js'></script>
									</TD>	
								</TR> 
								<TR>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
									<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
								</TR>
							</TABLE> 
						</FIELDSET> 
					</TD>	
				</TR>
			</TABLE>
		</DIV>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
    		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
	   			<TR>
					<TD WIDTH=10>&nbsp;</TD>  			
	        		<TD><BUTTON NAME="btnResult" CLASS="CLSMBTN" ONCLICK="vbscript:ResultClick()">��� ����</BUTTON></TD>
	        		<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:ShowGraphClick">�˻�Ư�� �׷���</A>&nbsp;|&nbsp;<A href="vbscript:ReturnClick()">������ �ý��� ����</A></TD>
	        		<TD WIDTH=10>&nbsp;</TD>
	   		</TR>
    		</TABLE>
    	</TD>
    </TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtSTDack" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsCri" tag="24" tabindex=-1>
</FORM> 
<DIV ID="MousePT" NAME="MousePT"> 
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe> 
</DIV> 
</BODY> 
</HTML>


