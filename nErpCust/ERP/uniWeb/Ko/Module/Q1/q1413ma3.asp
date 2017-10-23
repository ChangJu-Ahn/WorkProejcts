<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1413MA3
'*  4. Program Name         : 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>T��� ������ ���ø� �˻��� ����</TITLE>

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

Const BIZ_PGM_QRY_ID = "q1413Mb3.asp"
Const PGM_JUMP_ID1 = "q1411ma1"
Const PGM_JUMP_ID2 = "Q1441MA1.asp"

Dim lgNextNo					'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo					' ""

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgMpsFirmDate
Dim lgLlcGivenDt								
Dim IsOpenPop          
Dim gSelframeFlg

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    	lgIntFlgMode = Parent.OPMD_CMODE                                               	'��: Indicates that current mode is Create mode
    	lgIntGrpCount = 0
    	IsOpenPop = False
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	With frm1
		.txtLotSize.AllowNull = True
		.txtLotSize.Text = ""
	End With
End Sub

'------------------------------------------  OpenAQL()  -------------------------------------------------
'	Name : OpenAQL()
'	Description : AQL PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAQL()
	OpenAQL = false

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "AQL�˾�"					' �˾� ��Ī 
	arrParam(1) = "B_MINOR"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtAQL.value)				' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("Q0011", "''", "S") & ""				' Where Condition		
	arrParam(5) = "AQL"							' �����ʵ��� �� ��Ī	
	arrField(0) = "MINOR_CD"						' Field��(0)
	arrField(1) = "MINOR_NM"						' Field��(0)
    arrHeader(0) = "�ڵ�"					' Header��(0)
    arrHeader(1) = "��"						' Header��(0)
    
   	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtAQL.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtAQL.Value    = arrRet(0)
		frm1.txtAQLNm.Value  = arrRet(1)
		frm1.txtAQL.focus
	End If	

	Set gActiveElement = document.activeElement
	OpenAQL = true	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Sub ShowCondition(Byval vFlag)
	SELECT CASE vFlag
		CASE "S"
			Q3.Style.display = ""
			Q4.Style.display = "none"
		CASE "G"
			Q3.Style.display = "none"
			Q4.Style.display = ""
		CASE ELSE
			Q3.Style.display = "none"
			Q4.Style.display = "none"
	End SELECT
End Sub

'===========================================  2.3.1 ResultClick()  ==========================================
'=	Name : ResultClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function ResultClick()

	Dim strVal
	Dim DefectMode

								'�˻������ �˾ƺ���.
	If Not chkField(Document, "2") Then  '��: Check contents area
       		Exit Function
    	End If	
	
	If frm1.rdoRigor.rdoRigor3.checked = true then
		frm1.txtRigor.value = "1"			
	Elseif frm1.rdoRigor.rdoRigor1.checked = true then
		frm1.txtRigor.value = "2"
	Elseif frm1.rdoRigor.rdoRigor2.checked = true then
		frm1.txtRigor.value = "3"
	Else
	End if		
	
	'/* Issue: �Ϲݰ˻����, Ư���˻���ؿ� ���� �˻���� ���� ���� �ν��� �߸��� - START */
	If frm1.rdoDefectLevel.rdoDefectLevel1.checked = true then
		If frm1.rdoSpecial.rdoSpecial1.checked = true then
			frm1.txtDefectMode.value = "S-1"
		ElseIf frm1.rdoSpecial.rdoSpecial2.checked = true then
			frm1.txtDefectMode.value = "S-2"
		ElseIf frm1.rdoSpecial.rdoSpecial3.checked = true then
			frm1.txtDefectMode.value = "S-3"
		ElseIf frm1.rdoSpecial.rdoSpecial4.checked = true then
			frm1.txtDefectMode.value = "S-4"
		Else	
		End if		
	End if
	
	If frm1.rdoDefectLevel.rdoDefectLevel2.checked = true then
		If frm1.rdoNormal.rdoNormal1.checked = true then
			frm1.txtDefectMode.value = "I"
		ElseIf frm1.rdoNormal.rdoNormal2.checked = true then
			frm1.txtDefectMode.value = "II"
		ElseIf frm1.rdoNormal.rdoNormal3.checked = true then
			frm1.txtDefectMode.value = "III"
		Else	
		End if		
	End if
	'/* Issue: �Ϲݰ˻����, Ư���˻���ؿ� ���� �˻���� ���� ���� �ν��� �߸��� - END */
	
	IF frm1.txtRigor.Value = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'���û����� üũ�Ͻʽÿ� 
		Exit Function	
	ElseIF  frm1.txtDefectMode.Value = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'���û����� üũ�Ͻʽÿ� 
		Exit Function	
	End IF
	
	Call ggoOper.ClearField(Document, "1")										'��: Clear Contents  Field
	
	Call LayerShowHide(1)
		
	strVal = BIZ_PGM_QRY_ID & "?txtLotSize=" & frm1.txtLotSize.Text  		'��: '��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtAQL=" & Trim(frm1.txtAQL.Value)   
	strVal = strVal & "&txtRigor=" & Trim(frm1.txtRigor.Value)	
	strVal = strVal & "&txtDefectMode=" & Trim(frm1.txtDefectMode.Value)	
	
	Call RunMyBizASP(MyBizASP, strVal)
	
	
End Function

'=========================================  2.3.2 ShowGraphClick()  ========================================
'=	Name : ShowGraphClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function ShowGraphClick()
	Dim Replace
	Dim strVal
	
	IF frm1.txtSampleSize.Text = "" then
		Call DisplayMsgBox("229920", "X", "X", "X") 		'����׸��� �����ϴ� 
		Exit Function	
	End IF
	
	frm1.txtReplaceMode.Value = 0					'Passing for No Error
	
	strVal = PGM_JUMP_ID2 & "?txtLotSize=" & frm1.txtLotSize.Text
	strVal = strVal & "&txtDefectRate=" & Trim(frm1.txtAQL.Value)   		'�ҷ����� AQL�� �Ѵ�.
	strVal = strVal & "&txtSampleSize=" & frm1.txtSampleSize.Text   
	strVal = strVal & "&txtAcceptSize=" & frm1.txtAcceptSize.Text   
	strVal = strVal & "&txtReplaceMode=" & Trim(frm1.txtReplaceMode.Value)	'Replacement Passing for No Error
	'/* Issue: �˻��� �������� Return - START */
	strVal = strVal & "&txtPageCode=" & "AA"
	'/* Issue: �˻��� �������� Return - END */
	Location.href = strVal
End Function

'=============================================  2.3.3()  ======================================
'=	Event Name : ReturnClick
'=	Event Desc :
'========================================================================================================
Function ReturnClick()
	PgmJump(PGM_JUMP_ID1)
End Function

'------------------------------------------  OpenAQL()  -------------------------------------------------
'	Name : OpenAQL()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAQL()
	OpenAQL = false
	
	Dim arrRet
	Dim arrParam1, arrParam2
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = frm1.txtAQL.value
	arrParam2 = "Q0012"
	
	iCalledAspName = AskPRAspName("q1211pa3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2), _
	              "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	frm1.txtAQL.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtAQL.value = arrRet(0)
		frm1.txtAQL.focus
	End If
	
	Set gActiveElement = document.activeElement
	OpenAQL = true
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'��: Load table , B_numeric_format
	Call AppendNumberPlace("6", "5", "3")
	Call AppendNumberPlace("7", "11", "4")	
	Call ggoOper.LockField(Document, "N")                                   	'��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
		
	Call InitVariables																'��: Initializes local global variables          
	Call SetDefaultVal
	Call SetToolbar("10000000000111")	
	Call ShowCondition("")
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'==========================================================================================
'   Event Name : rdoDefectLevel1_onClick
'   Event Desc :
'==========================================================================================
Sub rdoDefectLevel1_onClick()
	Call ShowCondition("S")
End Sub

'==========================================================================================
'   Event Name : rdoDefectLevel2_onClick
'   Event Desc :
'==========================================================================================
Sub rdoDefectLevel2_onClick()
	Call ShowCondition("G")
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : Fnc
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
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then    				'��: Check contents area
       		Exit Function
    End If

	Call ResultClick()
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
	On Error Resume Next 
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
    Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
	On Error Resume Next                                                    					'��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	On Error Resume Next                                                    					'��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 Call parent.FncExport(Parent.C_SINGLE)					'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()
	On Error Resume Next
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : DbQueryOkOPEN

' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQueryOkOPEN

' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	On Error Resume Next
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()		
	On Error Resume Next
End Function

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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>��� ������ ���ø� �˻�</FONT></TD>
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
					<TD VALIGN="top"  WIDTH="100%" HEIGHT=*>
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
								<TR ID=Q1>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>���ݵ�</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRigor" TAG="2X" ID="rdoRigor1"><LABEL FOR="rdoRigor1">����˻�</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRigor" TAG="2X" ID="rdoRigor2"><LABEL FOR="rdoRigor2">��ٷο�˻�</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRigor" TAG="2X" ID="rdoRigor3"><LABEL FOR="rdoRigor3">�����Ѱ˻�</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q2>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>�˻���ر���</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDefectLevel" TAG="2X" ID="rdoDefectLevel1"><LABEL FOR="rdoDefectLevel1">Ư���˻����</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDefectLevel" TAG="2X" ID="rdoDefectLevel2"><LABEL FOR="rdoDefectLevel2">�Ϲݰ˻����</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q3>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>�˻����</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSpecial" TAG="2X" ID="rdoSpecial1"><LABEL FOR="rdoSpecial1">S-1</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSpecial" TAG="2X" ID="rdoSpecial2"><LABEL FOR="rdoSpecial2">S-2</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSpecial" TAG="2X" ID="rdoSpecial3"><LABEL FOR="rdoSpecial3">S-3</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSpecial" TAG="2X" ID="rdoSpecial4"><LABEL FOR="rdoSpecial4">S-4</LABEL></TD>
								</TR>
								<TR ID=Q4>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>�˻����</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoNormal" TAG="2X" ID="rdoNormal1"><LABEL FOR="rdoNormal1">��</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoNormal" TAG="2X" ID="rdoNormal2"><LABEL FOR="rdoNormal2">��</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoNormal" TAG="2X" ID="rdoNormal3"><LABEL FOR="rdoNormal3">��</LABEL></TD>
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
									<TD CLASS="TD5" NOWRAP>��Ʈũ��</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma3_fpDoubleSingle1_txtLotSize.js'></script>
									</TD>								
									<TD CLASS="TD5" NOWRAP>AQL</TD>
									
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma3_txtAQL_txtAQL.js'></script>
										<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAQL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenAQL()">
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
										<script language =javascript src='./js/q1413ma3_txtSampleSize_txtSampleSize.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�հ���������</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma3_txtAcceptSize_txtAcceptSize.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>���հ���������</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma3_txtRejectSize_txtRejectSize.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtRigor" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtDefectMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtReplaceMode" tag="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

