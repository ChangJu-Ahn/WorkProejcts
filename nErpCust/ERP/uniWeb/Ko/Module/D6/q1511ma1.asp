<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1511MA1
'*  4. Program Name         : Quality Configuration
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PD6G020,PD6G010
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
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "q1511mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "q1511mb2.asp"											 '��: �����Ͻ� ���� ASP�� 

Dim IsOpenPop

<!-- #Include file="../../inc/lgvariables.inc" -->	

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                                        '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
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
		.rdoPRYNBeforeReceipt2.checked = True
		.rdoSTYNAftereReceipt2.checked = True
		.rdoModifyYNAfterRelease2.checked = True
	End With
End Sub

'==========================================  2.2.3 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

    Call CommonQueryRs(" A.MINOR_CD, A.MINOR_NM ", " B_MINOR A INNER JOIN B_CONFIGURATION B ON A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD ", " A.MAJOR_CD = " & FilterVar("Q0026", "''", "S") & " AND B.REFERENCE = " & FilterVar("D", "''", "S") & " ORDER BY A.MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspDt , lgF0, lgF1, Chr(11))
	
	Call CommonQueryRs(" A.MINOR_CD, A.MINOR_NM ", " B_MINOR A INNER JOIN B_CONFIGURATION B ON A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD ", " A.MAJOR_CD = " & FilterVar("Q0026", "''", "S") & " AND B.REFERENCE = " & FilterVar("R", "''", "S") & " ORDER BY A.MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboReleaseDt , lgF0, lgF1, Chr(11))
		    
End Sub

'------------------------------------------  OpenPlant1()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd1.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd1.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "�����ڵ�"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd1.Value = arrRet(0)		
		frm1.txtPlantNm1.Value = arrRet(1)		
	End If
	frm1.txtPlantCd1.Focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenPlant2()  -------------------------------------------------
'	Name : OpenPlant2()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd2.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd2.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "�����ڵ�"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd2.Value = arrRet(0)		
		frm1.txtPlantNm2.Value = arrRet(1)		
		lgBlnFlgChgValue = True
	End If	
	frm1.txtPlantCd2.Focus
	Set gActiveElement = document.activeElement
	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029																'��: Load table , B_numeric_format
    Call SetToolbar("110010000000011")
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
	Call InitVariables
	Call InitComboBox
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd1.value = Parent.gPlant
		frm1.txtPlantNm1.value = Parent.gPlantNm

		Call MainQuery
	Else
		lgBlnFlgChgValue = False
		frm1.txtPlantCd1.focus 
	End If
	
	Set gActiveElement = document.activeElement
End Sub

'========================================== rdoPRYNBeforeReceipt1_OnClick()  ======================================
'	Name : rdoPRYNBeforeReceipt1_OnClick()
'	Description :
'========================================================================================================= 
Sub rdoPRYNBeforeReceipt1_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================== rdoPRYNBeforeReceipt2_OnClick()  ======================================
'	Name : rdoPRYNBeforeReceipt2_OnClick()
'	Description :
'========================================================================================================= 
Sub rdoPRYNBeforeReceipt2_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================== rdoSTYNAftereReceipt1_OnClick()  ======================================
'	Name : rdoSTYNAftereReceipt1_OnClick()
'	Description :
'========================================================================================================= 
Sub rdoSTYNAftereReceipt1_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================== rdoSTYNAftereReceipt2_OnClick()  ======================================
'	Name : rdoSTYNAftereReceipt2_OnClick()
'	Description :
'========================================================================================================= 
Sub rdoSTYNAftereReceipt2_OnClick()
	lgBlnFlgChgValue = True
End Sub  

'========================================== rdoSTYNAftereReceipt1_OnClick()  ======================================
'	Name : rdoSTYNAftereReceipt1_OnClick()
'	Description :
'========================================================================================================= 
Sub rdoModifyYNAfterRelease1_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================== rdoSTYNAftereReceipt2_OnClick()  ======================================
'	Name : rdoSTYNAftereReceipt2_OnClick()
'	Description :
'========================================================================================================= 
Sub rdoModifyYNAfterRelease2_OnClick()
	lgBlnFlgChgValue = True
End Sub  

'========================================== txtPlantCd1_KeyPress()  ======================================
'	Name : txtPlantCd1_KeyPress()
'	Description :
'========================================================================================================= 
Sub txtPlantCd1_KeyPress()
	If KeyAscii = 13 Then
		frm1.txtPlantNm1.value = ""
	End If
End Sub

'=======================================================================================================
'   Event Name : cboInspDt_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboInspDt_onchange()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : cboReleaseDt_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboReleaseDt_onchange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False															'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then Exit Function
    End If
    
	If frm1.txtPlantCd1.value = "" Then
		frm1.txtPlantNm1.value = ""
		Call DisplayMsgBox("169901","X","X","X")
		frm1.txtPlantCd1.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
    '-----------------------
    'Erase contents area
    '----------------------- 
	Call ggoOper.ClearField(Document, "2")
	Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    Call InitVariables															'��: Initializes local global variables
	
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then Exit Function    
       
    FncQuery = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    
	FncNew = False																'��: Processing is NG
	
	Err.Clear																	'��: Protect system from crashing
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
	Call InitVariables															'��: Initializes local global variables
	Call SetDefaultVal
	Call SetToolbar("110010000000011")
	
	frm1.txtPlantCd2.focus 
	Set gActiveElement = document.activeElement
	
	FncNew = True                                                            	'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False																'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                          '��: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If frm1.txtPlantCd2.value = "" Then
		frm1.txtPlantNm2.value = ""
		frm1.txtPlantCd2.focus
		Set gActiveElement = document.activeElement
		Call DisplayMsgBox("169901","X","X","X")
		Exit Function
	End If    
	
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function   
		
    FncSave = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim IntRetCD 
    
    FncPrev = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then										'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")									'��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If
    
	'-----------------------
    'Check condition area
    '-----------------------
 
    If frm1.txtPlantCd1.value = "" Then
		frm1.txtPlantNm1.value = ""
		frm1.txtPlantCd1.focus
		Set gActiveElement = document.activeElement
		Call DisplayMsgBox("169901","X","X","X")
		Exit Function
	End If
    
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbPrev = False Then Exit Function  
		           
	FncPrev = True
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim IntRetCD 
    
    FncNext = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then										'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")									'��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then Exit Function
    End If
    
	'-----------------------
    'Check condition area
    '----------------------- 
    If frm1.txtPlantCd1.value = "" Then
		frm1.txtPlantNm1.value = ""
		frm1.txtPlantCd1.focus
		Set gActiveElement = document.activeElement
		Call DisplayMsgBox("169901","X","X","X")
		Exit Function
	End If
    
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbNext = False Then Exit Function  
    
	FncNext = False
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	frm1.txtPlantCd2.value = ""
	frm1.txtPlantNm2.value = ""
	
	lgIntFlgMode = Parent.OPMD_CMODE														'��: Indicates that current mode is Crate mode
	lgBlnFlgChgValue = True
	Call ggoOper.ClearField(Document, "1")                                      			'��: Clear Condition Field
	Call ggoOper.LockField(Document, "N")													'��: This function lock the suitable field
	
	frm1.txtPlantCd2.focus
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
    Call parent.FncExport(Parent.C_SINGLE)													'��: ȭ�� ���� 
    FncExcel = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_SINGLE , False)											'��:ȭ�� ����, Tab ���� 
    FncFind = True
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")						'��: "Will you destory previous data"
		If IntRetCD = vbNo Then Exit Function
    End If

    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False																			'��: Processing is NG
    
    Dim strVal
    
    LayerShowHide(1)
       
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 & _
							  "&txtPlantCd=" & Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg=" & ""
    
	Call RunMyBizASP(MyBizASP, strVal)														'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True																			'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbPrev
' Function Desc : This function is the previous data query and display
'========================================================================================
Function DbPrev()
    DbPrev = False																			'��: Processing is NG
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 & _
							  "&txtPlantCd=" & Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg=" & "P"													'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)														'��: �����Ͻ� ASP �� ���� 
	
	DbPrev = True
End Function

'========================================================================================
' Function Name : DbNext
' Function Desc : This function is the previous data query and display
'========================================================================================
Function DbNext()
    DbNext = False																			'��: Processing is NG
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 & _
							  "&txtPlantCd=" & Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg=" & "N"													'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)														'��: �����Ͻ� ASP �� ���� 
	
	DbNext = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()																		'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")													'��: This function lock the suitable field
    Call SetToolbar("11101000111111")
    frm1.txtPlantCd1.focus
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 
	DbSave = False																			'��: Processing is NG

	LayerShowHide(1)
		
	With frm1
		.txtMode.value = Parent.UID_M0002													'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	End With
	
    DbSave = True																			'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()																			'��: ���� ������ ���� ���� 
	Call InitVariables
    frm1.txtPlantCd1.value = frm1.txtPlantCd2.value 
    Call MainQuery()
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ��ȯ�漳��</font></td>
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
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP HEIGHT=10></TD>
									<TD CLASS=TD656 NOWRAP HEIGHT=10></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD656><INPUT TYPE=TEXT NAME="txtPlantCd1" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="�����ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantPopup1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant1()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm1" SIZE=30 tag="14X" ALT="�����"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP HEIGHT=10></TD>
									<TD CLASS=TD656 NOWRAP HEIGHT=10></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT=30 WIDTH=100% COLSPAN = 2>
									<TABLE CLASS=TB2 CELLSPACING=0>
										<TR>
											<TD CLASS=TD5 NOWRAP HEIGHT=10></TD>
											<TD CLASS=TD656 NOWRAP HEIGHT=10></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>����</TD>
											<TD CLASS=TD656><INPUT TYPE=TEXT NAME="txtPlantCd2" SIZE=10 MAXLENGTH=4 tag="23XXXU" ALT="�����ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantPopup2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant2()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm2" SIZE=30 tag="24X" ALT="�����"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP HEIGHT=10></TD>
											<TD CLASS=TD656 NOWRAP HEIGHT=10></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%>></TD>
							</TR>
							<TR>
								<TD WIDTH= 100% valign=top>
									<FIELDSET><LEGEND>�Ϲ�����</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=16></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=16></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Release�� ��������</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoModifyYNAfterRelease ID=rdoModifyYNAfterRelease1 tag="22" VALUE="Y" ><LABEL FOR=rdoModifyYNAfterRelease1>��</LABEL>&nbsp;&nbsp;&nbsp;
																	   <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoModifyYNAfterRelease ID=rdoModifyYNAfterRelease2 tag="22" VALUE="N" CHECKED><LABEL FOR=rdoModifyYNAfterRelease2>�ƴϿ�</LABEL></TD>
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=20></TD>
											</TR>
										</TABLE>	
									</FIELDSET>
									<TABLE><TR><TD <%=HEIGHT_TYPE_02%>></TD></TR></Table>
									<FIELDSET><LEGEND>���԰˻�����</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�԰����˻� Release�� �����԰� �ڵ�ó��</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoPRYNBeforeReceipt ID=rdoPRYNBeforeReceipt1 tag="22" VALUE="Y" ><LABEL FOR=rdoPRYNBeforeReceipt1>��</LABEL>&nbsp;&nbsp;&nbsp;
																	   <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoPRYNBeforeReceipt ID=rdoPRYNBeforeReceipt2 tag="22" VALUE="N" CHECKED><LABEL FOR=rdoPRYNBeforeReceipt2>�ƴϿ�</LABEL></TD>
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�԰��İ˻� Release�� ����̵� �ڵ�ó��</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSTYNAftereReceipt ID=rdoSTYNAftereReceipt1 tag="22" VALUE="Y" ><LABEL FOR=rdoSTYNAftereReceipt1>��</LABEL>&nbsp;&nbsp;&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSTYNAftereReceipt ID=rdoSTYNAftereReceipt2 tag="22" VALUE="N" CHECKED><LABEL FOR=rdoSTYNAftereReceipt2>�ƴϿ�</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
										</TABLE>	
									</FIELDSET>
									<TABLE><TR><TD <%=HEIGHT_TYPE_02%>></TD></TR></Table>
									<FIELDSET><LEGEND>�⺻ ǥ�ð� ����</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�˻���</TD>
												<TD CLASS=TD6 NOWRAP><SELECT Name="cboInspDt" ALT="�˻���" STYLE="WIDTH: 150px" tag="22"></SELECT></TD>
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Release��</TD>
												<TD CLASS=TD6 NOWRAP><SELECT Name="cboReleaseDt" ALT="Release��" STYLE="WIDTH: 150px" tag="22"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
										</TABLE>	
									</FIELDSET>
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
	<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" SRC="../../blank.htm" tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
	</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
