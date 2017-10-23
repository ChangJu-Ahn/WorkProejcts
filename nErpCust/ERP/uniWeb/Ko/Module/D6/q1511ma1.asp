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
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
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

Const BIZ_PGM_QRY_ID = "q1511mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "q1511mb2.asp"											 '☆: 비지니스 로직 ASP명 

Dim IsOpenPop

<!-- #Include file="../../inc/lgvariables.inc" -->	

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                                        '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'☆: 사용자 변수 초기화 
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
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
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

	arrParam(0) = "공장 팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd1.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"		
    
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

	arrParam(0) = "공장 팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd2.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"		
    
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
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call SetToolbar("110010000000011")
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
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
    
    FncQuery = False															'⊙: Processing is NG
    
    Err.Clear																	'☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
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
	Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    Call InitVariables															'⊙: Initializes local global variables
	
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then Exit Function    
       
    FncQuery = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    
	FncNew = False																'⊙: Processing is NG
	
	Err.Clear																	'☜: Protect system from crashing
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
	Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
	Call InitVariables															'⊙: Initializes local global variables
	Call SetDefaultVal
	Call SetToolbar("110010000000011")
	
	frm1.txtPlantCd2.focus 
	Set gActiveElement = document.activeElement
	
	FncNew = True                                                            	'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False																'⊙: Processing is NG
    
    Err.Clear																	'☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                          '⊙: No data changed!!
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
		
    FncSave = True																'⊙: Processing is OK
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
        Call DisplayMsgBox("900002","X","X","X")									'☆: 밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"
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
        Call DisplayMsgBox("900002","X","X","X")									'☆: 밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"
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
	
	lgIntFlgMode = Parent.OPMD_CMODE														'⊙: Indicates that current mode is Crate mode
	lgBlnFlgChgValue = True
	Call ggoOper.ClearField(Document, "1")                                      			'⊙: Clear Condition Field
	Call ggoOper.LockField(Document, "N")													'⊙: This function lock the suitable field
	
	frm1.txtPlantCd2.focus
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
    Call parent.FncExport(Parent.C_SINGLE)													'☜: 화면 유형 
    FncExcel = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_SINGLE , False)											'☜:화면 유형, Tab 유무 
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")						'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then Exit Function
    End If

    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False																			'⊙: Processing is NG
    
    Dim strVal
    
    LayerShowHide(1)
       
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 & _
							  "&txtPlantCd=" & Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg=" & ""
    
	Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True																			'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbPrev
' Function Desc : This function is the previous data query and display
'========================================================================================
Function DbPrev()
    DbPrev = False																			'⊙: Processing is NG
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 & _
							  "&txtPlantCd=" & Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg=" & "P"													'☆: 조회 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
	
	DbPrev = True
End Function

'========================================================================================
' Function Name : DbNext
' Function Desc : This function is the previous data query and display
'========================================================================================
Function DbNext()
    DbNext = False																			'⊙: Processing is NG
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 & _
							  "&txtPlantCd=" & Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg=" & "N"													'☆: 조회 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
	
	DbNext = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()																		'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE														'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")													'⊙: This function lock the suitable field
    Call SetToolbar("11101000111111")
    frm1.txtPlantCd1.focus
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	DbSave = False																			'⊙: Processing is NG

	LayerShowHide(1)
		
	With frm1
		.txtMode.value = Parent.UID_M0002													'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	End With
	
    DbSave = True																			'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()																			'☆: 저장 성공후 실행 로직 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품질환경설정</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD656><INPUT TYPE=TEXT NAME="txtPlantCd1" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantPopup1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant1()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm1" SIZE=30 tag="14X" ALT="공장명"></TD>
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
											<TD CLASS=TD5 NOWRAP>공장</TD>
											<TD CLASS=TD656><INPUT TYPE=TEXT NAME="txtPlantCd2" SIZE=10 MAXLENGTH=4 tag="23XXXU" ALT="공장코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantPopup2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant2()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm2" SIZE=30 tag="24X" ALT="공장명"></TD>
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
									<FIELDSET><LEGEND>일반정보</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=16></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=16></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Release후 수정여부</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoModifyYNAfterRelease ID=rdoModifyYNAfterRelease1 tag="22" VALUE="Y" ><LABEL FOR=rdoModifyYNAfterRelease1>예</LABEL>&nbsp;&nbsp;&nbsp;
																	   <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoModifyYNAfterRelease ID=rdoModifyYNAfterRelease2 tag="22" VALUE="N" CHECKED><LABEL FOR=rdoModifyYNAfterRelease2>아니오</LABEL></TD>
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=20></TD>
											</TR>
										</TABLE>	
									</FIELDSET>
									<TABLE><TR><TD <%=HEIGHT_TYPE_02%>></TD></TR></Table>
									<FIELDSET><LEGEND>수입검사정보</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>입고전검사 Release시 구매입고 자동처리</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoPRYNBeforeReceipt ID=rdoPRYNBeforeReceipt1 tag="22" VALUE="Y" ><LABEL FOR=rdoPRYNBeforeReceipt1>예</LABEL>&nbsp;&nbsp;&nbsp;
																	   <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoPRYNBeforeReceipt ID=rdoPRYNBeforeReceipt2 tag="22" VALUE="N" CHECKED><LABEL FOR=rdoPRYNBeforeReceipt2>아니오</LABEL></TD>
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>입고후검사 Release시 재고이동 자동처리</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSTYNAftereReceipt ID=rdoSTYNAftereReceipt1 tag="22" VALUE="Y" ><LABEL FOR=rdoSTYNAftereReceipt1>예</LABEL>&nbsp;&nbsp;&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSTYNAftereReceipt ID=rdoSTYNAftereReceipt2 tag="22" VALUE="N" CHECKED><LABEL FOR=rdoSTYNAftereReceipt2>아니오</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
										</TABLE>	
									</FIELDSET>
									<TABLE><TR><TD <%=HEIGHT_TYPE_02%>></TD></TR></Table>
									<FIELDSET><LEGEND>기본 표시값 설정</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>검사일</TD>
												<TD CLASS=TD6 NOWRAP><SELECT Name="cboInspDt" ALT="검사일" STYLE="WIDTH: 150px" tag="22"></SELECT></TD>
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=5></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=5></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Release일</TD>
												<TD CLASS=TD6 NOWRAP><SELECT Name="cboReleaseDt" ALT="Release일" STYLE="WIDTH: 150px" tag="22"></SELECT></TD>
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
