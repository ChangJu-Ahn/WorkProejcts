<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Long-term Inv Analysis
'*  2. Function Name        : 
'*  3. Program ID           : I3111MA1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : PI3G111, PI3G110
'*  7. Modified date(First) : 2005/01/26
'*  8. Modified date(Last)  : 2006/09/01
'*  9. Modifier (First)     : Jaewoo Koh
'* 10. Modifier (Last)      : LEE SEUNG WOOK
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

Const BIZ_PGM_QRY_ID	= "i3111mb1.asp"
Const BIZ_PGM_SAVE_ID	= "i3111mb2.asp"

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
	Call LoadInfTB19029
    Call SetToolbar("110010000000011")
    Call ggoOper.LockField(Document, "N")
	Call InitVariables

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

'==========================================================================================
'   Event Name : cboplanflag1_OnClick
'   Event Desc : change flag setting
'==========================================================================================
'Sub cboplanflag1_OnClick()
'	lgBlnFlgChgValue = True
'	Call ggoOper.SetReqAttr(frm1.txtplanStockCalPeriod,"N")
'End Sub

'==========================================================================================
'   Event Name : cboplanflag2_OnClick
'   Event Desc : change flag setting
'==========================================================================================
'Sub cboplanflag2_OnClick()
'	lgBlnFlgChgValue = True
'	frm1.txtplanStockCalPeriod.value = ""
'	Call ggoOper.SetReqAttr(frm1.txtplanStockCalPeriod,"Q")
'End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False
    Err.Clear

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
       
    strVal = BIZ_PGM_QRY_ID & "?txtMode="		& Parent.UID_M0001 & _
							  "&txtPlantCd="	& Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg="	& ""
    
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
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode="		& Parent.UID_M0001 & _
							  "&txtPlantCd="	& Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg="	& "P"													'☆: 조회 조건 데이타 
    
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
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode="		& Parent.UID_M0001 & _
							  "&txtPlantCd="	& Trim(frm1.txtPlantCd1.value) & _
							  "&PrevNextFlg="	& "N"													'☆: 조회 조건 데이타 
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>장기재고분석환경설정</font></td>
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
									<TD CLASS=TD656><INPUT TYPE=TEXT NAME="txtPlantCd1" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantPopup1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant1()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm1" SIZE=30 tag="14X" ALT="공장명"></TD>
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
											<TD CLASS=TD5 NOWRAP HEIGHT=20></TD>
											<TD CLASS=TD6 NOWRAP HEIGHT=20></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>공장</TD>
											<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPlantCd2" SIZE=10 MAXLENGTH=4 tag="23XXXU" ALT="공장" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantPopup2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant2()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm2" SIZE=30 tag="24X" ALT="공장명"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP HEIGHT=20></TD>
											<TD CLASS=TD6 NOWRAP HEIGHT=20></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%>></TD>
							</TR>
							<TR>
								<TD WIDTH= 100% valign=top>
									<FIELDSET><LEGEND>판정 기준 설정</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=20></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>장기 재고</TD>
											    <TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtLongtermStockCalPeriod" SIZE=4 MAXLENGTH=4 tag="22" ALT="장기 보관 재고 기준기간" >&nbsp;개월 이상 사용(출고)실적이 없는 품목</TD>
												
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>악성 재고</TD>
												<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPerniciousStockCalPeriod" SIZE=4 MAXLENGTH=4 tag="22" ALT="악성 재고 기준기간" >&nbsp;개월 이상 사용(출고)실적이 없는 품목</TD>
												
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=20></TD>
											</TR>
										</TABLE>	
									</FIELDSET>
									
									<!--<FIELDSET><LEGEND>계획 기준 설정</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=20></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>생산 계획 반영 여부</TD>
												<TD CLASS="TD6">
												<SPAN STYLE="WIDTH: 50px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboplanflag" CHECKED ID="cboplanflag1" VALUE="Y" tag="22"><LABEL FOR="cboplanflag1">YES</LABEL></SPAN>
												<SPAN STYLE="WIDTH: 50px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboplanflag" ID="cboplanflag2" VALUE="N" tag="22"><LABEL FOR="cboplanflag2">NO</LABEL></SPAN></TD>
									
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>분석 년월 기준 향후</TD>
												<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtplanStockCalPeriod" SIZE=4 MAXLENGTH=4 tag="22" ALT="악성 재고 기준기간" >&nbsp;개월 간의 생산계획분을 감안하여 사용예정량을 계산한다</TD>
												
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20></TD>
												<TD CLASS=TD6 NOWRAP HEIGHT=20></TD>
											</TR>
										</TABLE>	
									</FIELDSET>-->
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
