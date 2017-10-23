<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : Cp,Cpk
'*  3. Program ID           : Q3111MA3
'*  4. Program Name         : 공정능력평가 
'*  5. Program Desc         : 
'*  6. Component List       : 
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "q3111mb3.asp"							'☆: Query 비지니스 로직 ASP명 
<!-- #Include file="../../inc/lgvariables.inc" -->						'☆: Query 비지니스 로직 ASP명 

Dim IsOpenPop        

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYMDFrom
Dim CompanyYMDTo

CompanyYMDTo = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)
CompanyYMDFrom = UNIDateAdd("M" , -1, CompanyYMDTo, Parent.gDateFormat)
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtYrDt1.Text = CompanyYMDFrom
	frm1.txtYrDt2.Text = CompanyYMDTo
	frm1.cboInspClassCd.value= "R"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","MA") %>
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Err.Clear    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " major_cd=" & FilterVar("Q0001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboInspClassCd ,lgF0  ,lgF1  ,Chr(11))
End Sub

'==========================================  2.2.6 InitChartFx()  =======================================
'	Name : InitChartFx()
'	Description : Initialize ChartFx
'========================================================================================================= 
Sub InitChartFx()
	With frm1.ChartFX1
		'Chart Title 및 Font 설정 
		.Title_(2) = "공정능력"
		.TopFont.Name = "굴림"
		
		'그래프의 GAP 지정 
		.TopGap = 20			'그래프의 위쪽 여백 지정 
		.BottomGap = 20		'그래프의 아래쪽 여백 지정 
		.RightGap = 5
		.LeftGap = 5
		
	End With
End Sub

'==========================================  2.2.6 ClearChartFx()  =======================================
'	Name : ClearChartFx()
'	Description : Clear Chart FX Datas
'========================================================================================================= 
Sub ClearChartFx()
	With frm1.ChartFX1
		' X축/Y축 눈금 및 값이 안보이게 함 
		.Axis(2).Visible = False
		.Axis(0).Visible = False
		
		' Clear  CD_CONSTANTLINES
		.ClearData  &H10		'CD_CONSTANTLINES
		
		'차트 FX와의 데이터 채널 초기화 
		.OpenDataEx 1, 1, 1
		.CloseData 1 Or &H800		'COD_VALUES Or COD_REMOVE
		
		'계열을 안보이게 함 
		.Series(0).Visible = False
		
		'Grid를 안보이게 함 
		.Grid = 0			'CHART_NOGRID
	End With
End Sub

'------------------------------------------  OpenPlant()-----------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
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
	
	frm1.txtPlantCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
		frm1.txtPlantCd.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD

	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(frm1.txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(frm1.txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(frm1.txtItemNm.Value)	' Item Name
	arrParam5 = Trim(frm1.cboInspClassCd.Value)
	
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	frm1.txtItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
		frm1.txtItemCd.Focus		
	End If	

	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

'------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : InspItem1 PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem()
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9, Param10, Param11, Param12
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	'품목코드가 있는 지 체크 
	If Trim(frm1.txtItemCd.Value) = "" then 
		Call DisplayMsgBox("229916","X","X","X") 		'품목정보가 필요합니다 
		frm1.txtItemCd.focus
		Exit Function
	End If
	IsOpenPop = True
	
	With frm1
		Param1 = Trim(.txtPlantCd.Value)
		Param2 = Trim(.txtPlantNm.Value)
		Param3 = Trim(.txtItemCd.Value)
		Param4 = Trim(.txtItemNm.Value)
		Param5 = Trim(.cboInspClassCd.Value)
		Param6 = Trim(.cboInspClassCd.Options(.cboInspClassCd.SelectedIndex).Text)
		Param7 = Trim(.txtInspItemCd.value)
		Param8 = ""
		Param9 = ""		'모든 검사방식 
		Param10 = ""
		Param11 = ""
		Param12 = ""
	End With
	
	iCalledAspName = AskPRAspName("q1211pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "q1211pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9, Param10, Param11, Param12), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtInspItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspItemCd.Value  = arrRet(1)
		frm1.txtInspItemNm.Value  = arrRet(2)
		frm1.txtInspItemCd.Focus		
	End If	
	
	Set gActiveElement = document.activeElement	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call InitVariables                                                      '⊙: Initializes local global variables
	
	'----------  Coding part  -------------------------------------------------------------
	
	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
	Call initCombobox
	Call SetDefaultVal
'	Call InitChartFX
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtYrDt1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtYrDt1.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtYrDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYrDt2_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtYrDt2.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtYrDt2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYrDt1_KeyPress(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrDt1_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtYrDt2_KeyPress(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	FncQuery = False                                                        '⊙: Processing is NG
	
	Err.Clear                                                               '☜: Protect system from crashing

	  '-----------------------
	'Erase contents area
	'----------------------- 
	Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
	Call InitVariables	
'	Call ClearChartFX
	
	  '-----------------------
	'Check condition area
	'----------------------- 

	If Not chkField(Document, "1") Then						'⊙: This function check indispensable field
		Exit Function
	End If
	
	If ValidDateCheck(frm1.txtYrDt1, frm1.txtYrDt2) = False Then
		Exit Function
	End If
	
	  '-----------------------
	'Query function call area
	'----------------------- 
	If DbQuery = False then
		Exit Function
	End If			
	
	FncQuery = True
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 

	Dim IntRetCD 
	
	FncNew = False                                                          					'⊙: Processing is NG
		
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                       		'⊙: Lock  Suitable  Field
	Call InitVariables																'⊙: Initializes local global variables
	Call SetDefaultVal
'	Call ClearChartFX
	
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
	FncNew = True
End Function

'========================================================================================
' Function Name : FncDelete
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
' Function Name : FncFind
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_SINGLE, False)                                     
    FncFind = True
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call Parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
    Call parent.FncExport(Parent.C_SINGLE)
    FncExcel = True
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	
	Dim strVal
    	
	Call LayerShowHide(1)

	Err.Clear              
	
	DbQuery = False                                                        					'⊙: Processing is NG
		
	strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.Value) 			'☜: 공장코드를 읽어 온다.
	strVal = strVal & "&cboInspClassCd=" & Trim(frm1.cboInspClassCd.value)			'☜: 검사분류코드를 읽어 온다.
	strVal = strVal & "&txtYrDt1=" & frm1.txtYrDt1.Text				'☜: 검사일자시작을 읽어 온다.
	strVal = strVal & "&txtYrDt2=" & frm1.txtYrDt2.Text					'☜: 검사일자종료를 읽어 온다.
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)					'☜: 품목코드를 읽어 온다.	
	strVal = strVal & "&txtInspItemCd=" & Trim(frm1.txtInspItemCd.value)				'☜: 검사항목코드를 읽어 온다.	

	Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 

	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	Call DrawEBChart
	Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어 
End Function


'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl, intChartNo)

	SetPrintCond = False

	StrEbrFile	= "Q3111MA3"

	StrUrl = ""

	SetPrintCond = True

End Function


Function DrawEBChart()
	Dim StrUrl, StrEbrFile, ObjName


	If Not SetPrintCond(StrEbrFile, strUrl, 1) Then
		Exit Function
	End If

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	EBActionA.menu.value = 0
    Call FncEBR5RC2(ObjName, "view", StrUrl,EBActionA,"EBR")
End Function 

Function MyBizASP1_onReadyStateChange()
		If LCase(MyBizASP1.Document.ReadyState) = "complete" Then
			Call LayerShowHide(0)
		End If
End Function



'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data Save and display
'========================================================================================
Function DbSave()
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백 %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공정능력평가</font></td>
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
						<FIELDSET CLASS=CLSFLD>
							<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_40%>>		
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=18 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
									<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWPAP>검사분류</TD>
									<TD CLASS="TD6" NOWPAP>
									<SELECT Name="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" tag="12"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="12XXXU"><IMG align=top height=20 name=btnItemCd onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
									<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>검사항목</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="검사항목" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspItem()">
									<INPUT TYPE=TEXT NAME="txtInspItemNm" SIZE=20 MAXLENGTH="40" tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>기간</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYrDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="기간FROM" tag="12X1" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYrDt2 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="기간TO" tag="12X1" id=fpDateTime2></OBJECT>');</SCRIPT>										
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
					<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=2>
						<TR>
							<TD HEIGHT=100% WIDTH=20%>
								<TABLE CLASS="TB3" WIDTH="100%" CELLSPACING=0 CELLPADDING=0>		
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TD CLASS="TD5" NOWRAP>Cp</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCp" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>Cpk</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCpk" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>검사규격</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspSpec" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>상한규격</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtUSL" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>하한규격</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLSL" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>시료수</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSampleQty" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>최대공차</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMaxTol" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>									<TR>
										<TD CLASS="TD5" NOWRAP>최소공차</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMinTol" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>최대값</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMAX" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>최소값</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMIN" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>									<TR>
										<TD CLASS="TD5" NOWRAP>평균</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAverage" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>범위</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRange" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>표준편차</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtStd" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>					
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>-3시그마</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtM3Sigma" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>+3시그마</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtP3Sigma" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>측정단위코드</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMeasmtUnitCd" SIZE=20 MAXLENGTH=20 tag="24" STYLE="Text-Align: Right"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5% NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5% NOWRAP></TD>
									</TR>
								</TABLE>
							</TD>
							<TD HEIGHT=100% WIDTH=80%>
								<IFRAME NAME="MyBizASP1"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=AUTO framespacing=0 marginwidth=0 marginheight=0 ></IFRAME> 	
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT=20>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" tabindex=-1  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1 >
</FORM>
<FORM NAME="EBActionA" ID="EBActionA" TARGET="MyBizASP1" METHOD="POST"  scroll=yes> 
	<input TYPE="HIDDEN" NAME="menu" value=0 > 
	<input TYPE="HIDDEN" NAME="id" > 
	<input TYPE="HIDDEN" NAME="pw" >
	<input TYPE="HIDDEN" NAME="doc" > 
	<input TYPE="HIDDEN" NAME="form" > 
	<input TYPE="HIDDEN" NAME="runvar" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

