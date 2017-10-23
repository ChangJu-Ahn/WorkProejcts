<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q3211MA3
'*  4. Program Name         : Worst 10 품질현황 
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
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "q3211mb3.asp"							'☆: Query 비지니스 로직 ASP명 
<!-- #Include file="../../inc/lgvariables.inc" -->

Const C_1=1
Const C_2=2
Const C_3=3
Const C_4=4
Const C_5=5
Const C_6=6
Const C_7=7
Const C_8=8
Const C_9=9
Const C_10=10

Dim C_Total
Dim Rm
Dim SpYrMnth1
Dim SpYrMnth2
Dim IsOpenPop
Dim strEbrErr
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------------- 
Dim CompanyYM
CompanyYM = UNIMonthClientFormat(UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gAPDateFormat))
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
	frm1.txtYrMnth1.text = CompanyYM
	frm1.txtYrMnth2.text = CompanyYM
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
End Sub

'==========================================  2.2.6 InitSpreadSheet()  =======================================
'	Name : InitSpreadSheet()
'	Description : 
'========================================================================================================= 
Sub InitSpreadSheet(Byval Unit)		

	With frm1.vspdData
	
		.ReDraw = false
		.MaxCols = 10

		.MaxRows = 8
		
		.Col = 0
		.Row = 1
		.Text = "공급처"
		.Row = 2
		.Text = "LOT수"
		.Row = 3
		.Text = "불합격LOT수"
		.Row = 4
		.Text = "검사의뢰수"
		.Row = 5
		.Text = "검사수"
		.Row = 6
		.Text = "불량수"

		If Unit <> "" then
			.Row = 7
			.Text = "불량률" & "(" & Unit & ")"
			.Row = 8
			.Text = "LOT불합격률" & "(%)"
		Else
			.Row = 7
			.Text = "불량률"
			.Row = 8
			.Text = "LOT불합격률"
		End If
				
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit
		
		.ColWidth(0) = 15
		
		ggoSpread.SSSetEdit C_1, "1", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_2, "2", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_3, "3", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_4, "4", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_5, "5", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_6, "6", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_7, "7", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_8, "8", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_9, "9", 11, 1, -1, 20
		ggoSpread.SSSetEdit C_10, "10", 11, 1, -1, 20
		Call SetSpreadLock
		.ReDraw = true
		
	End With
    
End Sub

'==========================================  2.2.6 SetSpreadLock()  =======================================
'	Name : SetSpreadLock()
'	Description : 
'========================================================================================================= 
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================  2.2.6 SetSpreadColor()  =======================================
'	Name : SetSpreadColor()
'	Description : 
'========================================================================================================= 
Sub SetSpreadColor(ByVal lRow)
End Sub

'==========================================  2.2.6 InitChartFx()  =======================================
'	Name : InitChartFx()
'	Description : Initialize ChartFx
'========================================================================================================= 
Sub InitChartFx()
	With frm1.ChartFX1
		'Chart Title 및 Font 설정 
		.Title_(0) = "LOT불합격률 추이도"
		.TopFont.Name = "굴림"
		.Axis(0).decimals = 4
		
		'Chart Legend Font 설정 
		.LegendBoxObj.Font.Name = "굴림"
		
		'그래프의 GAP 지정 
		.TopGap = 5			'그래프의 위쪽 여백 지정 
		.BottomGap = 20		'그래프의 아래쪽 여백 지정 
		.RightGap = 5
		.LeftGap = 50
	End With
End Sub

'==========================================  2.2.7 ClearChartFx()  =======================================
'	Name : ClearChartFx()
'	Description : Clear Chart FX Datas
'========================================================================================================= 
Sub ClearChartFx()
	With frm1.ChartFX1
		' X축/Y축 눈금 및 값이 안보이게 함 
		.Axis(2).Visible = False
		.Axis(0).Visible = False
		
		'범례 Clear
		.ClearLegend 1
		
		'차트 FX와의 데이터 채널 초기화 
		.OpenDataEx 1, 1, 1
		.CloseData 1 Or &H800		'COD_VALUES Or COD_REMOVE
		
		'계열을 안보이게 함 
		.Series(0).Visible = False
		.Title_(0) = "LOT불합격률 추이도"
	End With
End Sub

'------------------------------------------  OpenPlant1()  -------------------------------------------------
'	Name : OpenPlant1()
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
	arrParam5 = "R"
	
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

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables

    '----------  Coding part  -------------------------------------------------------------
    Call ggoOper.FormatDate(frm1.txtYrMnth1, Parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtYrMnth2, Parent.gDateFormat, 2)
    Call SetDefaultVal
    Call InitSpreadSheet("") 

    
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
    frm1.RadioDRType.rdoCase1.Checked = True
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
Sub txtYrMnth1_DblClick(Button)
    If Button = 1 Then
        frm1.txtYrMnth1.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtYrMnth1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYrMnth1_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrMnth1_Change()

End Sub

'=======================================================================================================
'   Event Name : txtYrMnth2_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrMnth2_DblClick(Button)
    If Button = 1 Then
        frm1.txtYrMnth2.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtYrMnth2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYrMnth1_KeyPress(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrMnth1_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtYrMnth2_KeyPress(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrMnth2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtYrMnth2_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrMnth2_Change()
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	Dim DefectRatioUnit	
	
	FncQuery = False                                                '⊙: Processing is NG
	strEbrErr = False
	Err.Clear                                                       '☜: Protect system from crashing
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then			'⊙: This function check indispensable field
		Exit Function
	End If	

	If ValidDateCheck(frm1.txtYrMnth1, frm1.txtYrMnth2) = False Then
		Exit Function
	End If
	
	If 	CommonQueryRs(" DEFECT_RATIO_UNIT_CD "," Q_DEFECT_RATIO_BY_INSP_CLASS ", " INSP_CLASS_CD = " & FilterVar("R", "''", "S") & "  AND PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("220401","X","X","X")
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	DefectRatioUnit = Trim(lgF0(0))
	
	'-----------------------
	'Erase contents area
	'----------------------- 
	Call ggoOper.ClearField(Document, "2")				'⊙: Clear Contents  Field
	Call InitVariables									'⊙: Initializes local global variables
	
	frm1.vspdData.focus
	ggoSpread.Source = frm1.vspdData
	
	Call InitSpreadSheet(DefectRatioUnit)

		
	Dim strITEM_NM	
	Call CommonQueryRs("ITEM_NM","B_ITEM", " ITEM_CD =  " & FilterVar(frm1.txtItemCd.Value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strITEM_NM = replace(lgF0,Chr(11),vbTab)
	frm1.txtItemNm.Value = strITEM_NM
		
	'-----------------------
	'Query function call area
	'----------------------- 
	If DbQuery = False then
		Exit Function
	End If			
									'☜: Query db data
	
	FncQuery = True									'⊙: Processing is OK
	
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function fncsave() 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow()
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow()
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                              
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
    On Error Resume Next                                                    '☜: Protect system from crashing
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
	Dim Yr1
	Dim Mnth1
	Dim Yr2
	Dim Mnth2
	
	Call LayerShowHide(1)

	Yr1 = Left(frm1.txtYrMnth1.DateValue,4)
	Mnth1 = Mid(frm1.txtYrMnth1.DateValue,5, 2)
	
	Yr2 = Left(frm1.txtYrMnth2.DateValue,4)
	Mnth2 = Mid(frm1.txtYrMnth2.DateValue,5, 2)
	
	If Yr1 > Yr2 Then	
		Call DisplayMsgBox("229901","X","X","X")		'기간 입력이 올바르지 않습니다 
		Exit Function
	ElseIf Yr1 = Yr2 and Mnth1> Mnth2 Then
		Call DisplayMsgBox("229901","X","X","X")		'기간 입력이 올바르지 않습니다 
		Exit Function
	ElseIf CInt(Mnth1) > 12 or CInt(Mnth2) > 12 Then
		Call DisplayMsgBox("229901","X","X","X")		'기간 입력이 올바르지 않습니다 
		Exit Function
	End If

	frm1.txtYr1.Value = Yr1
	frm1.txtMnth1.Value = Mnth1
	frm1.txtYr2.Value = Yr2
	frm1.txtMnth2.Value = Mnth2

	If frm1.RadioDRType.rdoCase1.Checked = True Then
		frm1.txtDataFlag.Value = "L"
	Else
		frm1.txtDataFlag.Value = "Q"
	End If
	
	Err.Clear                                                               					'☜: Protect system from crashing
	
	DbQuery = False                                                        					 '⊙: Processing is NG

	strVal = BIZ_PGM_QRY_ID & "?txtDataFlag=" & frm1.txtDataFlag.Value   '☜: '☆: 조회 조건 데이타 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)   '☜: '☆: 조회 조건 데이타 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.Value)
	strVal = strVal & "&txtYr1=" & Trim(frm1.txtYr1.Value)
	strVal = strVal & "&txtMnth1=" & Trim(frm1.txtMnth1.Value)
	strVal = strVal & "&txtYr2=" & Trim(frm1.txtYr2.Value)
	strVal = strVal & "&txtMnth2=" & Trim(frm1.txtMnth2.Value)

	Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True 
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()			
    CALL EBROk											'☆: 조회 성공후 실행로직 
    Call SetToolbar("11000000000111")
	Call SetSpreadLock
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
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   	Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어 
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
End Function


'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl)

	Dim strtitle
	Dim strYear,strMonth,strDay
	SetPrintCond = False

	Call ExtractDateFrom(frm1.txtYrMnth1.Text,frm1.txtYrMnth1.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	
	StrEbrFile	= "q3211ma3"
	
	if frm1.rdoCase1.checked = true Then 
	  strtitle = "LOT불합격률 추이도"
	Else
	  strtitle = "불량률 추이도"
	End If


	StrUrl =  "Title|"		& strtitle
	
    
	SetPrintCond = True
	
End Function



Sub MyBizASP1_onreadystatechange()
    	If LCase(MyBizASP1.Document.ReadyState) = "complete" Then
			Call LayerShowHide(0)
		End If
End Sub


Function EBROK()

		Dim StrUrl, StrEbrFile, ObjName
		Call LayerShowHide(1) 
		If Not SetPrintCond(StrEbrFile, strUrl) Then
			Exit Function
		End If
		ObjName = AskEBDocumentName(StrEbrFile,"ebr")
		
		lgEBProcessbarOut = "T"
		EBActionA.menu.value = 0
		Call FncEBR5RC2(ObjName, "view", StrUrl,EBActionA,"EBR")

 
End Function 
'========================================================================================
' Function Name : DBQueryErr
' Function Desc :  
'========================================================================================
Function DBQueryErr()														'☆: 조회 성공후 실행로직 

   strEbrErr = True 
   MyBizASP1.Document.location.href="../../blank.htm"
End Function   


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<!--' SPACE AREA-->
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!--' TAB, REFERENCE AREA -->
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Worst 10 품질현황</font></td>      
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>      
						    	</TR>      
						</TABLE>      
					</TD>      
					<TD WIDTH=*>&nbsp;</TD>      
				</TR>      
			</TABLE>      
		</TD>      
	</TR>
	<!--' CONDITION, CONTENT AREA -->
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<!--' SAPCE AREA -->
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<!--' CONDITION AREA -->
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS=CLSFLD>
							<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_40%>>		
								<TR>      
									<TD CLASS="TD5" NOWRAP>공장</TD>      
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=18 ALT="공장" tag="13XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">      
									<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>      
									<TD CLASS="TD5" NOWRAP>기간</TD>      
									<TD CLASS="TD6" NOWRAP>      
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYrMnth1 CLASS=FPDTYYYYMM title=FPDATETIME ALT="기간FROM" tag="13" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;      
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYrMnth2 CLASS=FPDTYYYYMM title=FPDATETIME ALT="기간TO" tag="13" id=fpDateTime2></OBJECT>');</SCRIPT>      
									</TD>      
								</TR>      
								<TR>      
									<TD CLASS="TD5" NOWRAP>품목</TD>      
									<TD CLASS="TD6" NOWRAP>      
										<INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="11XXXU"><IMG align=top height=20 name=btnItemCd onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">      
									<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>      
									<TD CLASS="TD5" NOWRAP></TD>      
									<TD CLASS="TD6" NOWRAP></TD>      
	      
								</TR>      
								<TR>      
									<TD CLASS="TD5" NOWRAP>DATA</TD>      
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioDRType" TAG="1X" ID="rdoCase1"><LABEL FOR="rdoCase1" >LOT불합격률</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioDRType" TAG="1X" ID="rdoCase2"><LABEL FOR="rdoCase2">불량률</LABEL>
									</TD>
									<td CLASS="TD5" NOWPAP HEIGHT=5></td>      
									<td CLASS="TD6" NOWPAP HEIGHT=5></td>      
								</TR>      
							</TABLE>      
						</FIELDSET>      
      
					</TD>      
				</TR>      
				<!--' SAPCE AREA -->
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<!--' CONTENT AREA(MULTI) -->      
				<TR>      
					<TD WIDTH=100% HEIGHT=44% valign=top>      
						<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>      
							<TR>	      
								<TD HEIGHT="100%" WIDTH="75%">      
									<IFRAME NAME="MyBizASP1"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=auto framespacing=0 marginwidth=0 marginheight=0 ></IFRAME>      
								</TD>      
							</TR>    
						</TABLE>      
					</TD>      
				</TR>      
				<TR>      
					<TD HEIGHT=5 WIDTH=100%></TD>      
				</TR>      
				<TR>      
					<TD WIDTH=100% HEIGHT=* valign=top>      
						<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>      
							<TR>      
								<TD HEIGHT="100%">      
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>      
								</TD>      
							</TR>      
						</TABLE>      
					</TD>      
				</TR>      
			</TABLE>      
		</TD>      
	</TR>
	<!--' SPACE AREA -->
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<!--' BATCH,JUMP AREA -->
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
	<!--' IFRAME AREA -->
	<TR>      
		      
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>      
		</TD>      
	</TR>      
</TABLE>      
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" rows="1" cols="20" tabindex=-1 ></TEXTAREA>      
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtYr1" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtMnth1" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtYr2" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtMnth2" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtDataFlag" tag="24" tabindex=-1 >      
      
</FORM>      

<FORM NAME="EBActionA" ID="EBAction" TARGET="MyBizASP1" METHOD="POST"  scroll=yes> 
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

