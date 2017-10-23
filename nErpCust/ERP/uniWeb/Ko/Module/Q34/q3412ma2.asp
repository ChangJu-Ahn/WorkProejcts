<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q3212MA1
'*  4. Program Name         : 불량유형파레토도 
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

Const BIZ_PGM_QRY_ID = "q3412mb2.asp"							'☆: Query 비지니스 로직 ASP명 
<!-- #Include file="../../inc/lgvariables.inc" -->								'☆: Query 비지니스 로직 ASP명 

Const C_Total=1
Const D_Total=1

Dim Col
Dim IsOpenPop               
Dim lgNoData1
Dim lgNoData2

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYM
CompanyYM = UNIMonthClientFormat(UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gAPDateFormat))
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
 Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE        'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False  	            'Indicates that no value changed
    lgIntGrpCount = 0        	            'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                       'initializes Previous Key
    lgLngCurRows = 0                        'initializes Deleted Rows Count
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtYrDt.Text = CompanyYM
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
'	Name : InitSpreadSheet1()
'	Description : 
'========================================================================================================= 
Sub InitSpreadSheet1()
	
	With frm1.vspdData1
		.ReDraw = false
		
		.MaxCols = C_Total				'이번달 합계 
		.MaxRows = 4
		
		.Col = 0
		.Row = 1
		.Text = "불량수"
		.Row = 2
		.Text = "점유율"
		.Row = 3
		.Text = "누적불량수"
		.Row = 4
		.Text = "누적점유율"
		
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit
		
		.ColWidth(0)=12
	
		ggoSpread.SSSetEdit C_Total, "계", 8, 1, -1, 20
		
		Call SetSpreadLock1
		
		.ReDraw = true
		
	End With
End Sub

'==========================================  2.2.6 InitSpreadSheet()  =======================================
'	Name : InitSpreadSheet2()
'	Description : 
'========================================================================================================= 
Sub InitSpreadSheet2()

	With frm1.vspdData2
	
		.ReDraw = false
		
		.MaxCols = D_Total  				' 지난달 합계 
		.MaxRows = 4
		
		.Col = 0
		.Row = 1
		.Text = "불량수"
		.Row = 2
		.Text = "점유율"
		.Row = 3
		.Text = "누적불량수"
		.Row = 4
		.Text = "누적점유율"
		
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit
		
		.ColWidth(0)=12
		
		ggoSpread.SSSetEdit D_Total, "계",  8, 1, -1, 20
		
		Call SetSpreadLock2
		.ReDraw = true
		
	End With
    
End Sub

'==========================================  2.2.6 SetSpreadLock()  =======================================
'	Name : SetSpreadLock1()
'	Description : 
'========================================================================================================= 
Sub SetSpreadLock1()
    ggoSpread.Source = frm1.vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================  2.2.6 SetSpreadLock()  =======================================
'	Name : SetSpreadLock2()
'	Description : 
'========================================================================================================= 
Sub SetSpreadLock2()
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================  2.2.6 InitChartFx()  =======================================
'	Name : InitChartFx()
'	Description : Initialize ChartFx
'========================================================================================================= 
Sub InitChartFx()
	With frm1.ChartFX1
		'Chart Title 및 Font 설정 
		.Title_(0) = "당월"
		.LeftFont.Name = "굴림"
		
		'Chart Series Legend Font 설정 
		.SerLegBoxObj.Font.Name = "굴림"
		
		'그래프의 GAP 지정 
		.TopGap = 5			'그래프의 위쪽 여백 지정 
		.BottomGap = 20		'그래프의 아래쪽 여백 지정 
		.RightGap = 5
		.LeftGap = 50
		
		.MultipleColors = False
		
	End With
	
	With frm1.ChartFX2
		'Chart Title 및 Font 설정 
		.Title_(0) = "전월"
		.LeftFont.Name = "굴림"

		'Chart Series Legend Font 설정 
		.SerLegBoxObj.Font.Name = "굴림"
		
		'그래프의 GAP 지정 
		.TopGap = 5			'그래프의 위쪽 여백 지정 
		.BottomGap = 20		'그래프의 아래쪽 여백 지정 
		.RightGap = 5
		.LeftGap = 50
		
		.MultipleColors = False
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
		.Axis(1).Visible = False
		
		'범례 Clear
		.ClearLegend 1	
		
		'차트 FX와의 데이터 채널 초기화 
		.OpenDataEx 1, 1, 1
		.CloseData 1 Or &H800		'COD_VALUES Or COD_REMOVE
		
		'계열을 안보이게 함 
		.Series(0).Visible = False
	End With
	
	With frm1.ChartFX2
		' X축/Y축 눈금 및 값이 안보이게 함 
		.Axis(2).Visible = False
		.Axis(0).Visible = False
		.Axis(1).Visible = False
		
		'범례 Clear
		.ClearLegend 1	
		
		'차트 FX와의 데이터 채널 초기화 
		.OpenDataEx 1, 1, 1
		.CloseData 1 Or &H800		'COD_VALUES Or COD_REMOVE
		
		'계열을 안보이게 함 
		.Series(0).Visible = False
	End With
End Sub

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
		Param5 = "F"
		Param6 = "최종검사"
		Param7 = Trim(.txtInspItemCd.value)
		Param8 = ""
		Param9 = ""
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
		frm1.txtInspItemCd.Value = arrRet(1)
		frm1.txtInspItemNm.Value = arrRet(2)	
		frm1.txtInspItemCd.Focus
	End If	

	Set gActiveElement = document.activeElement
End Function

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
	arrParam5 = "F"
	
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
	Call ggoOper.FormatDate(frm1.txtYrDt, Parent.gDateFormat, 2)
    Call InitVariables                                                      '⊙: Initializes local global variables

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitSpreadSheet1 
    Call InitSpreadSheet2
'    Call InitChartFX
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
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
'   Event Name : txtYrDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtYrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtYrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYrDt_KeyPress(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtYrDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtYrDt_Change()	
End Sub


'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl, intChartNo)

	Dim strInspClassCd, strInspItemCd, strInspYear, strInspMnth, strItemCd, strPlantCd
	Dim strYYYYMM, strChartTitle
		
	SetPrintCond = False

	If intChartNo = 2 Then
		strYYYYMM = DateAdd("m", -1, frm1.txtYrDt.Value)
		strChartTitle = "전월"
	Else
		strYYYYMM = frm1.txtYrDt.Value
		strChartTitle = "당월"
	End If
	
	strInspYear = Year(strYYYYMM)
	strInspMnth = Month(strYYYYMM)


	strInspClassCd	= FilterVar("F","","SNM")
	strInspYear		= FilterVar(strInspYear,"","SNM")
	strInspMnth	= FilterVar(strInspMnth,"","SNM")
	strItemCd		= FilterVar(frm1.txtItemCd.value,"","SNM")
	strPlantCd		= FilterVar(frm1.txtPlantCd.value,"","SNM")
	strInspItemCd	= FilterVar(frm1.txtInspItemCd.value, "", "SNM")

	If strInspItemCd = "" Then
		strInspItemCd = "%"
	End If 

	StrEbrFile	= "Q3412MA21"

	StrUrl = "insp_class_cd|" & strInspClassCd
	StrUrl = StrUrl & "|insp_year|"	& strInspYear
	StrUrl = StrUrl & "|insp_mnth|"	& strInspMnth
	StrUrl = StrUrl & "|item_cd|"		& strItemCd
	StrUrl = StrUrl & "|plant_cd|"		& strPlantCd
	StrUrl = StrUrl & "|insp_item_cd|"		& strInspItemCd
	StrUrl = StrUrl & "|ChartTitle|"		& strChartTitle

	SetPrintCond = True
	
End Function


Function CallDrawEBChart1()
	Dim StrUrl, StrEbrFile, ObjName

	Call LayerShowHide(1)

	If Not SetPrintCond(StrEbrFile, strUrl, 1) Then
		Exit Function
	End If

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	EBActionA.menu.value = 0
    Call FncEBR5RC2(ObjName, "view", StrUrl,EBActionA,"EBR")

End Function

Function CallDrawEBChart2()
	Dim StrUrl, StrEbrFile, ObjName

	Call LayerShowHide(1)

	If Not SetPrintCond(StrEbrFile, strUrl, 2) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	EBActionB.menu.value = 0
    Call FncEBR5RC2(ObjName, "view", StrUrl,EBActionB,"EBR")

End Function

Function MyBizASP1_onReadyStateChange()
	If lgNoData1 = False then
		If LCase(MyBizASP1.Document.ReadyState) = "complete" Then
			Call CallDrawEBChart2    			'☜: Query db data
		End If
	End If 
End Function


Function MyBizASP2_onReadyStateChange()
	If lgNoData2 = False then
		If LCase(MyBizASP2.Document.ReadyState) = "complete" Then
			If DbQuery = False then
				Exit Function
			End If										'☜: Query db data
		End If
	End If 
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	
	FncQuery = False                                                        '⊙: Processing is NG
	lgNoData1 = False
	lgNoData2 = False

	Err.Clear                                                               '☜: Protect system from crashing
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
	'Erase contents area
	'----------------------- 
	Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
	Call InitVariables									'⊙: Initializes local global variables
	Call InitSpreadSheet1
	Call InitSpreadSheet2
'	Call ClearChartFx
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then						'⊙: This function check indispensable field
		Exit Function
	End If
	
	With frm1
		.vspdData1.focus
		ggoSpread.Source = .vspdData1
		
		Call InitSpreadSheet1
    	End With

	With frm1
		.vspdData2.focus
		ggoSpread.Source = .vspdData2
		
		Call InitSpreadSheet2
    	End With
'	Call ClearChartfx
	
	'-----------------------
	'Query function call area
	'----------------------- 
	Call CallDrawEBChart1
'	If DbQuery = False then
'		Exit Function
'	End If										'☜: Query db data
'
'	FncQuery = True									'⊙: Processing is OK
	
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 

	With frm1
	 
		Dim IntRetCD 
		
		FncNew = False                                                          					'⊙: Processing is NG
		
		'-----------------------
		'Check previous data area
		'-----------------------
		
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
			
		'-----------------------
		'Erase condition area
		'Erase contents area
		'-----------------------
		Call ggoOper.ClearField(Document, "A")
		Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
		
		Call InitVariables                                                      '⊙: Initializes local global variables
		Call SetDefaultVal
		
		ggoSpread.Source = .vspdData1
		Call InitSpreadSheet1
		
		ggoSpread.Source = .vspdData2
		Call InitSpreadSheet2
'		Call ClearChartFX
	    	
	End With
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
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave()
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
    
 	Call LayerShowHide(1)
	
	frm1.txtYr.Value = Left(frm1.txtYrDt.DateValue,4)
	frm1.txtMnth.Value = Mid(frm1.txtYrDt.DateValue,5, 2)
	
	Err.Clear                                                               					'☜: Protect system from crashing
	
	DbQuery = False                                                        					 '⊙: Processing is NG
		
	strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.Value)  '☜:
	strVal = strVal & "&txtYr=" & Trim(frm1.txtYr.value)
	strVal = strVal & "&txtMnth=" & Trim(frm1.txtMnth.value)
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
	strVal = strVal & "&txtInspItemCd=" & Trim(frm1.txtInspItemCd.value)

	Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	Call SetSpreadLock1
	Call SetSpreadLock2
	Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>불량유형파레토표</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=18 tag="13XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
									<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>연월</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtYrDt CLASS=FPDTYYYYMM title=FPDATETIME ALT="연월" tag="12X1"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="13XXXU"><IMG align=top height=20 name=btnItemCd onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
									<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>검사항목</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="검사항목" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspItem()">
											<INPUT TYPE=TEXT NAME="txtInspItemNm" SIZE=20 MAXLENGTH="40" tag="14" >
									</TD>
								</TR>		
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=50% valign=top>
						<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
							<TR>	
								<TD HEIGHT="100%" WIDTH="49%">
									<IFRAME NAME="MyBizASP1"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=AUTO framespacing=0 marginwidth=0 marginheight=0 ></IFRAME>     
								</TD>
								<TD HEIGHT="100%" WIDTH="2%">
								</TD>
								<TD HEIGHT="100%" WIDTH="49%">
									<IFRAME NAME="MyBizASP2"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=AUTO framespacing=0 marginwidth=0 marginheight=0 ></IFRAME>     
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
								<TD HEIGHT="100%" WIDTH="49%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<TD HEIGHT="100%" WIDTH="2%">
								</TD>
								<TD HEIGHT="100%" WIDTH="49%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" rows="1" cols="20" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1>      
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1>      
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1>      
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>      
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1>      
<INPUT TYPE=HIDDEN NAME="txtYr" tag="24" tabindex=-1>      
<INPUT TYPE=HIDDEN NAME="txtMnth" tag="24" tabindex=-1>      
      
</FORM>      
<FORM NAME="EBActionA" ID="EBActionA" TARGET="MyBizASP1" METHOD="POST"  scroll=yes> 
	<input TYPE="HIDDEN" NAME="menu" value=0 > 
	<input TYPE="HIDDEN" NAME="id" > 
	<input TYPE="HIDDEN" NAME="pw" >
	<input TYPE="HIDDEN" NAME="doc" > 
	<input TYPE="HIDDEN" NAME="form" > 
	<input TYPE="HIDDEN" NAME="runvar" >
</FORM>

<FORM NAME="EBActionB" ID="EBActionB" TARGET="MyBizASP2" METHOD="POST"  scroll=yes> 
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

