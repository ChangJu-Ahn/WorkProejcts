<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q3311MA1
'*  4. Program Name         : 품질추이(일별)
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2004/07/27
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Lee Seung Wook
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

Const BIZ_PGM_QRY_ID = "q3311mb1.asp"	
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
Const C_11=11
Const C_12=12
Const C_13=13
Const C_14=14
Const C_15=15
Const C_16=16
Const C_17=17
Const C_18=18
Const C_19=19
Const C_20=20
Const C_21=21
Const C_22=22
Const C_23=23
Const C_24=24
Const C_25=25
Const C_26=26
Const C_27=27
Const C_28=28
Const C_29=29
Const C_30=30
Const C_31=31

Dim C_Total
Dim Rm
Dim IsOpenPop
dim strEbrErr

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

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	C_Total = 32
	frm1.txtYrMnth.Text = CompanyYM
End Sub

'==========================================  2.2.6 InitSpreadSheet()  =======================================
'	Name : InitSpreadSheet()
'	Description : 
'========================================================================================================= 
Sub InitSpreadSheet(Byval Unit)
	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_Total

		.MaxRows = 8
		
		.Col = 0
		.Row = 1
		.Text = "LOT수"
		.Row = 2
		.Text = "불합격LOT수"
		.Row = 3
		.Text = "검사의뢰수"
		.Row = 4
		.Text = "검사수"
		.Row = 5
		.Text = "불량수"

		If Unit <> "" then
			.Row = 6
			.Text = "불량률" & "(" & Unit & ")"
			.Row = 7
			.Text = "LOT불합격률" & "(%)"
		Else
			.Row = 6
			.Text = "불량률"
			.Row = 7
			.Text = "LOT불합격률"
		End If

		.Row = 8
		.Text = "목표" & "(" & Unit & ")"
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit
		
		.ColWidth(0) = 15
		
		ggoSpread.SSSetEdit C_1, "1일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_2, "2일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_3, "3일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_4, "4일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_5, "5일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_6, "6일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_7, "7일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_8, "8일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_9, "9일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_10, "10일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_11, "11일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_12, "12일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_13, "13일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_14, "14일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_15, "15일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_16, "16일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_17, "17일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_18, "18일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_19, "19일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_20, "20일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_21, "21일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_22, "22일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_23, "23일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_24, "24일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_25, "25일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_26, "26일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_27, "27일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_28, "28일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_29, "29일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_30, "30일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_31, "31일", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_Total, "합계", 8, 1, -1, 20
		
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
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
	End If	
	
	frm1.txtPlantCd.Focus
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
	arrParam5 = "P"
	
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
	End If	

	frm1.txtItemCd.Focus
	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

'------------------------------------------  OpenRoutNo()  -------------------------------------------------
'	Name : OpenRoutNo()
'	Description : RoutNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRoutNo()

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
	
	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "품목","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = "라우팅 팝업"					' 팝업 명칭 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtRoutNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				" AND ITEM_CD = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S")
		
	arrParam(5) = "라우팅"			
	
    arrField(0) = "ED10" & parent.gcolsep & "ROUT_NO"							
    arrField(1) = "DESCRIPTION"
    arrField(2) = "ITEM_CD"													
    arrField(3) = "ED10" & parent.gcolsep & "BOM_NO"							
    arrField(4) = "ED10" & parent.gcolsep & "MAJOR_FLG"						
   
    arrHeader(0) = "라우팅"						
    arrHeader(1) = "라우팅명"
    arrHeader(2) = "품목"											
    arrHeader(3) = "BOM Type"					
    arrHeader(4) = "주라우팅"				        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=640px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
    frm1.txtRoutNo.focus
	If arrRet(0) <> "" Then
		frm1.txtRoutNo.Value		= arrRet(0)		
		frm1.txtRoutNoDesc.Value	= arrRet(1)
	Else
		Exit Function
	End If		
	Set gActiveElement = document.activeElement
End Function


'------------------------------------------  OpenOprNo()  -------------------------------------------------
'	Name : OpenOprNo()
'	Description : OprNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function    

	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
	
	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "품목","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = "공정팝업"	
	arrParam(1) = "P_ROUTING_DETAIL A inner join P_WORK_CENTER B on A.wc_cd = B.wc_cd and A.plant_cd = B.plant_cd " & _
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = " & FilterVar("P1006", "''", "S") & "" & _
				  " and A.rout_order in (" & FilterVar("F", "''", "S") & " ," & FilterVar("I", "''", "S") & " ) "				
	arrParam(2) = UCase(Trim(frm1.txtOprNo.Value))
	arrParam(3) = ""
	If (Trim(frm1.txtItemCd.value) <> "" AND Trim(frm1.txtRoutNo.value) <> "") THEN
		arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
					  " and	A.item_cd = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") & _
					  " and	A.rout_no = " & FilterVar(UCase(frm1.txtRoutNo.value), "''", "S")
	ElseIf Trim(frm1.txtRoutNo.value) <> "" THEN
		arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
					  " and	A.rout_no = " & FilterVar(UCase(frm1.txtRoutNo.value), "''", "S")
	ElseIf (Trim(frm1.txtItemCd.value) <> "" AND Trim(frm1.txtRoutNo.value) = "") THEN
		arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
					  " and	A.item_cd = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S")
	Else 		
		arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 
	End If	
	
	arrParam(5) = "공정"			
	
	arrField(0) = "ED10" & parent.gcolsep & "A.OPR_NO"	
	arrField(1) = "ED15" & parent.gcolsep & "C.MINOR_NM"
	arrField(2) = "ED10" & parent.gcolsep & "A.ROUT_NO"
	arrField(3) = "A.ITEM_CD"
	arrField(4) = "ED10" & parent.gcolsep & "A.WC_CD"
	arrField(5) = "ED10" & parent.gcolsep & "A.INSIDE_FLG"
	arrField(6) = "ED10" & parent.gcolsep & "A.INSP_FLG"
	
	arrHeader(0) = "공정"
	arrHeader(1) = "공정작업명"
	arrHeader(2) = "라우팅"
	arrHeader(3) = "품목"		
	arrHeader(4) = "작업장"	
	arrHeader(5) = "사내구분"
	arrHeader(6) = "검사여부"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=640px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtOprNo.focus
	If arrRet(0) <> "" Then
		frm1.txtOprNo.Value	= arrRet(0)
		frm1.txtOprNoDesc.Value	= arrRet(1)
	Else
		Exit Function
	End If		
	Set gActiveElement = document.activeElement
	
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
    Call ggoOper.FormatDate(frm1.txtYrMnth, Parent.gDateFormat, 2)
    
    Call SetDefaultVal
    Call InitSpreadSheet("") 

    
    Call SetToolbar("11000000000011")					'⊙: 버튼 툴바 제어 
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
'   Event Name : txtYrMnth_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrMnth_DblClick(Button)
    If Button = 1 Then
        frm1.txtYrMnth.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtYrMnth.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYrMnth_KeyPress(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYrMnth_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	Dim DefectRatioUnit	
    Dim Yr
	Dim Mnth
	
	FncQuery = False                                                        '⊙: Processing is NG
	strEbrErr= False
	Err.Clear                                                               '☜: Protect system from crashing
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then	Exit Function
	End If

	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then	Exit Function


	With frm1
	
		' ******** SpreadSheet 불량율 단위 셋팅 (AJJ, 030726)	
		If 	CommonQueryRs(" DEFECT_RATIO_UNIT_CD "," Q_DEFECT_RATIO_BY_INSP_CLASS ", " INSP_CLASS_CD = " & FilterVar("R", "''", "S") & "  AND PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("220401","X","X","X")
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		DefectRatioUnit = Trim(lgF0(0))


		'-----------------------
		'Erase contents area
		'----------------------- 
		Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
		Call InitVariables									'⊙: Initializes local global variables


		' ******** SpreadSheet 날짜 셋팅 (AJJ, 030826)
		Yr = Left(.txtYrMnth.DateValue,4)
		Mnth = Mid(.txtYrMnth.DateValue,5, 2)
	
		If Mnth = "01" or Mnth = "03" or Mnth = "05" or Mnth = "07" or Mnth = "08" or Mnth = "10" or Mnth = "12" Then
			C_Total = 32
		ElseIf Mnth = "02" Then
			If CInt(Yr) Mod 4 = 0 Then				'윤년일 경우 2월은 29일로 처리 
				C_Total = 30
			Else
				C_Total = 29
			End If
		Else
			C_Total = 31
		End If

		.txtYr.Value = Yr
		.txtMnth.Value =  Mnth
		.txtCTotal.Value = C_Total
	
		.vspdData.focus
		ggoSpread.Source = .vspdData

	End With
	

	Call InitSpreadSheet(DefectRatioUnit)

	'-----------------------
	'Query function call area
	'----------------------- 
'	If DbQuery = False then	Exit Function	'☜: Query db data
	Call EBRok
    	
	FncQuery = True
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     
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


'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl)

	Dim strInspClassCd, strYYYYMM, strInspFrDt, strInspToDt, strItemCd, strPlantCd
	Dim strOprNo,strRoutNo
		
	SetPrintCond = False
	
	strInspFrDt = UNIGetFirstDay(frm1.txtYrMnth.text,Parent.gServerDateFormat)
	strInspToDt = UNIGetLastDay(frm1.txtYrMnth.text,Parent.gServerDateFormat)
	
	strYYYYMM		= FilterVar(frm1.txtYrMnth.Text,"","SNM")
	strInspClassCd	= FilterVar("P","","SNM")
	strInspFrDt		= FilterVar(strInspFrDt,"","SNM")
	strInspToDt		= FilterVar(strInspToDt,"","SNM")
	strItemCd		= FilterVar(frm1.txtItemCd.value,"","SNM")
	strPlantCd		= FilterVar(frm1.txtPlantCd.value,"","SNM")
	
	If Trim(frm1.txtOprNo.value)  = "" Then
	   strOprNo = "%"
	Else
	   strOprNo =FilterVar(frm1.txtOprNo.value,"","SNM")
	End IF
	
	If Trim(frm1.txtRoutNo.value)  = "" Then
	   strRoutNo = "%"
	Else
	   strRoutNo =FilterVar(frm1.txtRoutNo.value,"","SNM")
	End IF
	
	If frm1.RadioDRType.rdoCase1.Checked = True Then
		StrEbrFile	= "Q3311MA12"
	Else
		StrEbrFile	= "Q3311MA11"
	End If
	

	StrUrl = StrUrl & "YyyyMm|"			& strYYYYMM
	StrUrl = StrUrl & "|insp_class_cd|" & strInspClassCd
	StrUrl = StrUrl & "|insp_dt_fr|"	& strInspFrDt
	StrUrl = StrUrl & "|insp_dt_to|"	& strInspToDt
	StrUrl = StrUrl & "|item_cd|"		& strItemCd
	StrUrl = StrUrl & "|plant_cd|"		& strPlantCd
	StrUrl = StrUrl & "|Opr_No|"			& strOprNo
	StrUrl = StrUrl & "|Rout_No|"		& strRoutNo
	
	

	SetPrintCond = True
	
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal

	Call LayerShowHide(1)
	
	If frm1.RadioDRType.rdoCase1.Checked = True Then
		frm1.txtDataFlag.Value = "L"
	Else
		frm1.txtDataFlag.Value = "Q"
	End If
	
	Err.Clear                                                               					'☜: Protect system from crashing

	DbQuery = False                                                        					 '⊙: Processing is NG
		
	strVal = BIZ_PGM_QRY_ID & "?txtDataFlag=" & frm1.txtDataFlag.Value _
							& "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value) _
							& "&txtItemCd=" & Trim(frm1.txtItemCd.Value) _
							& "&txtYr=" & Trim(frm1.txtYr.Value) _
							& "&txtMnth=" & Trim(frm1.txtMnth.Value) _
							& "&txtRoutNo=" & Trim(frm1.txtRoutNo.value) _
							& "&txtOprNo=" & Trim(frm1.txtOprNo.value) _
							& "&txtCTotal=" & Trim(frm1.txtCTotal.value)
	
	Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True
End Function


Sub MyBizASP1_onreadystatechange()
      If    strEbrErr = False Then   
			if lcase(MyBizASP1.document.Readystate) = "complete" then
			  
			   If DbQuery = False then Exit Sub

			end if
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
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어 
	
	frm1.vspdData.Redraw = False
	Call SetSpreadLock
	frm1.vspdData.Redraw = True
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품질추이(일별)</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=18 ALT="공장" tag="13XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
									<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>연월</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtYrMnth CLASS=FPDTYYYYMM title=FPDATETIME ALT="연월" tag="13"></OBJECT>');</SCRIPT>
									</TD>									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="13XXXU"><IMG align=top height=20 name=btnItemCd onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
									<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>라우팅</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="11XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutNo()">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="11XXXU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()">&nbsp;<input TYPE=TEXT NAME="txtOprNoDesc" SIZE="30" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>DATA</TD>      
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioDRType" TAG="1X" ID="rdoCase1"><LABEL FOR="rdoCase1" >LOT불합격률</LABEL>
														   <INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioDRType" TAG="1X" ID="rdoCase2"><LABEL FOR="rdoCase2">불량률</LABEL>      
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
					<TD WIDTH=100% HEIGHT=58% valign=top>      
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
<INPUT TYPE=HIDDEN NAME="txtYr" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtMnth" tag="24" tabindex=-1 >      
<INPUT TYPE=HIDDEN NAME="txtCTotal" tag="24" tabindex=-1 >      
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

