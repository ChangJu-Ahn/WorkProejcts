<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1111MA1
'*  4. Program Name         : 연/월 품질목표등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG090,PQBG100
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_QRY_ID = "Q1115MB1.asp"
Const BIZ_PGM_SAVE_ID = "Q1115MB2.asp"								
Const BIZ_PGM_DEL_ID = "Q1115MB3.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop          
Dim strYr
Dim strMonth
Dim strDay

Call ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gServerDateType, strYr, strMonth, strDay)

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                       	               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                	              				'⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       		'⊙: Initializes Group View Size
    
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False							'☆: 사용자 변수 초기화 
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
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd1.value = UCase(Parent.gPlant)
		frm1.txtPlantNm1.value = Parent.gPlantNm
		frm1.txtPlantCd2.value = UCase(Parent.gPlant)
		frm1.txtPlantNm2.value = Parent.gPlantNm
	End If

	frm1.txtYr1.Text = strYr	
	frm1.txtYr2.Text = strYr
	frm1.cboInspClassCd1.value = "R"
	frm1.cboInspClassCd2.value = "R"
End Sub

'========================================================================================================= 
'	Name : SetFormatfpDoubleSingle()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetFormatfpDoubleSingle()
	frm1.txtYrTargetValue.DecimalPoint = Parent.gComNumDec
	frm1.txtYrTargetValue.Separator = Parent.gComNum1000
	
	frm1.txtMnthTargetValue1.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue2.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue3.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue4.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue5.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue6.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue7.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue8.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue9.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue10.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue11.DecimalPoint = Parent.gComNumDec
	frm1.txtMnthTargetValue12.DecimalPoint = Parent.gComNumDec
	
	frm1.txtMnthTargetValue1.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue2.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue3.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue4.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue5.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue6.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue7.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue8.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue9.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue10.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue11.Separator = Parent.gComNum1000
	frm1.txtMnthTargetValue12.Separator = Parent.gComNum1000
End Sub

'==========================================  2.2.2 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " major_cd=" & FilterVar("Q0001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    Call SetCombo2(frm1.cboInspClassCd1 ,lgF0  ,lgF1  ,Chr(11))
	Call SetCombo2(frm1.cboInspClassCd2 ,lgF0  ,lgF1  ,Chr(11))

    Call CommonQueryRs(" DEFECT_RATIO_UNIT_CD "," Q_DEFECT_RATIO_UNIT ","", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDefectRatioUnitCd ,lgF0  ,lgF0  ,Chr(11))
End Sub

'==========================2.2.3 YearlyToMonthly()==========================================
'   Name : YearlyToMonthly
'   Desc : Setting Monthly Values with the Yearly Value
'==========================================================================================
Sub YearlyToMonthly()
	Dim IntRetCD
	Dim strYearValue
	
	IntRetCD = DisplayMsgBox("220505", Parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then	Exit Sub
    
	With frm1
		strYearValue = Trim(.txtYrTargetValue.Text)
	
		.txtMnthTargetValue1.Text 	= strYearValue
		.txtMnthTargetValue2.Text 	= strYearValue
		.txtMnthTargetValue3.Text 	= strYearValue
		.txtMnthTargetValue4.Text 	= strYearValue
		.txtMnthTargetValue5.Text 	= strYearValue
		.txtMnthTargetValue6.Text 	= strYearValue
		.txtMnthTargetValue7.Text 	= strYearValue
		.txtMnthTargetValue8.Text 	= strYearValue
		.txtMnthTargetValue9.Text 	= strYearValue
		.txtMnthTargetValue10.Text 	= strYearValue
		.txtMnthTargetValue11.Text 	= strYearValue
		.txtMnthTargetValue12.Text 	= strYearValue
	End With
End Sub

'==========================2.2.4 MonthlyToYearly()==========================================
'   Name : MonthlyToYearly
'   Desc : Setting Yearly Values With the Average of Monthly Values
'==========================================================================================
Sub MonthlyToYearly()
	Dim dblSum
	Dim IntRetCD
	dblSum = 0
	IntRetCD = DisplayMsgBox("220605", Parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then	Exit Sub
    
	With frm1
		If Trim(.txtMnthTargetValue1.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue1.Text))
		End If
		If Trim(.txtMnthTargetValue2.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue2.Text))
		End If
		If Trim(.txtMnthTargetValue3.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue3.Text))
		End If
		If Trim(.txtMnthTargetValue4.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue4.Text))
		End If
		If Trim(.txtMnthTargetValue5.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue5.Text))
		End If
		If Trim(.txtMnthTargetValue6.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue6.Text))
		End If
		If Trim(.txtMnthTargetValue7.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue7.Text))
		End If
		If Trim(.txtMnthTargetValue8.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue8.Text))
		End If
		If Trim(.txtMnthTargetValue9.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue9.Text))
		End If
		If Trim(.txtMnthTargetValue10.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue10.Text))
		End If
		If Trim(.txtMnthTargetValue11.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue11.Text))
		End If
		If Trim(.txtMnthTargetValue12.Text) <> "" Then
			dblSum = dblSum + UNICDbl(Trim(.txtMnthTargetValue12.Text))
		End If
		
		.txtYrTargetValue.Text = UNIFormatNumber(CStr(dblSum / 12), 2, -2, 0, 3, 0)
	End With
End Sub

'------------------------------------------  OpenPlant1()  -------------------------------------------------
'	Name : OpenPlant1()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant1() 
	OpenPlant1 = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
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
		frm1.txtPlantCd1.Value    = arrRet(0)		
		frm1.txtPlantNm1.Value    = arrRet(1)		
	End If	
	
	frm1.txtPlantCd1.Focus
	Set gActiveElement = document.activeElement
	OpenPlant1 = True
End Function

'------------------------------------------  OpenPlant2()  -------------------------------------------------
'	Name : OpenPlant2()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant2()
	OpenPlant2 = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If UCase(frm1.txtPlantCd2.ClassName) = UCase(Parent.UCN_PROTECTED) Then
		Exit Function
	End If
	IsOpenPop = True

	arrParam(0) = "공장팝업"	
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
		frm1.txtPlantCd2.Value    = arrRet(0)		
		frm1.txtPlantNm2.Value    = arrRet(1)		
	End If	
	
	frm1.txtPlantCd2.Focus
	Set gActiveElement = document.activeElement
	OpenPlant2 = true
End Function

'==========================================  3.1.1 Form_load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6", "10", "2")
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtYr1, parent.gDateFormat, 3)	
	Call ggoOper.FormatDate(frm1.txtYr2, parent.gDateFormat, 3)
	Call InitVariables																			'⊙: Initializes local global variables
	Call InitComboBox
	Call SetToolBar("11101000000011")
	Call SetDefaultVal
	
	Call SetSingleFocus
	Call SetFormatfpDoubleSingle
	lgBlnFlgChgValue = False
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtYr1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYr1_DblClick(Button)
    If Button = 1 Then
        frm1.txtYr1.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtYr1.Focus
        Set gActiveElement = document.activeElement
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYr2_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtYr2_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtYr2_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYr2_DblClick(Button)
    If Button = 1 Then
        frm1.txtYr2.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtYr2.Focus
        Set gActiveElement = document.activeElement
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYr1_KeyPress(Button)
'   Event Desc : 달력을 호출한다.
'======================================================================================================
Function txtYr1_KeyPress(KeyAscii)
	txtYr1_KeyPress = false
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
	txtYr1_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtYr2_KeyPress(Button)
'   Event Desc : 달력을 호출한다.
'======================================================================================================
Function txtYr2_KeyPress(KeyAscii)
	txtYr2_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtYr2_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtYrTargetValue_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtYrTargetValue_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue1_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue1_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue2_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue2_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue3_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue3_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue4_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue4_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue5_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue5_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue6_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue6_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue7_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue7_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue8_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue8_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue9_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue9_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue10_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue10_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue11_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue11_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtMnthTargetValue12_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtMnthTargetValue12_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : cboInspClassCd2_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboInspClassCd2_onchange()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : cboInspClassCd2_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboDefectRatioUnitCd_onchange()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtYrTargetValue_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtYrTargetValue_KeyPress(KeyAscii)
	txtYrTargetValue_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtYrTargetValue_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtYrTargetValue_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtYrTargetValue_KeyPress(KeyAscii)
	txtYrTargetValue_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtYrTargetValue_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue1_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue1_KeyPress(KeyAscii)
	txtMnthTargetValue1_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue1_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue2_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue2_KeyPress(KeyAscii)
	txtMnthTargetValue2_KeyPress = FALSE
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue2_KeyPress = TRUE
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue3_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue3_KeyPress(KeyAscii)
	txtMnthTargetValue3_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue3_KeyPress = false
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue4_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue4_KeyPress(KeyAscii)
	txtMnthTargetValue4_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue4_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue5_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue5_KeyPress(KeyAscii)
	txtMnthTargetValue5_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue5_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue6_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue6_KeyPress(KeyAscii)
	txtMnthTargetValue6_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue6_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue7_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue7_KeyPress(KeyAscii)
	txtMnthTargetValue7_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue7_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue8_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue8_KeyPress(KeyAscii)
	txtMnthTargetValue8_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue8_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue9_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue9_KeyPress(KeyAscii)
	txtMnthTargetValue9_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue9_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue10_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue10_KeyPress(KeyAscii)
	txtMnthTargetValue10_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue10_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue11_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue11_KeyPress(KeyAscii)
	txtMnthTargetValue11_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue11_KeyPress = true
End Function

'=======================================================================================================
'   Event Name : txtMnthTargetValue12_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function txtMnthTargetValue12_KeyPress(KeyAscii)
	txtMnthTargetValue12_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtMnthTargetValue12_KeyPress = true
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	
	FncQuery = False                                                        '⊙: Processing is NG
	
	Err.Clear                                                            		   '☜: Protect system from crashing
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
    End If
	
	'-----------------------
	'Check condition area
	'-----------------------
	If Not ChkField(Document, "1") Then	Exit Function
	
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	Call InitVariables
	
	Call ggoOper.LockField(Document, "N")								'⊙: This function lock the suitable field
    
    '-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False then	Exit Function									'☜: Query db data
	
	FncQuery = True
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    
	FncNew = False                                                          '⊙: Processing is NG
	
	Err.Clear
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
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call SetDefaultVal
	lgBlnFlgChgValue = False                	              				'⊙: Indicates that no value changed
	Call SetToolBar("11101000000011")
	
	Call SetSingleFocus
	
	FncNew = True
End Function

'========================================================================================
' Function Name : FncNew2
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew2() 
	Dim IntRetCD 
    
	FncNew2 = False                                                          '⊙: Processing is NG
	
	Err.Clear
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")                                         '⊙: Clear Contents  Field
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call InitVariables                                                      '⊙: Initializes local global variables
	lgBlnFlgChgValue = False                	              				'⊙: Indicates that no value changed
	Call SetToolBar("11101000000011")
	
	Call SetSingleFocus
	
	FncNew2 = True
End Function

'========================================================================================
' Function Name : SetSingleFocus
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Sub SetSingleFocus()
	If Trim(frm1.txtPlantCd1.Value) = "" Then
		frm1.txtPlantCd1.focus 
	Else
		frm1.cboInspClassCd1.focus 
	End If
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 

	Dim IntRetCD 
	
	FncDelete = False                                                       '⊙: Processing is NG
	
	Err.Clear 
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then	Exit Function
	
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then Exit Function
	
	FncDelete = True                                                        '⊙: Processing is OK                   							'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 	
	Dim IntRetCD 
	
	FncSave = False                                                  		       '⊙: Processing is NG

	Err.Clear                                                            	 		  '☜: Protect system from crashing
	
	On Error Resume Next                                           	       '☜: Protect system from crashing
	   
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then Exit Function
    	
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False Then Exit Function                              '☜: Save db data
    
	FncSave = True
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
	Dim IntRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If

	lgIntFlgMode = Parent.OPMD_CMODE														'⊙: Indicates that current mode is Crate mode
	lgBlnFlgChgValue = True
	Call ggoOper.ClearField(Document, "1")                                      					'⊙: Clear Condition Field
	Call ggoOper.LockField(Document, "N")												'⊙: This function lock the suitable field

	Call SetToolBar("11101000000011")
	FncCopy = TRUE
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = false                                    					            		'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	FncInsertRow = false                                                 						'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = false                                                 						'☜: Protect system from crashing
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
	FncPrev = false
	Dim strVal
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	ElseIf lgPrevNo = "" Then
	 	Call DisplayMsgBox("900011", "X", "X", "X")
	 	Exit Function
	End If
	
	strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd1.value)	_
							& "&txtYr=" & frm1.txtYr1.Text _
							& "&cboInspClassCd=" & lgPrevNo									'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)
	FncPrev = true 
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	FncNext = false
	Dim strVal
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")
		Exit Function
	End If
	
	strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd1.value)	_
							& "&txtYr=" & frm1.txtYr1.Text _
							& "&cboInspClassCd=" & lgNextNo	
	
	Call RunMyBizASP(MyBizASP, strVal)
	FncNext = true
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLE)					 '☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
	FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	Err.Clear                                                               					'☜: Protect system from crashing
	
	Call LayerShowHide(1)
	
	DbDelete = False															'⊙: Processing is NG
	
	Dim strVal
	
	strVal = BIZ_PGM_DEL_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd2.value) _
							& "&txtYr=" & frm1.txtYr2.Text _
							& "&cboInspClassCd=" & Trim(frm1.cboInspClassCd2.value)
	
	Call RunMyBizASP(MyBizASP, strVal)				
	
	DbDelete = True
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()
	DbDeleteOk = false
	lgBlnFlgChgValue = False
	Call FncNew2()
	DbDeleteOk = true
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	
	DbQuery = False                                                        				'⊙: Processing is NG
	
	Err.Clear                                                               					'☜: Protect system from crashing
	
	Call LayerShowHide(1)
	
	strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd1.value)	_
							& "&txtYr=" & frm1.txtYr1.Text _
							& "&cboInspClassCd=" & Trim(frm1.cboInspClassCd1.value)
	
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()															'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE
	lgBlnFlgChgValue = False											'⊙: Indicates that current mode is Update mode
	Call ggoOper.LockField(Document, "Q")
	Call SetToolBar("11111000001111")	
	
	frm1.txtYrTargetValue.Focus
	Set gActiveElement = document.activeElement
			
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : 
'========================================================================================
Function DbSave() 
	Err.Clear																	'☜: Protect system from crashing
	Call LayerShowHide(1)
	DbSave = False															'⊙: Processing is NG

	Dim strVal
	With frm1
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	End With

	DbSave = True
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()	
	With frm1
		.txtPlantCd1.Value = .txtPlantCd2.Value
		.txtYr1.Text = .txtYr2.Text
		.cboInspClassCd1.Value = .cboInspClassCd2.Value
	End With
	
	Call InitVariables
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>연/월 품질목표</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
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
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT NAME="txtPlantCd1" SIZE="10" MAXLENGTH="4" ALT="공장코드" TAG="12XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnPlantCd ONCLICK=vbscript:OpenPlant1() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm1" TAG="14X" >
									</TD>
									<TD CLASS="TD5" NOWRAP>연도</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1115ma1_txtYr1_txtYr1.js'></script>
									</TD>
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>검사분류</TD>
									 <TD CLASS=TD6 NOWRAP><SELECT NAME="cboInspClassCd1" ALT="검사분류" STYLE="WIDTH: 150px" TAG="12"></SELECT></TD>
									 <TD CLASS=TD5 NOWRAP></TD>
									 <TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* vALIGN=top>
						<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD WIDTH=50% HEIGHT=* vALIGN=top>
									<FIELDSET CLASS="CLSFLD">
										<LEGEND>연목표</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<td CLASS="TD5" NOWPAP HEIGHT=5></td>
												<td CLASS="TD6" NOWPAP HEIGHT=5></td>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>공장</TD>
												<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlantCd2" SIZE="10" MAXLENGTH="4" ALT="공장코드" TAG="23XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnPlantCd ONCLICK=vbscript:OpenPlant2() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON" tag="25">&nbsp;<INPUT NAME="txtPlantNm2" TAG="24" ></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>연도</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q1115ma1_txtYr2_txtYr2.js'></script>
												</TD>
											</TR>
											<TR>
												 <TD CLASS=TD5 NOWRAP>검사분류</TD>
												 <TD CLASS=TD6 NOWRAP><SELECT NAME="cboInspClassCd2" ALT="검사분류" STYLE="WIDTH: 150px" TAG="23"></SELECT></TD>
											</TR>							
											<TR>
												<TD CLASS="TD5" NOWRAP>연목표치</TD>
												<TD CLASS="TD6" NOWRAP>
													<script language =javascript src='./js/q1115ma1_txtYrTargetValue_txtYrTargetValue.js'></script>
												</TD>
											</TR>
											<TR>
												 <TD CLASS=TD5 NOWRAP>불량률단위</TD>
												 <TD CLASS=TD6 NOWRAP><SELECT NAME="cboDefectRatioUnitCd" ALT="불량률단위" STYLE="WIDTH: 70px" TAG="22"></SELECT></TD>
											</TR>
											<!--											
											<TR>
												<TD CLASS=TD5 NOWRAP>불량률단위</TD>
												<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDefectRatioUnitCd" SIZE="10" STYLE="Text-Align: Right" ALT="불량률단위" TAG="24" ></TD>
											</TR>
											-->
											<TR>
												<td CLASS="TD5" NOWPAP HEIGHT=5></td>
												<td CLASS="TD6" NOWPAP HEIGHT=5></td>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
								<TD WIDTH=50% HEIGHT=* vALIGN=top>
									<FIELDSET CLASS="CLSFLD">
									<LEGEND>월목표</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
											<TR>
												<td CLASS="TD5" NOWPAP HEIGHT=5></td>
												<td CLASS="TD6" NOWPAP HEIGHT=5></td>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>1월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue1_txtMnthTargetValue1.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>2월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue2_txtMnthTargetValue2.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>3월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue3_txtMnthTargetValue3.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>4월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue4_txtMnthTargetValue4.js'></script>
												</TD>
											</TR>								
											<TR>
												<TD CLASS="TD5" NOWRAP>5월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue5_txtMnthTargetValue5.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>6월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue6_txtMnthTargetValue6.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>7월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue7_txtMnthTargetValue7.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>8월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue8_txtMnthTargetValue8.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>9월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue9_txtMnthTargetValue9.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>10월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue10_txtMnthTargetValue10.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>11월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue11_txtMnthTargetValue11.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>12월</TD>
												<TD CLASS="TD6" NOWRAP>
												<script language =javascript src='./js/q1115ma1_txtMnthTargetValue12_txtMnthTargetValue12.js'></script>
												</TD>
											</TR>
											<TR>
												<td CLASS="TD5" NOWPAP HEIGHT=5></td>
												<td CLASS="TD6" NOWPAP HEIGHT=5></td>
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
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
	      	<TD WIDTH="100%" >
	      		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
	        		<TR>
	        			<TD WIDTH=10>&nbsp;</TD>
	        			<TD>
	        				<BUTTON NAME="btnMonthlyToYearly" CLASS="CLSMBTN" ONCLICK="vbscript:MonthlyToYearly()">연목표 적용</BUTTON>&nbsp;
	        				<BUTTON NAME="btnYearlyToMonthly" CLASS="CLSMBTN" ONCLICK="vbscript:YearlyToMonthly()">월목표 적용</BUTTON>
	        			</TD>
	        			<TD WIDTH=*>&nbsp;</TD>
	        		</TR>
	      		</TABLE>
	      	</TD>
         </TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" TAG="24" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
