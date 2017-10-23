<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q3211MA1
'*  4. Program Name         : ǰ������(�Ϻ�)
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/26
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Ahn Jung Je
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
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit																'��: indicates that All variables must be declared in advance

Const BIZ_PGM_QRY_ID = "q3211mb1.asp"							'��: Query �����Ͻ� ���� ASP�� 
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

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
Dim CompanyYM,CompanyYMD, LastDate, FirstDate
CompanyYM = UNIMonthClientFormat(UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gAPDateFormat))
CompanyYMD = UniConvDateAToB(CompanyYM, Parent.gServerDateFormat, Parent.gAPDateFormat)

LastDate = UNIGetLastDay(CompanyYMD,Parent.gServerDateFormat)
FirstDate = UNIGetFirstDay(CompanyYMD,Parent.gServerDateFormat)

'--------------- ������ coding part(�������,End)------------------------------------------------------------- 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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
		.Text = "LOT��"
		.Row = 2
		.Text = "���հ�LOT��"
		.Row = 3
		.Text = "�˻��Ƿڼ�"
		.Row = 4
		.Text = "�˻��"
		.Row = 5
		.Text = "�ҷ���"
		
		If Unit <> "" then
			.Row = 6
			.Text = "�ҷ���" & "(" & Unit & ")"
			.Row = 7
			.Text = "LOT���հݷ�" & "(%)"
		Else
			.Row = 6
			.Text = "�ҷ���"
			.Row = 7
			.Text = "LOT���հݷ�"
		End If
				
		.Row = 8
		.Text = "��ǥ" & "(" & Unit & ")"
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit
		
		.ColWidth(0)=15
		
		ggoSpread.SSSetEdit C_1, "1��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_2, "2��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_3, "3��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_4, "4��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_5, "5��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_6, "6��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_7, "7��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_8, "8��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_9, "9��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_10, "10��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_11, "11��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_12, "12��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_13, "13��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_14, "14��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_15, "15��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_16, "16��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_17, "17��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_18, "18��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_19, "19��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_20, "20��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_21, "21��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_22, "22��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_23, "23��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_24, "24��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_25, "25��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_26, "26��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_27, "27��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_28, "28��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_29, "29��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_30, "30��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_31, "31��", 8, 1, -1, 20
		ggoSpread.SSSetEdit C_Total, "�հ�", 8, 1, -1, 20
		
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
		'Chart Title �� Font ���� 
		.Title_(0) = "LOT���հݷ� ���̵�"
		.LeftFont.Name = "����"
		.Axis(0).decimals = 4
		
		'Chart Series Legent Font ���� 
		.SerLegBoxObj.Font.Name = "����"
		
		'�׷����� GAP ���� 
		.TopGap = 5			'�׷����� ���� ���� ���� 
		.BottomGap = 20			'�׷����� �Ʒ��� ���� ���� 
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
		' X��/Y�� ���� �� ���� �Ⱥ��̰� �� 
		.Axis(2).Visible = False
		.Axis(0).Visible = False
		
		'���� Clear
		.ClearLegend 1
		
		'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.OpenDataEx 1, 1, 1
		.CloseData 1 Or &H800		'COD_VALUES Or COD_REMOVE
		
		'�迭�� �Ⱥ��̰� �� 
		.Series(0).Visible = False
		.Title_(0) = "LOT���հݷ� ���̵�"
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

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
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

	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'���������� �ʿ��մϴ� 
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
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
	End If	

	frm1.txtItemCd.Focus
	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

'------------------------------------------  OpenBp()  -------------------------------------------------
'	Name : OpenBp()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó �˾�"					' �˾� ��Ī 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBpCd.Value)					' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "(BP_TYPE = " & FilterVar("CS", "''", "S") & " Or BP_TYPE = " & FilterVar("S", "''", "S") & " )"			' Where Condition	
	arrParam(5) = "����ó"						' �����ʵ��� �� ��Ī	
	
    arrField(0) = "BP_CD"								' Field��(0)
    arrField(1) = "BP_NM"								' Field��(1)
    
    arrHeader(0) = "����ó�ڵ�"					' Header��(0)
    arrHeader(1) = "����ó��"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
	End If	

	frm1.txtBpCd.Focus
	Set gActiveElement = document.activeElement
	OpenBp = true
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitVariables                                                      '��: Initializes local global variables

    '----------  Coding part  -------------------------------------------------------------
    Call ggoOper.FormatDate(frm1.txtYrMnth, Parent.gDateFormat, 2)
    
    Call SetDefaultVal
    Call InitSpreadSheet("") 
    'Call InitChartFX
    
    Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 
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
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
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
	Dim StrUrl, StrEbrFile, ObjName
	
	FncQuery = False                                                        '��: Processing is NG
	
	Err.Clear                                                               '��: Protect system from crashing
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then Exit Function
	End If

	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then	Exit Function
	
	strEbrErr = False

	With frm1
	
		' ******** SpreadSheet �ҷ��� ���� ���� (AJJ, 030726)	
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
		Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
		Call InitVariables									'��: Initializes local global variables


		' ******** SpreadSheet ��¥ ���� (AJJ, 030826)
		Yr = Left(.txtYrMnth.DateValue,4)
		Mnth = Mid(.txtYrMnth.DateValue,5, 2)
	
		If Mnth = "01" or Mnth = "03" or Mnth = "05" or Mnth = "07" or Mnth = "08" or Mnth = "10" or Mnth = "12" Then
			C_Total = 32
		ElseIf Mnth = "02" Then
			If CInt(Yr) Mod 4 = 0 Then				'������ ��� 2���� 29�Ϸ� ó�� 
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
	
    'Call ClearChartfx	
	Call InitSpreadSheet(DefectRatioUnit)

	'-----------------------
	'Query function call area
	'----------------------- 
	'If DbQuery = False then Exit Function
	Call EbrOk
	
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
	FncPrint = False
    Call parent.FncPrint()                                              
    FncPrint = True
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel()
	FncExcel = False 
	Call parent.FncExport(Parent.C_MULTI)
	FncExcel = True 
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
		
	SetPrintCond = False
	
	strInspFrDt = UNIGetFirstDay(frm1.txtYrMnth.text,Parent.gServerDateFormat)
	strInspToDt = UNIGetLastDay(frm1.txtYrMnth.text,Parent.gServerDateFormat)
	
	strYYYYMM		= FilterVar(frm1.txtYrMnth.Text,"","SNM")
	strInspClassCd	= FilterVar("R","","SNM")
	strInspFrDt		= FilterVar(strInspFrDt,"","SNM")
	strInspToDt		= FilterVar(strInspToDt,"","SNM")
	strItemCd		= FilterVar(frm1.txtItemCd.value,"","SNM")
	strPlantCd		= FilterVar(frm1.txtPlantCd.value,"","SNM")
	
	If frm1.RadioDRType.rdoCase1.Checked = True Then
		StrEbrFile	= "Q3211MA12"
	Else
		StrEbrFile	= "Q3211MA11"
	End If
	

	StrUrl = StrUrl & "YyyyMm|"			& strYYYYMM
	StrUrl = StrUrl & "|insp_class_cd|" & strInspClassCd
	StrUrl = StrUrl & "|insp_dt_fr|"	& strInspFrDt
	StrUrl = StrUrl & "|insp_dt_to|"	& strInspToDt
	StrUrl = StrUrl & "|item_cd|"		& strItemCd
	StrUrl = StrUrl & "|plant_cd|"		& strPlantCd

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
	
	Err.Clear                                                               					'��: Protect system from crashing

	DbQuery = False                                                        					 '��: Processing is NG
	
	strVal = BIZ_PGM_QRY_ID & "?txtDataFlag=" & frm1.txtDataFlag.Value _
							& "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value) _
							& "&txtItemCd=" & Trim(frm1.txtItemCd.Value) _
							& "&txtYr=" & Trim(frm1.txtYr.Value) _
							& "&txtMnth=" & Trim(frm1.txtMnth.Value) _
							& "&txtBpCd=" & Trim(frm1.txtBpCd.value) _
							& "&txtCTotal=" & Trim(frm1.txtCTotal.value)
	
	Call RunMyBizASP(MyBizASP, strVal)							'��: �����Ͻ� ASP �� ���� 
	
	DbQuery = True                                                          					'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	Call SetToolbar("11000000000111")										'��: ��ư ���� ���� 
	
	frm1.vspdData.Redraw = False
	Call SetSpreadLock
	frm1.vspdData.Redraw = True
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
' Function Name : DBQueryErr
' Function Desc :  
'========================================================================================
Function DBQueryErr()														'��: ��ȸ ������ ������� 
   strEbrErr = True 
   MyBizASP1.Document.location.href="../../blank.htm"
   
End Function




</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<!-- SPACE AREA-->
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- TAB, REFERENCE AREA -->
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ������(�Ϻ�)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    	</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- CONDITION, CONTENT AREA -->
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<!-- SAPCE AREA -->
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<!-- CONDITION AREA -->
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS=CLSFLD>
							<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_40%>>		
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=18 ALT="����" tag="13XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
									<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtYrMnth CLASS=FPDTYYYYMM title=FPDATETIME ALT="����" tag="13"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="ǰ��" tag="13XXXU"><IMG align=top height=20 name=btnItemCd onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
									<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=20 ALT="����ó" tag="11XXXU"><IMG align=top height=20 name=btnBpCd onclick=vbscript:OpenBp() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
	
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>DATA</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioDRType" TAG="1X" ID="rdoCase1"><LABEL FOR="rdoCase1" >LOT���հݷ�</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioDRType" TAG="1X" ID="rdoCase2"><LABEL FOR="rdoCase2">�ҷ���</LABEL>      
									</TD>      
									<td CLASS="TD5" NOWPAP HEIGHT=5></td>      
									<td CLASS="TD6" NOWPAP HEIGHT=5></td>      
								</TR>      
							</TABLE>      
						</FIELDSET>      
					</TD>      
				</TR>      
				<!-- SAPCE AREA -->
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<!-- CONTENT AREA(MULTI) -->
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
				<!--<TR>
					<TD WIDTH=100% HEIGHT=58% valign=top>
						<TABLE width=1016  height=476 cellspacing=0 cellpadding=0 border=0>
							<TR>
								<TD><IFRAME NAME="MyBizASP1"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=auto framespacing=0 marginwidth=0 marginheight=0 ></IFRAME></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>-->      
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
	<!-- SPACE AREA -->
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<!-- BATCH,JUMP AREA -->
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
	<!-- IFRAME AREA -->
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

