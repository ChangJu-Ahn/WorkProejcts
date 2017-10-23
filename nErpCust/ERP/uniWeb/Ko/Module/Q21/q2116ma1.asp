<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2116MA1
'*  4. Program Name         : 불합격통지 등록 
'*  5. Program Desc         : Quality Configuration
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                       

Const BIZ_PGM_QRY_ID = "Q2116MB1.asp"
Const BIZ_PGM_SAVE_ID = "Q2116MB2.asp"
Const BIZ_PGM_DEL_ID = "Q2116MB3.asp"											 '☆: 비지니스 로직 ASP명 

Const BIZ_PGM_JUMP1_ID = "Q2111MA1"
Const BIZ_PGM_JUMP2_ID = "Q2117MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgNextNo					'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

Dim IsOpenPop          

<% 
'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYMD

CompanyYMD  = GetSvrDate
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 
%>

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                       				'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                	              		'⊙: Indicates that no value changed
    lgIntGrpCount = 0                                               '⊙: Initializes Group View Size
    
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
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
		
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If
		
	If ReadCookie("txtInspReqNo") <> "" Then
		frm1.txtInspReqNo.Value = ReadCookie("txtInspReqNo")
	End If
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtInspReqNo", ""	
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description :Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
	
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
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
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

'------------------------------------------  OpenInspReqNo()  -------------------------------------------------
'	Name : OpenInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspReqNo()      
	OpenInspReqNo = false     

	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)
	Param2 = Trim(frm1.txtPlantNm.Value)
	Param3 = Trim(frm1.txtInspReqNo.Value)
	Param4 = "R"
	Param5 = "R"
	Param6 = ""			'검사진행상태 
	
	iCalledAspName = AskPRAspName("Q4111pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "Q4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtInspReqNo.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspReqNo.Value    = arrRet(0)		
		frm1.txtInspReqNo.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspReqNo = true
End Function

'------------------------------------------  OpenInspReqNo2()  -------------------------------------------------
'	Name : OpenInspReqNo2()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspReqNo2()       
	OpenInspReqNo2 = false      
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtInspReqNo2.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo2.Value)	
	Param4 = "R"
	Param5 = "R"
	'Param6 = "D"			'검사진행상태 
	Param6 = ""			'검사진행상태 
		
	iCalledAspName = AskPRAspName("Q4111pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "Q4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtInspReqNo2.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspReqNo2.value = arrRet(0)
    	frm1.txtItemCd.value = arrRet(4)
    	frm1.txtItemNm.value = arrRet(5)
    	frm1.txtBpCd.value = arrRet(7)
    	frm1.txtBpNm.value = arrRet(8)
    	frm1.txtLotNo.value = arrRet(26)
    	frm1.txtLotSubNo.value = arrRet(27)
    	frm1.txtLotSize.Text = arrRet(28)
    	frm1.txtInspDt.Text = arrRet(25)
    	frm1.txtDecision.value = arrRet(24)	
		frm1.txtInspReqNo2.Focus
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspReqNo2 = true
End Function

'=============================================  2.5.1 LoadInspection()  ======================================
'=	Event Name : LoadInspection
'=	Event Desc :
'========================================================================================================
Function LoadInspection()
	Dim intRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 LoadRelease()  ======================================
'=	Event Name : LoadRelease
'=	Event Desc :
'========================================================================================================
Function LoadRelease()
	Dim intRetCD
	
        If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'==========================================  3.1.1 Form_load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029																	'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")												'⊙: Lock  Suitable  Field
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("11101000000011")
	Call SetDefaultVal
	Call InitVariables																		'⊙: Initializes local global variables
    
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtFrameDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFrameDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrameDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFrameDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFrameDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInspReqNo2_OnChange()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtInspReqNo2_OnChange()
	
	Call CommonQueryRs(" A.ITEM_CD, B.ITEM_NM, BP_CD, (SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD = A.BP_CD), A.LOT_NO, A.LOT_SUB_NO, A.LOT_SIZE"," Q_INSPECTION_RESULT AS A, B_ITEM AS B ", " A.ITEM_CD *= B.ITEM_CD AND A.PLANT_CD =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.INSP_REQ_NO =  " & FilterVar(frm1.txtInspReqNo2.Value, "''", "S") & " AND A.INSP_RESULT_NO = 1 AND A.INSP_CLASS_CD = " & FilterVar("R", "''", "S") & "  AND A.DECISION = " & FilterVar("R", "''", "S") & "  ORDER BY A.INSP_REQ_NO, A.INSP_RESULT_NO",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	lgF0 = Trim(Replace(lgF0,Chr(11),vbTab))
	
	If lgF0 = "" Then
		Call DisplayMsgBox("223709","X","X","X") 		'공장정보가 필요합니다 
		Exit Sub
	End If
	
	lgF1 = Trim(Replace(lgF1,Chr(11),""))
	lgF2 = Trim(Replace(lgF2,Chr(11),""))
	lgF3 = Trim(Replace(lgF3,Chr(11),""))
	lgF4 = Trim(Replace(lgF4,Chr(11),""))
	lgF5 = Trim(Replace(lgF5,Chr(11),""))
	lgF6 = Trim(Replace(lgF6,Chr(11),""))
	
	With frm1
		.txtItemCd.Value = lgF0
		.txtItemNm.Value = lgF1
		.txtBpCd.Value = lgF2
		.txtBpNm.Value = lgF3		
		.txtLotNo.Value = lgF4
		.txtLotSubNo.Value = lgF5
		.txtLotSize.Value = Trim(lgF6)		'Input Pro의 Value 속성 이용함.
	End With

	Call CommonQueryRs(" CONVERT(CHAR(10), A.INSP_DT, 20), (SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = " & FilterVar("Q0010", "''", "S") & " AND MINOR_CD = A.DECISION) AS DECISION_NM"," Q_INSPECTION_RESULT AS A, B_ITEM AS B ", " A.ITEM_CD *= B.ITEM_CD AND A.PLANT_CD =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.INSP_REQ_NO =  " & FilterVar(frm1.txtInspReqNo2.Value, "''", "S") & " AND A.INSP_RESULT_NO = 1 AND A.INSP_CLASS_CD = " & FilterVar("R", "''", "S") & "  AND A.DECISION = " & FilterVar("R", "''", "S") & "  ORDER BY A.INSP_REQ_NO, A.INSP_RESULT_NO",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	lgF0 = UNIDateClientFormat(Trim(Replace(lgF0,Chr(11),"")))
	lgF1 = Trim(Replace(lgF1,Chr(11),""))
	
	With frm1
		.txtInspDt.Text = lgF0
		.txtDecision.Value = lgF1
		.txtFrameDt.text = lgF0
	End With		
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	FncQuery = False                                                        '⊙: Processing is NG
	Err.Clear                                                               '☜: Protect system from crashing
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	'-----------------------
	'Erase contents area
	'----------------------- 
	Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
	Call InitVariables									'⊙: Initializes local global variables
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then						'⊙: This function check indispensable field
		Exit Function
	End If
	Call ggoOper.LockField(Document, "N")								'⊙: This function lock the suitable field
	'-----------------------
	'Query function call area
	'----------------------- 
	If DbQuery = False then
		Exit Function
	End If									'☜: Query db data
	FncQuery = True									'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
	FncNew = False                                                          					'⊙: Processing is NG

	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	 
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                       				'⊙: Lock  Suitable  Field
	Call InitVariables															'⊙: Initializes local global variables
	Call SetDefaultVal
	Call SetToolBar("11101000000011")
	
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If   
	lgBlnFlgChgValue = False
	FncNew = True
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD
	FncDelete = False									'⊙: Processing is NG
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then
		Exit Function
	End If									'☜: Delete db data
	FncDelete = True                                                        					'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	FncSave = False                                                         					'⊙: Processing is NG
	Err.Clear						                                                        '☜: Protect system from crashing
	
	 '-----------------------
	'Precheck area

	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
	
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                            				 '⊙: Check contents area
		Exit Function
	End If

	'-----------------------
	'Save function call area
	'-----------------------	
	If DbSave = False then	
		Exit Function
	End If			                                		                '☜: Save db data
	FncSave = True                                                        					  '⊙: Processing is OK
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
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
		
	lgIntFlgMode = Parent.OPMD_CMODE							'⊙: Indicates that current mode is Crate mode
	lgBlnFlgChgValue = True
	
	'Primary Key Field Clear
	With frm1
		.txtInspReqNo.Value = ""
		.txtInspReqNo2.Value = ""
		.txtItemCd.Value = ""
		.txtItemNm.Value = ""
		.txtBpCd.Value = ""
		.txtBpNm.Value = ""
		.txtLotNo.Value = ""
		.txtLotSubNo.Value = ""
		.txtLotSize.Text = ""
		.txtInspDt.Text = ""
		.txtDecision.Value = ""
	End With

	Call ggoOper.LockField(Document, "N")							'⊙: This function lock the suitable field
	
	FncCopy = true
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
	FncInsertRow = false
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = false
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
	FncPrev = false
	Dim strVal
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	ElseIf lgPrevNo = "" Then
	 	Call DisplayMsgBox("900011", "X", "X", "X")  '☜ 바뀐부분 
	 	Exit Function
	End If
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
	strVal = strVal & "&txtInspReqNo=" & lgPrevNo						'☆: 조회 조건 데이타 
	
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
		Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	End If
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'☜: 비지니스 처리 ASP의 상태값 
	strVal = strVal & "&txtInspReqNo=" & lgNextNo						'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)
	FncNext = true
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
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
    FncFind = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If

	End If
	
	FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	DbDelete = False																'⊙: Processing is NG
    Call LayerShowHide(1)	
	Dim strVal
	
	strVal = BIZ_PGM_DEL_ID & "?txtMode=" & Parent.UID_M0003						'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtplantcd=" & Trim(frm1.txtplantcd.value)			'☆: 조회 조건 데이타 
	strVal = strVal & "&txtInspReqNo=" & Trim(frm1.txtInspReqNo2.value)				'☜: 조회 조건 데이타 
	strVal = strVal & "&hPlantCd=" & Trim(frm1.hPlantCd.value)				'☜: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)				
	
	DbDelete = True			                                                   				'⊙: Processing is NG
End Function	

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()					
	DbDeleteOk = false
	lgBlnFlgChgValue = False
	Call MainNew()
	DbDeleteOk = true
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	DbQuery = False
	
	Dim strVal
	
	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtplantcd=" & Trim(frm1.txtplantcd.value)			'☆: 조회 조건 데이타 
	strVal = strVal & "&txtInspReqNo=" & Trim(frm1.txtInspReqNo.value)				'☜: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
		
	DbQuery = True                                                          					'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOkOPEN

' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()																		'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE														'⊙: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	
	Call ggoOper.LockField(Document, "Q")												'⊙: This function lock the suitable field
	
	Call SetToolBar("11111000001111")
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	Dim strVal
	DbSave = False															'⊙: Processing is NG

	Call LayerShowHide(1)
	
	With frm1
		.txtMode.Value = Parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.Value = lgIntFlgMode
		.txtUpdtUserId.Value = Parent.gUsrID
		.txtInsrtUserId.Value = Parent.gUsrID
		
		If lgIntFlgMode = Parent.OPMD_CMODE then
			.hPlantCd.value = .txtPlantCd.value  
		End If
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
	End With
	
	DbSave = True
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()
	DbSaveOk = false
	frm1.txtInspReqNo.value = frm1.txtInspReqNo2.value 
	Call InitVariables
	Call MainQuery()
	DbSaveOk = true
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>수입검사 불합격통지</FONT></TD>
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
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20  MAXLENGTH=18 ALT="검사의뢰번호" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspReqNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspReqNo()"></TD>							
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD VALIGN="top"  WIDTH="100%">
						<FIELDSET CLASS="CLSFLD"><LEGEND>검사결과내용</LEGEND>
							<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo2" SIZE=20 MAXLENGTH=18 ALT="검사의뢰번호" tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspReqNo2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspReqNo2()"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
	                				<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=10 ALT="작성자" tag="24">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="24" ></TD>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=4 ALT="공급처" tag="24">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24"></TD>
								</TR>
								<TR>
           							<TD CLASS="TD5" NOWRAP>로트번호</TD>
							    	<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE=15 MAXLENGTH=12 ALT="LOT NO" tag="24">
										<INPUT TYPE=TEXT NAME="txtLotSubNo" SIZE=10 MAXLENGTH=5 tag="24" STYLE="Text-Align: Right"></TD>
	                				<TD CLASS="TD5" NOWRAP>로트크기</TD>            
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLotSize name=txtLotSize CLASS=FPDS140 title=FPDOUBLESINGLE ALT="LOT SIZE" tag="24X3"> <PARAM Name="AllowNull" Value="-1"> <PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT>
									</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TD5" NOWRAP>검사일</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ID="txtInspDt" NAME="txtInspDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="검사일" TAG="24x1"></OBJECT>');</SCRIPT>
									</TD>
				                	<TD CLASS="TD5" NOWRAP>판정</TD>
				                	<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDecision" SIZE=20  ALT="판정" tag="24"></TD>
				                </TR>
							</TABLE>
						</FIELDSET>	
						<FIELDSET CLASS="CLSFLD"><LEGEND>불합격내용</LEGEND>
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS="TD5" NOWRAP>작성자</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFramer" SIZE=10 MAXLENGTH=10 ALT="작성자" tag="22"></TD>
									<TD CLASS="TD5" NOWRAP>작성일</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtFrameDt name=txtFrameDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="작성일" tag="22X1"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>불량현상</TD>
									<TD CLASS="TD6" NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtDefectComment" style="width:650px;" MAXLENGTH=200 TAG="22" ALT="불량현상"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>불량내용</TD>
									<TD CLASS="TD6" NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtDefectContents" style="width:650px;" MAXLENGTH=200 TAG="21" ALT="불량내용"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>개선요망사항</TD>
									<TD CLASS="TD6" NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtRequiredImprovement" style="width:650px;" MAXLENGTH=200 TAG="21" ALT="개선요망사항"></TD>
								</TR>
							</TABLE>	
						</FIELDSET>
					</TD>	
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
      		<TD WIDTH=100%>
      			<TABLE <%=LR_SPACE_TYPE_30%>>
        				<TR>        				
        					<TD WIDTH=10>&nbsp;</TD>
        					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspection">수입검사</A>&nbsp;|&nbsp;<A href="vbscript:LoadRelease">Release</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
       				</TR>
      			</TABLE>
      		</TD>
    	</TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hInspReqNo" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>


















