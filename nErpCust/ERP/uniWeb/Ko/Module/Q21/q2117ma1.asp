<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2117MA1
'*  4. Program Name         : Release
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

Const BIZ_PGM_QRY_ID = "Q2117MB1.asp"
Const BIZ_PGM_SAVE_ID = "Q2117MB2.asp"
Const BIZ_PGM_DEL_ID = "Q2117MB3.asp"											 '☆: 비지니스 로직 ASP명 

Const BIZ_PGM_JUMP1_ID = "Q2111MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgNextNo					'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

Dim lgMvmtMethod
Dim lgPRYNBeforeGR
Dim lgSTYNAfterGR
Dim strInspClass

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                       	              '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                	              	'⊙: Indicates that no value changed
    lgIntGrpCount = 0
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False							'☆: 사용자 변수 초기화 
    
    '###검사분류별 변경부분 Start###
    strInspClass = "R"
	'###검사분류별 변경부분 End###
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

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'======================================================================================================
'	Name : OpenPlant()
'	Description :Plant PopUp
'======================================================================================================
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
	'###검사분류별 변경부분 Start###	
	Param4 = strInspClass 		'검사분류 
	'###검사분류별 변경부분 End###
	Param5 = ""			'판정 
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

 '------------------------------------------  OpenSlForGood()  -------------------------------------------------
'	Name : OpenSlForGood()
'	Description :SlForGood PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSlForGood()
	Dim strCode
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If UCase(frm1.txtSlCdForGood.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function	
	End If
	
	IsOpenPop = True
	strCode = Trim(frm1.txtSlCdForGood.Value)
	arrParam(0) = "양품저장창고팝업"	
	arrParam(1) = "B_Storage_Location"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "Plant_Cd =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " And SL_TYPE >= " & FilterVar("I", "''", "S") & " "    ' Where Condition
	arrParam(5) = "양품저장창고"			
	
    arrField(0) = "SL_CD"	
    arrField(1) = "SL_NM"	
    
    arrHeader(0) = "창고코드"		
    arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtSlCdForGood.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtSlCdForGood.value = arrRet(0)   
		frm1.txtSlNmForGood.value = arrRet(1)  
		frm1.txtSlCdForGood.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenSlForGood = true
End Function

 '------------------------------------------  OpenSlForDefective()  -------------------------------------------------
'	Name : OpenSlForDefective()
'	Description :SlForDefective PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSlForDefective()
	Dim strCode
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If UCase(frm1.txtSlCdForDefective.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function	
	End If
	
	IsOpenPop = True
	strCode = Trim(frm1.txtSlCdForDefective.Value)
	arrParam(0) = "불량품저장창고팝업"	
	arrParam(1) = "B_Storage_Location"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "Plant_Cd =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " And SL_TYPE >= " & FilterVar("I", "''", "S") & " "    ' Where Condition
	arrParam(5) = "불량품저장창고"			
	
    arrField(0) = "SL_CD"	
    arrField(1) = "SL_NM"	
    
    arrHeader(0) = "창고코드"		
    arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtSlCdForDefective.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtSlCdForDefective.value = arrRet(0)   
		frm1.txtSlNmForDefective.value = arrRet(1)  
		frm1.txtSlCdForDefective.Focus 		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenSlForDefective = true
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

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function CheckSlCd()
	CheckSlCd = False
	With frm1
		If lgMvmtMethod = "A" Then 
			If lgSTYNAfterGR = "Y" Then
				If .txtGoodQty.Text > 0 Then
					If Trim(.txtSlCdForGood.Value) = "" Then
						'****** Msgbox : 양품을 저장할 창고를 선택하십시오.
						Call DisplayMsgBox("221326", "X", "X", "X")
						Exit Function
					End If
				End If
		
				If .txtDefectQty.Text > 0 Then
					If Trim(.txtSlCdForDefective.Value) = "" Then
						'****** Msgbox : 불량품을 저장할 창고를 선택하십시오.
						Call DisplayMsgBox("221327", "X", "X", "X")
						Exit Function
					End If
				End If
			End If
		ElseIf lgMvmtMethod = "B" Then 
			If lgPRYNBeforeGR = "Y" Then
				If .txtGoodQty.Text > 0 Then
					If Trim(.txtSlCdForGood.Value) = "" Then
						'****** Msgbox : 양품을 저장할 창고를 선택하십시오.
						Call DisplayMsgBox("221326", "X", "X", "X")
						Exit Function
					End If
				End If
			End If
		End If
	End With
	CheckSlCd = True
End Function

'==========================================================================================
'   Event Name : ProtectDetails
'   Event Desc :
'==========================================================================================
Sub ProtectDetails()
	With ggoOper
		Call .SetReqAttr(frm1.txtReleaseDt, "Q")
		Call .SetReqAttr(frm1.txtSlCdForGood, "Q")
		Call .SetReqAttr(frm1.txtSlCdForDefective, "Q")
	End With
End Sub

'==========================================================================================
'   Event Name : ReleaseDetails
'   Event Desc :
'==========================================================================================
Sub ReleaseDetails()
	With ggoOper
		Call .SetReqAttr(frm1.txtReleaseDt, "N")
		
		If lgMvmtMethod = "A" Then '입고 후 검사 
			If lgSTYNAfterGR = "Y" Then
				If frm1.txtGoodQty.Text > 0 Then
					Call .SetReqAttr(frm1.txtSlCdForGood, "N")
				Else
					Call .SetReqAttr(frm1.txtSlCdForGood, "Q")
				End If
		
				If frm1.txtDefectQty.Text > 0 Then
					Call .SetReqAttr(frm1.txtSlCdForDefective, "N")
				Else
					Call .SetReqAttr(frm1.txtSlCdForDefective, "Q")
				End If
			Else
				Call .SetReqAttr(frm1.txtSlCdForGood, "Q")
				Call .SetReqAttr(frm1.txtSlCdForDefective, "Q")
			End If
		ElseIf lgMvmtMethod = "B" Then  '입고 전 검사 
			If lgPRYNBeforeGR = "Y" Then
				If frm1.txtGoodQty.Text > 0 Then
					Call .SetReqAttr(frm1.txtSlCdForGood, "N")
				Else
					Call .SetReqAttr(frm1.txtSlCdForGood, "Q")
				End If
		
				Call .SetReqAttr(frm1.txtSlCdForDefective, "Q")
			Else
				Call .SetReqAttr(frm1.txtSlCdForGood, "Q")
				Call .SetReqAttr(frm1.txtSlCdForDefective, "Q")
			End If
		Else		
			Call .SetReqAttr(frm1.txtSlCdForGood, "Q")
			Call .SetReqAttr(frm1.txtSlCdForDefective, "Q")
		End If
	End With
End Sub

'==========================================  3.1.1 Form_load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029																	'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtInspReqDt, parent.gDateFormat, 1)		
	Call ggoOper.FormatDate(frm1.txtReleaseDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtInspDt, parent.gDateFormat, 1)		
	Call ggoOper.LockField(Document, "N")												'⊙: Lock  Suitable  Field
	'----------  Coding part  -------------------------------------------------------------
	Call ProtectDetails()	
	Call SetToolBar("11100000000011")
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
'   Event Name : txtReleaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReleaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReleaseDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReleaseDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReleaseDt_Change()
    lgBlnFlgChgValue = True
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
	Call ggoOper.LockField(Document, "2")						'⊙: Clear Contents  Field
	Call ProtectDetails()
	Call InitVariables									'⊙: Initializes local global variables
	
    '-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then						'⊙: This function check indispensable field
		Exit Function
	End If
	
	'-----------------------
	'Query function call area
	'----------------------- 
	If DbQuery = False then
		Exit Function
	End If										'☜: Query db data

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
	Call ProtectDetails()
	Call InitVariables															'⊙: Initializes local global variables
	Call SetToolBar("11100000000011")
	Call SetDefaultVal
	
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If
	
	FncNew = True
	
	
	
End Function

'========================================================================================
' Function Name : FncDelete()
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
	
	FncDelete = True
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim blnReturn
	
	FncSave = False                                                         					'⊙: Processing is NG
	
	Err.Clear						                                                        '☜: Protect system from crashing
	
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                            				 '⊙: Check contents area
		Exit Function
	End If
	
	blnReturn = CheckSlCd
	If blnReturn = False Then
		Exit Function
	End If
	
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then	
		Exit Function
	End If				                                		                '☜: Save db data
	
	FncSave = True
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = True
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next                                                   					'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = True
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
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
 Function FncNext() 
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
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 	Call parent.FncExport(Parent.C_SINGLE)		
End Function

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
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	Call parent.FncFind(Parent.C_SINGLE, False)     
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	Err.Clear                                                               						'☜: Protect system from crashing
	DbDelete = False																'⊙: Processing is NG
    Call LayerShowHide(1)	
	Dim strVal
	
	strVal = BIZ_PGM_DEL_ID & "?txtMode=" & Parent.UID_M0003						'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtplantcd=" & Trim(frm1.txtplantcd.value)			'☆: 조회 조건 데이타 
	strVal = strVal & "&txtInspReqNo=" & Trim(frm1.txtInspReqNo.value)				'☜: 조회 조건 데이타 
		
	Call RunMyBizASP(MyBizASP, strVal)				
	
	DbDelete = True
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()									'☆: 삭제 성공후 실행 로직 
    Call CancelRestoreToolBar()
	Call InitVariables
	Call MainQuery()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Err.Clear                                                               							'☜: Protect system from crashing
	Call LayerShowHide(1)
	DbQuery = False
	
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtInspReqNo=" & Trim(frm1.txtInspReqNo.value)				'☜: 조회 조건 데이타 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
		
	DbQuery = True                                                          					'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOkOPEN
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()			
		lgIntFlgMode = Parent.OPMD_UMODE           
    	lgBlnFlgChgValue = False															'☆: 조회 성공후 실행로직 
    	
	'검사진행상태가 "R"인 아닌 경우에만 수정이 가능하도록 처리   	
   	If frm1.hStatusFlag.value = "R" Then
   		Call ggoOper.LockField(Document, "Q")              '⊙: Lock  Suitable Field
   		Call SetToolBar("11110000000111")
   	Else
   		Call ggoOper.LockField(Document, "N")		
   		Call ReleaseDetails()
   		Call SetToolBar("11101000000111")		'⊙: 버튼 툴바 제어 
   	End If
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	Dim strVal
	Err.Clear																	'☜: Protect system from crashing
	
	DbSave = False															'⊙: Processing is NG
	
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value = Parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.Value = Parent.gUsrID
		.txtUpdtUserId.Value = Parent.gUsrID
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
	End With
	
	DbSave = True
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>수입검사 Release</FONT></TD>
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
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14" ></TD>
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
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
                				<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 ALT="품목" tag="24">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="24" ></TD>
								<TD CLASS="TD5" NOWRAP>Release일</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2117ma1_txtReleaseDt_txtReleaseDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 ALT="공급처" tag="24">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>검사의뢰일</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2117ma1_txtInspReqDt_txtInspReqDt.js'></script>
								</TD>
							</TR>							
							<TR>
                				<TD CLASS="TD5" NOWRAP>로트번호</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE=15 ALT="로트번호" tag="24">
									<INPUT TYPE=TEXT NAME="txtLotSubNo" SIZE=10 tag="24" STYLE="Text-Align: Right"></TD>
                				<TD CLASS="TD5" NOWRAP>로트크기</TD>            
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2117ma1_fpDoubleSingle1_txtLotSize.js'></script>
								</TD>
           					</TR>
   							<TR>
								<TD CLASS="TD5" NOWRAP>검사원</TD>
     								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspectorCd" SIZE=10 ALT="검사원" tag="24">
									<INPUT TYPE=TEXT NAME="txtInspectorNm" SIZE=20 tag="24" ></TD>								
								<TD CLASS="TD5" NOWRAP>검사일</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2117ma1_txtInspDt_txtInspDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>양품수</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2117ma1_fpDoubleSingle2_txtGoodQty.js'></script>
								</TD>
								<TD CLASS="TD5" NOWRAP>불량품수</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2117ma1_fpDoubleSingle3_txtDefectQty.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>판정</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDecision" SIZE=20 tag="24" "Text-Align: Center"></TD>
								<TD CLASS="TD5" NOWRAP>검사진행현황</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtStatusFlag" SIZE=20 tag="24" "Text-Align: Center"></TD>
							</TR>
                			<TR>
								<TD CLASS="TD5" NOWRAP>양품저장창고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSlCdForGood" SIZE=10 MAXLENGTH=7 ALT="양품저장창고" tag="25XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSlForGood()">
									<INPUT TYPE=TEXT NAME="txtSlNmForGood" SIZE=20 tag="24" ></TD>
								<TD CLASS="TD5" NOWRAP>불량품저장창고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSlCdForDefective" SIZE=10 MAXLENGTH=7 ALT="불량품저장창고" tag="25XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSlForDefective()">
									<INPUT TYPE=TEXT NAME="txtSlNmForDefective" SIZE=20 tag="24" ></TD>
							</TR>
							<% Call SubFillREmBodyTD5656(15)%>
						</TABLE>
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
	        				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspection">수입검사</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
       				</TR>
      			</TABLE>
      		</TD>
    	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  tabindex=-1 WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspReqNo" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hPlantCd" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hStatusFlag" TAG="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


