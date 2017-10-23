<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2611MA1
'*  4. Program Name         : 이상발생 보고서 정보등록 
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "Q2611MB1.asp"
Const BIZ_PGM_SAVE_ID = "Q2611MB2.asp"
Const BIZ_PGM_DEL_ID = "Q2611MB3.asp"

Const BIZ_PGM_JUMP_ID = "Q2612MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop      

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYMD	'#####
CompanyYMD = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gDateFormat)                                           '☆: 초기화면에 뿌려지는 시작 날짜 -----
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 

'==========================================  2.1.1 InitVariables()======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                       	              							'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                	              								'⊙: Indicates that no value changed
    lgIntGrpCount = 0
    IsOpenPop = False
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboInspClassCd.value= "R"
	frm1.cboCounterPlanFlag.value= ""
	frm1.txtFrameDt.text = CompanyYMD
	frm1.txtOccurDtFr.text = CompanyYMD
	frm1.txtOccurDtTo.text = CompanyYMD
End Sub

'==========================================  2.2.6 InitComboBox()=======================================
'	Name : InitComboBox
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspClassCd , lgF0, lgF1, Chr(11))
	
	Call SetCombo(frm1.cboCounterPlanFlag,"N","아니오")
	Call SetCombo(frm1.cboCounterPlanFlag,"Y","예")
End Sub

'------------------------------------------  OpenPlant() -------------------------------------------------
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

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam,arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)
		frm1.txtPlantNm.Value    = arrRet(1)
		lgBlnFlgChgValue = True
	End If	
	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true	
End Function

'------------------------------------------  OpenMgmtNo()  -------------------------------------------------
'	Name : OpenMgmtNo()
'	Description : OpenMgmtNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMgmtNo()        
	Dim arrRet
	Dim Param1, Param2, Param3, Param4
	Dim iCalledAspName, IntRetCD
	
	If UCase(frm1.txtMgmtNo1.ClassName) = UCase(Parent.UCN_PROTECTED) Then
		Exit Function
	End If
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True
	
	Param1 = ""		
	Param2 = ""
	Param3 = Trim(frm1.txtMgmtNo1.Value)
	Param4 = ""
	
	iCalledAspName = AskPRAspName("q2611pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q2611pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtMgmtNo1.Value    = arrRet(0)
	End If	
	
	frm1.txtMgmtNo1.Focus
	Set gActiveElement = document.activeElement
	OpenMgmtNo = true
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

	If UCase(frm1.txtItemCd.ClassName) = UCase(Parent.UCN_PROTECTED) Then
		Exit Function
	End If
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
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
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value    = arrRet(0)
		frm1.txtItemNm.Value    = arrRet(1)
		lgBlnFlgChgValue = True
	End If	
	
	frm1.txtItemCd.Focus
	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

'------------------------------------------  OpenWc()  -------------------------------------------------
'	Name : OpenWc()
'	Description : Wc PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenWc()
	Dim arrRet

	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	If UCase(frm1.txtWcCd.ClassName) = UCase(Parent.UCN_PROTECTED) Then
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "작업장"											' 팝업명칭	
	arrParam(1) = "P_WORK_CENTER"										' TABLE 명칭	
	arrParam(2) = Trim(frm1.txtWcCd.Value)										' Code Condition	
	arrParam(3) = ""												' Name Condition	
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "" 	' Where Condition
	arrParam(5) = "작업장"											' 조건필드의 라벨 명칭	

	arrField(0) = "WC_CD"	
	arrField(1) = "WC_NM"	
	
	arrHeader(0) = "작업장코드"		
	arrHeader(1) = "작업장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtWcCd.Value    = arrRet(0)
		frm1.txtWcNm.Value    = arrRet(1)		
		lgBlnFlgChgValue = True
	End If	
	
	frm1.txtWcCd.Focus
	Set gActiveElement = document.activeElement
End Function

'=============================================  2.5.1 JumpOccurResult()  ======================================
'=	Event Name : JumpOccurResult
'=	Event Desc :
'========================================================================================================
Function LoadOccurResult()
	Dim intRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then Exit Function
	End If
	
	WriteCookie "txtMgmtNo", UCase(Trim(frm1.txtMgmtNo2.value))
	WriteCookie "cboInspClassCd", UCase(Trim(frm1.cboInspClassCd.value))
	WriteCookie "txtItemCd", UCase(Trim(frm1.txtItemCd.value))
	WriteCookie "txtItemNm", UCase(Trim(frm1.txtItemNm.value))
	WriteCookie "txtPlantCd", UCase(Trim(frm1.txtPlantCd.value))
	WriteCookie "txtPlantNm", UCase(Trim(frm1.txtPlantNm.value))
	WriteCookie "txtFramer", UCase(Trim(frm1.txtFramer.value))
	WriteCookie "txtWcCd", UCase(Trim(frm1.txtWcCd.value))
	WriteCookie "txtWcNm", UCase(Trim(frm1.txtWcNm.value))
	WriteCookie "txtFrameDt", UCase(Trim(frm1.txtFrameDt.text))
	WriteCookie "txtOccurDtFr", UCase(Trim(frm1.txtOccurDtFr.text))
	WriteCookie "txtOccurDtTo", UCase(Trim(frm1.txtOccurDtTo.text))
		
	PgmJump(BIZ_PGM_JUMP_ID)
End Function

'==========================================  3.1.1 Form_load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029											'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field
	Call InitVariables
	Call InitComboBox
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolbar("11101000000011")
	Call SetDefaultVal
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
	
	Call Sub_cboInspClassCd_onchange
	
	frm1.txtMgmtNo1.focus 
	
	lgblnFlgChgValue = False
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtOccurDtFr_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtOccurDtFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtOccurDtFr.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtOccurDtFr_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtOccurDtFr_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtOccurDtTo_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtOccurDtTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtOccurDtTo.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInvClsDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtOccurDtTo_Change()
    lgBlnFlgChgValue = True
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
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtFrameDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : cboInspClassCd_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboInspClassCd_onchange()
	lgBlnFlgChgValue = True	
	Call Sub_cboInspClassCd_onchange
End Sub

'=======================================================================================================
'   Event Name : cboInspClassCd_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub Sub_cboInspClassCd_onchange()	
	If frm1.cboInspClassCd.Value = "P" Then
		Call ggoOper.SetReqAttr(frm1.txtWcCd, "N")
	Else 
		frm1.txtWcCd.value = ""
		frm1.txtWcNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtWcCd, "Q")
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    Dim IntRetCD 
	
	FncQuery = False                                                       								'⊙: Processing is NG
	
	Err.Clear                                                            		   						'☜: Protect system from crashing
	
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then
		Exit Function
	End If
		
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
	
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
	Call InitVariables
	Call ggoOper.LockField(Document, "N")								'⊙: This function lock the suitable field
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False then	Exit Function				'☜: Query db data
	
	FncQuery = True
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()
	Dim IntRetCD 
	
	FncNew = False             				'⊙: Processing is NG
	
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
	
	Call ggoOper.ClearField(Document, "A")
	
	With frm1
		ReleaseTag(.txtPlantCd)
		ReleaseTag(.txtItemCd)
		ReleaseTag(.txtWcCd)
		ReleaseTag(.cboInspClassCd)
		ReleaseTag(.txtFramer)
		ReleaseTag(.txtReasonForOccur)
		Call ggoOper.SetReqAttr(.txtOccurDtFr, "N")
		Call ggoOper.SetReqAttr(.txtOccurDtTo, "N")
		Call ggoOper.SetReqAttr(.txtFrameDt, "N")
		ReleaseTag(.txtContentsofAssignableOccur)
	End With
	
	Call ggoOper.LockField(Document, "N")                                       							'⊙: Lock  Suitable  Field
	Call InitVariables												'⊙: Initializes local global variables
	Call SetToolbar("11101000000011")
	Call SetDefaultVal
	Call Sub_cboInspClassCd_onchange 
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
	
	lgBlnFlgChgValue = False
	frm1.txtMgmtNo1.focus 
	
	FncNew = True
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()
	Dim IntRetCD 
	
	FncDelete = False												'⊙: Processing is NG
	
	  '-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then	Exit Function
	
	If DbDelete = False Then Exit Function				'☜: Delete db data
	
	FncDelete = True
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	
	Dim IntRetCD 
	
	FncSave = False                                                         								'⊙: Processing is NG
	
	Err.Clear						                                       					'☜: Protect system from crashing
	
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
	
	If Trim(frm1.txtWcCd.Value) = "" Then
		If frm1.cboInspClassCd.Value = "P" Then
			Call DisplayMsgBox("221805", "X", "X", "X") '작업장이 필요합니다. 
			Exit Function
		End If 
	Else
		If frm1.cboInspClassCd.Value <> "P" Then
			Call DisplayMsgBox("221806", "X", "X", "X") '작업장은 필요하지 않습니다. 
			Exit Function
		End If 
	End If
	
	If ValidDateCheck(frm1.txtOccurDtFr, frm1.txtOccurDtTo) = False Then
		Exit Function
	Else
		If ValidDateCheck(frm1.txtOccurDtTo, frm1.txtFrameDt) = False Then Exit Function
	End If
		
	If Len(Trim(frm1.txtContentsofAssignableOccur.Value)) > 200 Then
		Call MsgBox("이상발생내용은 200자를 초과할 수 없습니다", vbInformation)
		frm1.txtContentsofAssignableOccur.Focus
		Exit Function
	End If
	
	If Len(Trim(frm1.txtReasonForOccur.Value)) > 200 Then
		Call MsgBox("발생사유는 200자를 초과할 수 없습니다", vbInformation)
		frm1.txtReasonForOccur.Focus
		Exit Function
	End If
		   
	'-----------------------
	'Save function call area
	'-----------------------	
	If DbSave = False then Exit Function
	
	FncSave = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow()
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = True
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = True
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
	
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
							& "&txtMgmtNo1=" & lgPrevNo									'☆: 조회 조건 데이타 
	
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
	
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
							& "&txtMgmtNo1=" & lgNextNo									'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If

	lgIntFlgMode = Parent.OPMD_CMODE										'⊙: Indicates that current mode is Crate mode
	
	' 조건부 필드를 삭제한다. 
	Call ggoOper.ClearField(Document, "1")                         	'⊙: Clear Condition Field
	With frm1
		.txtMgmtNo2.value = ""
		ReleaseTag(.txtPlantCd)
		ReleaseTag(.txtItemCd)
		ReleaseTag(.txtWcCd)
		ReleaseTag(.cboInspClassCd)
		ReleaseTag(.txtFramer)
		ReleaseTag(.txtReasonForOccur)
		Call ggoOper.SetReqAttr(.txtOccurDtFr, "N")
		Call ggoOper.SetReqAttr(.txtOccurDtTo, "N")
		Call ggoOper.SetReqAttr(.txtFrameDt, "N")
		ReleaseTag(.txtContentsofAssignableOccur)
	End With
	Call ggoOper.LockField(Document, "N")                         	'⊙: Lock  Suitable  Field
	Call InitVariables												'⊙: Initializes local global variables
	Call SetToolbar("11101000000011")
	
	lgBlnFlgChgValue = True
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLE)		
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                              
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
		If IntRetCD = vbNo Then Exit Function
	End If
	
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery()
	Dim strVal
	
	Err.Clear     
	
	Call LayerShowHide(1)                                                          							'☜: Protect system from crashing
	
	DbQuery = False                                                        								'⊙: Processing is NG
		
	strVal = BIZ_PGM_QRY_ID & "?txtMgmtNo=" & Trim(frm1.txtMgmtNo1.Value)						'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)
	
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												'☆: 조회 성공후 실행로직 
	lgIntFlgMode = Parent.OPMD_UMODE   
	lgBlnFlgChgValue = False  
	Call ggoOper.LockField(Document, "Q")	
	
	If frm1.cboInspClassCd.Value = "P" Then
		Call ggoOper.SetReqAttr(frm1.txtWcCd, "N")
	Else 
		Call ggoOper.SetReqAttr(frm1.txtWcCd, "Q")
	End If
	
	With frm1
		If .cboCounterPlanFlag.Value = "Y" Then
			ProtectTag(.txtPlantCd)
			ProtectTag(.txtItemCd)
						
			ProtectTag(.txtWcCd)
			ProtectTag(.cboInspClassCd)
			ProtectTag(.txtFramer)
			ProtectTag(.txtReasonForOccur)
			Call ggoOper.SetReqAttr(.txtOccurDtFr, "Q")
			Call ggoOper.SetReqAttr(.txtOccurDtTo, "Q")
			Call ggoOper.SetReqAttr(.txtFrameDt, "Q")			
			
			ProtectTag(.txtContentsofAssignableOccur)
			
			Call SetToolbar("11100000001111")
		Else
			Call SetToolbar("11111000001111")
		End If
	End With
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
	Dim strVal
	
	Err.Clear                                                               								'☜: Protect system from crashing
	
	DbDelete = False												'⊙: Processing is NG
			
	strVal = BIZ_PGM_DEL_ID & "?txtMgmtNo=" & Trim(frm1.txtMgmtNo2.value)						'☆: 삭제 조건 데이타 
		
	Call RunMyBizASP(MyBizASP, strVal)				
	
	DbDelete = True
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
	lgBlnFlgChgValue = False
	Call MainNew()
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave()
	Err.Clear													'☜: Protect system from crashing
	
	Call LayerShowHide(1)
	
	DbSave = False												'⊙: Processing is NG
	
	frm1.txtMode.value = Parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
	frm1.txtFlgMode.value = lgIntFlgMode
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
	DbSave = True
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()												'☆: 저장 성공후 실행 로직 
	Call InitVariables
	Call MainQuery()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><font color=white>이상발생보고서</font></TD>
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
									<TD CLASS="TD5" NOWRAP>관리번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtMgmtNo1" SIZE=20  MAXLENGTH=18 ALT="관리번호" TAG="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgmtNo1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMgmtNo()"></TD>
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
								<TD CLASS="TD5" NOWRAP>관리번호</TD>
							    	<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMgmtNo2" SIZE=20 MAXLENGTH=18 ALT="관리번호" TAG="25XXXU"></TD>	
							    	<TD CLASS="TD5" NOWRAP>대책여부</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="cboCounterPlanFlag" ALT="대책여부" STYLE="WIDTH: 150px" TAG="24"></SELECT></TD>							
							 </TR>
							 <TR>
							    	<TD CLASS="TD5" NOWRAP>공장</TD>
     								<TD CLASS="TD6" NOWRAP>
     									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
									<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 TAG="24" ></TD>								
								<TD CLASS="TD5" NOWRAP>검사분류</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" TAG="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtItemCd" SIZE="15" MAXLENGTH="18" ALT="품목" TAG="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" ALIGN=top HEIGHT=20 name=btnItemCd TYPE="BUTTON" ONCLICK=vbscript:OpenItem()>
									<INPUT TYPE=TEXT NAME="txtItemNm" SIZE="20" MAXLENGTH="20" TAG="24" ></TD>
								<TD CLASS="TD5" NOWRAP>작업장</TD>
								<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtWcCd" SIZE=10 MAXLENGTH=10 ALT="작업장" TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWcCd" ALIGN=top TYPE="BUTTON" ONCLICK=vbscript:OpenWc()>
									<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 MAXLENGTH=20 TAG="24" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발생기간</TD>
							   	<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2611ma1_fpDateTime1_txtOccurDtFr.js'></script>&nbsp;~&nbsp;
									<script language =javascript src='./js/q2611ma1_fpDateTime2_txtOccurDtTo.js'></script>									
								</TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
                				<TD CLASS="TD5" NOWRAP>작성자</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFramer" SIZE=10 MAXLENGTH=10 ALT="작성자" TAG="22"></TD>
                				<TD CLASS="TD5" NOWRAP>작성일</TD>
								<TD CLASS="TD6" NOWRAP>	
									<script language =javascript src='./js/q2611ma1_fpDateTime3_txtFrameDt.js'></script>
								</TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>이상발생내용</TD>
								<TD CLASS="TD6" NOWRAP Colspan=3><INPUT TYPE=TEXT NAME="txtContentsofAssignableOccur" style="width:650px;" MAXLENGTH=200 TAG="22" ALT="이상발생내용"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발생사유</TD>
								<TD CLASS="TD6" NOWRAP Colspan=3><INPUT TYPE=TEXT NAME="txtReasonForOccur" style="width:650px;" MAXLENGTH=200 TAG="21" ALT="발생사유"></TD>
							</TR>
							<% Call SubFillREmBodyTD5656(14)%>
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
          				<TD WIDTH="*" ALIGN="right"><A href="vbscript:LoadOccurResult">이상대책보고&nbsp;&nbsp;</A></TD>
          				<TD WIDTH=10>&nbsp;</TD>
        				</TR>
      			</TABLE>
      		</TD>
    	</TR>    	
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="InsrtUserID" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="InsrtDt" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="UpdtDt" TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
