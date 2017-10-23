<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1212MA1
'*  4. Program Name         : 기타검사조건 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG010,PD6G020
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_QRY_ID="q1212mb1.asp"
Const BIZ_PGM_SAVE_ID="q1212mb2.asp"
Const BIZ_PGM_JUMP_ID = "q1211ma1"					           '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_InspSeries			'= 1									'☆: Spread Sheet의 Column별 상수 
Dim C_SampleQty				'= 2
Dim C_AccptDecisionQty		'= 3
Dim C_RejtDecisionQty		'= 4
Dim C_AccptDecisionDiscrete '= 5
Dim C_MaxDefectRatio		'= 6

Dim IsOpenPop						' Popup

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

	lgIntFlgMode = Parent.OPMD_CMODE        		     'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False 		                   'Indicates that no value changed
	lgIntGrpCount = 0                           	     'initializes Group View Size
	
	'---- Coding part--------------------------------------------------------------------
	
	lgStrPrevKey = ""                           		     'initializes Previous Key
	lgLngCurRows = 0                            	     'initializes Deleted Rows Count
	lgSortKey    = 1                            '⊙: initializes sort direction
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
	
	frm1.cboInspClassCd.value		= "R"
	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If	
	
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
	End If	
	

	If ReadCookie("txtItemNm") <> "" Then
		frm1.txtItemNm.Value = ReadCookie("txtItemNm")
	End If	
	
	If ReadCookie("txtInspClassCd") <> "" Then
		frm1.cboInspClassCd.Value = ReadCookie("txtInspClassCd")
	End If	
	
	If ReadCookie("txtInspItemCd") <> "" Then
		frm1.txtInspItemCd.Value = ReadCookie("txtInspItemCd")
	End If	
	
	If ReadCookie("txtInspItemNm") <> "" Then
		frm1.txtInspItemNm.Value = ReadCookie("txtInspItemNm")
	End If	
	
	If ReadCookie("txtInspMthdCd") <> "" Then
		frm1.txtInspMthdCd.Value = ReadCookie("txtInspMthdCd")
	End If	
		
	If ReadCookie("txtInspMthdNm") <> "" Then
		frm1.txtInspMthdNm.Value = ReadCookie("txtInspMthdNm")
	End If	
	
	If ReadCookie("txtInspClassCd") = "P" Then
		If ReadCookie("txtRoutNo") <> "" Then
			frm1.txtRoutNo.Value = ReadCookie("txtRoutNo")
		End If
		
		If ReadCookie("txtRoutNoDesc") <> "" Then
			frm1.txtRoutNoDesc.Value = ReadCookie("txtRoutNoDesc")
		End If
		
		If ReadCookie("txtOprNo") <> "" Then
			frm1.txtOprNo.Value = ReadCookie("txtOprNo")
		End If
	End If
		
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
	WriteCookie "txtInspClassCd", ""
	WriteCookie "txtInspItemCd", ""
	WriteCookie "txtInspItemNm", ""
	WriteCookie "txtInspMthdCd", ""
	WriteCookie "txtInspMthdNm", ""
	WriteCookie "txtRoutNo", ""
	WriteCookie "txtRoutNoDesc", ""
	WriteCookie "txtOprNo", ""

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	 Call InitSpreadPosVariables()
	     
     With frm1.vspdData
     
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20040518", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
    		
    	.MaxCols = C_MaxDefectRatio + 1				'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
    		
		Call GetSpreadColumnPos("A")
			
		Call AppendNumberPlace("6", "3","0")
		Call AppendNumberPlace("7", "15","4")
		ggoSpread.SSSetFloat C_InspSeries, "차수", 7, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "P"
		ggoSpread.SSSetFloat C_SampleQty, "시료수", 21, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat C_AccptDecisionQty, "합격판정개수", 22, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat C_RejtDecisionQty, "불합격판정개수", 22, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat C_AccptDecisionDiscrete, "합격판정계수", 22, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_MaxDefectRatio, "최대허용불량률", 22, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		
 		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)

    	Call SetSpreadLock 
    			
		.ReDraw = true
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	  With frm1
		.vspdData.ReDraw = False
		Call ggoSpread.SpreadLock(C_InspSeries, -1, C_InspSeries)
		Call ggoSpread.SSSetRequired(C_SampleQty, -1)
		Call ggoSpread.SpreadLock(frm1.vspdData.MaxCols, -1, frm1.vspdData.MaxCols)
		.vspdData.ReDraw = True
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1.vspdData 
    	.Redraw = False
    	ggoSpread.Source = frm1.vspdData
    	ggoSpread.SSSetRequired C_InspSeries, pvStartRow, pvEndRow
    	ggoSpread.SSSetRequired C_SampleQty, pvStartRow, pvEndRow
    	.Redraw = True
    End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspClassCd  , lgF0, lgF1, Chr(11))
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_InspSeries			= 1									'☆: Spread Sheet의 Column별 상수 
	C_SampleQty				= 2
	C_AccptDecisionQty		= 3
	C_RejtDecisionQty		= 4
	C_AccptDecisionDiscrete = 5
	C_MaxDefectRatio		= 6
End Sub
 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
		C_InspSeries			= iCurColumnPos(1)									'☆: Spread Sheet의 Column별 상수 
		C_SampleQty				= iCurColumnPos(2)
		C_AccptDecisionQty		= iCurColumnPos(3)
		C_RejtDecisionQty		= iCurColumnPos(4)
		C_AccptDecisionDiscrete = iCurColumnPos(5)
		C_MaxDefectRatio		= iCurColumnPos(6)
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"									' 팝업 명칭 
	arrParam(1) = "B_PLANT"									' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.value)							' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "공장"									' 조건필드의 라벨 명칭 
		
	arrField(0) = "PLANT_CD"									' Field명(0)
	arrField(1) = "PLANT_NM"									' Field명(1)
	
	arrHeader(0) = "공장코드"								' Header명(0)
	arrHeader(1) = "공장명"								' Header명(1)
    
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
'	Description : Inspection Item By Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem()
	OpenInspItem = false
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9, Param10, Param11, Param12
	Dim iCalledAspName, IntRetCD
		
	
	If IsOpenPop = True Then Exit Function
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	'검사분류가 있는 지 체크 
	If Trim(frm1.cboInspClassCd.Value) = "" then 
		Call DisplayMsgBox("229915", "X", "X", "X") 		'검사분류정보가 필요합니다 
		frm1.cboInspClassCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	'품목코드가 있는 지 체크 
	If Trim(frm1.txtItemCd.Value) = "" then 
		Call DisplayMsgBox("229916", "X", "X", "X") 		'품목정보가 필요합니다 
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	If Trim(frm1.cboInspClassCd.Value) = "P" then 
		'RoutNo가 있는 지 체크 
		If Trim(frm1.txtRoutNo.Value) = "" then 
			Call DisplayMsgBox("220735", "X", "X", "X") 		'라우팅정보가 필요합니다 
			frm1.txtRoutNo.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
		
		'OprNo가 있는 지 체크 
		If Trim(frm1.txtOprNo.Value) = "" then 
			Call DisplayMsgBox("220736", "X", "X", "X") 		'공정정보가 필요합니다 
			frm1.txtOprNo.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If
	
	IsOpenPop = True
	
	With frm1
		Param1 = Trim(.txtPlantCd.Value)
		Param2 = Trim(.txtPlantNm.Value)
		Param3 = Trim(.txtItemCd.Value)
		Param4 = Trim(.txtItemNm.Value)
		Param5 = Trim(.cboInspClassCd.Value)
		Param6 = Trim(.cboInspClassCd.Options(.cboInspClassCd.SelectedIndex).Text)
		Param7 = Trim(.txtRoutNo.Value)
		Param8 = Trim(.txtRoutNoDesc.value)
		Param9 = Trim(.txtOprNo.Value)
		Param10 = Trim(.txtInspItemCd.value)
		Param11 = ""
		Param12 = "0000"		'조정형과 선별형을 제외한 모든 검사방식 -- 규준형, 체크검사 등 
	End With
	
	iCalledAspName = AskPRAspName("q1211pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9, Param10, Param11, Param12), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	frm1.txtInspItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspItemCd.Value = arrRet(1)
		frm1.txtInspItemNm.Value = arrRet(2)	
		frm1.txtInspMthdCd.Value = arrRet(3)
		frm1.txtInspMthdNm.Value = arrRet(4)
		frm1.txtInspItemCd.Focus
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspItem = true
End Function


'====================  OpenRoutNo  ======================================
' Function Name : OpenRoutNo
' Function Desc : OpenRoutNo Reference Popup
'==========================================================================
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
				  "And P_ROUTING_HEADER.ITEM_CD = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") 	
	arrParam(5) = "라우팅"			
	
    arrField(0) = "ROUT_NO"							' Field명(0)
    arrField(1) = "DESCRIPTION"						' Field명(1)
    arrField(2) = "BOM_NO"							' Field명(1)
    arrField(3) = "MAJOR_FLG"						' Field명(1)
   
    arrHeader(0) = "라우팅"						' Header명(0)
    arrHeader(1) = "라우팅명"					' Header명(1)
    arrHeader(2) = "BOM Type"					' Header명(1)
    arrHeader(3) = "주라우팅"					' Header명(1)        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
	If arrRet(0) <> "" Then
		frm1.txtRoutNo.Value = arrRet(0)		
		frm1.txtRoutNoDesc.Value = arrRet(1)		
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function



'**************************** Function OpenOprNo() ***********************************8
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
	
	If frm1.txtRoutNo.value= "" Then
		Call DisplayMsgBox("971012","X", "라우팅","X")
		frm1.txtRoutNo.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	

	arrParam(0) = "공정팝업"	
	arrParam(1) = "P_ROUTING_DETAIL A inner join P_WORK_CENTER B on A.wc_cd = B.wc_cd and A.plant_cd = B.plant_cd " & _
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = " & FilterVar("P1006", "''", "S") & ""				
	arrParam(2) = UCase(Trim(frm1.txtOprNo.Value))
	arrParam(3) = ""
	arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				  " and	A.item_cd = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") & _
				  " and	A.rout_no = " & FilterVar(UCase(frm1.txtRoutNo.value), "''", "S") & _
				  "	and	A.rout_order in (" & FilterVar("F", "''", "S") & " ," & FilterVar("I", "''", "S") & " ) "	
	arrParam(5) = "공정"			
	
	arrField(0) = "A.OPR_NO"	
	arrField(1) = "A.WC_CD"
	arrField(2) = "B.WC_NM"
	arrField(3) = "C.MINOR_NM"
	arrField(4) = "A.INSIDE_FLG"
	arrField(5) = "A.MILESTONE_FLG"
	arrField(6) = "A.INSP_FLG"
	
	arrHeader(0) = "공정"		
	arrHeader(1) = "작업장"	
	arrHeader(2) = "작업장명"
	arrHeader(3) = "공정작업명"
	arrHeader(4) = "사내구분"
	arrHeader(5) = "Milestone"
	arrHeader(6) = "검사여부"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtOprNo.focus
		Exit Function
	Else
		frm1.txtOprNo.Value = arrRet(0)
		frm1.txtOprNoDesc.Value	= arrRet(3)
	End If	
	
End Function


'=============================================  2.5.2 LoadInspStand()  ======================================
'=	Event Name : LoadInspStand
'=	Event Desc :
'========================================================================================================
Function LoadInspStand()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		'공장코드/명/품목코드/명/검사분류코드 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtItemCd", Trim(.txtItemCd.value)
		WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
		
		If Trim(.cboInspClassCd.value) = "P" Then
			WriteCookie "txtRoutNo", Trim(.txtRoutNo.value)
			WriteCookie "txtRoutNoDesc", Trim(.txtRoutNoDesc.value)
			WriteCookie "txtOprNo", Trim(.txtOprNo.value)
		End if

		
	End With
	PgmJump(BIZ_PGM_JUMP_ID)
End Function


'============================================= EnableField()  ======================================
'=	Event Name : EnableField
'=	Event Desc :
'========================================================================================================
Sub EnableField(Byval strInspClass)
	If	strInspClass = "P" Then
		Process.style.display	= ""
		Call ggoOper.SetReqAttr(frm1.txtRoutNo, "N")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "N")
	Else	
		Process.style.display	= "none"
		Call ggoOper.SetReqAttr(frm1.txtRoutNo, "Q")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "Q")
	End if
End Sub



'============================================= cboInspClassCd_onchange()  ======================================
'=	Event Name : cboInspClassCd_onchange()
'=	Event Desc :
'========================================================================================================
Sub cboInspClassCd_onchange()
	Call EnableField(frm1.cboInspClassCd.value)
End Sub


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolBar("11101101001011")							'⊙: 버튼 툴바 제어 
	Call EnableField(frm1.cboInspClassCd.value)
	If Trim(frm1.txtPlantCd.value) =  "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtItemCd.focus 
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
   	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
 	
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
    
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If   
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
	Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
End Sub 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    
    	Dim IntRetCD 
    
    	FncQuery = False                                                        						'⊙: Processing is NG
    
    	Err.Clear                                                               						'☜: Protect system from crashing

    	'-----------------------
    	'Check previous data area
    	'-----------------------
    	ggoSpread.Source = frm1.vspdData
        If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    	End If
    
    	'-----------------------
    	'Erase contents area
    	'-----------------------
    	Call ggoOper.ClearField(Document, "2")  
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
	
    	Call InitVariables										'⊙: Initializes local global variables
    
    	'-----------------------
    	'Check condition area
    	'-----------------------
    	If Not chkField(Document, "1") Then							'⊙: This function check indispensable field
       		Exit Function
    	End If
    
    	'-----------------------
    	'Query function call area
    	'-----------------------
    	
		If DbQuery = False then
			Exit Function
		End If											'☜: Query db data
       
    	FncQuery = True
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
	FncNew = False                                                          '⊙: Processing is NG
	
	Err.Clear                                                               '☜: Protect system from crashing
	'On Error Resume Next                                                    '☜: Protect system from crashing
	ggoSpread.Source = frm1.vspdData
	'-----------------------
	'Check previous data area
	'-----------------------
	If ggoSpread.SSCheckChange = True Then
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
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call SetDefaultVal
	Call SetToolBar("11101101001011")							'⊙: 버튼 툴바 제어 
	Call EnableField(frm1.cboInspClassCd.value)
	
	If Trim(frm1.txtPlantCd.value) =  "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtItemCd.focus 
	End If
	FncNew = True
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    
    Dim IntRetCD 
    
    FncDelete = False                                                      						'⊙: Processing is NG
    
    Err.Clear                                                               						'☜: Protect system from crashing
    On Error Resume Next                                                	
    
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
	End If
    
   	FncDelete = True
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
	Dim i
	Dim iRet
	FncSave = False                                                  		       '⊙: Processing is NG

	Err.Clear                                                            	 		  '☜: Protect system from crashing
	
	On Error Resume Next                                           	       '☜: Protect system from crashing
	   
	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then
    	IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
    	Exit Function
    End If
    
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "1") Then                                  '⊙: Check contents area
		Exit Function
	End If
    	
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSDefaultCheck = False Then                                  '⊙: Check contents area
    	Exit Function
    End If
    	
 	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then	
		Exit Function
	End If			                                                  '☜: Save db data
    
	FncSave = True
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
	With frm1
		If .vspdData.MaxRows < 1 then
	    	Exit function
    	End if
		
		.vspdData.ReDraw = False
		ggoSpread.Source = .vspdData	
		
		ggoSpread.CopyRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow
	    .vspdData.Row = .vspdData.ActiveRow
	    .vspdData.Col = C_InspSeries
	    .vspdData.value = ""

	    .vspdData.ReDraw = True                                   					            '☜: Protect system from crashing
	End With

	Call SetActiveCell(frm1.vspdData,C_InspSeries,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement		
	FncCopy = true
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = false
    	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End If
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  						'☜: Protect system from crashing
    FncCancel = true
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow
	
	On Error Resume Next
	
	FncInsertRow = false
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)

	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End If
	End If
		
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		'.vspdData.EditMode = True
		.vspdData.ReDraw = False
		ggoSpread.InsertRow .vspdData.ActiveRow, imRow
    	SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1

		.vspdData.ReDraw = True
    End With
    
	Call SetActiveCell(frm1.vspdData,C_InspSeries,.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.ActiveElement
    
    If Err.number = 0 Then FncInsertRow = True
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = false
	Dim lDelRows
	Dim iDelRowCnt, i
    
    With frm1
		If .vspdData.MaxRows < 1 then
			Exit function
		End if	
		.vspdData.focus
		ggoSpread.Source = .vspdData 
	    
	    '----------  Coding part  -------------------------------------------------------------   
	
		lDelRows = ggoSpread.DeleteRow
	End With
	FncDeleteRow = true
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
    FncPrev = false                                                   						'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    FncNext = false                                                 						'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)					'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
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
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	Call parent.FncFind(Parent.C_MULTI, False)     
End Function


'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
	FncScreenSave = false
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
	FncScreenRestore = false
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
 	Dim strVal
	Err.Clear                                                               					'☜: Protect system from crashing
	Call LayerShowHide(1)

	DbQuery = False
	With frm1	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal	= BIZ_PGM_QRY_ID	& "?txtMode="			& Parent.UID_M0001 _
										& "&txtPlantCd="		& .hPlantCd.value _
										& "&txtItemCd="			& .hItemCd.value _
										& "&cboInspClassCd="	& .hInspClassCd.value _
										& "&txtInspItemCd="		& .hInspItemCd.value _
										& "&txtRoutNo="			& .hRoutNo.value _
										& "&txtOprNo="			& .hOprNo.value _
										& "&lgStrPrevKey="		& lgStrPrevKey _
										& "&txtMaxRows="		& .vspdData.MaxRows
		Else
			strVal	= BIZ_PGM_QRY_ID	& "?txtMode="			& Parent.UID_M0001 _
										& "&txtPlantCd="		& Trim(.txtPlantCd.Value) _
										& "&txtItemCd="			& Trim(.txtItemCd.value) _
										& "&cboInspClassCd="	& Trim(.cboInspClassCd.Value) _
										& "&txtInspItemCd="		& Trim(.txtInspItemCd.value) _
										& "&txtRoutNo="			& Trim(.txtRoutNo.value) _
										& "&txtOprNo="			& Trim(.txtOprNo.value) _
										& "&lgStrPrevKey="		& lgStrPrevKey _
										& "&txtMaxRows="		& .vspdData.MaxRows
		End If
	End With

	Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 

    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE			'⊙: Indicates that current mode is Update mode
	Call SetToolBar("11101111001111")							'⊙: 버튼 툴바 제어 
	Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
	Call EnableField(frm1.cboInspClassCd.value)
'	Call SetActiveCell(frm1.vspdData,C_SampleQty,1,"M","X","X")
'	Set gActiveElement = document.activeElement
	DbQueryOk = true	
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpCnt
	Dim lGrpInsCnt
	Dim lGrpDelCnt 
	Dim strDel
	Dim strVal

	Dim iLoop
	Dim iColSep
	Dim iRowSep
	Dim iMaxRows
	Dim iInsertFlag
	Dim iUpdateFlag
	Dim iDeleteFlag
	Dim arrVal
	Dim arrDel

	Dim strInspSeries
	Dim strSampleQty
	Dim strAccptDecisionQty
	Dim strRejtDecisionQty
	Dim strAccptDecisionDiscrete
	Dim strMaxDefectRatio
	
	Call LayerShowHide(1)
	
	DbSave = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing
	
	iLoop       = 1 
	iColSep     = Parent.gColSep
	iRowSep     = Parent.gRowSep
	iMaxRows    = frm1.vspdData.MaxRows
	iInsertFlag = ggoSpread.InsertFlag
	iUpdateFlag = ggoSpread.UpdateFlag
	iDeleteFlag = ggoSpread.DeleteFlag                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1   
		lGrpInsCnt = 1
		lGrpDelCnt = 1 
		strVal = ""
    	strDel = ""
    		
    	'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    		.vspdData.Row = lRow
			.vspdData.Col = 0
			
			Select Case .vspdData.Text
				Case iInsertFlag					'☜: 신규 
					.vspdData.Col = C_InspSeries
					strInspSeries = Trim(.vspdData.Text)
					.vspdData.Col = C_SampleQty
					strSampleQty = Trim(.vspdData.Text)
					.vspdData.Col = C_AccptDecisionQty
					strAccptDecisionQty = Trim(.vspdData.Text)
					.vspdData.Col = C_RejtDecisionQty
					strRejtDecisionQty = Trim(.vspdData.Text)
					.vspdData.Col = C_AccptDecisionDiscrete
					strAccptDecisionDiscrete = Trim(.vspdData.Text)
					.vspdData.Col = C_MaxDefectRatio
					strMaxDefectRatio = Trim(.vspdData.Text)
					
					strVal = strVal & "C" & iColSep & _
									strInspSeries				& iColSep & _
									strSampleQty				& iColSep & _
									strAccptDecisionQty			& iColSep & _
									strRejtDecisionQty			& iColSep & _
									strAccptDecisionDiscrete	& iColSep & _
									strMaxDefectRatio			& iColSep & _
									CStr(lRow) & iRowSep
					lGrpCnt = lGrpCnt + 1
					lGrpInsCnt = lGrpInsCnt + 1
					ReDim Preserve arrVal(lGrpInsCnt - 1)
					arrVal(lGrpInsCnt - 1) = strVal		

				Case iUpdateFlag				'☜: 수정 
					.vspdData.Col = C_InspSeries
					strInspSeries = Trim(.vspdData.Text)
					.vspdData.Col = C_SampleQty
					strSampleQty = Trim(.vspdData.Text)
					.vspdData.Col = C_AccptDecisionQty
					strAccptDecisionQty = Trim(.vspdData.Text)
					.vspdData.Col = C_RejtDecisionQty
					strRejtDecisionQty = Trim(.vspdData.Text)
					.vspdData.Col = C_AccptDecisionDiscrete
					strAccptDecisionDiscrete = Trim(.vspdData.Text)
					.vspdData.Col = C_MaxDefectRatio
					strMaxDefectRatio = Trim(.vspdData.Text)
					
					strVal = strVal & "U" & iColSep & _
									strInspSeries				& iColSep & _
									strSampleQty				& iColSep & _
									strAccptDecisionQty			& iColSep & _
									strRejtDecisionQty			& iColSep & _
									strAccptDecisionDiscrete	& iColSep & _
									strMaxDefectRatio			& iColSep & _
									CStr(lRow) & iRowSep
					lGrpCnt = lGrpCnt + 1
					lGrpInsCnt = lGrpInsCnt + 1
					ReDim Preserve arrVal(lGrpInsCnt - 1)
					arrVal(lGrpInsCnt - 1) = strVal	
				Case iDeleteFlag				'☜: 삭제 
					.vspdData.Col = C_InspSeries
					strInspSeries = Trim(.vspdData.Text)
					
					strDel = strDel & "D" & iColSep & _
									strInspSeries & iColSep & _
									CStr(lRow) & iRowSep
					lGrpCnt = lGrpCnt + 1
					lGrpDelCnt = lGrpDelCnt + 1
					ReDim Preserve arrDel(lGrpDelCnt - 1)
					arrDel(lGrpDelCnt - 1) = strDel						
			End Select
		Next
	
		strVal = Join(arrVal,iRowSep)
		strDel = Join(arrDel,iRowSep)
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'☜: 비지니스 ASP 를 가동 
	End With
	
	DbSave = True                                                          '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()									'☆: 저장 성공후 실행 로직 
	DbSaveOk = true
	
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
	DbSaveOk = true
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	DbDelete = false
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기타 검사조건</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
	        						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14" ></TD>								
	        						<TD CLASS="TD5" NOWRAP>검사분류</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" tag="12"></SELECT></TD>
	        							</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
	        						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE="20" MAXLENGTH="20" tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR ID="Process">
					      			<TD CLASS="TD5" NOWRAP>라우팅</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="12XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutNo()">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>공정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()">&nbsp;<input TYPE=TEXT NAME="txtOprNoDesc" SIZE="30" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사항목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="검사항목" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspItem()">
										<INPUT TYPE=TEXT NAME="txtInspItemNm" SIZE=20 MAXLENGTH="40" tag="14" ></TD>
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
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE WIDTH="100%" HEIGHT=100% <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>검사방식</TD>
								<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtInspMthdCd" SIZE="10" MAXLENGTH="4" ALT="검사방식" tag="14">
									<INPUT TYPE=TEXT NAME="txtInspMthdNm" SIZE="40" MAXLENGTH="40" tag="14" ></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT="100%" COLSPAN=2>
									<script language =javascript src='./js/q1212ma1_I271712821_vspdData.js'></script>
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
	   	<TD WIDTH="100%">
   	  		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
   	  			<TR>
   	  				<TD WIDTH=10>&nbsp;</td>
    					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspStand">검사기준</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
    				</TR>
    			</TABLE>
    		</TD>
    	</TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" tabindex=-1 ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hInspClassCd" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hInspItemCd" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hOprNo" tag="24" tabindex=-1 >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
