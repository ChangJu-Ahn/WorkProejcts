<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2213QA1
'*  4. Program Name         : 불량유형조회 
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit														'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim	lgTopLeft_A	'@@@변경 
Dim	lgTopLeft_B	'@@@변경 
Dim lgStrPrevKey_A
Dim lgStrPrevKey_B

Dim strInspClass
Dim IsOpenPop

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYMD
CompanyYMD = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gDateFormat)                                           '☆: 초기화면에 뿌려지는 시작 날짜 -----
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID          = "q2213qb1.asp"             '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1		  = "q2213qb2.asp"             '☆: Biz logic spread sheet for #2

Const C_SHEETMAXROWS_A    = 50                         '☆: Spread sheet에서 보여지는 row for #1
Const C_SHEETMAXROWS_D_A  = 100                        '☆: Server에서 한번에 fetch할 최대 데이타 건수 

Const C_SHEETMAXROWS_B    = 50                         '☆: Spread sheet에서 보여지는 row for #2
Const C_SHEETMAXROWS_D_B  = 100                        '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey  = 4                                    '☆☆☆☆: Max key value
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
	IsOpenPop = False
    '###검사분류별 변경부분 Start###
    strInspClass = "P"
	'###검사분류별 변경부분 End###	                               'Indicates that no value changed
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtDtFr.Text	= CompanyYMD
	frm1.txtDtTo.Text	= CompanyYMD
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","QA") %>
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet(ByVal pvGridId)
    If pvGridId = "A" Then                                   ' 초기화 Spreadsheet #1 
        Call SetZAdoSpreadSheet("Q2213QA1", "S", "A", "V20021125", parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
    End If

    Call SetZAdoSpreadSheet("Q2213QA1", "S", "B", "V20021125", parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X")
    Call SetSpreadLock(pvGridId)
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(Byval pvGridId)
    If pvGridId = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
    End If

    ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.id)      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_Plant"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Condition
	arrParam(4) = ""
	arrParam(5) = "공장"							' TextBox 명칭 

    arrField(0) = "B_Plant.Plant_Cd"					' Field명(0)
    arrField(1) = "B_Plant.Plant_NM"					' Field명(1)
        
    arrHeader(0) = "공장코드"						' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtPlantCd.Focus
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	Else
		Exit Function
	End If	
	Set gActiveElement = document.activeElement
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
	Param6 = "R"			'검사진행상태 
	
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

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "품목팝업"							' 팝업 명칭 
	arrParam(1) = "B_Item_By_Plant,B_Item"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtItemCd.Value)				' Code Condition
	arrParam(3) = ""										' Name Condition
	arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd"
	arrParam(4) = arrParam(4) & "  And B_Item_By_Plant.Plant_Cd =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " " 			' Where Condition
	arrParam(5) = "품목"								' TextBox 명칭 
	
	arrField(0) = "B_Item_By_Plant.Item_Cd"					' Field명(0)
	arrField(1) = "B_Item.Item_NM"							' Field명(1)
	arrField(2) = "B_Item.SPEC"								' Field명(2)
	
	arrHeader(0) = "품목코드"							' Header명(0)
	arrHeader(1) = "품목명"								' Header명(1)
	arrHeader(2) = "규격"								' Header명(2)
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtItemCd.Focus
	If Trim(arrRet(0)) <> "" Then
		frm1.txtItemCd.Value = Trim(arrRet(0))
		frm1.txtItemNm.Value = Trim(arrRet(1))
	Else
		Exit Function
	End If
	Set gActiveElement = document.activeElement
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

	arrParam(0) = "라우팅 팝업"					' 팝업 명칭 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtRoutNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	If Trim(frm1.txtItemCd.value) <> "" Then
		arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
					" AND ITEM_CD = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S")
	Else
		arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
	End if		
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
	ElseIf (Trim(frm1.txtItemCd.value) = "" AND Trim(frm1.txtRoutNo.value) <> "") THEN
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



'------------------------------------------  OpenWc()  -------------------------------------------------
'	Name : OpenWc()
'	Description : Supplier PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenWc()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True

	arrParam(0) = "작업장팝업"						' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWcCd.Value)		' Code Condition
	arrParam(3) = ""						' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "	' Where Condition
	arrParam(5) = "작업장"						' TextBox 명칭 
	
	arrField(0) = "WC_CD"						' Field명(0)
    arrField(1) = "WC_NM"						' Field명(1)
    
    arrHeader(0) = "작업장코드"						' Header명(0)
    arrHeader(1) = "작업장명"					' Header명(1)
    	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtWcCd.Focus
	If Trim(arrRet(0)) <> "" Then
		frm1.txtWcCd.Value = Trim(arrRet(0))
		frm1.txtWcNm.Value = Trim(arrRet(1))
	Else
		Exit Function
	End If
	Set gActiveElement = document.activeElement
	OpenWc = true
End Function

'------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : Inspection Item By Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem()
	OpenInspItem = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	'품목코드가 있는 지 체크 
	If Trim(frm1.txtItemCd.Value) = "" then 
		Call DisplayMsgBox("229916", "X", "X", "X") 		'품목정보가 필요합니다 
		frm1.txtItemCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목별 검사항목팝업"						' 팝업 명칭 
	arrParam(1) = "Q_Inspection_Standard_By_Item, Q_Inspection_Item"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtInspItemCd.Value)		' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "Q_Inspection_Standard_By_Item.Insp_Item_Cd = Q_Inspection_Item.Insp_Item_Cd"
	arrParam(4) = arrParam(4) & "  And Q_Inspection_Standard_By_Item.Plant_Cd =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " "
	arrParam(4) = arrParam(4) & "  And Q_Inspection_Standard_By_Item.Item_Cd =  " & FilterVar(frm1.txtItemCd.Value, "''", "S") & " " 			' Where Condition
	arrParam(4) = arrParam(4) & "  And Q_Inspection_Standard_By_Item.insp_class_cd=" & FilterVar("P", "''", "S") & "   "
	arrParam(5) = "검사항목"						' TextBox 명칭 
	
	arrField(0) = "Q_Inspection_Standard_By_Item.INSP_ITEM_CD"							' Field명(0)
	arrField(1) = "Q_Inspection_Item.INSP_ITEM_NM"							' Field명(1)
	
	arrHeader(0) = "검사항목코드"						' Header명(0)
	arrHeader(1) = "검사항목명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtInspItemCd.Focus
	If Trim(arrRet(0)) <> "" Then
		frm1.txtInspItemCd.Value = Trim(arrRet(0))
		frm1.txtInspItemNm.Value = Trim(arrRet(1))
	Else
		Exit Function
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspItem = true
End Function


'==================================================================================
' Name : PopZAdoConfigGrid
' Desc :
'==================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	Call OpenOrderByPopup(gActiveSpdSheet.Id)
End Sub

'===========================================================================
' Function Name : OpenOrderByPopup
' Function Desc : OpenOrderByPopup Reference Popup
'===========================================================================
Function OpenOrderByPopup(ByVal pvGridId)

	Dim arrRet
	
	On Error Resume Next
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp", Array(ggoSpread.GetXMLData(pvGridId), gMethodText), _
	         "dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData(pvGridId, arrRet(0), arrRet(1))
		Call InitVariables
		Call InitSpreadSheet(pvGridId)
		If pvGridId = "B" Then
			Call DbqueryOnLeaveCell(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow)
		End If
   End If
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal
	Call InitSpreadSheet("A")
    Call SetToolbar("11000000000011")
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
	   	frm1.txtPlantNm.value = Parent.gPlantNm
	End IF
	frm1.txtPlantCd.focus
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode )
End Sub

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtFr.Action = 7
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtTo.Action = 7
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX KeyPress
'==========================================================================================
Function  txtDtFr_KeyPress(KeyAscii)
	txtDtFr_KeyPress = false
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
	txtDtFr_KeyPress = true
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX KeyPress
'==========================================================================================
Function txtDtTo_KeyPress(KeyAscii)
	txtDtTo_KeyPress = false
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
	txtDtTo_KeyPress = true
End Function

'==========================================================================================
'   Event Name : txtPlantCd
'   Event Desc : Change
'==========================================================================================
Function  txtPlantCd_onChange()
	txtPlantCd_onChange = false
	If Trim(frm1.txtPlantCd.Value) = "" Then
			frm1.txtPlantNm.Value = ""
	End If
	txtPlantCd_onChange = true
End Function

'==========================================================================================
'   Event Name : txtItemCd
'   Event Desc : Change
'==========================================================================================
Function  txtItemCd_onChange()
	txtItemCd_onChange = true
	If Trim(frm1.txtItemCd.Value) = "" Then
		frm1.txtItemNm.Value = ""
	End If
	txtItemCd_onChange = true
End Function

'==========================================================================================
'   Event Name : txtBpCd
'   Event Desc : Change
'==========================================================================================
Function  txtWcCd_onChange()
	txtWcCd_onChange = false
	If Trim(frm1.txtWcCd.Value) = "" Then
		frm1.txtWcNm.Value = ""
	End If
	txtWcCd_onChange = true
End Function

'==========================================================================================
'   Event Name : txtInspItemCd
'   Event Desc : Change
'==========================================================================================
Function  txtInspItemCd_onChange()
	txtInspItemCd_onChange = false
	If Trim(frm1.txtInspItemCd.Value) = "" Then
		frm1.txtInspItemNm.Value = ""
	End If
	txtInspItemCd_onChange = true
End Function

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row = NewRow Or NewRow <= 0 Then Exit Sub

	If CheckRunningBizProcess = True Then
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = Col	
		frm1.vspdData.Action = 0
		Exit Sub
	End If
	
	Call DbqueryOnLeaveCell(NewCol, NewRow)

End Sub

Sub DbqueryOnLeaveCell(ByVal Col, ByVal Row)
	lgStrPrevKey_B = ""	'@@@변경 
	
	Call DisableToolBar(parent.TBC_QUERY)  
	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)	
	
	frm1.vspdData2.MaxRows = 0	'@@@변경 
	
	If DbQuery("B") = False Then
	   Call RestoreToolBar()
	   Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("00000000001")
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
	Call SetPopupMenuItemInf("00000000001")
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

    If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData2
        Exit Sub
    End If

    Call SetSpreadColumnValue("B", frm1.vspdData2, Col, Row)
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
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
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	'☜: 재쿼리 체크'
      	If CheckRunningBizProcess = True Then	'@@@변경 
			Exit Sub
		End If
				
		If lgStrPrevKey_A <> "" Then                           '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음	
			lgTopLeft_A = "Y"	'@@@변경 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery("A") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
   End If    
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows =< NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then	'☜: 재쿼리 체크'
      	If CheckRunningBizProcess = True Then	'@@@변경 
			Exit Sub
		End If
		
		If lgStrPrevKey_B <> "" Then                        '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			lgTopLeft_B = "Y"		'@@@변경 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery("B") = False Then	'@@@변경 
				Call RestoreToolBar()
				Exit Sub
			End If

		End If
   End If    
End Sub

'===========================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'===========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'===========================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'===========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is 
'========================================================================================
Function FncQuery()
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear     


    If lgBlnFlgChgValue = True Then
		IntRetCD = .DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
    
    If ValidDateCheck(frm1.txtDtFr, frm1.txtDtTo) = False Then
		Exit Function
    End If

    If Name_check("A") = False Then
		Set gActiveElement = document.activeElement
		Exit Function
	End If

    '-----------------------
    'Query function call area
    '-----------------------
    lgStrPrevKey_A = ""
    
	If DbQuery("A") = False Then   
		Exit Function           
    End If														'☜: Query db data
	
    FncQuery = True		
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = true
End Function

'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function DbQuery(ByVal pOpt) 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)

    With frm1
        If pOpt = "A" Then
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
			strVal	= BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value) _
					& "&txtDtFr=" & Trim(.txtDtFr.Text) _
					& "&txtDtTo=" & Trim(.txtDtTo.Text) _
					& "&txtInspReqNo=" & Trim(.txtInspReqNo.value) _
    				& "&txtItemCd=" & Trim(.txtItemCd.value) _
    				& "&txtRoutNo=" & Trim(.txtRoutNo.value) _
    				& "&txtOprNo=" & Trim(.txtOprNo.value) _
					& "&txtWcCd=" & Trim(.txtWcCd.value) _
					& "&txtInspItemCd=" & Trim(.txtInspItemCd.value) _
					& "&iOpt=" & pOpt
        Else   
			strVal	= BIZ_PGM_ID1 & "?txtPlantCd=" & Trim(.txtPlantCd.value) _
					& "&txtInspReqNo=" & GetKeyPosVal("A", 1) _
					& "&txtInspResultNo=" & GetKeyPosVal("A", 2) _
					& "&txtInspItemCd=" & GetKeyPosVal("A", 3) _
					& "&txtInspSeries=" & GetKeyPosVal("A", 4) _
					& "&iOpt=" & pOpt
        End If   
   
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        If pOpt = "A" Then
           strVal	= strVal & "&lgStrPrevKey="   & lgStrPrevKey_A _
					& "&lgSelectListDT=" & GetSQLSelectListDataType(pOpt) _
					& "&lgTailList="     & MakeSQLGroupOrderByList(pOpt) _
					& "&lgSelectList="   & EnCoding(GetSQLSelectList(pOpt)) _
					& "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D_A)
        Else   
           strVal	= strVal & "&lgStrPrevKey="   & lgStrPrevKey_B _
					& "&lgSelectListDT=" & GetSQLSelectListDataType(pOpt) _
					& "&lgTailList="     & MakeSQLGroupOrderByList(pOpt) _
					& "&lgSelectList="   & EnCoding(GetSQLSelectList(pOpt)) _
					& "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D_B)
        End If  

        Call RunMyBizASP(MyBizASP, strVal)
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(ByVal pOpt)														'☆: 조회 성공후 실행로직 
	DbQueryOk = false

    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11000000000111")							'⊙: 버튼 툴바 제어 

	If pOpt = "A" Then
		If lgTopLeft_A <> "Y" Then	'@@@변경 
			Call DbqueryOnLeaveCell(1, 1)
		End If
		lgTopLeft_A = "N"	'@@@변경 
    End If
    
    frm1.vspdData.focus
    DbQueryOk = true
End Function

'========================================================================================
' Function Name : Name_Check
'========================================================================================
Function Name_Check(ByVal Check)
	Name_Check = False

	With frm1
		If Check = "A" Then
			'-----------------------
			'Check Rout_No	 
			'-----------------------
			If Trim(.txtRoutNo.Value) <> "" Then
				If 	CommonQueryRs(" DESCRIPTION "," P_ROUTING_HEADER ", " ROUT_NO = " & FilterVar(.txtRoutNo.Value, "''", "S") & " AND PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
							
					lgF0 = Split(lgF0, Chr(11))
					.txtRoutNoDesc.Value = lgF0(0)
				Else
					.txtRoutNoDesc.Value = ""
					Call DisplayMsgBox("181300","X","X","X")
					.txtRoutNo.focus 
					Exit Function
				End If
			End If
			'-----------------------
			'Check Opr_No	 
			'-----------------------
			If Trim(.txtOprNo.Value) <> "" Then
				If 	CommonQueryRs(" B.MINOR_NM "," P_ROUTING_DETAIL A,B_MINOR B ", " A.JOB_CD = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("P1006", "''", "S") & "" & " AND A.OPR_NO = " & FilterVar(.txtOprNo.Value, "''", "S") & _
								"AND A.PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					lgF0 = Split(lgF0, Chr(11))
					.txtOprNoDesc.Value = lgF0(0)
				End If
			End If
		End If
	End With
	
	Name_Check = True

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공정검사불량유형조회</font></td>
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
        									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14" ></TD>								
        									<TD CLASS="TD5" NOWRAP>기간</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/q2213qa1_fpDateTime5_txtDtFr.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/q2213qa1_fpDateTime6_txtDtTo.js'></script>										
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
        							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20 MAXLENGTH=18 ALT="검사의뢰번호" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnInspReqNo align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenInspReqNo()"></TD>
        							<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnItemCd align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenItem()">
														   <INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>라우팅</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="11XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutNo()">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>공정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="11XXXU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()">&nbsp;<input TYPE=TEXT NAME="txtOprNoDesc" SIZE="30" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=12 MAXLENGTH=20 ALT="작업장" tag="11XXXU"><IMG align=top height=20 name=btnBpCd onclick="vbscript:OpenWc()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
														   <INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>검사항목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="검사항목" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspItem()">
														   <INPUT TYPE=TEXT NAME="txtInspItemNm" SIZE=20 MAXLENGTH="40" tag="14" ></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
										<TR HEIGHT="*">
											<TD WIDTH="60%">
												<script language =javascript src='./js/q2213qa1_A_vspdData.js'></script>
											</TD>
											<TD WIDTH=10>&nbsp;</TD>
											<TD WIDTH="40%">
												<script language =javascript src='./js/q2213qa1_B_vspdData2.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
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
        					<!--<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">공정검사불량유형등록</a></TD>-->
					<TD WIDTH="*" ALIGN="RIGHT">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>
