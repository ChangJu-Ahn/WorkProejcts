<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2211QA1
'*  4. Program Name         : 공정검사일보조회 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2004/07/14
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit												'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim strInspClass
Dim IsOpenPop
'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID	= "Q2211qb1.asp"                         '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "Q2211ma1"                     '☆: Cookie에서 사용할 상수 

Dim C_InspReqNo
Dim C_InspResultNo
Dim C_ReleaseDt
Dim C_InspDt
Dim C_ItemCd
Dim C_ItemNm
Dim C_RoutNo
Dim C_RoutDesc
Dim C_OprNo
Dim C_OprNm
Dim C_WCCd
Dim C_WCNm
Dim C_MinorCd
Dim C_MinorNm
Dim C_LotSize
Dim C_InspQty
Dim C_DefectQty
Dim C_DefectRatio
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYMD
CompanyYMD = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gDateFormat)                                           '☆: 초기화면에 뿌려지는 시작 날짜 -----
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------------- 
                 
'==========================================  InitComboBox()  ======================================
'	Name : InitComboBox()
'	Description : Init ComboBox
'==================================================================================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'Q0010' ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboDecision , lgF0, lgF1, Chr(11))
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
	lgBlnFlgChgValue = False
	IsOpenPop = False
    '###검사분류별 변경부분 Start###
    strInspClass = "P"
	'###검사분류별 변경부분 End###	
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

    arrField(0) = "Plant_Cd"					' Field명(0)
    arrField(1) = "Plant_NM"					' Field명(1)
        
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
	
	If IsOpenPop = True Then Exit Function
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If	
	
	IsOpenPop = True
	
	arrParam(0) = "품목팝업"							' 팝업 명칭 
	arrParam(1) = "B_Item_By_Plant a, B_Item b "					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtItemCd.Value)				' Code Condition
	arrParam(3) = ""										' Name Condition
	arrParam(4) = "a.Item_Cd = b.Item_Cd And a.Plant_Cd = " & FilterVar(Trim(frm1.txtPlantCd.Value), "", "S")  			' Where Condition
	arrParam(5) = "품목"								' TextBox 명칭 
	
	arrField(0) = "a.Item_Cd"					' Field명(0)
	arrField(1) = "b.Item_NM"							' Field명(1)
	arrField(2) = "b.SPEC"								' Field명(2)
	
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
		arrParam(4) = "P_ROUTING_HEADER.PLANT_CD =" & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S") & _
					" AND ITEM_CD =" & FilterVar(Ucase(Trim(frm1.txtItemCd.value)),"''","S")
	Else
		arrParam(4) = "P_ROUTING_HEADER.PLANT_CD =" & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S")
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
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = 'P1006'" & _
				  " and A.rout_order in ('F','I') "				
	arrParam(2) = UCase(Trim(frm1.txtOprNo.Value))
	arrParam(3) = ""
	If (Trim(frm1.txtItemCd.value) <> "" AND Trim(frm1.txtRoutNo.value) <> "") THEN
		arrParam(4) = "A.plant_cd =" & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S") & _
					  " and	A.item_cd =" & FilterVar(Ucase(Trim(frm1.txtItemCd.value)),"''","S") & _
					  " and	A.rout_no =" & FilterVar(Ucase(Trim(frm1.txtRoutNo.value)),"''","S")
	ElseIf (Trim(frm1.txtItemCd.value) = "" AND Trim(frm1.txtRoutNo.value) <> "") THEN
		arrParam(4) = "A.plant_cd =" & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S") & _
					  " and	A.rout_no =" & FilterVar(Ucase(Trim(frm1.txtRoutNo.value)),"''","S")
	ElseIf (Trim(frm1.txtItemCd.value) <> "" AND Trim(frm1.txtRoutNo.value) = "") THEN
		arrParam(4) = "A.plant_cd =" & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S") & _
					  " and	A.item_cd =" & FilterVar(Ucase(Trim(frm1.txtItemCd.value)),"''","S")
	Else 		
		arrParam(4) = "A.plant_cd =" & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S") 
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

	arrParam(0) = "작업장팝업"							' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWcCd.Value)					' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = "PLANT_CD =" & FilterVar(Trim(frm1.txtPlantCd.value), "", "S") 	' Where Condition
	arrParam(5) = "작업장"								' TextBox 명칭 
	
	arrField(0) = "WC_CD"									' Field명(0)
    arrField(1) = "WC_NM"									' Field명(1)
    
    arrHeader(0) = "작업장코드"							' Header명(0)
    arrHeader(1) = "작업장명"							
    	
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
End Function

'==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'=================================================================================================
Function CookiePage(Byval Kubun)

	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strCookie
	Dim ii,jj,kk
	Dim iSeq
	Dim IntRetCD
    Dim strTemp
    Dim arrVal
         	
    If Kubun = 1 Then								'Jump로 화면을 이동할 경우 
		If  lgSaveRow <  1 Then
			IntRetCD = DisplayMsgBox("900002",Parent.VB_YES_NO,"X","X")
			Exit Function
		End If	
		
		Redim  lgMark(UBound(lgFieldNM)) 
		
		strCookie  = ""
		iSeq       = 0
		
		For ii = 0 to Parent.C_MaxSelList - 1 
			For jj = 0 to UBound(lgFieldNM) -1
				If lgPopUpR(ii,0) = lgFieldCD(jj) Then
					iSeq = iSeq + 1
					lgMark(jj) = "X"
					strCookie = strCookie & "" & TRIM(LGFIELDNM(JJ)) & "" & Parent.gRowSep
					frm1.vspdData.Row = lgSaveRow
					frm1.vspdData.Col = iSeq
					strCookie = strCookie & frm1.vspdData.Text & Parent.gRowSep
				
					kk = CInt(lgNextSeq(jj)) 
					If kk > 0 And kk <= UBound(lgFieldNM) Then 
						lgMark(kk - 1) = "X"
						iSeq = iSeq + 1
						
						strCookie = strCookie & "" & TRIM(LGFIELDNM(KK-1)) & "" & Parent.gRowSep
						frm1.vspdData.Row = lgSaveRow
						frm1.vspdData.Col = iSeq
						strCookie = strCookie & frm1.vspdData.Text & Parent.gRowSep
					End If    
					jj =  UBound(lgFieldNM)  + 100
				End If    
			Next
		Next      
		
		WriteCookie CookieSplit , strCookie
		
		'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------

		'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
		
		Call PgmJump(BIZ_PGM_JUMP_ID)
	
	ElseIf Kubun = 0 Then							'Jump로 화면이 이동해 왔을경우 
		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		Dim iniSep
			
		If ReadCookie("txtPlantCd") <> "" Then
			frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
		End If
		
		If ReadCookie("txtPlantNm") <> "" Then
			frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
		End If	
				
		WriteCookie "txtPlantCd", ""
		WriteCookie "txtPlantNm", ""

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF
End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030804", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_DefectRatio + 1
		.MaxRows = 0

 		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("7", "13", "2")
		
		ggoSpread.SSSetEdit  C_InspReqNo,	"검사의뢰번호",	15
		ggoSpread.SSSetEdit  C_InspResultNo,"SEQ",		   5
		ggoSpread.SSSetDate  C_ReleaseDt,	"Release일",  10, 2, Parent.gDateFormat  
		ggoSpread.SSSetDate  C_InspDt,		"검사일",     10, 2, Parent.gDateFormat  
		ggoSpread.SSSetEdit  C_ItemCd,		"품목코드",   15
		ggoSpread.SSSetEdit  C_ItemNm,		"품목명",     20
		ggoSpread.SSSetEdit  C_RoutNo,		"라우팅",	  10
		ggoSpread.SSSetEdit  C_RoutDesc,	"라우팅명",   15
		ggoSpread.SSSetEdit  C_OprNo,		"공정",       10
		ggoSpread.SSSetEdit  C_OprNm,		"공정작업명", 15
		ggoSpread.SSSetEdit  C_WCCd,		"작업장코드", 10
		ggoSpread.SSSetEdit  C_WCNm,		"작업장명",   15
		ggoSpread.SSSetEdit  C_MinorCd,     "판정",		   5
		ggoSpread.SSSetEdit  C_MinorNm,		"판정",		  10
		ggoSpread.SSSetFloat C_LotSize,		"로트크기",	  15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_InspQty,     "검사수",	  15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_DefectQty,	"불량수",	  15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_DefectRatio,	"불량률(%)",  15, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
 		
 		Call ggoSpread.SSSetColHidden(C_InspResultNo, C_InspResultNo, True)
 		Call ggoSpread.SSSetColHidden(C_MinorCd, C_MinorCd, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
	    ggoSpread.SSSetSplit2(2)  
		
		.ReDraw = true
		
    End With
End Sub

'==========================================  2.6.1 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()

	C_InspReqNo		= 1
	C_InspResultNo  = 2
	C_ReleaseDt	= 3
	C_InspDt = 4
	C_ItemCd = 5
	C_ItemNm = 6
	C_RoutNo = 7
	C_RoutDesc = 8
	C_OprNo = 9
	C_OprNm = 10
	C_WCCd = 11
	C_WCNm = 12
	C_MinorCd = 13
	C_MinorNm = 14
	C_LotSize = 15
	C_InspQty = 16
	C_DefectQty = 17
	C_DefectRatio = 18

End Sub

'==========================================  2.6.2 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		
		C_InspReqNo		= iCurColumnPos(1)
		C_InspResultNo  = iCurColumnPos(2)
		C_ReleaseDt		= iCurColumnPos(3)
		C_InspDt		= iCurColumnPos(4)
		C_ItemCd		= iCurColumnPos(5)									
		C_ItemNm		= iCurColumnPos(6)
		C_RoutNo		= iCurColumnPos(7)
		C_RoutDesc		= iCurColumnPos(8)
		C_OprNo			= iCurColumnPos(9)
		C_OprNm			= iCurColumnPos(10)
		C_WCCd			= iCurColumnPos(11)
		C_WCNm			= iCurColumnPos(12)
		C_MinorCd		= iCurColumnPos(13)
		C_MinorNm		= iCurColumnPos(14)
		C_LotSize		= iCurColumnPos(15)
		C_InspQty		= iCurColumnPos(16)
		C_DefectQty		= iCurColumnPos(17)
		C_DefectRatio	= iCurColumnPos(18)
				
 	End Select
End Sub


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitComboBox()
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어	
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
   	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = Ucase(Parent.gPlant)
	   	frm1.txtPlantNm.value = Parent.gPlantNm
	End IF
	frm1.txtPlantCd.focus
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode ) 
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		
 			lgSortKey = 1
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

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
Sub PopRestoreSpreadColumnInf()	'###그리드 컨버전 주의부분###
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)	
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	 
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If    
End Sub

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtFr.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtDtFr.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtTo.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtDtTo.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtDtFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function txtDtTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtPlantCd_onChange()
	If Trim(frm1.txtPlantCd.Value) = "" Then
			frm1.txtPlantNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtItemCd_onChange()
	If Trim(frm1.txtItemCd.Value) = "" Then
		frm1.txtItemNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtWcCd
'   Event Desc : txtWcCd Change
'==========================================================================================
Function  txtWcCd_onChange()
	If Trim(frm1.txtWcCd.Value) = "" Then
		frm1.txtWcNm.Value = ""
	End If
End Function

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function								'⊙: This function check indispensable field
    
    If ValidDateCheck(frm1.txtDtFr, frm1.txtDtTo) = False Then
   		frm1.txtDtFr.focus 
		Set gActiveElement = document.activeElement
		Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData 

	If Name_check("A") = False Then
		Set gActiveElement = document.activeElement
		Exit Function
	End If
								'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then	Exit Function

    FncQuery = True															'⊙: Processing is OK
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
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1
		'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------------------------------
		strVal = BIZ_PGM_ID & "?txtPlantCd="	& Trim(.txtPlantCd.value) & _
							  "&txtDtFr="		& Trim(.txtDtFr.Text) & _
							  "&txtDtTo="		& Trim(.txtDtTo.Text) & _
							  "&txtInspReqNo="	& Trim(.txtInspReqNo.value) & _
							  "&txtItemCd="		& Trim(.txtItemCd.value) & _
							  "&txtRoutNo="		& Trim(.txtRoutNo.value) & _
							  "&txtOprNo="		& Trim(.txtOprNo.value) & _
							  "&txtWcCd="		& Trim(.txtWcCd.value) & _
							  "&cboDecision="	& Trim(.cboDecision.value) & _
							  "&txtMaxRows="	& .vspdData.MaxRows & _
							  "&lgStrPrevKey="	& lgStrPrevKey                      '☜: Next key tag
		Call RunMyBizASP(MyBizASP, strVal)
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11000000000111")							'⊙: 버튼 툴바 제어 
	lgBlnFlgChgValue = False
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : Name_Check
'========================================================================================
Function Name_Check(ByVal Check)

	Name_Check = False
	
	With frm1

		'-----------------------
		'Check Plant_Cd	 
		'-----------------------
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(.txtPlantCd.Value,"","S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			.txtPlantNm.Value = ""
			Call DisplayMsgBox("125000","X","X","X")
			.txtPlantCd.focus 
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		.txtPlantNm.Value = lgF0(0)

		If Check = "A" Then
			'-----------------------
			'Check Item_Cd	 
			'-----------------------
			If Trim(.txtItemCd.value) <> "" Then
				
				If 	CommonQueryRs(" b.ITEM_NM "," B_ITEM_BY_PLANT a inner join B_ITEM b on a.ITEM_CD = b.ITEM_CD " , _
								" a.ITEM_CD = " & FilterVar(.txtItemCd.Value,"","S") & " AND a.PLANT_CD = " & FilterVar(.txtPlantCd.Value,"","S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
							
					lgF0 = Split(lgF0, Chr(11))
					.txtItemNm.Value = lgF0(0)
				Else
				
					If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(.txtItemCd.Value,"","S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
								
						lgF0 = Split(lgF0, Chr(11))
						.txtItemNm.Value = lgF0(0)
						Call DisplayMsgBox("122700","X","X","X")
						.txtItemCd.focus 
					Else
						.txtItemNm.Value = ""
						Call DisplayMsgBox("122600","X","X","X")
						.txtItemCd.focus 
					End If
					Exit Function
				End If
			End If
			
			'-----------------------
			'Check Rout_No	 
			'-----------------------
			If Trim(.txtRoutNo.Value) <> "" Then
				If 	CommonQueryRs(" DESCRIPTION "," P_ROUTING_HEADER ", " ROUT_NO = " & FilterVar(.txtRoutNo.Value,"","S") & " AND PLANT_CD = " & FilterVar(.txtPlantCd.Value,"","S"), _
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
				If 	CommonQueryRs(" B.MINOR_NM "," P_ROUTING_DETAIL A,B_MINOR B ", " A.JOB_CD = B.MINOR_CD AND B.MAJOR_CD = 'P1006'"  & " AND A.OPR_NO = " & FilterVar(.txtOprNo.Value,"","S") & _
								"AND A.PLANT_CD = " & FilterVar(.txtPlantCd.Value,"","S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					lgF0 = Split(lgF0, Chr(11))
					.txtOprNoDesc.Value = lgF0(0)
				End If
			End If
			 
			'-----------------------
			'Check WC_Cd	 
			'-----------------------
			If Trim(.txtWCCd.Value) <> "" Then
				If 	CommonQueryRs(" WC_NM "," P_WORK_CENTER ", " WC_CD = " & FilterVar(.txtWCCd.Value,"","S") & " AND PLANT_CD = " & FilterVar(.txtPlantCd.Value,"","S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
							
					lgF0 = Split(lgF0, Chr(11))
					.txtWCNm.Value = lgF0(0)
				Else
					.txtWCNm.Value = ""
					Call DisplayMsgBox("182100","X","X","X")
					.txtWCCd.focus 
					Exit Function
				End If
			End If
		End If
	End With
	
	Name_Check = True

End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공정검사일보조회</font></TD>
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
										<script language =javascript src='./js/q2211qa1_fpDateTime5_txtDtFr.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/q2211qa1_fpDateTime6_txtDtTo.js'></script>
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
									<TD CLASS="TD5" NOWRAP>판정</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboDecision" ALT="판정" STYLE="WIDTH: 150px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>작업장</TD>
	        						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=12 MAXLENGTH=20 ALT="작업장" tag="11XXXU"><IMG align=top height=20 name=btnWcCd onclick="vbscript:OpenWc()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
														<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>	
	     							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TR>
							<TD HEIGHT=100% WIDTH=100% Colspan=2>
								<script language =javascript src='./js/q2211qa1_I590472012_vspdData.js'></script>
							</TD>	
						</TR>	
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
        					<!--<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">공정검사</a></TD>-->
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
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
    </DIV>
</BODY>
</HTML>
