<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m1311ma1
'*  4. Program Name         : PL등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2003/06/04
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit		
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
Const BIZ_PGM_ID = "m1311mb1.asp"												'☆: 비지니스 로직 ASP명 
Const C_OpenRef_file_A = "m1311ra1.asp"
Const C_OpenRef_file_B = "m1311ra2.asp"

Dim lblnWinEvent
Dim lgOpenFlag
Dim lgRefABflag

Dim C_ChdItemCd 		
Dim C_BtnItemPopUp		
Dim C_ChdItemNm 		
Dim C_ParItemQty		
Dim C_ParItemUnit		
Dim C_ParItemUnitPop 	
Dim C_ChdItemQty  		
Dim C_ChdItemUnit		
Dim C_ChdItemUnitPop	
Dim C_Paytype		 	
Dim C_PaytypeNm			
Dim C_PlSeqNo			
Dim C_Loss		 		

Dim IsOpenPop          

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

'========================================================================================================
Sub ReadCookiePage()

	Dim strTemp, arrVal
	
	If (Parent.ReadCookie("m1311Plant") = "") or (Parent.ReadCookie("m1311Item") = "") or (Parent.ReadCookie("m1311Supplier") = "") then Exit sub

	frm1.txtPlantCd.value = Parent.ReadCookie("m1311Plant")
	frm1.txtItemCd.value = Parent.ReadCookie("m1311Item")
	frm1.txtSpplCd.value = Parent.ReadCookie("m1311Supplier")
				
	Call Parent.WriteCookie("m1311Plant" , "")
	Call Parent.WriteCookie("m1311Item" , "")
	Call Parent.WriteCookie("m1311Supplier" , "")
	
	Call MainQuery()	
End Sub
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE   
    lgBlnFlgChgValue = False    
    lgIntGrpCount = 0           
    
    lgStrPrevKey = ""           
    lgLngCurRows = 0            
    frm1.vspdData.MaxRows = 0
    
End Sub
'========================================================================================================
Sub SetDefaultVal()
	lgOpenFlag = False    
    lgRefABflag = ""
    frm1.txtPlantCd.value=Parent.gPlant
	frm1.txtPlantNm.value=Parent.gPlantNm	
	frm1.txtPlantCd2.value=Parent.gPlant
	frm1.txtPlantNm2.value=Parent.gPlantNm	
	frm1.txtFrDt.text = StartDate
    Call SetToolbar("1110111100101111")
  
    frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement
End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ChdItemCd 		= 1
	C_BtnItemPopUp		= 2
	C_ChdItemNm 		= 3
	C_ParItemQty		= 4 
	C_ParItemUnit		= 5
	C_ParItemUnitPop 	= 6
	C_ChdItemQty  		= 7
	C_ChdItemUnit		= 8
	C_ChdItemUnitPop	= 9
	C_Paytype		 	= 10
	C_PaytypeNm			= 11
	C_PlSeqNo			= 12
	C_Loss		 		= 13
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20030520",, parent.gAllowDragDropSpread
       .ReDraw = false
	
		.MaxCols = C_Loss+1					
		.MaxRows = 0
	
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_ChdItemCd,		"자품목", 15,,,18,2
		ggoSpread.SSSetButton 	C_BtnItemPopUp
		ggoSpread.SSSetEdit 	C_ChdItemNm,		"자품목명", 25
		SetSpreadFloatLocal	C_ParItemQty,		"모품목수량", 15,1,6
		ggoSpread.SSSetEdit 	C_ParItemUnit,		"모품목단위", 10,,,3,2
		ggoSpread.SSSetButton 	C_ParItemUnitPop
		SetSpreadFloatLocal	C_ChdItemQty,		"자품목소요수량", 15,1,6
		ggoSpread.SSSetEdit 	C_ChdItemUnit,		"자품목단위",10,,,3,2
		ggoSpread.SSSetButton 	C_ChdItemUnitPop
		ggoSpread.SSSetCombo 	C_Paytype,		"",10,0,False
		ggoSpread.SSSetCombo 	C_PaytypeNm,		"지급구분",10,0,False
		ggoSpread.SSSetEdit 	C_PlSeqNo,		"PLSeqNo", 20
		SetSpreadFloatLocal	C_Loss,			"Loss율(%)", 12,1,5
		
		call ggoSpread.MakePairsColumn(C_ChdItemCd,C_BtnItemPopUp)
		call ggoSpread.MakePairsColumn(C_ParItemUnit,C_ParItemUnitPop)
		call ggoSpread.MakePairsColumn(C_ChdItemUnit,C_ChdItemUnitPop)
		
		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)	
		Call ggoSpread.SSSetColHidden(C_PayType,	C_PayType,	True)
		Call ggoSpread.SSSetColHidden( C_PlSeqNo,C_PlSeqNo,	True)
		
		.ReDraw = true
		Call SetSpreadLock 
	
    End With
End Sub
'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
     
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_ChdItemCd 		= iCurColumnPos(1)
			C_BtnItemPopUp		= iCurColumnPos(2)
			C_ChdItemNm 		= iCurColumnPos(3)
			C_ParItemQty		= iCurColumnPos(4) 
			C_ParItemUnit		= iCurColumnPos(5)
			C_ParItemUnitPop 	= iCurColumnPos(6)
			C_ChdItemQty  		= iCurColumnPos(7)
			C_ChdItemUnit		= iCurColumnPos(8)
			C_ChdItemUnitPop	= iCurColumnPos(9)
			C_Paytype		 	= iCurColumnPos(10)
			C_PaytypeNm			= iCurColumnPos(11)
			C_PlSeqNo			= iCurColumnPos(12)
			C_Loss		 		= iCurColumnPos(13)
			
    End Select    
End Sub
'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.sssetrequired	C_ChdItemCd,	-1,			-1
    ggoSpread.SSSetProtected C_ChdItemNm,	-1,			-1	
	ggoSpread.spreadunlock 	C_ParItemQty,	-1,			C_ParItemQty,		-1
	ggoSpread.sssetrequired  C_ParItemQty,	-1,			-1
	ggoSpread.sssetrequired  C_ParItemUnit,	-1,			-1
	ggoSpread.sssetrequired  C_ChdItemQty,	-1,			-1
	ggoSpread.sssetrequired  C_ChdItemUnit,	-1,			-1
	ggoSpread.sssetrequired  C_PayTypeNm,	-1,			-1
	ggoSpread.spreadunlock 	C_Loss,			-1,			C_Loss,				-1
	ggoSpread.SSSetProtected frm1.vspdData.MaxCols, -1
End Sub
'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    frm1.vspdData.ReDraw = False
	ggoSpread.sssetrequired C_ChdItemCd,			pvStartRow,			pvEndRow
	ggoSpread.SSSetProtected C_ChdItemNm,			pvStartRow,			pvEndRow
	ggoSpread.sssetrequired C_ParItemQty,			pvStartRow,			pvEndRow
	ggoSpread.sssetrequired C_ParItemUnit,			pvStartRow,			pvEndRow
	ggoSpread.sssetrequired C_ChdItemQty,			pvStartRow,			pvEndRow
	ggoSpread.sssetrequired C_ChdItemUnit,			pvStartRow,			pvEndRow
	ggoSpread.sssetrequired C_PayTypeNm,			pvStartRow,			pvEndRow
	ggoSpread.SSSetProtected frm1.vspdData.MaxCols, pvStartRow,			pvEndRow
	frm1.vspdData.ReDraw = True
End Sub
'========================================================================================================
Sub InitComboBox()
	
    Dim strDataCd, strDataNm
    Dim strCboDataCd 
    Dim strCboDataNm
    Dim iStrWhere 
    Dim i
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear    
    
    iStrWhere = " MAJOR_CD = " & FilterVar("M2201", "''", "S") & " "                                   
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ",iStrWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	

	strDataCd = split(lgF0,chr(11))
	strDataNm = split(lgF1,chr(11))
	
	For i = 0 to Ubound(strDataCd,1) - 1
		strCboDataCd = strCboDataCd & strDataCd(i) & vbTab
		strCboDataNm = strCboDataNm & strDataNm(i) & vbTab
	Next
	
    ggoSpread.SetCombo strCboDataCd, C_Paytype
	ggoSpread.SetCombo strCboDataNm, C_PaytypeNm
End Sub
'========================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_Plant"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "Plant_CD"	
    arrField(1) = "Plant_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)
		frm1.txtPlantNm.Value= arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'========================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	If UCase(frm1.txtItemCd.ClassName) = UCase(Parent.UCN_PROTECTED) Then Exit Function
	 
	If Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If
	 
	IsOpenPop = True
	'***2003.3월 패치분 수정(2003.02.26-Lee,Eun Hee)-유효일추가*****
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "12!MO"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "20!M"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus	
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)  
		frm1.txtItemNm.Value    = arrRet(1)  
		frm1.txtItemCd.focus	
		Set gActiveElement = document.activeElement   
		
	End If 
End Function
'========================================================================================================
Function OpenPlant2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd2.ClassName)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_Plant"				
	arrParam(2) = Trim(frm1.txtPlantCd2.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "Plant_CD"	
    arrField(1) = "Plant_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd2.focus
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtPlantCd2.Value= arrRet(0)
		frm1.txtPlantNm2.Value= arrRet(1)
		frm1.txtPlantCd2.focus
		Set gActiveElement = document.activeElement	
	End If	
	
End Function
'========================================================================================================
Function OpenItem2()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	If UCase(frm1.txtItemCd2.ClassName) = UCase(Parent.UCN_PROTECTED) Then Exit Function
	 
	If Trim(frm1.txtPlantCd2.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd2.focus
		Exit Function
	End If
	 
	IsOpenPop = True
	'***2003.3월 패치분 수정(2003.02.26-Lee,Eun Hee)-유효일추가*****
	arrParam(0) = Trim(frm1.txtPlantCd2.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd2.value)		' Item Code
	arrParam(2) = "12!MO"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "20!O"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus	
		Exit Function
	Else
		frm1.txtItemCd2.Value    = arrRet(0)  
		frm1.txtItemNm2.Value    = arrRet(1)  
		frm1.txtItemCd2.focus	
		Set gActiveElement = document.activeElement   
		
	End If 
End Function
'========================================================================================================
Function OpenChdItem()
	
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspddata.col = C_ChdItemCd
	'***2003.3월 패치분 수정(2003.02.26-Lee,Eun Hee)-유효일추가*****

	arrParam(0) = Trim(frm1.txtPlantCd2.value)	' Plant Code
	arrParam(1) = Trim(frm1.vspddata.text)
'@@반제품 포함 수정[060522]
'	arrParam(2) = "35!PP"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
'	arrParam(3) = "30!P"
	arrParam(2) = "25!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_ChdItemCd,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call frm1.vspdData.SetText(C_ChdItemNm,		frm1.vspdData.ActiveRow,	arrRet(1))
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow frm1.vspdData.ActiveRow
	End If	
	
	Call changeItemPlant()	
End Function
'========================================================================================================
Function OpenParItemUnit()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "모품목단위"	
	arrParam(1) = "B_Unit_OF_MEASURE"				
	arrParam(2) = GetSpreadText(frm1.vspdData,C_ParItemUnit,frm1.vspdData.ActiveRow,"X","X")
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "모품목단위"			
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "모품목단위"		
    arrHeader(1) = "모품목단위명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_ParItemUnit,		frm1.vspdData.ActiveRow,	arrRet(0))
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow frm1.vspdData.ActiveRow
	End If	
	
End Function
'========================================================================================================
Function OpenChdItemUnit()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ChdItemUnit
	
	arrParam(0) = "자품목단위"	
	arrParam(1) = "B_Unit_OF_MEASURE"				
	arrParam(2) = GetSpreadText(frm1.vspdData,C_ChdItemUnit,frm1.vspdData.ActiveRow,"X","X")
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "자품목단위"			
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "자품목단위"		
    arrHeader(1) = "자품목단위명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_ChdItemUnit,		frm1.vspdData.ActiveRow,	arrRet(0))
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow frm1.vspdData.ActiveRow
	End If	
	
End Function
'========================================================================================================
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtSpplCd.ClassName) = UCase(Parent.UCN_PROTECTED) then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "외주처"					
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.txtSpplCd.value)	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "				
	arrParam(5) = "외주처"					

	arrField(0) = "BP_CD"						
	arrField(1) = "BP_NM"						

	arrHeader(0) = "외주처"					
	arrHeader(1) = "외주처명"				

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtSpplCd.Value  = arrRet(0)		
		frm1.txtSpplNm.Value  = arrRet(1)
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement	
	End If
End Function	
'========================================================================================================
Function OpenSppl2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtSpplCd2.ClassName) = UCase(Parent.UCN_PROTECTED) then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "외주처"						
	arrParam(1) = "B_BIZ_PARTNER"					
	arrParam(2) = Trim(frm1.txtSpplCd2.value)		
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "외주처"						

	arrField(0) = "BP_CD"							
	arrField(1) = "BP_NM"							

	arrHeader(0) = "외주처"						
	arrHeader(1) = "외주처명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSpplCd2.focus
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtSpplCd2.Value  = arrRet(0)		
		frm1.txtSpplNm2.Value  = arrRet(1)
		frm1.txtSpplCd2.focus
		Set gActiveElement = document.activeElement	
	End If
End Function	
'========================================================================================================
Function OpenRefA()

	Dim strRet
	Dim arrParam(5)
	Dim iCalledAspName
		
	If lblnWinEvent = True  Then Exit Function
				
	if Trim(frm1.txtPlantCd2.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공장","X")
		frm1.txtPlantCd2.focus
		Set gActiveElement = document.activeElement
		Exit Function 
	end if
	
	if Trim(frm1.txtItemCd2.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "모품목","X")
		frm1.txtItemCd2.focus
		Set gActiveElement = document.activeElement	
		Exit Function 
	end if
	
	if Trim(frm1.txtSpplCd2.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "외주처","X")
		frm1.txtSpplCd2.focus
		Set gActiveElement = document.activeElement	
		Exit Function 
	end if
	
	
	if Trim(frm1.txtFrDt.Text) = "" or Trim(frm1.txtToDt.Text) = "" then
		Call DisplayMsgBox("17A002","X" , "적용유효일","X")
	'	frm1.txtSpplCd2.focus
		Set gActiveElement = document.activeElement	
		Exit Function 
	end if
	
	
	If lgOpenFlag = False Then 
		Call changeItem()		
		Exit Function
	End If
		
	If Trim(frm1.hdnProcure.value)	<>	"O" and  Trim(frm1.hdnProcure.value)	<>	"M" then 
		
		Call DisplayMsgBox("17A012","X" , "이 품목" , "BOM복사참조")
		lgOpenFlag	= False
		frm1.txtItemCd2.focus
		Set gActiveElement = document.activeElement	
		
		Exit Function 
	end if
		
	lblnWinEvent = True	
	
	arrParam(0) = lgIntFlgMode
	arrParam(1) = Trim(frm1.txtPlantCd2.value)
	arrParam(2) = Trim(frm1.txtItemCd2.Value)
        '200704 KSJ 추가(BOM적용유효일추가)
	arrParam(3) = Trim(frm1.txtFrDt.text)
	arrParam(4) = Trim(frm1.txtToDt.text)

	arrParam(5) = Trim(frm1.txtBomNo.Value)
	
	Call SetAflag()
	
	iCalledAspName = AskPRAspName("m1311ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1311ra1", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lblnWinEvent = False
	
	lgOpenFlag	= False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetRefA(strRet)
		Call ChangeTag(true)
	End If	
		
End Function
'========================================================================================================
Function SetRefA(strRet)

	Dim Index1,Count1
	Dim boolExist
	Dim ilngRow
	Dim iMaxRows
	Dim iCurRow
	Dim intIndex
	
	Const C_ChdItemCd_Ref 		= 0
	Const C_ChdItemNm_Ref 		= 1
	Const C_ParItemQty_Ref		= 2 
	Const C_ParItemUnit_Ref		= 3
	Const C_ChdItemQty_Ref  	= 4
	Const C_ChdItemUnit_Ref		= 5
	Const C_Paytype_Ref		 	= 6
	Const C_Loss_Ref		 	= 7
	Const C_ValidFrDt			= 8
	Const C_ValidToDt			= 9
			
	Count1 		= Ubound(strRet,1)
	
	boolExist	= False
	ilngRow		= 0
	
	With frm1
		.vspdData.ReDraw = False
		
		For index1 = 0 to Count1-2
			
			iMaxRows = .vspdData.MaxRows
			'중복된 자품목 참조 Check 20040423 주석 처리 
			'ilngRow = .vspdData.SearchCol(C_ChdItemCd, ilngRow, iMaxRows, strRet(index1,C_ChdItemCd_Ref), 0)
			
			'If ilngRow <> -1 Then
			 '  boolExist = True
			  ' Call DisplayMsgBox("17a005","X",strRet(Index1,C_ChdItemCd_Ref) & ";","자품목")
			   .vspdData.ReDraw = True
			 '  Exit Function
			'End If
			
			If boolExist <> True then
				Call fncinsertrow(1)
				iCurRow	= .vspdData.ActiveRow 
				
				Call .vspdData.SetText(C_ChdItemCd,		iCurRow, strRet(index1,C_ChdItemCd_Ref))
				Call .vspdData.SetText(C_ChdItemNm,		iCurRow, strRet(index1,C_ChdItemNm_Ref))
				Call .vspdData.SetText(C_ParItemQty,	iCurRow, strRet(index1,C_ParItemQty_Ref))
				Call .vspdData.SetText(C_ParItemUnit,	iCurRow, strRet(index1,C_ParItemUnit_Ref))
				Call .vspdData.SetText(C_ChdItemQty,	iCurRow, strRet(index1,C_ChdItemQty_Ref))
				Call .vspdData.SetText(C_ChdItemUnit,	iCurRow, strRet(index1,C_ChdItemUnit_Ref))
				Call .vspdData.SetText(C_Paytype,		iCurRow, strRet(index1,C_Paytype_Ref))
				
				.vspdData.Row = iCurRow
				.vspdData.Col = C_Paytype
				intIndex = .vspdData.Value
				.vspdData.Col = C_PayTypeNm
				.vspdData.Value = intIndex
				
				Call .vspdData.SetText(C_Loss,	iCurRow, strRet(index1,C_Loss_Ref))
			Else
				boolExist = False
			End if 
		Next
	
		frm1.txtBomNo.Value = strRet(Count1,4)	
		.vspdData.ReDraw = True
	End with
End Function
'========================================================================================================
Function OpenRefB()

	Dim strRet
	Dim arrParam(3)
	Dim iCalledAspName
	
	If lblnWinEvent = True  Then Exit Function
	
	if Trim(frm1.txtPlantCd2.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공장","X")
		frm1.txtPlantCd2.focus
		Set gActiveElement = document.activeElement
		Exit Function 
	end if
	
	if Trim(frm1.txtItemCd2.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "모품목","X")
		frm1.txtItemCd2.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	
	if Trim(frm1.txtSpplCd2.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "외주처","X")
		frm1.txtSpplCd2.focus
		Set gActiveElement = document.activeElement	
		Exit Function 
	end if
	
	if Trim(frm1.txtFrDt.Text) = "" or Trim(frm1.txtToDt.Text) = "" then
		Call DisplayMsgBox("17A002","X" , "적용유효일","X")
		Set gActiveElement = document.activeElement	
		Exit Function 
	end if
	
	If Trim(frm1.hdnProcure.value)	<>	"O" and Trim(frm1.hdnProcure.value)	<>	"M" then 
		Call DisplayMsgBox("17A012","X" , "이 품목" , "타업체복사참조")
		lblnWinEvent	= False
		frm1.txtItemCd2.focus
		Set gActiveElement = document.activeElement	
		Exit Function 
	end if
	
	lblnWinEvent = True
			
	arrParam(0) = lgIntFlgMode
	arrParam(1) = Trim(frm1.txtPlantCd2.Value)
	arrParam(2) = Trim(frm1.txtItemCd2.Value)
	arrParam(3) = Trim(frm1.txtSpplCd2.Value)
	
	Call SetBflag()

	iCalledAspName = AskPRAspName("m1311ra2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1311ra2", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetRefB(strRet)
		Call ChangeTag(true)
	End If	
		
End Function
'========================================================================================================
Function SetRefB(strRet)

	Dim Index1,Count1
	Dim boolExist
	Dim ilngRow
	Dim iMaxRows
	Dim iCurRow
	Dim intIndex

	Const C_ChdItemCd_Ref 		= 0
	Const C_ChdItemNm_Ref 		= 1
	Const C_ParItemQty_Ref		= 2 
	Const C_ParItemUnit_Ref		= 3
	Const C_ChdItemQty_Ref  	= 4
	Const C_ChdItemUnit_Ref		= 5
	Const C_PaytypeCd_Ref		= 6
	Const C_PaytypeNm_Ref	 	= 7
	Const C_Loss_Ref		 	= 8


	Count1 		= Ubound(strRet,1)
	boolExist = False
	
	with frm1
		
		For index1 = 0 to Count1-2
	
			iMaxRows = .vspdData.MaxRows
			'중복된 자품목 참조 Check 20040423 주석처리 
			'ilngRow = .vspdData.SearchCol(C_ChdItemCd, ilngRow, iMaxRows, strRet(index1,C_ChdItemCd_Ref), 0)
			
			'If ilngRow <> -1 Then
			'   boolExist = True
			'   Call DisplayMsgBox("17a005","X",strRet(Index1,C_ChdItemCd_Ref) & ";","자품목")
			 '  Call fncCancel()
			'   .vspdData.ReDraw = True
			'   Exit Function
			'End If
			
			If boolExist <> True then
				
				Call fncinsertrow(1)
				iCurRow	= .vspdData.ActiveRow 
				
				Call .vspdData.SetText(C_ChdItemCd,		iCurRow, strRet(index1,C_ChdItemCd_Ref))
				Call .vspdData.SetText(C_ChdItemNm,		iCurRow, strRet(index1,C_ChdItemNm_Ref))
				Call .vspdData.SetText(C_ParItemQty,	iCurRow, strRet(index1,C_ParItemQty_Ref))
				Call .vspdData.SetText(C_ParItemUnit,	iCurRow, strRet(index1,C_ParItemUnit_Ref))
				Call .vspdData.SetText(C_ChdItemQty,	iCurRow, strRet(index1,C_ChdItemQty_Ref))
				Call .vspdData.SetText(C_ChdItemUnit,	iCurRow, strRet(index1,C_ChdItemUnit_Ref))
				Call .vspdData.SetText(C_Paytype,		iCurRow, strRet(index1,C_PaytypeCd_Ref))
				
				.vspdData.Row = iCurRow
				.vspdData.Col = C_Paytype			
				intIndex = .vspdData.Value
				.vspdData.Col = C_PayTypeNm
				.vspdData.Value = intIndex
				
				Call .vspdData.SetText(C_Loss,			iCurRow, strRet(index1,C_Loss_Ref))
			Else
				boolExist = False
			End if 
		Next
		.vspdData.ReDraw = True
	End with
End Function
'========================================================================================================
Function SetAflag()
	lgRefABflag = "A"
End Function
'========================================================================================================
Function SetBflag()
	lgRefABflag = "B"
End Function
'========================================================================================================
Function ResetABflag()
	lgRefABflag = ""	
End Function
'========================================================================================================
Sub ChangeTag(ByVal flg)
	
	if flg = true then
		ggoOper.SetReqAttr	frm1.txtPlantCd2, "Q"
		ggoOper.SetReqAttr	frm1.txtItemCd2, "Q"
		'ggoOper.SetReqAttr	frm1.txtSpplCd2, "Q"
	else
		ggoOper.SetReqAttr	frm1.txtPlantCd2, "N"
		ggoOper.SetReqAttr	frm1.txtItemCd2, "N"
		ggoOper.SetReqAttr	frm1.txtSpplCd2, "N"
	end if
	
End Sub
'========================================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 6                                                              
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6",		            ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"    
   End Select
         
End Sub
'========================================================================================================
Sub SetSpreadLockAfterQuery()

	frm1.vspdData.ReDraw = False
	
    ggoSpread.sssetrequired	C_ChdItemCd, -1, -1
	ggoSpread.SSSetProtected C_ChdItemNm, -1, -1
	ggoSpread.spreadunlock 	C_ParItemQty, -1, C_ParItemQty, -1
	ggoSpread.sssetrequired C_ParItemQty, -1, -1
	ggoSpread.sssetrequired C_ParItemUnit, -1, -1
	ggoSpread.sssetrequired C_ChdItemQty, -1, -1
	ggoSpread.sssetrequired C_ChdItemUnit, -1, -1
	ggoSpread.sssetrequired C_PayTypeNm, -1, -1
	ggoSpread.spreadunlock 	C_Loss, -1, C_Loss, -1
	frm1.vspdData.ReDraw = True
	
End Sub
'========================================================================================================
Sub Form_Load()
        
    Call LoadInfTB19029                             
    Call ggoOper.LockField(Document, "N")           
    
    Call AppendNumberPlace("6","9","4")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                                                    
    Call SetDefaultVal
    Call InitVariables
    
    Call ReadCookiePage()  
    Call InitComboBox()
       
End Sub
'========================================================================================================
Function changeItem()

	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                               
 
    If Trim(frm1.txtPlantCd2.Value) = "" OR Trim(frm1.txtItemCd2.Value) = "" Then
    	Exit Function
    End if
    
    changeItem = False                 
    
    With frm1
		Call CommonQueryRs(" c.plant_cd, c.plant_nm, b.item_cd, b.item_nm, a.Procur_Type,b.basic_unit "," b_item_by_plant a (noLock), b_item b (noLock),  b_plant c (noLock) ", " a.plant_cd = c.plant_cd AND a.item_cd = b.item_cd AND c.plant_cd =  " & FilterVar(.txtPlantCd2.value, "''", "S") & "  AND b.item_cd =  " & FilterVar(.txtItemCd2.value, "''", "S") & "  "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
		if lgF2 = "" or lgF2 = Null then
			Call DisplayMsgBox("122700","X","X","X")
			.txtItemCd2.value = ""
			.txtItemNm2.value = ""
			.txtItemCd2.focus 
			Set gActiveElement = document.activeElement
			Exit Function
		END IF
		
		If Ubound(split(lgF0,chr(11)),1) > 0 then	
			.txtPlantCd2.Value   = split(lgF0,chr(11))(0)
			.txtPlantNm2.Value   = split(lgF1,chr(11))(0)
			.txtItemCd2.Value    = split(lgF2,chr(11))(0)
			.txtItemNm2.Value    = split(lgF3,chr(11))(0)
		    .hdnProcure.Value    = split(lgF4,chr(11))(0)
            .hdnUnit.Value       = split(lgF5,chr(11))(0)
		    If lgRefABflag = "A" Then
				 lgOpenFlag = "True"
				 lgRefABflag = ""
				 OpenRefA()
		    ElseIf lgRefABflag = "B" Then
				 lgOpenFlag = "True"
				 lgRefABflag = ""
				 OpenRefB()
		    End If
		End If
    End With
	
    changeItem = True                  

End Function
'========================================================================================================
Function changeItemPlant()
    Dim Itemcd
    Dim strVal
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                               
    
    Itemcd = UCase(Trim(GetSpreadText(frm1.vspdData,C_ChdItemCd,frm1.vspdData.ActiveRow,"X","X")))
    
    if Trim(frm1.txtPlantCd2.Value) = "" or Itemcd = "" then
        Exit Function
    End if
    
    changeItemPlant = False                 
    
    If LayerShowHide(1) = False Then Exit Function
        
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeItemPlant"
    strVal = strVal & "&txtItemCd= " & FilterVar(Itemcd, " " ,  "SNM")
    strVal = strVal & "&txtPlantCd=" & FilterVar(Trim(frm1.txtPlantCd2.Value), " " ,"SNM")

    Call RunMyBizASP(MyBizASP, strVal)
	
    changeItemPlant = True                  

End Function
'========================================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtFrDt.focus
	End if
End Sub
'========================================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtToDt.focus
	End if
End Sub
'========================================================================================================
Sub txtFrDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'========================================================================================================
Sub txtToDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	 gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("1101111111")

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	frm1.vspdData.Row = Row
End Sub
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.source= frm1.vspdData     
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)   

    if C_ChdItemCd  = Col then
        Call changeItemPlant()
    end if      
       
End Sub
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = Col
	intIndex = frm1.vspdData.Value

	frm1.vspdData.Col = C_PayTypeNm-1
	frm1.vspdData.Value = intIndex
End Sub
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	if frm1.vspdData.ActiveCol = C_ParItemUnitPop then
		Call OpenParItemUnit()
	elseif frm1.vspdData.ActiveCol = C_ChdItemUnitPop then
		Call OpenChdItemUnit()
	elseif frm1.vspdData.ActiveCol = C_BtnItemPopUp then
	    Call OpenChdItem()	
	End if
	
End Sub
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then	
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
    End if
    
End Sub
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    ggoSpread.Source = frm1.vspdData
    
    Err.Clear                                               
	
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    														
    If Not chkField(Document, "1") Then						
       Exit Function
    End If
    
    If DbQuery = False Then Exit Function
       
    FncQuery = True											
    Set gActiveElement = document.activeElement
End Function
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    On Error Resume Next                                   
	Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData    
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call ChangeTag(false)
    Call InitVariables                                      
    Call SetDefaultVal
 
    FncNew = True                                           
	Set gActiveElement = document.activeElement
End Function

'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
	    
    FncSave = False 

    Err.Clear       

    ggoSpread.Source = frm1.vspdData 
         
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  
        IntRetCD = DisplayMsgBox("900001","X","X","X")            
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then               
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData                  
    If Not ggoSpread.SSDefaultCheck Then              
       Exit Function
    End If
    
    with frm1
        If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" Then
			Call DisplayMsgBox("17a003","X","적용유효일","X")			
			Exit Function
		End if   
	End with

    If DbSave = False Then Exit Function
    
    FncSave = True                                    
    Set gActiveElement = document.activeElement
End Function

'========================================================================================================

Function FncDelete() 
    
	Dim IntRetCD

    FncDelete = False
    
    ggoSpread.Source = frm1.vspdData  
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function
    						
    If lgIntFlgMode <> Parent.OPMD_UMODE Then 
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    FncDelete = True    
    
End Function

'========================================================================================================

Function FncCancel() 

	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                
    
    if frm1.vspdData.MaxRows < 1 then
    	Call ChangeTag(False)
    End if
    Set gActiveElement = document.activeElement
End Function
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		
		If imRow = "" Then
			Exit Function
		End if
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	
	If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    if frm1.vspdData.Maxrows < 1	then exit function
    
    frm1.vspdData.focus
    ggoSpread.Source = frm1.vspdData 
        
	lDelRows = ggoSpread.DeleteRow
    
    Set gActiveElement = document.activeElement
End Function
'========================================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'========================================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_SINGLE)	
    Set gActiveElement = document.activeElement				
End Function
'========================================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_SINGLE , False)     
    Set gActiveElement = document.activeElement      
End Function
'========================================================================================================
Function FncExit()
	
	Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
     Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet() 
    Call ggoSpread.ReOrderingSpreadData()
    ggoSpread.SSSetProtected C_ChdItemCd , -1	
End Sub
'========================================================================================================
Function FncCopy()
	frm1.vspdData.ReDraw = False
	if frm1.vspdData.Maxrows < 1	then exit function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
    Call frm1.vspdData.SetText(1,	frm1.vspdData.ActiveRow,	"")
    Call frm1.vspdData.SetText(2,	frm1.vspdData.ActiveRow,	"")
    
	frm1.vspdData.ReDraw = True
	Set gActiveElement = document.activeElement
End Function
'========================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim strVal
	
	lgRefABflag	=""
	DbQuery = False
    
    if LayerShowHide(1) = False then
       Exit Function 
    end if
    
    Err.Clear                                                           

	With frm1
    
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtPlantCd=" & .hdnPlant.value
		    strVal = strVal & "&txtitemCd=" & .hdnItem.value
			strVal = strVal & "&txtSpplCd=" & .hdnSppl.Value
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows	    
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		    strVal = strVal & "&txtitemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtSpplCd=" & Trim(.txtSpplCd.value)
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows	    
		end if 
	
		.hdnmaxrow.value = .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)
    End With
    
    DbQuery = True
End Function
'========================================================================================================
Function DbQueryOk()													
	Dim ii
	
    lgBlnFlgChgValue = False

    if frm1.vspdData.MaxRows > 0 then
    	Call SetToolbar("1110111100101111")
		lgIntFlgMode = Parent.OPMD_UMODE
		
		Call ggoOper.LockField(Document, "Q")
		
		Call SetSpreadLockAfterQuery()		
    else
    	Call SetToolbar("1111111100101111")
		lgIntFlgMode = Parent.OPMD_UMODE   
	
		Call ggoOper.LockField(Document, "N")
	end if
	
	Call RemovedivTextArea
	
	ggoSpread.SSSetProtected C_ChdItemCd , -1		
	ggoSpread.SSSetProtected C_BtnItemPopUp , -1	
End Function
'========================================================================================================
Function DbQueryOkhdr()													
	
    lgBlnFlgChgValue = False

    Call ggoOper.LockField(Document, "Q")								
    
	frm1.txtPlantCd2.value = frm1.txtPlantCd.value
	frm1.txtPlantNm2.value = frm1.txtPlantNm.value
	frm1.txtItemCd2.value = frm1.txtItemCd.value
	frm1.txtItemNm2.value = frm1.txtItemNm.value
	frm1.txtSpplCd2.value = frm1.txtSpplCd.value
	frm1.txtSpplNm2.value = frm1.txtSpplNm.value
	Call ggoOper.LockField(Document, "N")
    Call SetToolbar("1110111100101111")
	lgIntFlgMode = Parent.OPMD_CMODE   
	  
End Function
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim strVal,strDel
	Dim iColSep, iRowSep
	
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size

    DbSave = False                                                      
    
    Call DisableToolBar(Parent.TBC_SAVE)                                          '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)   
    
    On Error Resume Next                                               
	Err.Clear
	
	iColSep = Parent.gColSep													
	iRowSep = Parent.gRowSep													
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]
	
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

	With frm1
		if lgIntFlgMode = Parent.OPMD_UMODE then
			.txtMode.value = Parent.UID_M0005
		else
			.txtMode.value = Parent.UID_M0002
		end if
		
	    strVal = ""
	    strDel = ""
	    
	    For lRow = 1 To .vspdData.MaxRows

	        .vspdData.Row = lRow
	        .vspdData.Col = 0		
			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag	       
				.vspdData.Col = C_ParItemQty		'4
				If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
					Call DisplayMsgBox("970021","X","모품목수량","X")
					Call LayerShowHide(0)
					Exit Function
				End if
			
				.vspdData.Col = C_ChdItemQty		'7
				If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
					Call DisplayMsgBox("970021","X","자품목소요수량","X")
					Call LayerShowHide(0)
					Exit Function
				End if
			End Select			
			
			.vspdData.Col = 0		
			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag		
				
					strVal = "C"																				& iColSep				
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ChdItemCd,lRow, "X","X"))				& iColSep
					strVal = strVal & ""																		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ChdItemNm,lRow,"X","X"))				& iColSep
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_ParItemQty,lRow,"X","X"),0)		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ParItemUnit,lRow,"X","X"))				& iColSep
					strVal = strVal & ""																		& iColSep
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_ChdItemQty,lRow,"X","X"),0)		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ChdItemUnit,lRow,"X","X"))				& iColSep
					strVal = strVal & ""																		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Paytype,lRow,"X","X"))					& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PaytypeNm,lRow,"X","X"))				& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlSeqNo,lRow,"X","X"))					& iColSep
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_Loss,lRow,"X","X"),0)			& iColSep
					strVal = strVal & lRow & iRowSep
	                
	            Case ggoSpread.UpdateFlag		
					
					strVal = "U"																				& iColSep				
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ChdItemCd,lRow, "X","X"))				& iColSep
					strVal = strVal & ""																		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ChdItemNm,lRow,"X","X"))				& iColSep
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_ParItemQty,lRow,"X","X"),0)		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ParItemUnit,lRow,"X","X"))				& iColSep
					strVal = strVal & ""																		& iColSep
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_ChdItemQty,lRow,"X","X"),0)		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ChdItemUnit,lRow,"X","X"))				& iColSep
					strVal = strVal & ""																		& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Paytype,lRow,"X","X"))					& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PaytypeNm,lRow,"X","X"))				& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlSeqNo,lRow,"X","X"))					& iColSep
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_Loss,lRow,"X","X"),0)			& iColSep
					strVal = strVal & lRow & iRowSep
	                
	            Case ggoSpread.DeleteFlag
				
					strDel = "D"																				& iColSep				
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ChdItemCd,lRow, "X","X"))				& iColSep
					strDel = strDel & ""																		& iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ChdItemNm,lRow,"X","X"))				& iColSep
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_ParItemQty,lRow,"X","X"),0)		& iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ParItemUnit,lRow,"X","X"))				& iColSep
					strDel = strDel & ""																		& iColSep
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_ChdItemQty,lRow,"X","X"),0)		& iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ChdItemUnit,lRow,"X","X"))				& iColSep
					strDel = strDel & ""																		& iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_Paytype,lRow,"X","X"))					& iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_PaytypeNm,lRow,"X","X"))				& iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_PlSeqNo,lRow,"X","X"))					& iColSep
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_Loss,lRow,"X","X"),0)			& iColSep
					strDel = strDel & lRow & iRowSep
	                
	        End Select
	        
	        .vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
			       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			End Select    
	    Next
	End With
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   
	'------ Developer Coding part (End ) -------------------------------------------------------------- 
    Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'==============================================================================================================================
Function DbSaveOk()												
   
	Call InitVariables
    lgIntFlgMode = Parent.OPMD_UMODE		
	Call ggoOper.LockField(Document, "Q")
	
	lgBlnFlgChgValue = False
	
	Call MainQuery()
	
End Function
'========================================================================================================

Function DbDelete() 
    Err.Clear                       
    
    DbDelete = False				
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003					
    strVal = strVal & "&txtPLNo=" & Trim(frm1.hdnPLNo.value)		
    strVal = strVal & "&txtSppl=" & Trim(frm1.hdnSppl.value)
    strVal = strVal & "&txtItem=" & Trim(frm1.hdnItem.value)
    strVal = strVal & "&txtPlant=" & Trim(frm1.hdnPlant.value)
    strVal = strVal & "&txtUpdtUserId=" & Parent.gUsrID
    
    if LayerShowHide(1) = False then
       Exit Function 
    end if   
	
	Call RunMyBizASP(MyBizASP, strVal)								

    DbDelete = True                                                 

End Function

'========================================================================================

Function DbDeleteOk()
	lgBlnFlgChgValue = False
	Call FncNew()
End Function

'========================================================================================

Function changePlntCd(inwhere)
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                        
    
    changePlntCd = False           
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    Dim strVal
        
    if inwhere = "1" then
		strVal = BIZ_PGM_ID & "?txtMode=" & "changePlntCd"
		strVal = strVal & "&FLG=1"
		strVal = strVal & "&txtPlantCd=" & FilterVar(Trim(frm1.txtPlantCd.Value),"","SNM")
	ELSE
		strVal = BIZ_PGM_ID & "?txtMode=" & "changePlntCd"
		strVal = strVal & "&FLG=2"
		strVal = strVal & "&txtPlantCd=" & FilterVar(Trim(frm1.txtPlantCd2.Value),"","SNM")
	END IF
		    
    Call RunMyBizASP(MyBizASP, strVal)
	
	changePlntCd = True            

End Function
'========================================================================================================
Function changeItemCd(inwhere)
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                        
    
    changeItemCd = False           
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    Dim strVal    
           
    if inwhere = "1" then 
		strVal = BIZ_PGM_ID & "?txtMode=" & "changeItemCd"
		strVal = strVal & "&FLG=1"
		strVal = strVal & "&txtPlantCd=" & FilterVar(Trim(frm1.txtPlantCd.Value),"","SNM")
		strVal = strVal & "&txtItemCd=" & FilterVar(Trim(frm1.txtItemCd.Value),"","SNM")
    ELSE
		strVal = BIZ_PGM_ID & "?txtMode=" & "changeItemCd"
		strVal = strVal & "&FLG=2"
		strVal = strVal & "&txtPlantCd=" & FilterVar(Trim(frm1.txtPlantCd2.Value),"","SNM")
		strVal = strVal & "&txtItemCd=" & FilterVar(Trim(frm1.txtItemCd2.Value),"","SNM")
	END IF
    Call RunMyBizASP(MyBizASP, strVal)
	
    changeItemCd = True            

End Function
'========================================================================================================
Function changeSpplCd(inwhere)
	Dim strVal
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                        
    
    changeSpplCd = False           
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    if inwhere = "1" then
		strVal = BIZ_PGM_ID & "?txtMode=" & "changeSpplCd"
		strVal = strVal & "&FLG=1"
		strVal = strVal & "&txtSupplierCd=" & FilterVar(Trim(frm1.txtSpplCd.Value),"","SNM")
	Else 
		strVal = BIZ_PGM_ID & "?txtMode=" & "changeSpplCd"
		strVal = strVal & "&FLG=2"
		strVal = strVal & "&txtSupplierCd=" & FilterVar(Trim(frm1.txtSpplCd2.Value),"","SNM")
	End if
	
    
    Call RunMyBizASP(MyBizASP, strVal)
	
    changeSpplCd = True            
End Function
'==============================================================================================================================
Function RemovedivTextArea()
	Dim ii
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function
'==============================================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 border="0">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>외주P/L</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenRefA" onMouseOver="vbscript:SetAflag" onMouseOut="vbscript:ResetABflag" onFocus="vbscript:SetAflag" onBlur="vbscript:ResetABflag">BOM복사</A> | 
											<A href="vbscript:OpenRefB" onMouseOver="vbscript:SetBflag" onMouseOut="vbscript:ResetABflag" onFocus="vbscript:SetBflag" onBlur="vbscript:ResetABflag">타업체복사</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													<INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 tag="14x"></TD>
									<TD CLASS="TD5" NOWRAP>모품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="모품목" NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
													<INPUT TYPE=TEXT ALT="모품목" NAME="txtItemNm" SIZE=20 tag="14x"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>외주처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="외주처" NAME="txtSpplCd"  SIZE=10 MAXLENGTH=10 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSppl()">
														   <INPUT TYPE=TEXT ALT="외주처" NAME="txtSpplNm" SIZE=20 tag="14x"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd2" SIZE=10 MAXLENGTH=4 tag="23NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant2()">
												<INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm2" SIZE=20 tag="24x"></TD>
								<TD CLASS="TD5" NOWRAP>모품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="모품목" NAME="txtItemCd2" SIZE=15 MAXLENGTH=18 tag="23NXXU" ONCHANGE="vbscript:changeItem()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem2()">
												<INPUT TYPE=TEXT ALT="모품목" NAME="txtItemNm2" SIZE=20 tag="24x"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>외주처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="외주처" NAME="txtSpplCd2" SIZE=10 MAXLENGTH=10 tag="23NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSppl2()">
													   <INPUT TYPE=TEXT ALT="외주처" NAME="txtSpplNm2" SIZE=20 tag="24x"></TD>
								<TD CLASS="TD5" NOWRAP>적용유효일</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td NOWRAP>
												<OBJECT ALT=적용유효일 NAME="txtFrDt" classid=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22N1" Title="FPDATETIME"></OBJECT>
											</td>
											<td NOWRAP>
												~
											</td>
											<td NOWRAP>
												<OBJECT ALT=적용유효일 NAME="txtToDt" classid=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="22N1" Title="FPDATETIME"></OBJECT>
											</td>
										</tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>BOM Type</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="BOM Type" NAME="txtBomNo" SIZE=34 tag="24X">
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<OBJECT classid=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  >
										<PARAM NAME="MaxCols" VALUE="0">
										<PARAM NAME="MaxRows" VALUE="0">
									</OBJECT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSppl" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnProcure" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnUnit" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow"  tag="14">
<P ID="divTextArea"></P>
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
