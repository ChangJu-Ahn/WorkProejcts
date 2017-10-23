<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1211MA2
'*  4. Program Name         : 공급처별배분비등록 
'*  5. Program Desc         : 공급처별배분비등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/01/09
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID	= "m1211mb2.asp"
CONST BIZ_PGM_ID2	= "m1211mb201.asp"												'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'spdData
Dim C_PlantCd 	             '공장 
Dim C_PlantNm 	             '공장명 
Dim C_ItemCd 	             '품목 
Dim C_ItemNm 	             '품목명 
Dim C_SpplSpec 	             '품목규격 

'spdData2
Dim C_SpplCd                 '공급처 
Dim C_SpplNm 	             '공급처명 
Dim C_Quota_Rate             '배분비율 
Dim C_Purpriority            '발주배정가중치 
Dim C_Defflg                 '주공급업체여부 
Dim C_SpplDlvylt             '구매L/T
Dim C_GrpCd 	             '구매그룹 
Dim C_GrpNm 	             '구매그룹명 
Dim C_ParentPlantCd
Dim C_ParentItemCd
Dim C_ParentRowNo
Dim C_RecordCnt

Dim lgIntFlgModeM           'Variable is for Operation Status
Dim lgStrPrevKeyM()			'Multi에서 재쿼리를 위한 변수 
Dim lglngHiddenRows()		'Multi에서 재쿼리를 위한 변수	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgSortKey1
Dim lgSortKey2

Dim lgPageNo1
Dim lgCurrRow
Dim lgSpdHdrClicked

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE   
    lgIntFlgModeM = Parent.OPMD_CMODE		'Indicates that current mode is Create mode3            
    lgBlnFlgChgValue = False                
    lgIntGrpCount = 0                       
    lgStrPrevKey1 = ""						'initializes Previous Key
    lgStrPrevKey2 = ""						'initializes Previous Key
    
    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
    lgSortKey2 = 2
    lgPageNo = 0
    lgPageNo1 = 0
    
    frm1.vspdData.MaxRows = 0
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtPlantCd.Value = Parent.gPlant
	frm1.txtPlantNm.Value = Parent.gPlantNm
	Call SetToolbar("1110000000001111")
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
	Set gActiveSpdSheet = frm1.vspdData
End Sub
'====================================================================================================
Sub ReadCookiePage()
	
	if Trim(ReadCookie("m1211qa1_plantcd")) = "" then Exit Sub
	
	frm1.txtPlantCd.Value	 = ReadCookie("m1211qa1_plantcd")
	frm1.txtItemCd.Value	 = ReadCookie("m1211qa1_itemcd")
	
	Call MainQuery()
	
	Call WriteCookie("m1211qa1_plantcd","")
	Call WriteCookie("m1211qa1_itemcd","")
	Call WriteCookie("m1211qa1_suppliercd","")
End Sub
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub
'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "A" Then	
		C_PlantCd 	  = 1             '공장 
		C_PlantNm 	  = 2             '공장명 
		C_ItemCd 	  = 3             '품목 
		C_ItemNm 	  = 4             '품목명 
		C_SpplSpec 	  = 5             '품목규격 
		
	Else
		C_SpplCd      = 1             '공급처 
		C_SpplNm 	  = 2             '공급처명 
		C_Quota_Rate  = 3             '배분비율 
		C_Purpriority = 4             '발주배정가중치 
		C_Defflg      = 5            '주공급업체여부 
		C_SpplDlvylt  = 6            '구매L/T
		C_GrpCd 	  = 7            '구매그룹 
		C_GrpNm 	  = 8            '구매그룹명 
		C_ParentPlantCd = 9
		C_ParentItemCd	= 10
		C_ParentRowNo	= 11
		C_RecordCnt		= 12
	End If
End Sub
'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
    Call InitSpreadPosVariables(pvSpdNo)
	
	If pvSpdNo = "A" Then
		With frm1.vspdData	
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20030219",,Parent.gAllowDragDropSpread
	
			.ReDraw = false
			.MaxCols = C_SpplSpec + 1							
			.Col = .MaxCols:	.ColHidden = True
			.MaxRows = 0
    
			Call GetSpreadColumnPos("A")
 
			ggoSpread.SSSetEdit 	C_PlantCd, "공장", 15
			ggoSpread.SSSetEdit 	C_PlantNm,"공장명",20
			ggoSpread.SSSetEdit 	C_ItemCd,"품목",20
			ggoSpread.SSSetEdit 	C_ItemNm, "품목명", 25
			ggoSpread.SSSetEdit 	C_SpplSpec, "품목규격", 25
				
			Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantNm)
			Call ggoSpread.MakePairsColumn(C_ItemCd,C_SpplSpec)

			Call SetSpreadLock("A") 

			.ReDraw = true
		End With
	
	Elseif  pvSpdNo = "B" Then
		With frm1.vspdData2	
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20030219",,Parent.gAllowDragDropSpread

			.ReDraw = false
			.MaxCols = C_RecordCnt + 1							
			.MaxRows = 0
    
			Call GetSpreadColumnPos("B")
 
			ggoSpread.SSSetEdit		C_SpplCd		,"공급처"			, 10,,,10,2
			ggoSpread.SSSetEdit 	C_SpplNm		,"공급처명"			, 18
			SetSpreadFloatLocal		C_Quota_Rate	,"배분비율(%)"		,15,1,5
			ggoSpread.SSSetEdit		C_Purpriority	,"발주배정가중치"	,15
			ggoSpread.SSSetEdit 	C_Defflg		,"주공급업체여부"	,15, 2
			ggoSpread.SSSetEdit		C_SpplDlvylt	,"구매L/T"			,15
			ggoSpread.SSSetEdit 	C_GrpCd			,"구매그룹"			,15,,,4,2
			ggoSpread.SSSetEdit 	C_GrpNm			,"구매그룹명"		,20
			ggoSpread.SSSetEdit 	C_ParentPlantCd	, ""		, 10
			ggoSpread.SSSetEdit 	C_ParentItemCd	, ""		, 10
			ggoSpread.SSSetEdit     C_ParentRowNo	, ""		, 25,2,,,2
			ggoSpread.SSSetEdit     C_RecordCnt		, ""		, 25,2,,,2
	
			Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
			Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)
			
			Call ggoSpread.SSSetColHidden(C_ParentPlantCd,	C_RecordCnt,	True)		
			Call ggoSpread.SSSetColHidden(.MaxCols,			.MaxCols,	True)		
			
			Call SetSpreadLock("B") 
    
			.ReDraw = true
		End with
	End If
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
		If pvSpdNo = "A" Then
			ggoSpread.SpreadLock		-1 , -1
			Call SetSpreadColor(-1,-1)
		Else
			.vspdData.ReDraw = False
			ggoSpread.SpreadLock		-1 , -1
			ggoSpread.SpreadLock		C_PlantCd , -1
			ggoSpread.SpreadUnLock		C_Quota_Rate , -1, -1
			ggoSpread.SSSetRequired		C_Quota_Rate, -1, -1                  '배분비 
			.vspdData.ReDraw = True
		End IF
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected		C_PlantCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_PlantNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_ItemCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_ItemNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_SpplSpec, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd		= iCurColumnPos(1)
			C_PlantNm 		= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_SpplSpec		= iCurColumnPos(5)
			
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SpplCd 		= iCurColumnPos(1)
			C_SpplNm		= iCurColumnPos(2)
			C_Quota_Rate	= iCurColumnPos(3)
			C_Purpriority	= iCurColumnPos(4)
			C_Defflg        = iCurColumnPos(5)
			C_SpplDlvylt    = iCurColumnPos(6)
			C_GrpCd         = iCurColumnPos(7)
			C_GrpNm         = iCurColumnPos(8)
			C_ParentPlantCd = iCurColumnPos(9)
			C_ParentItemCd  = iCurColumnPos(10)
			C_ParentRowNo   = iCurColumnPos(11)
			C_RecordCnt     = iCurColumnPos(12)
	End Select
End Sub	

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp  공장 
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value   = arrRet(0)		
		frm1.txtPlantNm.value	= arrret(1)
		frm1.txtPlantCd.focus
	End If	
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Plant PopUp 품목 
'===================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)

	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	= arrRet(0)
		frm1.txtItemNm.Value	= arrRet(1)
		frm1.txtItemCd.focus
	End If
End Function

'------------------------------------------  OpenBP()  ---------------------------------------------
'	Name : OpenBP()
'	Description : SpplCd PopUp 공급처 
'---------------------------------------------------------------------------------------------------------
Function OpenBP()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_SpplCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "공급처"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "공급처"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "공급처"		
    arrHeader(1) = "공급처명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_SpplCd
		frm1.vspdData.Row = frm1.vspdData.ActiveRow

		frm1.vspdData.Text = arrRet(0)		
		frm1.vspdData.Col  = C_SpplNm
		frm1.vspdData.Text = arrret(1)
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	
		Call SpplChange()	
	End If		
End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 그리드의 숫자 부분이 변경된면 이 함수를 변경 해야함.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )

   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"   
        Case 7                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "7" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","99"  
    End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                
    Call ggoOper.LockField(Document, "N")              
    Call InitSpreadSheet("A")
    Call InitSpreadSheet("B")
'    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)                               
    Call SetDefaultVal
    Call InitVariables  
    Call ReadCookiePage()
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   

 	If Row <= 0 Then
 		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
    Else
		Call SetPopupMenuItemInf("0001111111")         '화면별 설정 
    End IF

	Set gActiveSpdSheet = frm1.vspdData
	    
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey1 = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey1 = 2
		Else
			ggoSpread.SSSort Col, lgSortKey1	'Sort in Descending
			lgSortkey1 = 1
		End If
	Else
 		lgSpdHdrClicked = 0		'2003-03-01 Release 추가 
 		Call Sub_vspdData_ScriptLeaveCell(0, 0, Col, frm1.vspdData.ActiveRow, False)
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If    			
End Sub

'========================================================================================
' Function Name : vspdData2_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
 	Dim strShowDataFirstRow
 	Dim strShowDataLastRow
 	Dim lngStartRow
 	Dim i,k
 	Dim strFlag,strFlag1
 	Dim iActiveRow
 	
 	gMouseClickStatus = "SP2C"   

 	Set gActiveSpdSheet = frm1.vspdData2
 	
 	If Row <= 0 Then
 		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
    Else
		Call SetPopupMenuItemInf("0001111111")         '화면별 설정 
    End IF
 	
 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 AND Col <> 0 Then	'2003-03-01 Release 추가 
 		ggoSpread.Source = frm1.vspdData2

 		frm1.vspdData.Row = frm1.vspdData.ActiveRow
 		frm1.vspdData.Col = frm1.vspdData.MaxCols

 		iActiveRow = CInt(frm1.vspdData.Text)
 		
 		frm1.vspdData2.Redraw = False
		lngStartRow = CInt(ShowFromData(iActiveRow, CInt(lglngHiddenRows(iActiveRow - 1))))
		frm1.vspdData2.Redraw = True
		
		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Ascending
 			lgSortKey2 = 2
 		ElseIf lgSortKey2 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Descending
 			lgSortKey2 = 1
		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
 	
 	With frm1.vspdData2
 		For i = 1 to .MaxRows
 			.Row = i
 			.col = 0	
 			If .Rowhidden = False Then
 				k = K + 1
 				if .text <> ggoSpread.UpdateFlag  then
 					.text = k
 				end if
 			End If
 		Next
 	End With 	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If y<20 Then			'2003-03-01 Release 추가 
	    lgSpdHdrClicked = 1 
	End If
	
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData2_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
   
   If Button = 2 And gMouseClickStatus = "SP2C" Then
      gMouseClickStatus = "SP2CR"
   End If
End Sub    

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData2_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

''========================================================================================
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
Sub PopRestoreSpreadColumnInf()	'###그리드 컨버전 주의부분###
	Dim lngRangeFrom
	Dim lngRangeTo	
	Dim lRow

    ggoSpread.Source = gActiveSpdSheet
    
    If gActiveSpdSheet.Name = "vspdData" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("A")
		Call ggoSpread.ReOrderingSpreadData
    ElseIf gActiveSpdSheet.Name = "vspdData2" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("B")
		frm1.vspdData2.Redraw = False
		
		Call ggoSpread.ReOrderingSpreadData("F")

		Call DbQuery2(frm1.vspdData.ActiveRow,False)
		
		lngRangeFrom = Clng(ShowDataFirstRow2)
		lngRangeTo = Clng(ShowDataLastRow2)
		
		lRow = frm1.vspdData.ActiveRow	'###그리드 컨버전 주의부분###
		frm1.vspdData2.Redraw = True
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo lRow
    End If
    
 	'------ Developer Coding part (Start)	
 	'------ Developer Coding part (End) 	
End Sub

'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)
	Dim strMark
	Dim iparentrow

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	
	With frm1.vspdData2
		.Row = Row
		.Col = C_ParentRowNo
		iparentrow = .text
		.Col = 0
		strMark = .Text
		.Col = C_RecordCnt 
		.Text = strMark
	
		Call QuotaRateChange(Row)   
	End With
	
	With frm1.vspdData
		If strMark = ggoSpread.UpdateFlag Then
			.Row = iparentrow
			.Col = 0
			.Text = ggoSpread.UpdateFlag
		End if
	End With
End Sub	
'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release 추가 
		Exit Sub
	End If
	
	Call Sub_vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)	
End Sub

'=======================================================================================================
'   Event Name : Sub_vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub Sub_vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	
	Dim lRow
	if Row = 0 then exit sub
	If Row <> NewRow And NewRow > 0 Then
		With frm1        
			If CheckRunningBizProcess = True Then
				Call SetActiveCell(frm1.vspdData,1,Row,"M","X","X")
				Exit Sub
			End If
			lgCurrRow = NewRow	
		End With
		
		With frm1.vspdData2
			.ReDraw = False
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.RowHidden = True
			.BlockMode = False
			.ReDraw = True
		End With
		If DbQuery2(lgCurrRow, False) = False Then	Exit Sub
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '/* 해상도에 상관없이 재쿼리되도록 수정 - START */
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '☜: 재쿼리 체크 
    '/* 해상도에 상관없이 재쿼리되도록 수정 - END */
		if Trim(lgPageNo) = "" then exit sub
		If lgPageNo > 0   Then            '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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

'=======================================================================================================
' Function Name : DefaultCheck
' Function Desc : 
'=======================================================================================================
Function DefaultCheck()
	DefaultCheck = False
	Dim i
	Dim j
	Dim RequiredColor 

	ggoSpread.Source = frm1.vspdData2
	RequiredColor = ggoSpread.RequiredColor
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				.Col = 0
				If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Then
					For j = 1 To .MaxCols
						.Col = j
						If .BackColor = RequiredColor Then
							If Len(Trim(.Text)) < 1 Then
								.Row = 0
								Call DisplayMsgBox("970021","X",.Text,"")
								Call SetActiveCell(frm1.vspdData2,j,i,"M","X","X")
								Exit Function
							End If
						End If			
					Next
				End If
			End If
		Next
	End With
	DefaultCheck = True
End Function

'==========================================   QuotaRateChange()  ======================================
'	Name : QuotaRateChange()
'	Description : 
'=================================================================================================
Sub QuotaRateChange(ByVal Row)
    Dim iparentrow
    Dim iReqQty,iApportionQty,iquotarate 
    Dim totalquotarate,totalApportionQty
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim index 

	with frm1.vspdData2
		.Row		= Row    
		.Col		= C_ParentRowNo
		iparentrow  = Trim(.text)
		
		lngRangeFrom = DataFirstRow(iparentrow)
	    lngRangeTo   = DataLastRow(iparentrow)
		
		totalquotarate = 0
		
		.Row		= Row    
		.Col		= 0
		
		for index = lngRangeFrom  to lngRangeTo
		    .Row = index
		    .Col = 0 
		    if Trim(.Text) <> ggoSpread.DeleteFlag  then
				.Col = C_Quota_Rate
				totalquotarate = totalquotarate + Unicdbl(.text)
		    end if
		next 
	End with
End Sub

'=======================================================================================================
' Function Name : ShowDataFirstRow
' Function Desc : 
'=======================================================================================================
Function ShowDataFirstRow()
	Dim i
	ShowDataFirstRow = 0
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataFirstRow2
' Function Desc : 
'=======================================================================================================
Function ShowDataFirstRow2()
	ShowDataFirstRow2 = 0
	Dim i
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow2 = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataLastRow
' Function Desc : 
'=======================================================================================================
Function ShowDataLastRow()
	Dim i
	ShowDataLastRow = 0
	
	With frm1.vspdData
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataLastRow2
' Function Desc : 
'=======================================================================================================
Function ShowDataLastRow2()
	ShowDataLastRow2 = 0
	Dim i
	
	With frm1.vspdData2
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow2 = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : DataFirstRow
' Function Desc : 
'=======================================================================================================
Function DataFirstRow(ByVal Row)
	Dim i
	DataFirstRow = 0
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = C_ParentRowNo
			If Clng(.text) = Clng(Row) Then
				DataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : DataLastRow
' Function Desc : 
'=======================================================================================================
Function DataLastRow(ByVal Row)
	Dim i
	DataLastRow = 0
	
	With frm1.vspdData2
		For i = .MaxRows To 1 Step -1
			.Row = i
			.Col = C_ParentRowNo
			If Clng(.text) = Clng(Row) Then
				DataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
'   Function Name : ShowFromData
'   Function Desc : 
'=======================================================================================================
Function ShowFromData(Byval Row, Byval lngShowingRows)	'###그리드 컨버전 주의부분###
'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 3을 리턴하는 기능을 수행하는 함수다.
	ShowFromData = 0
	
	Dim lngRow
	Dim lngStartRow
	
	With frm1.vspdData2
		
		Call SortSheet()
		'------------------------------------
		' Find First Row
		'------------------------------------ 
		lngStartRow = 0
		If .MaxRows < 1 Then Exit Function
		
		For lngRow = 1 To .MaxRows
			.Row = lngRow
			.Col = C_ParentRowNo
			If Row = CInt(.Text) Then
				lngStartRow = lngRow
				ShowFromData = lngRow
				Exit For
			End If    
		Next
		'------------------------------------
		' Show Data
		'------------------------------------ 
		
		If lngStartRow > 0 Then
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.Col = C_RecordCnt
			.Col2 = C_RecordCnt
			.DestCol = 0
			.DestRow = 1
			.Action = 19	'SS_ACTION_COPY_RANGE
			.RowHidden = False
			
			.BlockMode = False
			
			'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 첫번째 부터 2번째 까지의 Row를 숨긴다.
			If lngStartRow > 1 Then
				.BlockMode = True
				.Row = 1
				.Row2 = lngStartRow - 1
				.RowHidden = True
				.BlockMode = False
			End If

			'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 7번째 부터 마지막 까지의 Row를 숨긴다.
			If lngStartRow < .MaxRows Then
				If lngStartRow + lngShowingRows <= .MaxRows Then
					.BlockMode = True
					.Row = lngStartRow + lngShowingRows
					.Row2 = .MaxRows
					.RowHidden = True
					.BlockMode = False
				End If
			End If
			
			.BlockMode = False
			.Row = lngStartRow	'2003-03-01 Release 추가 
			.Col = 0			'2003-03-01 Release 추가 
			.Action = 0			'2003-03-01 Release 추가 
		End If
	End With	
End Function

'======================================================================================================
' Function Name : SortSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function SortSheet()
	SortSheet = false
    With frm1.vspdData2
        .BlockMode = True
        .Col = 0
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .SortBy = 0 'SS_SORT_BY_ROW

        .SortKey(1) = C_ParentRowNo
        .SortKey(2) = C_RecordCnt
        
        .SortKeyOrder(1) = 0 'SS_SORT_ORDER_ASCENDING
        .SortKeyOrder(2) = 0 'SS_SORT_ORDER_ASCENDING

        .Col = 1	'C_SupplierCd	'###그리드 컨버전 주의부분###
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .Action = 25 'SS_ACTION_SORT
        
        .BlockMode = False
    End With     

    SortSheet = true
End Function

'=======================================================================================================
' Function Name : ChangeCheck
' Function Desc : 
'=======================================================================================================
Function ChangeCheck()
	Dim i
	ChangeCheck = False
	
	ggoSpread.Source = frm1.vspdData2
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text = ggoSpread.UpdateFlag Then
				ChangeCheck = True
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : CheckDataExist
' Function Desc : 
'=======================================================================================================
Function CheckDataExist()
	Dim i
	CheckDataExist = False
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				CheckDataExist = True
				Exit Function
			End IF
		Next
	End With
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Err.Clear                                                          
    
    FncQuery = False                                                   
    
	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call InitVariables
    
    ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    												
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
	Set gActiveElement = document.activeElement
    FncQuery = True									
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    Err.Clear                                                           
    
    FncNew = False                                                      
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                              
    Call ggoOper.LockField(Document, "N")                               
    
    ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    Call SetDefaultVal
    Call InitVariables                                                  
	Set gActiveElement = document.activeElement
    FncNew = True                                                       
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Err.Clear         

    FncSave = False                                                         
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                                      
    
    ggoSpread.Source = frm1.vspdData
    
    If ChangeCheck = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           
        Exit Function
    End If
    
    If DefaultCheck = False Then
    	Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then	
		Exit Function
	End If			
	  
	Set gActiveElement = document.activeElement
    FncSave = True                                                       
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
	FncCancel = false
	Dim lRow
	Dim i,k,iCnt
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iActiveRow
	Dim iConvActiveRow
	Dim strFlag
	
	iActiveRow = frm1.vspdData.ActiveRow
	frm1.vspdData.Row = iActiveRow
	frm1.vspdData.Col = frm1.vspdData.MaxCols
	iConvActiveRow = frm1.vspdData.Text
	
	If frm1.vspdData.MaxRows < 1 then
	    FncCancel = True
		Exit function
	End If
	
	'Check Spread2 Data Exists for the keys
	If CheckDataExist = False Then
	    FncCancel = True
    	Exit function
    End If
	
	If gActiveSpdSheet.ID = "B" Then

		ggoSpread.Source = frm1.vspdData2	
		With frm1.vspdData2
			
			'범위가 보이지 않는 곳까지 넘어갔을 경우에 대한 처리 - START	    
		    lngRangeFrom = .SelBlockRow
		    .Row = lngRangeFrom

			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()
			
			.Redraw = False
			ggoSpread.EditUndo                                                 '☜: Protect system from crashing
			.Redraw = True

			iCnt=0
			For k=lngRangeFrom To lngRangeTo
				.Row=k
				.col=0
				if .text = ggoSpread.UpdateFlag then
					iCnt = iCnt + 1
				End if	
			Next
			
			If iCnt = 0 Then
				ggoSpread.Source = frm1.vspdData
				ggoSpread.EditUndo iActiveRow                                                '☜: Protect system from crashing
			End If	
		End With
	Else
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo                                                  '☜: Protect system from crashing

		ggoSpread.Source = frm1.vspdData2	
		With frm1.vspdData2
			'범위가 보이지 않는 곳까지 넘어갔을 경우에 대한 처리 - START	    
		    lngRangeFrom = .SelBlockRow
		    .Row = lngRangeFrom
			.Redraw = False
			
			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()
			
			iCnt=1
			For k=lngRangeFrom to lngRangeTo
				.Row=k
				ggoSpread.EditUndo k                                                 '☜: Protect system from crashing
			Next
			.Redraw = True
		End WIth	
	End If
	
	lRow = frm1.vspdData.ActiveRow
	If lngRangeTo = 0 Then
		lglngHiddenRows(lRow - 1) = 0
	Else
		lglngHiddenRows(lRow - 1) = lngRangeTo - lngRangeFrom  + 1
	End If
	'**********///// END
	'********** START
	If lglngHiddenRows(lRow - 1) = 0 Then
		frm1.cmdInsertSampleRows.Disabled = False
	End If
	
	k = 0 
	For i = lngRangeFrom To lngRangeTo
	    frm1.vspdData2.Row = i 
	    frm1.vspdData2.Col = 0
	    strFlag = Trim(frm1.vspdData2.Text)
	    If strFlag = ggoSpread.UpdateFlag Then 
	        k = 1
	        Exit For
	    End If
	next 
	
	Call vspdData2_Click(frm1.vspdData2.ActiveCol,frm1.vspdData2.ActiveRow)

	Set gActiveElement = document.activeElement
	FncCancel = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()                        
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
 	Call parent.FncExport(Parent.C_MULTI)		
	Set gActiveElement = document.activeElement
 	FncExcel = True
 End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
	Set gActiveElement = document.activeElement
    FncFind = True
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = False
	
	Dim IntRetCD
	
    If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	Set gActiveElement = document.activeElement
    FncExit = True    
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* %>
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    Err.Clear 

    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
    
    With frm1

    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
	    strVal = strVal & "&txtPlantCd=" & .hdnPlant.value            '공장 
	    strVal = strVal & "&txtItemCd=" & .hdnItem.value              '품목 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
    Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
    End If

	Call RunMyBizASP(MyBizASP, strVal)	
    
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(byVal intARow,byVal intTRow)							
	Dim i
	Dim lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows
	Dim ii
	
	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolbar("1110100100011111")				'버튼 툴바 제어 

	With frm1
		'-----------------------
		'Reset variables area
		'-----------------------
		lRow = .vspdData.MaxRows
		If lRow > 0 And intARow > 0 Then
			If intTRow<=0 Then 
				ReDim lgStrPrevKeyM(intARow)	
				ReDim lglngHiddenRows(intARow)			'lRow = .vspdData.MaxRows	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
			Else
				TmpArrPrevKey=lgStrPrevKeyM
				TmpArrHiddenRows=lglngHiddenRows

				ReDim lgStrPrevKeyM(intTRow+intARow)	
				ReDim lglngHiddenRows(intTRow+intARow)			'lRow = .vspdData.MaxRows	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
				For i = 0 To intTRow
					lgStrPrevKeyM(i) = TmpArrPrevKey(i)
					lglngHiddenRows(i) = TmpArrHiddenRows(i)
				Next 
			End If

			For i = intTRow To intTRow+intARow
				lglngHiddenRows(i) = 0
			Next 

			if lgIntFlgModeM = Parent.OPMD_CMODE then
			    If DbQuery2(1, false) = False Then	Exit Function
		    end if
	
		    lgIntFlgModeM = Parent.OPMD_UMODE
		End If
	End With
	
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
	DbQueryOk = true
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbQueryOk2(Byval DataCount)
	DbQueryOk2 = false
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim Index
	
	With frm1.vspdData2
		
		lngRangeFrom = ShowDataFirstRow2()
		lngRangeTo = ShowDataLastRow2()
		
		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = C_RecordCnt
		
		.Col2 = C_RecordCnt
		.DestCol = 0
		.DestRow = lngRangeFrom
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False
	End With
	
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement
	
	DbQueryOk2 = true
End Function

'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row, Byval NextQueryFlag)
	Dim strVal
	Dim lngRet
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim txtPlantCd, txtItemCd
	
	DbQuery2 = False

	'/* 9월 정기패치: 좌측 스프레드의 행간 이동 시 이미 조회된 자료나 입력된 자료를 읽어 들일 때에도 '' 창 띄우기 - START */
	Call LayerShowHide(1)
	
	With frm1
		.vspdData.Row = CInt(Row)
		.vspdData.Col = .vspdData.MaxCols
		Row = CInt(.vspdData.Text)	
		If lglngHiddenRows(Row - 1) <> 0 And NextQueryFlag = False Then
			.vspdData2.ReDraw = False
			 lngRet = ShowFromData(Row, lglngHiddenRows(Row - 1))	'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 3을 리턴하는 기능을 수행하는 함수다.
			
			Call SetToolbar("1110100100011111")				'버튼 툴바 제어 
			Call LayerShowHide(0)
			
			lngRangeFrom = ShowDataFirstRow
			lngRangeTo = ShowDataLastRow		
						
			.vspdData2.ReDraw = True
			DbQuery2 = True
			Exit Function
		End If

		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		.vspdData.Row = CInt(Row)
		.vspdData.Col = C_PlantCd		    
		strVal = strVal & "&txtPlantCd=" & Trim(.vspdData.text)
		.vspdData.Col = C_ItemCd		    
		strVal = strVal & "&txtItemCd=" & Trim(.vspdData.text)
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'☜: Next key tag 
		strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(Row - 1)
		strVal = strVal & "&lRow=" & CStr(Row)
	End With

	Call RunMyBizASP(MyBizASP, strVal)
	
	DbQuery2 = True
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow
	Dim lGrpCnt
	Dim strVal 
	Dim lngRangeFrom
    Dim lngRangeTo
    Dim parentRow
    Dim totalRate
	Dim Zsep
	Dim iColSep
	Dim iRowSep

	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size
	Dim ii
	
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep
	
	DbSave = False                                                          '⊙: Processing is NG
    
	Call LayerShowHide(1)

	frm1.txtMode.value = Parent.UID_M0002

	lGrpCnt = 1
	strVal = ""
    Zsep = "@"
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0

	'-----------------------
	'Data manipulate area
	'-----------------------
	With frm1
	    For parentRow = 1 To .vspdData.MaxRows
			If Trim(GetSpreadText(.vspdData,0,parentRow,"X","X")) = ggoSpread.UpdateFlag Then
				
			    lngRangeFrom = DataFirstRow(parentRow)
			    lngRangeTo   = DataLastRow(parentRow)
			   
			    totalRate = 0
			    for lRow = lngRangeFrom To lngRangeTo
					totalRate = totalRate + UNICDbl(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))
				Next
				
			    If UniCdbl(totalRate) <> 100  then       '같은 공장 같은 품목에 배분비율이 100인지 
				    Call DisplayMsgBox("171325", "X", Trim(GetSpreadText(.vspdData,C_ItemCd,parentRow,"X","X")) & "(" & parentRow & "Row)" , "X")
				    Call LayerShowHide(0)
				    Call RemovedivTextArea
				    Exit Function
				End if 
					
			    for lRow = lngRangeFrom To lngRangeTo
			        If Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) = ggoSpread.UpdateFlag Then

						strVal = strVal & "U" & iColSep		
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_ParentPlantCd,lRow,"X","X")) & iColSep
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_ParentItemCd,lRow,"X","X")) & iColSep
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_SpplCd,lRow,"X","X")) & iColSep
			
						If Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))="" Then
							strVal = strVal & "0" & iColSep
						Else
							strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X")),0) & iColSep
						End If
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_ParentRowNo,lRow,"X","X")) & iColSep
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_RecordCnt,lRow,"X","X")) & iColSep & iRowSep
							
						lGrpCnt = lGrpCnt + 1
				
					End If			
   			    Next
				
				strVal = strVal & Zsep
				Select Case Trim(GetSpreadText(.vspdData,0,parentRow,"X","X"))
				    Case ggoSpread.UpdateFlag
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
				End Select   
			End If
			strVal  = ""
		Next     
	End With
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			'☜: 비지니스 ASP 를 가동 

	DbSave = True	
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()								
	Call InitVariables
	
    lgIntFlgMode	 = Parent.OPMD_UMODE		
	lgBlnFlgChgValue = False
	Call MainQuery		    
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<!--########################################################################################################
'       					6. Tag부 
'######################################################################################################## 
-->
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공급처별배분비</font></td>
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
					<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
                            <TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU" ALT="공 장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X" ALT="공 장"></TD>
							    <TD CLASS="TD5" NOWRAP>품목</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 STYLE="text-transform:uppercase" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
												   <INPUT TYPE=TEXT ALT="품목" NAME="txtItemNm" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
							</TR>
						</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="A"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			
			<TR>
			 <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
    
			<TR HEIGHT= 40%>
			 <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			  <TABLE <%=LR_SPACE_TYPE_60%>>
			   <TR>
			    <TD HEIGHT=100% WIDTH=100% COLSPAN=4>
			     <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="B"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
			    </TD>
			   </TR>
			  </TABLE>
			 </TD>
			</TR>
		</TABLE></TD>
	</TR>
    <tr>
      <td <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SRC="m1211mb2.asp" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtColsep" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRowsep" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
