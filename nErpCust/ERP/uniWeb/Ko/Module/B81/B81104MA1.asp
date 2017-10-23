<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81104MA1
'*  4. Program Name         : 품목구성코드등록
'*  5. Program Desc         : 품목구성코드등록
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee wol san
'* 10. Modifier (Last)      : 
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

<!-- #Include file="../../inc/incSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" --> 

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "B81comm.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">
Option Explicit											

CONST BIZ_PGM_ID = "B81104MB1.ASP"
<!-- #Include file="../../inc/lgvariables.inc" -->	

'@Grid_Column
Dim C_ITEM_ACCT
Dim C_ITEM_KIND
Dim C_C_CODE
Dim C_SPEC_ORDER
Dim C_SPEC_NAME
Dim C_SPEC_UNIT
Dim C_LENGTH
Dim C_EXAM
Dim C_REMARK
Dim C_HSPEC_ORDER

'@Global_Var
Dim lgSortKey1
Dim IsOpenPop
Dim lgitem_lvl
Dim EndDate, StartDate
Dim lgChk

EndDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gDateFormat)

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
	lgPageNo = ""
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
	Call SetToolBar("110011010011111")				'버튼 툴바 제어 
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		.ReDraw = false

		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20050301",, parent.gAllowDragDropSpread

	   .MaxCols = C_Remark + 2
	   .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
        ggoSpread.SSSetEdit C_ITEM_ACCT,		"", 10
        ggoSpread.SSSetEdit C_ITEM_KIND,		"", 10
		ggoSpread.SSSetEdit C_C_CODE,			"Category", 18
		ggoSpread.SSSetEdit C_SPEC_ORDER,		"순번", 7
		ggoSpread.SSSetEdit   C_SPEC_NAME,		"특성명", 10,,,40,2
		ggoSpread.SSSetEdit   C_SPEC_UNIT,		"특성단위", 10,,,5,2
		
		ggoSpread.SSSetFloat  C_LENGTH,		"자릿수", 9, "6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","10"
		ggoSpread.SSSetEdit   C_EXAM,			"예제", 20
		ggoSpread.SSSetEdit   C_REMARK,			"비고", 30,,,,2
		ggoSpread.SSSetEdit   C_HSPEC_ORDER,	"", 3,,,,2
		
		Call ggoSpread.SSSetColHidden(C_ITEM_ACCT,C_ITEM_KIND,True)
		Call ggoSpread.SSSetColHidden(C_C_CODE,C_C_CODE,True)
		Call ggoSpread.SSSetColHidden(.MaxCols-1,	.MaxCols-1,	True)	
		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)	
		
		.ReDraw = True
    End With

    Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
		.ReDraw = False
    
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.SpreadLock C_C_CODE, -1, C_C_CODE
		ggoSpread.SpreadLock C_SPEC_ORDER, -1, C_SPEC_ORDER   
		ggoSpread.SSSetRequired C_SPEC_UNIT, -1, C_SPEC_UNIT  
		ggoSpread.SSSetRequired C_SPEC_NAME, -1, C_SPEC_NAME  
		ggoSpread.SSSetRequired  C_LENGTH, -1, C_LENGTH 
		//ggoSpread.SpreadLock		-1,			-1
   		
		.ReDraw = True
    End With    
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData
		.ReDraw = False

  		ggoSpread.Source = frm1.vspdData
  
   		ggoSpread.SpreadUnLock		1, pvStartRow, ,pvEndRow
   		ggoSpread.SSSetProtected  C_C_CODE,			    pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected  C_SPEC_ORDER,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired   C_SPEC_NAME,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired   C_SPEC_UNIT,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired   C_LENGTH,			    pvStartRow,	pvEndRow
		
				
    
	    .ReDraw = True
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_ITEM_ACCT		= 1
	C_ITEM_KIND		= 2
	C_C_CODE		= 3
	C_SPEC_ORDER	= 4
	C_SPEC_NAME		=	5
	C_SPEC_UNIT		=	6	
	C_LENGTH		=	7
	C_EXAM			=	8
	C_REMARK		=	9
	C_HSPEC_ORDER   = 10
	
	
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
			
			C_ITEM_ACCT			=	iCurColumnPos(1)
			C_ITEM_KIND			=	iCurColumnPos(2)
			C_C_CODE			=	iCurColumnPos(3)
			C_SPEC_ORDER		=	iCurColumnPos(4)
			C_SPEC_NAME			=	iCurColumnPos(5)
			C_SPEC_UNIT			=	iCurColumnPos(6)	
			C_LENGTH			=	iCurColumnPos(7)
			C_EXAM				=	iCurColumnPos(8)
			C_REMARK			=	iCurColumnPos(9)
			
	End Select    
End Sub

'------------------------------------------  OpenItem_lvl()  ------------------------------------------------
'	Name : OpenItem_lvl()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItem_lvl(pItem_lvl)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	Dim sItem_acct,sItem_kind, sAddCondi
	dim sTitle1,sTitle2
    
   if frm1.txtItem_acct.value ="" then
     Call DisplayMsgBox("800489","X","품목계정","X")
	 frm1.txtItem_acct.focus()
	 Exit Function
	end if
   if frm1.txtItem_kind.value ="" then
	 Call DisplayMsgBox("800489","X","품목구분","X")
	 frm1.txtItem_kind.focus()
	 Exit Function
   End if 	 
 
	Select Case pItem_lvl
		Case "L1"
			arrParam(0) = frm1.txtItem_lvl1.alt &" POPUP"
			sTitle1="대분류"
			sTitle2="대분류명"
				
		Case "L2"
			sTitle1="중분류"
			sTitle2="중분류명"
			
			if frm1.txtItem_lvl1.value ="" then
				 Call DisplayMsgBox("800489","X",frm1.txtItem_lvl1.alt,"X")
				 frm1.txtItem_lvl1.focus()
				 Exit Function
			End if 	
			sAddCondi=" AND PARENT_CLASS_CD=" & filtervar(frm1.txtItem_lvl1.value,"''","S")
			arrParam(0) = frm1.txtItem_lvl2.alt &" POPUP"
		Case "L3"
			sTitle1="소분류"
			sTitle2="소분류명"
			if frm1.txtItem_lvl1.value ="" then
				 Call DisplayMsgBox("800489","X",frm1.txtItem_lvl1.alt,"X")
				 frm1.txtItem_lvl1.focus()
				 Exit Function
			End if 
			if frm1.txtItem_lvl2.value ="" then
				 Call DisplayMsgBox("800489","X",frm1.txtItem_lvl2.alt,"X")
				 frm1.txtItem_lvl2.focus()
				 Exit Function
			End if 		
			sAddCondi=" AND PARENT_CLASS_CD=" & filtervar(frm1.txtItem_lvl2.value,"''","S")
			arrParam(0) = frm1.txtItem_lvl3.alt &" POPUP"
	End Select
	
	sItem_acct = frm1.txtItem_acct.value
	sItem_kind = frm1.txtItem_kind.value
	
	If IsOpenPop = True Then Exit Function		 
	IsOpenPop = True

	
	arrParam(1) = "B_CIS_ITEM_CLASS"						<%' TABLE 명칭 %>

	arrParam(2) = eval("frm1.txtItem_LV"&pItem_lvl).value 	<%' Code Condition%>
	arrParam(4) = " ITEM_ACCT =N'"&sItem_acct&"'  AND ITEM_KIND=N'"&sItem_kind&"' AND ITEM_LVL='"&pItem_lvl&"' " &sAddCondi 	<%' Where Condition%>
	arrParam(5) = sTitle1					<%' 조건필드의 라벨 명칭 %>
	arrParam(3) = ""								<%' Name Cindition%>
	
    arrField(0) = "CLASS_CD"						<%' Field명(0)%>
    arrField(1) = "CLASS_NAME"					<%' Field명(1)%>
    
    arrHeader(0) = sTitle1						<%' Header명(0)%>
    arrHeader(1) = sTitle2							<%' Header명(1)%>

	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtItem_acct.focus
		Exit Function
	Else
	
	Select Case pItem_lvl
		Case "L1"
			frm1.txtItem_lvl1.focus()
			frm1.txtItem_lvl1.Value= arrRet(0)  
			frm1.txtItem_lvl1_nm.Value= arrRet(1)
		Case "L2"
			 frm1.txtItem_lvl2.focus()
			frm1.txtItem_lvl2.Value= arrRet(0)  
			frm1.txtItem_lvl2_nm.Value= arrRet(1)
		Case "L3"
		    frm1.txtItem_lvl3.focus()
			frm1.txtItem_lvl3.Value= arrRet(0)  
			frm1.txtItem_lvl3_nm.Value= arrRet(1)
	End Select
	
	
		
		
		Set gActiveElement = document.activeElement
	End If  
End Function



'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'==========================================================================================================
Sub Form_Load()	'###그리드 컨버전 주의부분###
	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet                                                        'Setup the Spread sheet1
	call InitComboBox2
	
	frm1.txtItem_acct.focus()
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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row


	frm1.vspdData.Row = Row
	frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)        '  <------변경된 표준 라인 
	
	Select Case Col
		Case C_SPEC_ORDER
			SetSpreadValue frm1.vspdData,Col+1,Row,lgItem_lvl(Trim(GetSpreadvalue(frm1.vspdData,Col,Row,"X","X"))),"",""
			ggoSpread.SSSetEdit   C_SPEC_UNIT, "코드", 10,,Row,lgItem_lvl(Trim(GetSpreadvalue(frm1.vspdData,Col,Row,"X","X"))),2
		
	End Select
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
 	
 	gMouseClickStatus = "SPC"   
	 	 	
 	Set gActiveSpdSheet = frm1.vspdData
 	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 			
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		Exit Sub
 	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     

    If OldLeft <> NewLeft Then Exit Sub
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '☜: 재쿼리 체크 
		If Trim(lgPageNo) = "" Then Exit Sub
		If lgPageNo > 0 Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End If
End Sub

'========================================================================================
' Function Name : vspdData_ButtonClicked
' Function Desc : 팝업버튼 선택시 
'========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
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


'========================================================================================================
'   Event Name : vspdData_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
     
    End Select
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
  
End Sub 

'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================
Sub txtAppFrDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtAppFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtAppFrDt.Focus
	End If
End Sub

Sub txtAppToDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtAppToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtAppToDt.Focus
	End If
End Sub

Sub txtAppFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtAppToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub InitSpreadComboBox()
    Dim strCboData 
    Dim strCboData2
    
    strCboData = ""
    strCboData2 = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0005", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	
    strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboData, C_DimensionCd
    ggoSpread.SetCombo strCboData2, C_Dimension
End Sub

Sub InitComboBox2()
    Dim    iCodeArr
    Dim    iNameArr
     ggoSpread.SetCombo "L1" & vbtab & "L2" & vbtab & "L3" , C_C_CODE
     ggoSpread.SetCombo "대분류" & vbtab & "중분류" & vbtab & "소분류" , C_SPEC_ORDER
 
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex
 
    With frm1.vspdData
		.Row = Row
		Select Case Col
		    Case C_SPEC_ORDER
		        .Col = Col
		        intIndex = .Value 
				.Col = C_C_CODE
				.Value = intIndex
		
		End Select
    End With

     ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() '###그리드 컨버전 주의부분###
    Dim IntRetCD     
    FncQuery = False                                                        

    
    
	If ggoSpread.SSCheckChange = True Then 'lgBlnFlgChgValue = True Or lgBtnClkFlg = True Or 
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")		'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '품목만으로 조회시 메시지 호출(2003.09.04)
   // If  Trim(frm1.txtItem_acct.Value) = "" and Trim(frm1.txtItem_kind.Value) <> "" then
	//	Call DisplayMsgBox("17A002", "X", "공장", "X")
	//	frm1.txtItem_acct.focus
	//	Exit Function
	//End if
	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData
    Call InitVariables															'Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
  	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If

   	//Call SetToolbar("10000000000111")
   
   // If Check_Input = False Then 
   // 	Call SetToolBar("110011010011111")				'버튼 툴바 제어 
	//    Exit Function
	//End If

	If DbQuery = False then	Exit Function
		      
    FncQuery = True	
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    Err.Clear                                                               '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData	   
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    
    FncNew = True  
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    FncSave = False                                                         
    
    If frm1.vspdData.maxrows < 1 then exit function    

	'-----------------------
    'Precheck area
    '-----------------------
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           
        Exit Function
    End If
    
    '----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then Exit Function

   	Call SetToolbar("10000000000111")

    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then Exit Function
	  
    FncSave = True                                                       
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
  Dim sSpec_order 
  sSpec_order= getSpecOrder((frm1.hmaxspec_order_nm.value ) ) 
    With frm1.vspdData
		If .maxrows < 1 then exit function
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor .ActiveRow, .ActiveRow
		.Row = .ActiveRow
		.Col = C_SPEC_ORDER
		.Text =sSpec_order
	
		.ReDraw = True
	End With
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	frm1.vspdData.Redraw = False
    If frm1.vspdData.maxrows < 1 then exit function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo 
 	frm1.vspdData.Redraw = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim imRow,sCategory
	Dim sSpec_order
	Dim item_acct
	Dim item_kind 
	dim lvl1_cd,lvl2_cd,lvl3_cd


	On Error Resume Next
	

	FncInsertRow = False
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
 

	item_acct= FilterVar(frm1.txtItem_acct.Value, "''", "S")
	item_kind= FilterVar(frm1.txtItem_kind.Value, "''", "S")
	lvl1_cd =FilterVar(frm1.txtItem_lvl1.Value, "''", "S") 
	lvl2_cd =FilterVar(frm1.txtItem_lvl2.Value, "''", "S") 
	lvl3_cd =FilterVar(frm1.txtItem_lvl3.Value, "''", "S") 

    if frm1.hmaxspec_order_nm.value="" then
		if CommonQueryRs (" isnull(MAX(SPEC_ORDER),'') "," b_cis_item_class_category ", _
			" ITEM_ACCT = " & item_acct & " AND ITEM_KIND=" & item_kind & " AND ITEM_LVL1_CD = " &lvl1_cd & " AND ITEM_LVL2_CD = " &lvl2_cd & " AND ITEM_LVL3_CD = " &lvl3_cd , _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6 ) =false then
	
		end if	
		frm1.hmaxspec_order_nm.value = lgF0
	end if	
   
	 sCategory= frm1.txtItem_lvl1.value  & frm1.txtItem_lvl2.value & frm1.txtItem_lvl3.value
	 sSpec_order=getSpecOrder(frm1.hmaxspec_order_nm.value )
	
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End If 
    
    With frm1.vspdData	

		.ReDraw = False
		.focus
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow .ActiveRow, imRow
		SetSpreadColor .ActiveRow, .ActiveRow + imRow - 1
		SetSpreadValue frm1.vspdData,1,.ActiveRow,frm1.txtItem_acct.value,"",""
		SetSpreadValue frm1.vspdData,2,.ActiveRow,frm1.txtItem_kind.value,"",""
		SetSpreadValue frm1.vspdData,3,.ActiveRow,sCategory,"",""
		SetSpreadValue frm1.vspdData,4,.ActiveRow,sSpec_order,"",""
		SetSpreadValue frm1.vspdData,.MaxCols -1 ,1,sSpec_order,"",""
		//GetSpreadvalue(frm1.vspdData,.MaxCols,.ActiveRow,"X","X")
		.ReDraw = True
    End With
    

    Set gActiveElement = document.ActiveElement
    
    If Err.number = 0 Then FncInsertRow = True
End Function

Function getSpecOrder(pOrder)
	dim temp
	if pOrder="" or pOrder=null then
		getSpecOrder="01"
	else 
		pOrder = cInt(pOrder) + 1 
		
		if len(pOrder)=1 then
			getSpecOrder="0"&pOrder
		else
			
			getSpecOrder =pOrder 
		end if
		
	end if
	frm1.hmaxspec_order_nm.value =getSpecOrder
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows 
	Dim lTempRows 

	If frm1.vspdData.maxrows < 1 then exit function
	
 '----------  Coding part  ------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData	
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 	Call parent.FncExport(Parent.C_MULTI)		
 End Function
 
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()

	FncExit = False
	
	Dim IntRetCD
	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If
    
    FncExit = True    
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
	Dim strVal

	DbQuery = False                                                             

	Call LayerShowHide(1)
	
	With frm1
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
	 	
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtItem_acct=" & .txtItem_acct.value
	    strVal = strVal & "&txtItem_kind=" & .txtItem_kind.value
	    strVal = strVal & "&txtItem_lvl1=" & .txtItem_lvl1.value
	    strVal = strVal & "&txtItem_lvl2=" & .txtItem_lvl2.value
	    strVal = strVal & "&txtItem_lvl3=" & .txtItem_lvl3.value
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else	
    
    
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtItem_acct=" & .txtItem_acct.value
	    strVal = strVal & "&txtItem_kind=" & .txtItem_kind.value
	    strVal = strVal & "&txtItem_lvl1=" & .txtItem_lvl1.value
	    strVal = strVal & "&txtItem_lvl2=" & .txtItem_lvl2.value
	    strVal = strVal & "&txtItem_lvl3=" & .txtItem_lvl3.value
	    strVal = strVal & "&lgPageNo="	 & lgPageNo						'☜: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If 
    
	End With

	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 
	

	DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbQueryOk()
	Dim ii
	
    lgIntFlgMode = Parent.OPMD_UMODE							'⊙: Indicates that current mode is Update mode
	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("110011110011111")				'버튼 툴바 제어 

    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		 
	End If
	lgChk=true
	Set gActiveElement = document.activeElement
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow
	Dim lGrpCnt     
	Dim strVal,strDel
	Dim ColSep, RowSep


	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
	Dim ii
	
	ColSep = parent.gColSep               
	RowSep = parent.gRowSep               
	
    DbSave = False                                                          '⊙: Processing is NG
	Call LayerShowHide(1)
	
	frm1.txtMode.value = Parent.UID_M0002
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 0
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferMaxCount = -1 
	iTmpDBufferMaxCount = -1 
	    
	
	
	With frm1
		.txtMode.value = parent.UID_M0002
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	ggoSpread.source = frm1.vspdData

      For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		'☜: C=Create
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		'☜: U=Update
				Case ggoSpread.DeleteFlag
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep		'☜: U=Delete
						
			End Select			
 
		    Select Case .vspdData.Text 
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'☜: 신규, 수정 
		      
					
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_ITEM_ACCT,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_ITEM_KIND,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(frm1.txtItem_lvl1.value) & ColSep
					strVal = strVal & Trim(frm1.txtItem_lvl2.value) & ColSep
					strVal = strVal & Trim(frm1.txtItem_lvl3.value) & ColSep
	
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_SPEC_ORDER,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_SPEC_NAME,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_SPEC_UNIT,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_LENGTH,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_EXAM,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_REMARK,lRow,"X","X")) & ColSep  & lRow & ColSep & RowSep
					lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag							'☜: 삭제 
					
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_ITEM_ACCT,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_ITEM_KIND,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(frm1.txtItem_lvl1.value) & ColSep
					strDel = strDel & Trim(frm1.txtItem_lvl2.value) & ColSep
					strDel = strDel & Trim(frm1.txtItem_lvl3.value) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_SPEC_ORDER,lRow,"X","X")) & ColSep & lRow & ColSep & RowSep
					
  		            lGrpCnt = lGrpCnt + 1
		    End Select
		 
		Next
		
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	End With


	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			'☜: 비지니스 ASP 를 가동 

	DbSave = True                                                      
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
End Function

'========================================================================================
' Function Name : Check_Input
' Function Desc : 
'========================================================================================
Function Check_Input()
	Check_Input = False
	frm1.txtItem_acct_nm.Value = ""
	frm1.txtitem_kind_nm.Value = ""
	frm1.txtSupplierNm.Value = ""

	If Trim(frm1.txtItem_acct.Value) <> "" And Trim(frm1.txtItem_kind.Value) <> "" Then

		If 	CommonQueryRs(" B.PLANT_NM, C.ITEM_NM, C.PHANTOM_FLG "," B_ITEM_BY_PLANT A, B_PLANT B, B_ITEM C ", _
		                " A.PLANT_CD = B.PLANT_CD AND A.ITEM_CD = C.ITEM_CD AND A.ITEM_CD = " & FilterVar(frm1.txtItem_kind.Value, "''", "S") & " AND A.PLANT_CD = " & FilterVar(frm1.txtItem_acct.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtItem_acct.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("125000","X","X","X")
				frm1.txtItem_acct_nm.Value = ""
				frm1.txtItem_acct.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItem_acct_nm.Value = lgF0(0)

			If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItem_kind.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("122600","X","X","X")
				frm1.txtitem_kind_nm.Value = ""
				frm1.txtItem_kind.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtitem_kind_nm.Value = lgF0(0)

			Call DisplayMsgBox("122700","X","X","X")
			frm1.txtItem_kind.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		frm1.txtItem_acct_nm.Value = lgF0(0)
		frm1.txtitem_kind_nm.Value = lgF1(0)
		
		If Trim(lgF2(0)) <> "N" Then
			Call DisplayMsgBox("181315","X","X","X")
			frm1.txtItem_kind.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
	
	ElseIf Trim(frm1.txtItem_acct.Value) <> "" Then
		'-----------------------
		'Check Plant CODE		'공장코드가 있는 지 체크 
		'-----------------------
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtItem_acct.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtItem_acct_nm.Value = ""
			frm1.txtItem_acct.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtItem_acct_nm.Value = lgF0(0)
	
	ElseIf Trim(frm1.txtItem_kind.Value) <> "" Then
		'-----------------------
		'Check Item CODE	 '품목코드가 있는 지 체크  
		'-----------------------
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItem_kind.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("122600","X","X","X")
			frm1.txtitem_kind_nm.Value = ""
			frm1.txtItem_kind.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtitem_kind_nm.Value = lgF0(0)
	End If

	If Trim(frm1.txtSuppliercd.Value) <> "" Then
		'-----------------------
		'Check BPt CODE		'공급처코드가 있는 지 체크 
		'-----------------------
		If 	CommonQueryRs(" BP_NM, BP_TYPE, usage_flag "," B_Biz_Partner ", " BP_CD = " & FilterVar(frm1.txtSuppliercd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("229927","X","X","X")
			frm1.txtSupplierNm.Value = ""
			frm1.txtSuppliercd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		frm1.txtSupplierNm.Value = lgF0(0)

		If Trim(lgF2(0)) <> "Y" Then
			Call DisplayMsgBox("179021","X","X","X")
			frm1.txtSuppliercd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		If Trim(lgF1(0)) <> "S" And Trim(lgF1(0)) <> "CS" Then
			Call DisplayMsgBox("179020","X","X","X")
			frm1.txtSuppliercd.focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
	End If
	
    //If frm1.txtAppFrDt.text <> "" And frm1.txtAppToDt.text <> "" Then
	//	If ValidDateCheck(frm1.txtAppFrDt, frm1.txtAppToDt) = False Then 
   	//		frm1.txtAppToDt.focus 
	//		Set gActiveElement = document.activeElement
	//		Exit Function
	//	End If
	//End If
	
	Check_Input = True
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow)
	With frm1.vspdData
		.Row = lRow
		.Col = C_SPEC_UNIT
		.Action = 0
		Call SetFocusToDocument("M") 
		.focus
	End With
End Function


Sub CookiePageLoad(ByVal ChkVal)
'	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp	

	if ChkVal="FORM_LOAD" then

			frm1.txtItem_acct.Value = ReadCookie("ITEM_ACCT")
			frm1.txtItem_kind.Value = ReadCookie("ITEM_KIND")

			Call FncQuery()


	else
	
	end if

End Sub


'========================================================================================
' Function Name : txtitem_acct_cd_OnChange
' Function Desc : 
'========================================================================================
Function txtitem_acct_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtitem_acct.value = "" Then
        frm1.txtitem_acct_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd='P1001' and minor_cd="&filterVar(frm1.txtitem_acct.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtitem_acct_nm.value=""
        Else
            frm1.txtitem_acct_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
     call txtItem_kind_OnChange()
    call txtItem_lvl1_OnChange()
    call txtItem_lvl2_OnChange()
    call txtItem_lvl3_OnChange()
End Function


'========================================================================================
' Function Name : txtItem_kind_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_kind_OnChange()
    Dim iDx
    Dim IntRetCd
 

    If frm1.txtItem_kind.value = "" Then
        frm1.txtItem_kind_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm ","  B_MINOR A, B_CIS_CONFIG B "," major_cd='Y1001' AND A.MINOR_CD = B.ITEM_KIND AND B.ITEM_ACCT = "&filtervar(frm1.txtitem_acct.value,"''","S")&" and minor_cd="&filterVar(frm1.txtItem_kind.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_kind_nm.value=""
        Else
            frm1.txtItem_kind_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
   
    call txtItem_lvl1_OnChange()
    call txtItem_lvl2_OnChange()
    call txtItem_lvl3_OnChange()
    
End Function
'========================================================================================
' Function Name : txtItem_lvl1_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_lvl1_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtItem_lvl1.value = "" Then
        frm1.txtItem_lvl1_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" CLASS_NAME "," B_CIS_ITEM_CLASS "," ITEM_ACCT="&filterVar(frm1.txtitem_acct.value,"''","S") & " AND ITEM_KIND="&filterVar(frm1.txtitem_kind.value,"''","S") & " AND CLASS_CD="&filterVar(frm1.txtItem_lvl1.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_lvl1_nm.value=""
        Else
            frm1.txtItem_lvl1_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
      call txtItem_lvl2_OnChange()
      //call txtItem_lvl3_OnChange()
End Function

'========================================================================================
' Function Name : txtItem_lvl2_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_lvl2_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtItem_lvl2.value = "" Then
        frm1.txtItem_lvl2_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" CLASS_NAME "," B_CIS_ITEM_CLASS "," ITEM_ACCT="&filterVar(frm1.txtitem_acct.value,"''","S") & " AND ITEM_KIND="&filterVar(frm1.txtitem_kind.value,"''","S") & "  AND PARENT_CLASS_CD="&filterVar(frm1.txtItem_lvl1.value,"''","S") & " AND CLASS_CD="&filterVar(frm1.txtItem_lvl2.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_lvl2_nm.value=""
        Else
            frm1.txtItem_lvl2_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
     call txtItem_lvl3_OnChange()
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 																		#
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
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strAspMnuMnunm")%></font></td>
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
					<FIELDSET CLASS="CLSFLD" >
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>품목계정</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목계정" NAME="txtItem_acct" SIZE=10 MAXLENGTH=5  tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK='vbscript:OpenPopupw "item_acct","txtItem_acct"'>&nbsp;<INPUT TYPE=TEXT NAME="txtItem_acct_nm" SIZE=20 MAXLENGTH=20 ALT="품목계정" tag="14X">
							<TD CLASS="TD5" NOWRAP>품목구분</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목구분" NAME="txtItem_kind" SIZE=10 MAXLENGTH=18 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK='vbscript:OpenPopupw "item_kind","txtItem_kind"'>&nbsp;<INPUT TYPE=TEXT ALT="품목구분" NAME="txtItem_kind_Nm" SIZE=20 tag="14X"></TD>
							
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>대분류</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="대분류" readOnly NAME="txtItem_lvl1" SIZE=10 MAXLENGTH=8  tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem_lvl('L1')">&nbsp;<INPUT TYPE=TEXT NAME="txtItem_lvl1_nm" SIZE=20 MAXLENGTH=20 ALT="대분류" tag="14X">
					
							<TD CLASS="TD5" NOWRAP>중분류</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="중분류" readOnly NAME="txtItem_lvl2" SIZE=10 MAXLENGTH=8 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem_lvl('L2')">&nbsp;<INPUT TYPE=TEXT ALT="중분류" NAME="txtItem_lvl2_nm" SIZE=20 tag="14X"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>소분류</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="소분류" readOnly NAME="txtItem_lvl3" SIZE=10 MAXLENGTH=8  tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem_lvl('L3')">&nbsp;<INPUT TYPE=TEXT NAME="txtItem_lvl3_nm" SIZE=20 MAXLENGTH=20 ALT="소분류" tag="14X">
							
							<TD CLASS="TD5" NOWRAP></TD>
							<TD CLASS="TD6" NOWRAP></TD>
							
						</TR>
						
					</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/b81104ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> ></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1">
			</IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnItem_acct"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnitem_kind"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnItem_lvl"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hmaxspec_order_nm"  tag="24" TABINDEX="-1">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>
 
