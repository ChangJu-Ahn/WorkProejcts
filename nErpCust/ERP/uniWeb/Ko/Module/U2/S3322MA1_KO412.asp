<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : S3322MA1_KO412
'*  4. Program Name         : 품의서문서관리(S)
'*  5. Program Desc         : 품의서문서관리(S)
'*  6. Component List       : 
'*  7. Modified date(First) : 2007/07/04
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee wol san
'* 10. Modifier (Last)      : Lee Ho Jun
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit											

CONST BIZ_PGM_ID = "S3322MB1_KO412.ASP"
Const BIZ_PGM_REG_ID = "S3322MA1_KO412"
<!-- #Include file="../../inc/lgvariables.inc" -->	

'@Grid_Column
Dim C_PROJECT_CODE		'프로젝트코드
Dim C_REPORT_NO			
Dim C_REPORT_NM
Dim C_INS_USER
Dim C_INS_DT
Dim C_REPORT_ABBR

Dim C_CurrPopup
Dim C_CostDt
Dim C_SupplierCd
Dim C_SupplierPopup
Dim C_SupplierNm
Dim C_Cost

'@Global_Var
Dim lgSortKey1
Dim IsOpenPop
Dim lgitem_lvl
Dim EndDate, StartDate
Dim lgAcct_item_cd
Dim lgAcct_kind_cd
   
EndDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------

StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)

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
	
	Call SetToolBar("110000010011111")				'버튼 툴바 제어 

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
        ggoSpread.Spreadinit "V20050103",, parent.gAllowDragDropSpread

	   .MaxCols = C_REPORT_ABBR + 1
	   .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit		C_PROJECT_CODE,		"프로젝트번호", 10
		ggoSpread.SSSetCombo	C_REPORT_NO,		"품의제목",		10
		ggoSpread.SSSetCombo	C_REPORT_NM,		"품의제목",		15
		ggoSpread.SSSetEdit		C_INS_USER,			"등록자", 12
		ggoSpread.SSSetEdit		C_INS_DT,			"등록일", 12
		ggoSpread.SSSetEdit		C_REPORT_ABBR,		"요약설명", 30,,,100
		
		Call ggoSpread.MakePairsColumn(C_REPORT_NO, C_REPORT_NM, "1")

		Call ggoSpread.SSSetColHidden(C_PROJECT_CODE,C_PROJECT_CODE,True)
		Call ggoSpread.SSSetColHidden(C_REPORT_NO,C_REPORT_NO,True)
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
		ggoSpread.SpreadLock		-1,			-1
   		
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
		ggoSpread.SSSetRequired  C_REPORT_ABBR,			pvStartRow,	pvEndRow

	    .ReDraw = True
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	
	C_PROJECT_CODE  =   1
	C_REPORT_NO		=	2
	C_REPORT_NM		=	3
	C_INS_USER      =   4
	C_INS_DT		=   5
	C_REPORT_ABBR	=	6
	
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
			
			C_PROJECT_CODE	=	iCurColumnPos(1)
			C_REPORT_NO		=	iCurColumnPos(2)
			C_REPORT_NM		=	iCurColumnPos(3)
			C_INS_USER		=	iCurColumnPos(4)
			C_INS_DT		=	iCurColumnPos(5)
			C_REPORT_ABBR	=	iCurColumnPos(6)	

	End Select    
End Sub


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
	Call InitSpreadSheet
	Call InitComboBox
	Call SetToolbar("11000000000111")
	'Call dbQuery()
	
End Sub

'==========================================  2.2.6 InitComboBox()  ========================================
' Name : InitComboBox()
' Desc : Combo Display
'==========================================================================================================
Sub InitComboBox()

    Dim strCboCd
    Dim strCboNm

	'// 구분
	Call CommonQueryRs(" UD_MINOR_CD,UD_MINOR_NM "," B_USER_DEFINED_MINOR ", " UD_MAJOR_CD = " & FilterVar("SX006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_REPORT_NO
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_REPORT_NM
   
	
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

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_REPORT_NO         ' 시스템구분
			intIndex = .value
			.col = C_REPORT_NM
			.value = intindex					
		Next	
	End With

End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###
	
	Dim sProjectCode,sReportNo
	Dim strval
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
 	
 	gMouseClickStatus = "SPC" 
 	  
 	sProjectCode = frm1.txtProjectCode.value
 	
' 	frm1.vspddata.Row = Row
' 	frm1.vspddata.col = C_REPORT_NO
' 	sReportNo = frm1.vspddata.text
' 	
' 	with frm1.vspddata
' 		.Row = Row
' 		.Col = C_REPORT_NO
' 		sReportNo = .text 
' 	End With

 	sReportNo = GetSpreadText(frm1.vspdData,C_REPORT_NO,Row,"X","X")
 	
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

 	with frm1
 	    strVal = BIZ_PGM_ID & "?txtMode=view"
	    strVal = strVal & "&project_code=" & sProjectCode
	    strVal = strVal & "&report_no=" & sReportNo
	    strVal = strVal & "&lgPageNo="	 & lgPageNo						'☜: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	End with 

	  MyBizASP1.location.href = "S3322RA1_KO412.asp?project_code=" & sProjectCode & "&Report_No=" & sReportNo

 	//Call RunMyBizASP(MyBizASP, strVal)	//zerry 잘안됨..수정할것.	

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
        Case C_Cost
            Call EditModeCheck(frm1.vspdData, Row, C_Curr, C_Cost,    "C" ,"I", Mode, "X", "X")
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
    Call InitData()
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================
Sub txtProjectCode_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
  
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

	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData
    Call InitVariables															'Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
'    If ValidDateCheck(frm1.txtFromInsrtDt, frm1.txtToInsrtDt)	=	False	Then Exit	Function
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
   	
'    If CompareDateByFormat(frm1.txtFromInsrtDt.text,frm1.txtToInsrtDt.text,frm1.txtFromInsrtDt.Alt,frm1.txtToInsrtDt.Alt, _
'        	               "970024",frm1.txtFromInsrtDt.UserDefinedFormat,parent.gComDateType, true) = False Then
'	   frm1.txtFromInsrtDt.focus
'	   Exit Function
'	End If
	
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
    With frm1.vspdData
		If .maxrows < 1 then exit function
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor .ActiveRow, .ActiveRow
		.Row = .ActiveRow
		.Col = C_REPORT_ABBR
		.Text = ""
	
		.ReDraw = True
	End With
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

    Dim iDx

	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
     Frm1.vspdData.Row = frm1.vspdData.ActiveRow
     
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If
    
    Call Initdata()

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	
	'On Error Resume Next
	
	FncInsertRow = False
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
   
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
		.ReDraw = True
    End With

    Set gActiveElement = document.ActiveElement
    
    If Err.number = 0 Then FncInsertRow = True
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
' Function Name : FncDelete
' Function Desc : 
'========================================================================================
Function FncDelete()	
	
	If frm1.vspdData.maxRows >= 1 then
		If DisplayMsgBox("210034", parent.VB_YES_NO, "x", "x") = vbYes Then '삭제하시겠습니까?
		 
		End If   
	End If

    'MyBizASP.location.href = "S3322MA1_KO412_frwriteBiz.asp?txtMode=" & UID_M0003  & "&txtKeyNo=" & MyBizAsp.frTitle.intKeyNo
    //MyBizASPForDelete.location.href = "S3322MA1_KO412_frwriteBiz.asp?txtMode=" & UID_M0003  & "&txtKeyNo=" & MyBizAsp.frTitle.intKeyNo
	    
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
	
	//MyBizASP1.location.href = "../../blank.htm"
	If lgIntFlgMode = Parent.OPMD_UMODE Then
	 	
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtProjectCode=" & .hdnProjectCode.value
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else	
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtProjectCode=" & .txtProjectCode.value
'	    strVal = strVal & "&txtTitle=" & .txtTitle.value
'       strVal = strVal & "&txtToUseDt=" & .txtToUseDt.text
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
	Call SetToolBar("110000000001111")				'버튼 툴바 제어 

    If frm1.vspdData.MaxRows > 0 Then
		Call SetToolBar("110010110001111")
		frm1.vspddata.focus
		MyBizASP1.location.href = "S3322RA1_KO412.asp?project_code=" & ""
	End If
	
	Call InitData()
	Set gActiveElement = document.activeElement
	call vspdData_click (1,1)
	frm1.txtProjectCode.focus()
	
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow
	Dim lGrpCnt     
	Dim strVal,strDel
	Dim ColSep, RowSep
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
	
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_PROJECT_CODE,lRow,"X","X"))  & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_REPORT_NO,lRow,"X","X"))  & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_REPORT_ABBR,lRow,"X","X")) & RowSep
					lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag							'☜: 삭제 
					
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_PROJECT_CODE,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_REPORT_NO,lRow,"X","X")) & RowSep

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


'=======================================================================================================
' Function Name : FncWrite
' Function Desc : 
'========================================================================================================
Function FncWrite()

		Dim arrRet,parm
		
		If UCase(Trim(frm1.txtProjectCode.value)) = "" Then
			Call DisplayMsgBox("900002", "x", "x", "x")		 '⊙: "Will you destory previous data"
			Exit Function
		End If
		
		If IsOpenPop = True Then Exit Function
         reDim parm(3)
		IsOpenPop = True

		arrRet = window.showModalDialog ("S3322PA1_KO412.asp?strMode=" & parent.UID_M0001 & "&project_code=" & frm1.txtProjectCode.value,Array(window.parent,parm(0),parm(1)), _
       "dialogWidth=600px; dialogHeight=470px; center: Yes; help: No; resizable: No; status: No;")			//"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	

		If arrRet = True Then
			call dbsaveOK()
			MyBizASP1.location.reload						
		End If
				
		IsOpenPop = False

	End Function

'=======================================================================================================
' Function Name : FncModify
' Function Desc : 
'========================================================================================================
Function FncModify()

	Dim arrRet,sReport_no
		
	If IsOpenPop = True Then Exit Function
	
	If UCase(Trim(frm1.txtProjectCode.value)) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x")		 '⊙: "Will you destory previous data"
		Exit Function
	End If
	
	if frm1.vspdData.maxRows = 0 then
		 call DisplayMsgBox("900025", "X", "X", "X") 
		 Exit Function
	end if
		
	sReport_no=GetSpreadText(frm1.vspdData,2,frm1.vspdData.ActiveRow,"X","X")
		
	IsOpenPop = True
		
	arrRet = window.showModalDialog ("S3322PA1_KO412.asp?strMode=" & parent.UID_M0002 & "&project_code=" & frm1.txtProjectCode.value  & "&Report_no=" & sReport_no,Array(window.parent,sReport_no), _
	"dialogWidth=600px; dialogHeight=470px; center: Yes; help: No; resizable: No; status: No;")	

	If arrRet = True Then
		call dbsaveOK()
		MyBizASP1.location.reload
	End If
		
	IsOpenPop = False

End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow)
	With frm1.vspdData
		.Row = lRow
		.Col = C_REPORT_ABBR
		.Action = 0
		Call SetFocusToDocument("M") 
		.focus
	End With
End Function


Function Jump()	

    Dim iRet
    Dim iRet2
    Dim strVal
    Dim iArr
    Dim eisWindow ,strPgmID
    'On Error Resume Next
   
	CookiePage("")
    PgmJump(BIZ_PGM_REG_ID)
  
End Function

'------------------------------------------  OpenProject()  -------------------------------------------------
'	Name : OpenProject()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProject()
	
	OpenProject = False
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtProjectCode.readOnly = True Then
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = "프로젝트팝업"
	arrParam(1) = "PMS_PROJECT"
	arrParam(2) = Trim(frm1.txtProjectCode.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "프로젝트"

	arrField(0) = "PROJECT_CODE"
	arrField(1) = "PROJECT_NM"

	arrHeader(0) = "프로젝트코드"
	arrHeader(1) = "프로젝트명"

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtProjectCode.Focus
	If arrRet(0) <> "" Then
		frm1.txtProjectCode.Value = arrRet(0)
		frm1.txtProjectNm.Value   = arrRet(1)
		frm1.txtProjectCode.Focus
	End If

	Set gActiveElement = document.activeElement
	OpenProject = True
	
End Function

'=================================================================================================
'   Event Name :vspddata_ComboSelChange
'   Event Desc :Combo Change Event
'==================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row

		Select Case Col

			'// 시스템 구분
			Case C_REPORT_NO
				.Col = Col
				intIndex = .Value
				.Col = C_REPORT_NO
				.Value = intIndex

			Case C_REPORT_NM
				.Col = Col
				intIndex = .Value
				.Col = C_REPORT_NM
				.Value = intIndex

		End Select
    End With
End Sub

'=====================================================================================================
'   Event Name : txtProjectCode_OnChange
'   Event Desc :
'=====================================================================================================
Sub txtProjectCode_OnChange()

	Call CommonQueryRs(" PROJECT_NM "," pms_project ", " PROJECT_CODE = " & FilterVar(frm1.txtProjectCode.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	frm1.txtProjectNm.value = Replace(Trim(lgF0), Chr(11), "")
	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%> >
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
				
					<TD WIDTH=* Align=right>
					<A onclick="vbscript:FncWrite()">등록</A>&nbsp;|&nbsp;<A onclick="vbscript:FncModify()">수정</A>
					<A onclick="vbscript:FncDelete()"></A></TD>
					<TD WIDTH=10>&nbsp;</TD>
					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=55%>
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
							<TD CLASS="TD5" NOWRAP>프로젝트번호</TD>
        					<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtProjectCode" SIZE="18" MAXLENGTH="25" ALT="프로젝트번호" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenProject()"  OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
        											 <INPUT TYPE=TEXT NAME="txtProjectNm" SIZE="26" MAXLENGTH=120 tag="14">
							</TD>
							
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		<TD <%=HEIGHT_TYPE_02%> ></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
			<IFRAME NAME="MyBizASP1" SRC="S3322RA1_KO412.asp" WIDTH=100% HEIGHT=90% FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
			<IFRAME NAME="MyBizASPForDelete" SRC="../../blank.htm" WIDTH=10% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
		</TD>
	</TR>	
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnProjectCode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hCurrency" tag="24">
</FORM>
</BODY>
</HTML>
 
