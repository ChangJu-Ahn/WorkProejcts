<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81102MA1
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

CONST BIZ_PGM_ID = "B81102MB1.ASP"
<!-- #Include file="../../inc/lgvariables.inc" -->	

'@Grid_Column
Dim C_ITEM_LVL_CD
Dim C_ITEM_LVL_NM
Dim C_Len
Dim C_Class_Cd
Dim C_Class_Name
Dim C_PARENT_CLASS_NM
Dim C_PARENT_CLASS_CD
Dim C_PARENT_CLASS_Popup
Dim C_REMARK
Dim C_CurrPopup

'@Global_Var
Dim lgSortKey1
Dim IsOpenPop
Dim lgitem_lvl
Dim EndDate, StartDate
   
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
	Call SetToolBar("111011010011111")				'버튼 툴바 제어 

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

	   .MaxCols = C_Remark + 1
	   .MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCombo  C_ITEM_LVL_CD,		"레벨코드", 10
		ggoSpread.SSSetCombo  C_ITEM_LVL_NM,		"레벨", 10
		ggoSpread.SSSetEdit   C_Len,				"길이", 8,2
		ggoSpread.SSSetEdit   C_Class_Cd,			"코드", 8,,,5,2
		ggoSpread.SSSetEdit   C_Class_Name,			"코드명", 15,,,40
		ggoSpread.SSSetEdit   C_PARENT_CLASS_CD,	"상위코드",10,,,10,2
		ggoSpread.SSSetButton C_PARENT_CLASS_Popup
		ggoSpread.SSSetEdit   C_PARENT_CLASS_NM,	"상위코드명", 15,,,100
		ggoSpread.SSSetEdit   C_REMARK,				"비고", 30,,,100
		
		Call ggoSpread.SSSetColHidden(C_ITEM_LVL_CD,C_ITEM_LVL_CD,True)
	
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
		ggoSpread.SpreadLock C_ITEM_LVL_NM, -1, C_ITEM_LVL_NM   
		ggoSpread.SpreadLock C_Class_Cd, -1, C_Class_Cd  
		ggoSpread.SpreadLock C_Len, -1, C_Len 
		ggoSpread.SSSetProtected C_PARENT_CLASS_CD, -1, C_PARENT_CLASS_Popup 
		ggoSpread.SpreadLock C_PARENT_CLASS_NM, -1, C_PARENT_CLASS_NM 
		ggoSpread.SSSetRequired  C_Class_Name, -1, C_Class_Name 
		
   		
		.ReDraw = True
    End With    
End Sub


Sub SetSpreadLock2()
	dim i
	for i=0 to frm1.vspdData.MaxRows
		
		  
			With frm1
			.vspdData.ReDraw = False
			
				if  Trim(GetSpreadText(frm1.vspdData,C_PARENT_CLASS_CD,i,"X","X")) ="*" then 
					ggoSpread.SSSetProtected		C_PARENT_CLASS_CD,		i,		i
				else
					ggoSpread.SSSetRequired		C_PARENT_CLASS_CD,		i,		i
				end if	
			.vspdData.ReDraw = True	
			End With
		
	next
	
	
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
		ggoSpread.SSSetRequired  C_ITEM_LVL_NM,			pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected C_Len,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired  C_Class_Cd,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired C_Class_Name,			pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired C_PARENT_CLASS_CD,			pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected C_PARENT_CLASS_NM,			pvStartRow,	pvEndRow
		
				
    
	    .ReDraw = True
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()

	C_ITEM_LVL_CD		=	1
	C_ITEM_LVL_NM		=	2
	C_Len				=	3
	C_Class_Cd			=	4	
	C_Class_Name		=	5
	C_PARENT_CLASS_CD	=	6
	C_PARENT_CLASS_Popup=   7
	C_PARENT_CLASS_NM	=	8
	C_REMARK			=	9
	
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
			C_ITEM_LVL_CD			=	iCurColumnPos(1)
			C_ITEM_LVL_NM			=	iCurColumnPos(2)
			C_Len					=	iCurColumnPos(3)
			C_Class_Cd				=	iCurColumnPos(4)	
			C_Class_Name			=	iCurColumnPos(5)
			C_PARENT_CLASS_CD		=	iCurColumnPos(6)
			C_PARENT_CLASS_Popup	=	iCurColumnPos(7)
			C_PARENT_CLASS_NM		=	iCurColumnPos(8)
			C_REMARK				=	iCurColumnPos(9)
			
	End Select    
End Sub

'------------------------------------------  OpenParent_cd()  -------------------------------------------------
' Name : OpenParent_cd()
' Description : SpreadItem PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenParent_cd(prow,colPos)
	Dim arrRet,strTemp_lvl
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim cd_temp
	dIM item_acct,item_kind

	item_acct = filtervar(frm1.txtItem_acct.value ,"''","S")
	item_kind = filtervar(frm1.txtitem_kind.value  ,"''","S")


   
   if GetSpreadtext(frm1.vspdData,2,cint(prow),"X","X")="중분류" THEN 
		strTemp_lvl="L1"
   elseif GetSpreadtext(frm1.vspdData,2,cint(prow),"X","X")="소분류" THEN 
		strTemp_lvl="L2"
   else 
	exit function 		
   end if
   
   
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "상위코드 팝업"						<%' 팝업 명칭 %>
	arrParam(1) = "B_CIS_ITEM_CLASS"						<%' TABLE 명칭 %>
	arrParam(2) = ""	<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = " ITEM_ACCT = "&item_acct&" AND ITEM_KIND="&item_kind&" AND ITEM_LVL='"&strTemp_lvl&"'"	<%' Where Condition%>
	arrParam(5) = "상위코드"					<%' 조건필드의 라벨 명칭 %>

    arrField(0) = "CLASS_CD"						<%' Field명(0)%>
    arrField(1) = "CLASS_NAME"					<%' Field명(1)%>
    
    arrHeader(0) = "상위코드"							<%' Header명(0)%>
    arrHeader(1) = "상위코드명"							<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.vspdData.Col = colPos-1
			.vspdData.Text = arrRet(0)
			.vspdData.Col = colPos+1
			.vspdData.Text = arrRet(1)
		end with
		ggoSpread.UpdateRow prow
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
	call CookiePageLoad("FORM_LOAD")
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
	dim intIndex
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)        '  <------변경된 표준 라인 
	
	Select Case Col
		Case C_ITEM_LVL_NM
			SetSpreadValue frm1.vspdData,Col+1,Row,lgItem_lvl(Trim(GetSpreadvalue(frm1.vspdData,Col,Row,"X","X"))),"",""
			SetSpreadValue frm1.vspdData,C_CLASS_CD,Row,"","",""
			if getSpreadtext (frm1.vspdData,C_ITEM_LVL_CD,Row,"X","X")="L1" then 
			   	SetSpreadValue frm1.vspdData,C_PARENT_CLASS_CD,Row,"*","",""
				ggoSpread.SSSetProtected C_PARENT_CLASS_CD, Row,Row
				
			else 
				SetSpreadValue frm1.vspdData,C_PARENT_CLASS_CD,Row,"","",""	
				ggoSpread.SSSetRequired C_PARENT_CLASS_CD, Row,Row
				
			end if
			 frm1.vspdData.Col = Col
		        intIndex = frm1.vspdData.Value 
				frm1.vspdData.Col = C_ITEM_LVL_CD
				frm1.vspdData.Value = intIndex
				
			
			ggoSpread.SSSetEdit   C_Class_Cd, "코드", 10,,Row,lgItem_lvl(Trim(GetSpreadvalue(frm1.vspdData,Col,Row,"X","X"))),2
		Case C_PARENT_CLASS_CD
			
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
	Dim strTemp_lvl
	Dim intPos1

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_PARENT_CLASS_Popup Then
		
			.Col = C_PARENT_CLASS_Popup
			.Row = Row
			
			Call OpenParent_cd(row,.Col)
		
	
		End if 
	End With
	
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
     ggoSpread.SetCombo "L1" & vbtab & "L2" & vbtab & "L3" , C_ITEM_LVL_CD
     ggoSpread.SetCombo "대분류" & vbtab & "중분류" & vbtab & "소분류" , C_ITEM_LVL_NM
 
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
		    Case C_ITEM_LVL_NM
		        .Col = Col
		        intIndex = .Value 
				.Col = C_ITEM_LVL_CD
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
    
  
	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData
    Call InitVariables															'Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
   

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

   	Call SetToolbar("11000111001111")

    '-----------------------
    'Save function call area
    '-----------------------
    
    If CheckLength = False then Exit Function
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
		 if  lvlSet=false then exit function
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor .ActiveRow, .ActiveRow
		.Row = .ActiveRow
		.Col = C_CLASS_CD
		.Text = ""
		.Col = C_CLASS_NAME
		.Text = ""
	
		if  Trim(GetSpreadText(frm1.vspdData,C_PARENT_CLASS_CD,.ActiveRow,"X","X")) ="*" then 
			ggoSpread.SSSetProtected		C_PARENT_CLASS_CD,		.ActiveRow,		.ActiveRow
		else
			ggoSpread.SSSetRequired		C_PARENT_CLASS_CD,		.ActiveRow,		.ActiveRow
		end if	
				
		.ReDraw = True
	End With
End Function
'========================================================================================
' Function Name : lvlSet
' Function Desc : 각 항목의 자리수를 배열에 담아놓기 
'========================================================================================
function CheckLength()
	Dim i
	CheckLength=false
	With frm1
	
	ggoSpread.source = frm1.vspdData

      For i = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = i
		    .vspdData.Col =0

		    Select Case .vspdData.Text 
		        Case ggoSpread.InsertFlag		'☜: 신규, 수정 
					if cint(GetSpreadText(.vspdData,C_LEN,i,"X","X"))<> cInt(len(GetSpreadText(.vspdData,C_CLASS_CD,i,"X","X"))) then
						Call DisplayMsgBox("970029","X","자릿수","X")
						Call SetToolBar("111011110011111")
						
						.vspdData.focus()	
						.vspdData.Action=0
						exit Function
					end if
		    end select 
	next 
	
	End With
	
		    
	CheckLength=true
End function

'========================================================================================
' Function Name : lvlSet
' Function Desc : 각 항목의 자리수를 배열에 담아놓기 
'========================================================================================
function lvlSet()

	ReDim lgitem_lvl(3)
	
	lvlSet=false
	
	if CommonQueryRs (" ITEM_LVL1,ITEM_LVL2,ITEM_LVL3"," B_CIS_CONFIG ", " ITEM_ACCT = " & FilterVar(frm1.txtItem_acct.Value, "''", "S") & " AND ITEM_KIND=" & FilterVar(frm1.txtItem_kind.Value, "''", "S"), _
	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6 ) =false then
	
	
			Call DisplayMsgBox("800489","X","품목코드구성설정","")
			frm1.txtItem_acct.focus
			Set gActiveElement = document.activeElement
			
			Exit function
	else
		
		lgitem_lvl(0) = split(lgF0,chr(11))(0)
		lgitem_lvl(1) = split(lgF1,chr(11))(0)
		lgitem_lvl(2) = split(lgF2,chr(11))(0)
		
	end if
		
	lvlSet=true
	
End function
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
	Dim imRow
	
	'On Error Resume Next
	
	FncInsertRow = False
	
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
   
   if  lvlSet=false then exit function
	
   
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
	

	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtItem_acct=" & .txtItem_acct.value
	    strVal = strVal & "&txtItem_kind=" & .txtItem_kind.value
	    strVal = strVal & "&cboItem_lvl=" & .cboItem_lvl.value
	    strVal = strVal & "&lgPageNo="	 & lgPageNo						'☜: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
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
	Call SetToolBar("111011110011111")				'버튼 툴바 제어 

    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		//frm1.txtItem_acct.focus 
	End If
	'call SetSpreadLock2
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
Function DbSave() 
    Dim lRow
	Dim lGrpCnt     
	Dim strVal,strDel
	Dim ColSep, RowSep
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
					if cint(GetSpreadText(.vspdData,C_LEN,lRow,"X","X"))<> cInt(len(GetSpreadText(.vspdData,C_CLASS_CD,lRow,"X","X"))) then
						Call DisplayMsgBox("970029","X","자릿수","X")
						.vspdData.Row=lRow
						.vspdData.col=C_CLASS_CD
						.action=0
						exit Function
					end if
					
					strVal = strVal & .txtItem_acct.value  & ColSep
					strVal = strVal & .txtItem_kind.value  & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_ITEM_LVL_CD,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Class_Cd,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Class_Name,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Parent_class_cd,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Parent_class_nm,lRow,"X","X")) & ColSep
					
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_REMARK,lRow,"X","X")) & ColSep  & lRow & ColSep & RowSep
					lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag							'☜: 삭제 
					
					strDel = strDel & .txtItem_acct.value  & ColSep
					strDel = strDel & .txtItem_kind.value  & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_ITEM_LVL_CD,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_Class_Cd,lRow,"X","X")) & ColSep 
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_Class_Name,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_Parent_class_cd,lRow,"X","X")) & ColSep & lRow & ColSep & RowSep
					
					
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

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow)
	With frm1.vspdData
		.Row = lRow
		.Col = C_Class_Cd
		.Action = 0
		Call SetFocusToDocument("M") 
		.focus
	End With
End Function


Sub CookiePageLoad(ByVal ChkVal)
'	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp	
	
	if	ReadCookie("ITEM_ACCT")="" then
	else
		if ChkVal="FORM_LOAD" then

			frm1.txtItem_acct.Value = ReadCookie("ITEM_ACCT")
			frm1.txtItem_kind.Value = ReadCookie("ITEM_KIND")
		
			Call WriteCookie("ITEM_ACCT","")
		    Call WriteCookie("ITEM_KIND","")
		
			Call FncQuery()
		end if


	end if
End Sub

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
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목계정" NAME="txtItem_acct" SIZE=10 MAXLENGTH=4  tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK='vbscript:OpenPopupw "item_acct","txtItem_acct"'>&nbsp;<INPUT TYPE=TEXT NAME="txtItem_acct_nm" SIZE=20 MAXLENGTH=20 ALT="품목계정" tag="14X">
							<TD CLASS="TD5" NOWRAP>품목구분</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목구분" NAME="txtItem_kind" SIZE=10 MAXLENGTH=18 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK='vbscript:OpenPopupw "item_kind","txtItem_kind"'>&nbsp;<INPUT TYPE=TEXT ALT="품목구분" NAME="txtitem_kind_Nm" SIZE=20 tag="14X"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>레벨</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
							<SELECT NAME="cboItem_lvl" tag="11X" STYLE="WIDTH: 87px;">
							
							<OPTION value="*"></OPTION>
							<OPTION value="L1">대분류</OPTION>
							<OPTION value="L2">중분류</OPTION>
							<OPTION value="L3">소분류</OPTION>
							</SELECT>
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
									<script language =javascript src='./js/b81102ma1_OBJECT1_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hdnAppFrDt"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnAppToDt"  tag="24" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>
 
