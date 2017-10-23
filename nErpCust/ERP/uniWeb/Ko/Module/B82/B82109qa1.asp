<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B82109MQ1
'*  4. Program Name         : 픔명/규격 변경의뢰조회
'*  5. Program Desc         : 픔명/규격 변경의뢰조회
'*  6. Component List       : 
'*  7. Modified date(First) :  2005/01/30
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../B81/B81COMM.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit											

CONST BIZ_PGM_ID = "B82109QB1.ASP"

<!-- #Include file="../../inc/lgvariables.inc" -->	

'@Grid_Column
Dim C_REQ_NO
Dim C_REQ_GBN

Dim C_ITEM_CD
Dim C_ITEM_NM
Dim C_SPEC_NAME
Dim C_INSERT_USER_ID
Dim C_REMARK
Dim C_CurrPopup

Dim C_REQ_DT
DIM C_REQ_ID
DIM C_STATUS	
DIM C_ITEM_R
DIM C_ITEM_T
DIM C_ITEM_P
DIM C_ITEM_Q
DIM C_TRANS_DT
DIM C_END_DT

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
	Call SetToolBar("111011010011111")				'버튼 툴바 제어 
	
	
	//frm1.txtFromEndDt.text = StartDate	
	//frm1.txtToEndDt.text = EndDate
	frm1.txtFromReqDt.text = StartDate	
	frm1.txtToReqDt.text = EndDate
	
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
        ggoSpread.Spreadinit "V20050201",, parent.gAllowDragDropSpread

	   .MaxCols = C_TRANS_DT + 1
	   .MaxRows = 0
		'Call GetSpreadColumnPos("A")


		ggoSpread.SSSetEdit C_REQ_NO, "의뢰번호", 10  
		ggoSpread.SSSetEdit C_REQ_ID, "의뢰자", 10  
		ggoSpread.SSSetEdit C_REQ_DT, "의뢰일자", 10  
		ggoSpread.SSSetEdit C_STATUS, "STATUS", 10  
		ggoSpread.SSSetEdit C_ITEM_CD, "품목코드", 10  
		ggoSpread.SSSetEdit C_ITEM_NM, "품목명", 15    
		ggoSpread.SSSetEdit C_SPEC_NAME, "규격", 14   
		ggoSpread.SSSetEdit C_ITEM_R, "접수", 8,2
		ggoSpread.SSSetEdit C_ITEM_T, "기술", 8,2
		ggoSpread.SSSetEdit C_ITEM_P, "구매", 8,2
		ggoSpread.SSSetEdit C_ITEM_Q, "품질", 8 ,2
		ggoSpread.SSSetEdit C_END_DT, "완료일자", 12
		ggoSpread.SSSetEdit C_TRANS_DT, "이관일자", 12
		//ggoSpread.SSSetEdit   C_REMARK, "비고", 20,,,,2 
	
		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)	
		
		.ReDraw = True
    End With

    Call SetSpreadLock()
End Sub


'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()

	C_REQ_NO     =   1
	C_REQ_ID     =   2
    C_REQ_DT	 =   3
    C_STATUS	 =   4
	C_ITEM_CD    =   5
	C_ITEM_NM    =   6
	C_SPEC_NAME	 =	 7
	C_ITEM_R     =   8
	C_ITEM_T     =   9
	C_ITEM_P     =   10
	C_ITEM_Q     =   11
	C_END_DT     =   12
	C_TRANS_DT   =   13
	C_REMARK	 =	 14

End Sub



'---------------------------------------------------------------------------------------------------------
'	Name : SetSpreadColor()
'	Description : SetSpreadColor
'---------------------------------------------------------------------------------------------------------
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	call SetSpreadLock()
End Sub

Sub SetSpreadLock()
     With frm1.vspdData
		.ReDraw = False
    
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLock		-1,			-1
   		
		.ReDraw = True
    End With    
    //SetSpreadColor frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow 
  ggoSpread.SpreadLockWithOddEvenRowColor()
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
	Call SetToolbar("11000000000111")
	  frm1.txtreq_user.focus()
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
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
 	
 	gMouseClickStatus = "SPC" 
 	  
 	lgAcct_item_cd=GetSpreadText(frm1.vspdData,C_REQ_NO,Row,"X","X")
 	lgAcct_kind_cd=GetSpreadText(frm1.vspdData,C_ITEM_CD,Row,"X","X")
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
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
	
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
    //Call GetSpreadColumnPos("A")
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
Sub txtFromReqDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtFromReqDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtFromReqDt.Focus
	End If
End Sub

Sub txtToReqDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtToReqDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtToReqDt.Focus
	End If
End Sub

Sub txtFromReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtToReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub



Sub txtFromEndDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtFromEndDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtFromEndDt.Focus
	End If
End Sub

Sub txtToEndDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtToEndDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtToEndDt.Focus
	End If
End Sub

Sub txtFromEndDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtToEndDt_KeyDown(KeyCode, Shift)
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
    If ValidDateCheck(frm1.txtFromReqDt, frm1.txttoReqDt)	=	False	Then Exit	Function
    IF ValidDateCheck(frm1.txtFromEndDt, frm1.txtTOEndDt)	=	False	Then Exit	Function
  
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
		//.Col = C_INSERT_USER_ID
		.Text = ""
	
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
		    strVal = strVal & "&txtItem_kind=" & .hdnitemcd.value
		    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else	
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&insrt_user_id=" & .txtreq_user.value 
		    strVal = strVal & "&lgPageNo="	 & lgPageNo						'☜: Next key tag 
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		    
		End If 
   
	End With
	Call ExecMyBizASP(frm1, strVal)									'☜: 비지니스 ASP 를 가동 
	

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
	Call SetToolbar("11000000000111")				'버튼 툴바 제어 

    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		
	End If
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
                                                          
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow)
	With frm1.vspdData
		.Row = lRow
		.Col = C_INSERT_USER_ID
		.Action = 0
		Call SetFocusToDocument("M") 
		.focus
	End With
End Function


Function Jump(strPgmID)	

    Dim iRet
    Dim iRet2
    Dim strVal
    Dim iArr
    Dim eisWindow
    On Error Resume Next
   
	CookiePage("")
    PgmJump(strPgmID)
  
End Function

Sub CookiePage(ByVal ChkVal)
	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp

	If frm1.vspdData.ActiveRow > 0 Then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_REQ_NO
		WriteCookie CookieSplit , frm1.vspdData.Text
	Else
		WriteCookie CookieSplit , ""
	End If
	
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
							<TD CLASS="TD5" NOWRAP>의뢰일자</TD>
							<TD CLASS="TD6" NOWRAP >
							<script language =javascript src='./js/b82109qa1_OBJECT2_txtFromReqDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/b82109qa1_OBJECT3_txtToReqDt.js'></script>
							<TD CLASS="TD5" NOWRAP>STATUS</TD>
							<TD CLASS="TD6" NOWRAP>
							<INPUT TYPE="RADIO" NAME="rbo_status" ID="rbo_status1" VALUE="'*'" CLASS="RADIO" TAG="11" CHECKED><LABEL FOR="rbo_status1">전체</LABEL>&nbsp;
                            <INPUT TYPE="RADIO" NAME="rbo_status" ID="rbo_status2" VALUE="'R','A','D'" CLASS="RADIO" TAG="11"><LABEL FOR="rbo_status2">진행중</LABEL>&nbsp;
                            <INPUT TYPE="RADIO" NAME="rbo_status" ID="rbo_status3" VALUE="'E','S','T'" CLASS="RADIO" TAG="11"><LABEL FOR="rbo_status3">완료</LABEL></td>
							
                            
							</TD>
						
						</TR>
						
						<TR>
							<TD CLASS="TD5" NOWRAP>의뢰자</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="의뢰자" NAME="txtreq_user" SIZE=10 MAXLENGTH=10  tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK='vbscript:OpenPopupw "req_user","txtreq_user"'>&nbsp;<INPUT TYPE=TEXT NAME="txtreq_user_nm" SIZE=20 MAXLENGTH=20 ALT="의뢰자" tag="14X"></td>
							<TD CLASS="TD5" NOWRAP>완료기간</TD>
							<TD CLASS="TD6" NOWRAP >
							<script language =javascript src='./js/b82109qa1_OBJECT2_txtFromEndDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/b82109qa1_OBJECT3_txtToEndDt.js'></script>
							</TD>
						</TR>
					
						
						<TR>
							<TD CLASS="TD5" NOWRAP>품목코드</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목코드" NAME="txtItem_cd" SIZE=18 MAXLENGTH=18  tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK='vbscript:OpenPopupw"item_cd","txtItem_cd"'>&nbsp;<INPUT TYPE=TEXT NAME="txtItem_cd_nm" SIZE=20 MAXLENGTH=20 ALT="품목코드" tag="14X"></td>
							
							        
							<TD CLASS="TD5" NOWRAP>규격</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="규격" NAME="txtItem_spec" SIZE=35 MAXLENGTH=18 tag="11NXX"></TD>
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
									<script language =javascript src='./js/b82109qa1_OBJECT1_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>					
					<TD></TD>
					
					<TD WIDTH=* ALIGN=RIGHT>
					
					&nbsp;<A href='vbscript:Jump("B82107MA1")'>품명/규격변경의뢰등록</a>
					&nbsp;<A href='vbscript:Jump("B82108MA1")'>품명/규격변경의뢰승인</a>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
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
 
