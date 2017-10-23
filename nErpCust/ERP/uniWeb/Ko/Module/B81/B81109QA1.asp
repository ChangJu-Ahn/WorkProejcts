<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81108MQ1
'*  4. Program Name         : 통합코드조회
'*  5. Program Desc         : 통합코드조회
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/30
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

CONST BIZ_PGM_ID = "B81109QB1.ASP"


<!-- #Include file="../../inc/lgvariables.inc" -->	

'@Grid_Column

Dim C_REQ_GBN
Dim C_ITEM_CD
Dim C_REQ_REASON
Dim C_REQ_ID
Dim C_REQ_DT
dIM C_SEQ

'@Global_Var
Dim lgSortKey1
Dim IsOpenPop
Dim lgitem_lvl
Dim EndDate, StartDate
Dim lgAcct_item_cd
   
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

	   .MaxCols = C_REQ_REASON + 1
	   .MaxRows = 0
		'Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit C_ITEM_CD,		"", 10  '1
		ggoSpread.SSSetEdit C_SEQ,			"순번", 10  '1
		ggoSpread.SSSetEdit C_REQ_ID,		"변경의뢰자", 15  '2
		ggoSpread.SSSetEdit C_REQ_DT,		"의뢰일자", 18    '3
		ggoSpread.SSSetEdit C_REQ_REASON,	"변경의뢰사유", 40    '4
		
		Call ggoSpread.SSSetColHidden(C_ITEM_CD,	C_ITEM_CD,	True)
		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)	
		
		.ReDraw = True
    End With

    Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================	
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
'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()

	C_ITEM_CD       =   1
	C_SEQ			=   2
	C_REQ_ID		=   3
	C_REQ_DT			=	4
	C_REQ_REASON		=	5
	
	
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
	Call CookiePage(0)
	
	Call SetToolbar("11000000000111")
	frm1.txtItem_cd.focus()
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
	Dim seq_no,strVal
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
 	
 	gMouseClickStatus = "SPC" 
 	 
 	lgAcct_item_cd=GetSpreadText(frm1.vspdData,1,Row,"X","X")
 	seq_no=GetSpreadText(frm1.vspdData,2,Row,"X","X")
 	
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

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex
 
    With frm1.vspdData
		.Row = Row
		Select Case Col
		    //Case C_ITEM_LVL_NM
		    //    .Col = Col
		     //   intIndex = .Value 
			//	.Col = C_REQ_REASON
			//	.Value = intIndex
		
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
   

   	//Call SetToolbar("10000000000111")
   
   // If Check_Input = False Then 
   // 	Call SetToolBar("111011010011111")				'버튼 툴바 제어 
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
   	
    If CompareDateByFormat(frm1.txtFromReqDt.text,frm1.txtToReqDt.text,frm1.txtFromReqDt.Alt,frm1.txtToReqDt.Alt, _
        	               "970024",frm1.txtFromReqDt.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtFromReqDt.focus
	   Exit Function
	End If
	
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
		    strVal = strVal & "&txtItem_cd=" & .hdnitemcd.value
		    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else	
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&Item_cd=" & .txtItem_cd.value 
		    strVal = strVal & "&seq_no=0" 
		    strVal = strVal & "&lgPageNo="	 & lgPageNo						'☜: Next key tag 
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		    
		End If 
   
	End With
	Call RunMyBizASP(MyBizASP, strVal)	
	//Call ExecMyBizASP(frm1, strVal)									'☜: 비지니스 ASP 를 가동 
	

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
		//.Col = C_INSERT_USER_ID
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

'========================================================================================================
' Function Name : CookiePage
' Function Desc : 
'========================================================================================================
Function CookiePage(ByVal Kubun)

       On Error Resume Next

       Const CookieSplit = 4877                            
       Dim strTemp, arrVal

       If Kubun = 1 Then
       
          WriteCookie CookieSplit , frm1.txtItem_cd.value & parent.gRowSep 
       
       ElseIf Kubun = 0 Then

              strTemp = ReadCookie(CookieSplit)
                     
              If strTemp = "" then Exit Function
                     
              arrVal = Split(strTemp, parent.gRowSep)

              If arrVal(0) = "" Then Exit Function
              
              frm1.txtItem_cd.value = arrVal(0)
              
              If Err.number <> 0 Then
                 Err.Clear
                 WriteCookie CookieSplit , ""
                 Exit Function 
              End If
              
              Call MainQuery()
                                   
              WriteCookie CookieSplit , ""
              
       End If

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
							<TD CLASS="TD5" NOWRAP>품목코드</TD>
							<TD CLASS="TD6" NOWRAP>
							<INPUT TYPE=TEXT ALT="품목코드" NAME="txtItem_cd" SIZE=15 MAXLENGTH=18  tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK='vbscript:OpenPopupw"item_cd","txtItem_cd"'>&nbsp;<INPUT TYPE=TEXT NAME="txtItem_cd_nm" SIZE=20 MAXLENGTH=20 ALT="품목코드" tag="14X"></td>
							
							        							<TD CLASS="TD5" NOWRAP>규격</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="규격" NAME="txtItem_spec" SIZE=35 MAXLENGTH=18 tag="14NXXU"></TD>
						</TR>
						
					</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				
				<TR>		
					<TD WIDTH="100%">
					<DIV ID="TabDiv"  SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							
							<TR>
							<TD CLASS="TD5" NOWRAP>의뢰일자</TD>
							<TD CLASS="TD6" NOWRAP >
							
							<script language =javascript src='./js/b81109qa1_OBJECT2_txtReqDt.js'></script>
							
								<TD CLASS="TD5" NOWRAP>완료일자</TD>
							<TD CLASS="TD6" NOWRAP >
							<script language =javascript src='./js/b81109qa1_OBJECT2_txtEndDt.js'></script>
							
							</TD>
						
						<TR>
							<TD CLASS="TD5" NOWRAP>대분류</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="대분류" NAME="txtItem_lvl1" SIZE=10 MAXLENGTH=4  tag="14NXXU" onKeyup="test()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" >&nbsp;<INPUT TYPE=TEXT NAME="txtItem_lvl1_nm" SIZE=20 MAXLENGTH=20 ALT="대분류" tag="14X"></td>
							<TD CLASS="TD5" NOWRAP>중분류</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="중분류" NAME="txtItem_lvl2" SIZE=10 MAXLENGTH=4 tag="14NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" >&nbsp;<INPUT TYPE=TEXT ALT="중분류" NAME="txtItem_lvl2_nm" SIZE=20 tag="14X"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>소분류</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="소분류" NAME="txtItem_lvl3" SIZE=10 MAXLENGTH=4  tag="14NXXU" onKeyup="test()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" >&nbsp;<INPUT TYPE=TEXT NAME="txtItem_lvl3_nm" SIZE=20 MAXLENGTH=20 ALT="소분류" tag="14X"></td>
							<TD CLASS="TD5" NOWRAP>공급처</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtPur_vendor" SIZE=10 MAXLENGTH=4  tag="14NXXU" onKeyup="test()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtPur_vendor_nm" SIZE=20 MAXLENGTH=20 ALT="공급처" tag="14X"></td>
						</TR>
						<TR>
								<TD CLASS=TD5 NOWRAP>조달구분</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPur_type" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="조달구분"> </TD>																						
								<TD CLASS=TD5 NOWRAP>재고단위</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItem_unit" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="재고단위"> </TD>
						</TR>
						
						<TR>
							<TD CLASS="TD5" NOWRAP>이관일자</TD>
							<TD CLASS="TD6" NOWRAP>
							<INPUT TYPE=TEXT NAME="trans_date" SIZE=25 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="이관일자">
							</td>
							
							<TD CLASS="TD5" NOWRAP>의뢰자</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="의뢰자" NAME="txtReq_id" SIZE=10 MAXLENGTH=4 tag="14NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" >&nbsp;<INPUT TYPE=TEXT ALT="의뢰자" NAME="txtReq_id_nm" SIZE=20 tag="14X"></TD>
						</TR>
						
						<TR>
								<TD CLASS=TD5 NOWRAP>Status</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStatus" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="Status"> </TD>																						
								<TD CLASS=TD5 NOWRAP>도면번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtdoc_no" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="도면번호"> </TD>
						</TR>
						
						
							
							<TR>
								<TD CLASS="TD5" NOWRAP>신고의뢰사유</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtReq_reason" SIZE=102 MAXLENGTH=70 tag="24XXX" ALT="신고의뢰사유"></TD>
							</TR>						
												
							<TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<script language =javascript src='./js/b81109qa1_OBJECT4_vspdData.js'></script>
								</TD>											
							</TR>						
						
						</TABLE>						
					</DIV>
					
					<td>
				</tr>	
				
			
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
 
