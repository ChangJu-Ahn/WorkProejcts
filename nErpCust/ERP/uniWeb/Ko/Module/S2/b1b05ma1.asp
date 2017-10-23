<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales														                *
'*  2. Function Name        : 품목기준정보												                *
'*  3. Program ID           : B1B05MA1     														        *
'*  4. Program Name         : 품목별 공장배분비															*
'*  5. Program Desc         : 품목별 공장배분비															*
'*  6. Comproxy List        : PB3CS90.dll, PB3CS91.dll				                                    *
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2002/11/15																*
'*  9. Modifier (First)     : son bum yeol    															*
'* 10. Modifier (Last)      : Ahn Tae Hee    															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 2002/11/15 : UI성능 적용						                            *
'*                            2002/11/27 : Grid성능 적용, Kang Jun Gu
'*                            2002/12/02 : Grid 추가 성능 적용, Kang Jun Gu
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                 '☜: indicates that All variables must be declared in advance

'========================================================================================================
Const BIZ_PGM_ID = "b1b05mb1.asp"            '☆: Head Query 비지니스 로직 ASP명 
Const BIZ_JUMP_ID = "S2211BA4"				 '☆: JUMP시 비지니스 로직 ASP명(품목별공장배분비일괄생성)
'========================================================================================================
Dim C_PlantCode  '공장코드 
Dim C_PlantCodeBtn  '공장코드팝업 
Dim C_PlantName  '공장명 
Dim C_Rate  '배분비 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
Dim IsOpenPop      ' Popup
'========================================================================================================
Sub initSpreadPosVariables()  
	C_PlantCode  = 1  '공장코드 
	C_PlantCodeBtn = 2  '공장코드팝업 
	C_PlantName  = 3  '공장명 
	C_Rate   = 4  '배분비 

End Sub
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub
'=========================================================================================================
Sub SetDefaultVal()
 frm1.txtConItemCode.focus
 frm1.txtRate.value = 0
End Sub
'==========================================================================================================
Sub LoadInfTB19029()
 <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub
'==========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
	    ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    

	    .MaxCols   = C_Rate	+ 1													' ☜:☜: Add 1 to Maxcols
	    .MaxRows = 0       
																				' ☜: Clear spreadsheet data 
	    Call GetSpreadColumnPos("A")
		.ReDraw = false
     
		ggoSpread.SSSetEdit C_PlantCode, "공장",30,,,4,2
		ggoSpread.SSSetButton C_PlantCodeBtn
		ggoSpread.SSSetEdit C_PlantName, "공장명",55,,,40
		Call AppendNumberPlace("6","3","2")
		ggoSpread.SSSetFloat C_Rate,"배분율(%)",30,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		call ggoSpread.MakePairsColumn(C_PlantCode,C_PlantCodeBtn)
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)					'☜: 공통콘트롤 사용 Hidden Column
		.ReDraw = true
    End With
    
End Sub
'===========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_PlantCode, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlantName, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Rate, pvStartRow, pvEndRow
		ggoSpread.SpreadUnLock C_PlantCodeBtn, pvStartRow , C_PlantCodeBtn, pvEndRow
		.vspdData.ReDraw = True
    End With

End Sub
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PlantCode  = iCurColumnPos(1)  '공장코드 
			C_PlantCodeBtn = iCurColumnPos(2)  '공장코드팝업 
			C_PlantName  = iCurColumnPos(3)  '공장명 
			C_Rate   = iCurColumnPos(4)  '배분비 
    End Select    
End Sub
'===========================================================================
Function OpenItemPopup(ByVal strPopUpStyle,objEleTagCode,objEleTagName)
  Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function
 
 If objEleTagCode.ReadOnly = True Then Exit Function

 IsOpenPop = True

 arrParam(0) = "품목"					<%' 팝업 명칭 %>
 arrParam(1) = "b_item "					<%' TABLE 명칭 %>
 arrParam(2) = Trim(objEleTagCode.Value)    <%' Code Condition%>
 arrParam(3) = ""							<%' Name Cindition%>
 arrParam(4) = ""							<%' Where Condition%>
 arrParam(5) = "품목"					<%' TextBox 명칭 %>

 arrField(0) = "item_cd"					<%' Field명(0)%>
 arrField(1) = "item_nm"					<%' Field명(1)%>
    
 arrHeader(0) = "품목"					<%' Header명(0)%>
 arrHeader(1) = "품목명"				<%' Header명(1)%>

 objEleTagCode.focus
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
   "dialogWidth=450px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else

  Select Case strPopUpStyle
  Case 1
   Call SetItemPopup(arrRet,"Condition")
  Case 2
   Call SetItemPopup(arrRet,"Content")
  End Select
 End If 
 
End Function

'===========================================================================
Function OpenSpreadPop(Byval strCode)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 If Trim(frm1.txtItemCode.value) = "" Then
  Msgbox "품목을 먼저 입력하세요",vbExclamation, Parent.gLogoName
  frm1.txtItemCode.focus
  Exit Function
 End If

 IsOpenPop = True

 arrParam(0) = "공장"									<%' 팝업 명칭 %>
 arrParam(1) = "B_PLANT Plant,B_ITEM_BY_PLANT ItemPlant"    <%' TABLE 명칭 %>
 arrParam(2) = ""											<%' Code Condition%>
 arrParam(3) = ""											<%' Name Cindition%>
 arrParam(4) = "(Plant.PLANT_CD = ItemPlant.PLANT_CD) and (itemplant.item_cd= "_
    & FilterVar(Trim(frm1.txtItemCode.value), "''", "S") & ")"     <%' Where Condition%>
 arrParam(5) = "공장"									<%' TextBox 명칭 %>
 
 arrField(0) = "ItemPlant.PLANT_CD"							<%' Field명(0)%>
 arrField(1) = "Plant.PLANT_NM"								<%' Field명(1)%>
    
 arrHeader(0) = "공장"									<%' Header명(0)%>
 arrHeader(1) = "공장명"								<%' Header명(1)%>
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetSpreadPop(arrRet)
 End If 
 
End Function

'========================================================================================================= 
Function SetItemPopup(Byval arrRet, ByVal PopUpStyle)

 Select Case PopUpStyle
 Case "Condition"
  frm1.txtConItemCode.value = arrRet(0)
  frm1.txtConItemName.value = arrRet(1)
 Case "Content"
 Call ggoOper.ClearField(Document, "2")       <%'⊙: Clear Contents  Field%>

  frm1.txtItemCode.value = arrRet(0)
  frm1.txtItemName.value = arrRet(1)

  lgBlnFlgChgValue = True

  Call ItemByPlant()

 End Select
  
End Function

'========================================================================================================= 
Function SetSpreadPop(Byval arrRet)

 With frm1

  .vspdData.Col = C_PlantCode
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_PlantName
  .vspdData.Text = arrRet(1)
  
  Call vspdData_Change(.vspdData.Col, .vspdData.Row)  <% ' 변경이 읽어났다고 알려줌 %>

 End With
 
End Function

'========================================================================================================= 
Function ItemByPlantOk()              <%'☆: 조회 성공후 실행로직 %>
 
 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄/찾기 

 Call SetToolbar("11101101001111")
 Call SetQuerySpreadColor(-1)
 Call SumRateGrid

 With frm1
  Dim PRow

  ggoSpread.Source = .vspdData

  .vspdData.ReDraw = False

  For PRow = 0 To .vspdData.MaxRows - 1
   .vspdData.Col = 0
   .vspdData.Row = PRow + 1
   .vspdData.Text = ggoSpread.InsertFlag
  Next

  .vspdData.ReDraw = True

 End With

End Function


'========================================================================================================= 
Sub SetQuerySpreadColor(ByVal lRow)
 
    With frm1

		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_PlantCode, lRow, lRow
		ggoSpread.SSSetProtected C_PlantName, lRow, lRow
		ggoSpread.SSSetRequired C_Rate, lRow, lRow
		ggoSpread.SpreadLock C_PlantCodeBtn, -1, C_PlantCodeBtn
		.vspdData.ReDraw = True
    
    End With

End Sub

<% '======================================================================================================
' Description : Query 후 배분비 자동합계 
'========================================================================================================= %>
Sub SumRateGrid()
 
 Dim SumTotal, iMonth, lRow
 
 SumTotal = 0
    
    ggoSpread.Source = frm1.vspdData
 For lRow = 1 To frm1.vspdData.MaxRows 
  frm1.vspdData.Row = lRow
  frm1.vspdData.Col = 0
  If frm1.vspdData.Text <> ggoSpread.DeleteFlag then
   frm1.vspdData.Col = C_Rate
   SumTotal = UNICDbl(frm1.vspdData.Text) +SumTotal
  End If
 Next
   
    frm1.txtRate.value = SumTotal
End Sub

'==========================================================================================================
Function CookiePage(Byval pvKubun)

	On Error Resume Next
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim iStrItemCd

	With frm1
		If pvKubun = 1 Then
			WriteCookie CookieSplit , .txtConItemCode.value
		ElseIf pvKubun = 0 Then
			iStrItemCd = Trim(ReadCookie(CookieSplit))
			
			If iStrItemCd = "" then Exit Function
			.txtConItemCode.value = iStrItemCd			
			WriteCookie CookieSplit , ""
		End If
	End With
End Function

'==========================================================================================================
Function JumpChgCheck(byVal pvStrJumpPgmId)
	Call CookiePage(1)
	Call PgmJump(pvStrJumpPgmId)
End Function

'========================================================================================================= 
Sub Form_Load()

 Call InitVariables              '⊙: Initializes local global variables
 Call LoadInfTB19029              '⊙: Load table , B_numeric_format
'   
 Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
 Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
 Call InitSpreadSheet
 Call SetDefaultVal
 Call CookiePage(0) 
 Call SetToolbar("11000000000011")          '⊙: 버튼 툴바 제어 
    
End Sub
'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub

'========================================================================================================= 
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

 With frm1.vspdData 
 
  ggoSpread.Source = frm1.vspdData
   
  If Row > 0 And Col = C_PlantCodeBtn Then
   .Col = Col-1
   .Row = Row
   Call OpenSpreadPop(.Text)
   
   Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")
  End If
    
 End With

End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col , ByVal Row)

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
	
	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")


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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub
'========================================================================================================= 
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

 lgBlnFlgChgValue = True

 Call SumRateGrid

End Sub
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================= 
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then  Exit Sub
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess Then   Exit Sub
  
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    
End Sub

'========================================================================================================= 
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

<%    '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then
      Exit Function
  End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    
    Call ggoOper.ClearField(Document, "2")          <%'⊙: Clear Contents  Field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>
   

<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
    End If

 Call ggoOper.LockField(Document, "N")         <%'⊙: Lock  Suitable  Field%>
    Call SetToolbar("11000000000011")          '⊙: 버튼 툴바 제어 

<%  '-----------------------
    'Query function call area
    '----------------------- %>
    Call DbQuery                <%'☜: Query db data%>

    FncQuery = True                <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================= 
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Check previous data area
    '-----------------------%>
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then
      Exit Function
  End If
    End If
    
<%  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------%>
    Call ggoOper.ClearField(Document, "A")                                      <%'⊙: Clear Condition,Contents Field%>
    Call ggoOper.LockField(Document, "N")                                       <%'⊙: Lock  Suitable  Field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>

 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄 
    Call SetToolbar("11000000000011")          '⊙: 버튼 툴바 제어 
    Call SetDefaultVal

    Set gActiveElement = document.ActiveElement   
    
    FncNew = True                <%'⊙: Processing is OK%>

End Function

'========================================================================================================= 
Function FncDelete() 
    
    Exit Function
    Err.Clear                                                               '☜: Protect system from crashing    
    
    FncDelete = False              <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","x","x","x")
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition,Contents Field
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

'========================================================================================================= 
Function FncSave() 
    
    Dim lRow  
    Dim ICnt 
    
	Call SumRateGrid

    ICnt = 0
  
	For lRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Col = 0
		frm1.vspdData.Row = lRow
 
		If frm1.vspdData.Text = ggoSpread.DeleteFlag Then
			ICnt = ICnt + 1
		End If     
	Next

	If ICnt <> frm1.vspdData.MaxRows  and frm1.txtRate.value < 100 Then
  
		Msgbox "배분비의 합이 100% 이어야 합니다.",vbExclamation, Parent.gLogoName
		frm1.vspdData.focus
		Exit Function
	End If
 
	If frm1.txtRate.value > 100 Then
	 
		Msgbox "배분비의 합이 100%를 초과합니다.",vbExclamation, Parent.gLogoName
		frm1.vspdData.focus
		Exit Function
 
	End If

    Dim IntRetCD 

    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '-----------------------%>
 <%'⊙: If MULTI/SINGLEMULTI %>
    If Not chkField(Document, "2") OR ggoSpread.SSDefaultCheck = False Then     <%'⊙: Check contents area%>
       Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll DbSave                                                    <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================= 
Function FncCopy() 
 Dim IntRetCD

 With frm1
  .vspdData.ReDraw = False
 
  ggoSpread.Source = frm1.vspdData 
  ggoSpread.CopyRow
  SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

  Call SumRateGrid

  .vspdData.ReDraw = True
 End With

 lgBlnFlgChgValue = True
    
End Function

'========================================================================================================= 
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData 
    ggoSpread.EditUndo  
 Call SumRateGrid
End Function

'========================================================================================================= 
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
  With frm1
		FncInsertRow = False                                                         '☜: Processing is NG

		If Not chkField(Document, "2") Then
		Exit Function
		End If

		If IsNumeric(Trim(pvRowCnt)) Then
		    imRow = CInt(pvRowCnt)
		Else
		    imRow = AskSpdSheetAddRowCount()
		    If imRow = "" Then
		        Exit Function
		    End If
		End If
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow

		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.Row= .vspdData.ActiveRow
		.vspdData.ReDraw = True
  End With
 	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    Set gActiveElement = document.ActiveElement   
   
   lgBlnFlgChgValue = True
End Function

'========================================================================================================= 
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
	lDelRows = ggoSpread.DeleteRow
 
    lgBlnFlgChgValue = True
    
    Call SumRateGrid
    
    End With
    
End Function

'========================================================================================================= 
Function FncPrint() 
    ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
End Function

'========================================================================================================= 
Function FncExcel() 
 Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================================= 
Function FncFind() 
 Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================================= 
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================= 
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================= 
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetQuerySpreadColor(-1)
End Sub

'========================================================================================================= 
Function FncExit()
 Dim IntRetCD
 FncExit = False

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================================= 
Function DbDelete() 
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>
End Function

'========================================================================================================= 
Function DbDeleteOk()              <%'☆: 삭제 성공후 실행 로직 %>
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>
End Function

'========================================================================================================= 
Function DbQuery() 

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>

	If   LayerShowHide(1) = False Then
		Exit Function 
	End If

     
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtConItemCode=" & Trim(frm1.txtHConItemCode.value)   <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtConItemCode=" & Trim(frm1.txtConItemCode.value)   <%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If    
    
	Call RunMyBizASP(MyBizASP, strVal)            <%'☜: 비지니스 ASP 를 가동 %>
 
    DbQuery = True                 <%'⊙: Processing is NG%>

End Function

'========================================================================================================= 
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>
  
 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄/찾기 
    Call SetToolbar("11101111001111")
	Call ggoOper.LockField(Document, "Q")
	Call SetQuerySpreadColor(-1)
	Call SumRateGrid

End Function

'========================================================================================================= 
Function DbSave() 

    Err.Clear                <%'☜: Protect system from crashing%>
 
    Dim lRow        
    Dim lGrpCnt     
 Dim strVal, strDel
 
    DbSave = False                                                          '⊙: Processing is NG
    
 If   LayerShowHide(1) = False Then
      Exit Function 
 End If

 
 With frm1
  .txtMode.value = Parent.UID_M0002
  .txtUpdtUserId.value = Parent.gUsrID
  .txtInsrtUserId.value = Parent.gUsrID
    
  '-----------------------
  'Data manipulate area
  '-----------------------
  lGrpCnt = 0
  strVal = ""
  strDel = ""

  '-----------------------
  'Data manipulate area
  '-----------------------
  For lRow = 1 To .vspdData.MaxRows
    
	.vspdData.Row = lRow
	.vspdData.Col = 0

   Select Case .vspdData.Text
    Case ggoSpread.InsertFlag       '☜: 신규 
     strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep'☜: C=Create
    Case ggoSpread.UpdateFlag       '☜: 수정 
     strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep'☜: U=Update
    Case ggoSpread.DeleteFlag       '☜: 삭제 
     strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep'☜: D=Delete
   End Select

   Select Case .vspdData.Text
    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
     '--- 품목코드 
     strVal = strVal & Trim(frm1.txtItemCode.value) & Parent.gColSep
     '--- 공장코드 
     .vspdData.Col = C_PlantCode              
     strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '--- 배분비 
     .vspdData.Col = C_Rate              
     strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep
                     
     lGrpCnt = lGrpCnt + 1

    Case ggoSpread.DeleteFlag
     '--- 품목코드 
     strDel = strDel & Trim(frm1.txtItemCode.value) & Parent.gColSep
     '--- 공장코드 
     .vspdData.Col = C_PlantCode              
     strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
     '--- 배분비 
     .vspdData.Col = C_Rate              
     strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep

     lGrpCnt = lGrpCnt + 1

   End Select

  Next
   
  .txtMaxRows.value = lGrpCnt 
  .txtSpread.value = strDel & strVal
      Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 
 
 End With
 
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================================= 
Function DbSaveOk()               <%'☆: 저장 성공후 실행 로직 %>

 Call ggoOper.ClearField(Document, "2")
    Call InitVariables
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별공장배분비</font></td>
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
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 COLSPAN = 4><INPUT NAME="txtConItemCode" ALT="품목" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemPopup 1,frm1.txtConItemCode,frm1.txtConItemName">&nbsp;<INPUT NAME="txtConItemName" TYPE="Text" SIZE=40 tag="24"></TD>
									<TD CLASS=TDT NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
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
								<TD CLASS=TD5 NOWRAP>품목</TD>
								<TD CLASS=TD6 COLSPAN = 4><INPUT NAME="txtItemCode" ALT="품목" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24XXXU">&nbsp;<INPUT NAME="txtItemName" TYPE="Text" SIZE=40 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>배분비의합</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRate" TYPE="Text" SIZE=20 STYLE="Text-Align:Right" tag="24XXX">&nbsp;<B>%</B>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/b1b05ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE WIDTH=100%>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck(BIZ_JUMP_ID)">품목별공장배분비일괄생성</a></TD>
					<TD WIDTH=10>&nbsp;</TD> 
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">

<INPUT TYPE=HIDDEN NAME="txtHConItemCode" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV> 
</BODY>
</HTML>
