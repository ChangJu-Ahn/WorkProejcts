<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : B1b04ma1
'*  4. Program Name         : HS부호등록 
'*  5. Program Desc         : HS부호등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/18
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : An Chang Hwan
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
<!--
 '******************************************  1.1 Inc 선언   **********************************************
-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit                 '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->

'******************************************  1.2 Global 변수/상수 선언  ***********************************
Const BIZ_PGM_QRY_ID   = "b1b04mb1.asp"            '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID  = "b1b04mb2.asp"            '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID   = "b1b04mb2.asp"

'==========================================  1.2.1 Global 상수 선언  ======================================
Dim C_HsCd
Dim C_HsNm
Dim C_HsSpec
Dim C_HsUnit
Dim C_HsUnit_Pb
  
Const C_SHEETMAXROWS = 100
 
Dim lgQuery
Dim lgCopyRow
Dim gblnWinEvent     '~~~ ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False				'⊙: Indicates that no value changed
	lgIntGrpCount = 0						'⊙: Initializes Group View Size
	lgStrPrevKey = ""						'initializes Previous Key
	lgLngCurRows = 0						'initializes Deleted Rows Count
	  
	gblnWinEvent = False
End Function

'========================================================================================================
Sub initSpreadPosVariables()  
	C_HsCd		= 1
	C_HsNm		= 2
	C_HsSpec	= 3
	C_HsUnit    = 4
	C_HsUnit_Pb = 5
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtHsCd.focus
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = False
	Call SetToolbar("11101100000011")
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread    
	
	With frm1
		.vspdData.ReDraw = False   
		.vspdData.MaxCols = C_HsUnit_Pb + 1
		.vspdData.MaxRows = 0
		   
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit  C_HsCd, "HS부호", 15,,,12,2
		ggoSpread.SSSetEdit  C_HsNm, "HS부호명", 45,,,50,1
		ggoSpread.SSSetEdit  C_HsSpec, "규격", 25,,,120,1
		ggoSpread.SSSetEdit	C_HsUnit, "단위", 15,,,3,2
		ggoSpread.SSSetButton C_HsUnit_Pb

		Call ggoSpread.MakePairsColumn(C_HsUnit,C_HsUnit_Pb)
		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
		
		Call SetSpreadLock()
		
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
			C_HsCd		= iCurColumnPos(1)
			C_HsNm		= iCurColumnPos(2)
			C_HsSpec	= iCurColumnPos(3)
			C_HsUnit    = iCurColumnPos(4)
			C_HsUnit_Pb	= iCurColumnPos(5)
	End Select    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
   With frm1
   	.vspdData.ReDraw = False
   	ggoSpread.spreadlock		 C_HsCd, -1, C_HsCd, -1
   	ggoSpread.spreadUnlock    C_HsNm , -1,  -1
   	ggoSpread.spreadUnlock    C_HsSpec , -1,  -1
   	ggoSpread.spreadUnlock    C_HsUnit  , -1,  -1
   	.vspdData.ReDraw = True
   End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStarRow, Byval pvEndRow)
   ggoSpread.Source = frm1.vspdData
   With frm1.vspdData
       .Redraw = False
   	ggoSpread.SSSetRequired   C_HsCd, pvStarRow, pvEndRow
   	ggoSpread.SSSetRequired   C_HsNm, pvStarRow, pvEndRow
   	ggoSpread.SSSetRequired   C_HsSpec, pvStarRow, pvEndRow
   	ggoSpread.SSSetRequired   C_HsUnit, pvStarRow, pvEndRow
   	ggoSpread.spreadUnlock    C_HsUnit_Pb , -1, -1
   	ggoSpread.spreadUnlock    C_HsSpec , -1, -1
   	.ReDraw = True
   End With
End Sub

'================================== 2.2.5 SetSpreadColor1() ==================================================
Sub SetSpreadColor1(ByVal pvStarRow, Byval pvEndRow)
 
   Dim Index
	  
   ggoSpread.Source = frm1.vspdData
	 
   With frm1.vspdData
	     
   	.Redraw = False

   	ggoSpread.SSSetprotected   C_HsCd, pvStarRow, pvEndRow
   	ggoSpread.SSSetRequired   C_HsNm, pvStarRow, pvEndRow
   	ggoSpread.SSSetRequired   C_HsSpec, pvStarRow, pvEndRow
   	ggoSpread.SSSetRequired   C_HsUnit, pvStarRow, pvEndRow
   	ggoSpread.spreadUnlock    C_HsSpec , -1, -1
		   
   	For Index = 1 to .MaxRows 
   		.Row = Index
   		.Col = 0
   		If .Text = ggoSpread.InsertFlag Then
   			Call SetSpreadColor(Index, Index)
   			ggoSpread.spreadUnlock    C_HsCd , Index, C_HsCd, Index
   			ggoSpread.SSSetRequired   C_HsCd, Index, Index
   		End if
   	Next
		    
   	.ReDraw = True
	   
   End With
End Sub
 
'------------------------------------------  OpenItemInfo()  -------------------------------------------------
Function OpenHsCd()
   On Error Resume Next
	  
   Dim strRet
   Dim iCalledAspName
	   
   If gblnWinEvent = True Or UCase(frm1.txtHsCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	   
   gblnWinEvent = True
	   
   iCalledAspName = AskPRAspName("B1b04pa1")
   If Trim(iCalledAspName) = "" Then
   	IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1b04pa1", "X")
   	lgIsOpenPop = False
   	Exit Function
   End If
	
   strRet = window.showModalDialog(iCalledAspName, Array(parent.window, ""), _
   	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
   gblnWinEvent = False
	     
   If strRet(0) = "" Then   
		frm1.txtHsCd.focus
		Exit Function
   Else
		frm1.txtHsCd.value = strRet(0)
		frm1.txtHsNm.value = strRet(1)
		frm1.txtHsCd.focus
   End If 
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
Function OpenUnit(Byval StrCode,Byval iWhere)'
   Dim arrRet
   Dim arrParam(5), arrField(6), arrHeader(6)

   If gblnWinEvent = True Then Exit Function

   gblnWinEvent = True

   arrParam(0) = "단위"				' 팝업 명칭 
   arrParam(1) = "B_UNIT_OF_MEASURE"		' TABLE 명칭 
   arrParam(2) = strCode					' Code Condition
   arrParam(3) = ""						' Name Cindition
   arrParam(4) = ""						' Where Condition
   arrParam(5) = "단위"				' TextBox 명칭 

   arrField(0) = "UNIT"					' Field명(0)
   arrField(1) = "UNIT_NM"					' Field명(1)

   arrHeader(0) = "단위"				' Header명(0)
   arrHeader(1) = "단위명"				' Header명(1)

   arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
   "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

   gblnWinEvent = False

   If arrRet(0) = "" Then
   		Exit Function
   Else
		With frm1
			If iWhere = 1 then
				.vspdData.Col = C_HsUnit
				.vspdData.Text = arrRet(0)
				Call vspdData_Change(.vspdData.Col, .vspdData.Row) 
			End if
			lgBlnFlgChgValue = True
		End With
   End If
End Function


'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	call LoadInfTB19029  
'	Call ggoOper.FormatField(Document, "2", CInt(ggAmtOfMoney.DecPoint), CInt(ggQty.DecPoint), _ 
'	                  CInt(ggUnitCost.DecPoint), CInt(ggExchRate.DecPoint), gDateFormat)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field     
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call SetDefaultVal
	Call InitVariables                                                      '⊙: Initializes local global variables
	frm1.txtHsCd.focus
    Call SetToolbar("11101101000011") '추가 
End Sub

'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	IF lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    		
	
End Sub

'========================================================================================
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
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor1(1, frm1.vspdData.MaxRows)
End Sub

'========================================================================================
Function FncSplitColumn()
    
   If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

   Call CheckMinNumSpread(frm1.vspdData, Col, Row)   '추가 
	
    lgBlnFlgChgValue = True
     
End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
  Dim strTemp
  Dim intPos1
  
  if lgQuery = true then Exit Sub
  if lgCopyRow = true then Exit Sub
 
  With frm1.vspdData 
  
     ggoSpread.Source = frm1.vspdData
     If Row > 0 And Col = C_HsUnit_Pb Then
         .Col = C_HsUnit
         .Row = Row
         
         Call OpenUnit(.text,1)
     End If
     
     End With
 End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
   If OldLeft <> NewLeft Then
   	Exit Sub
   End If
	    
   If frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS And lgStrPrevKey <> "" Then
   	If CheckRunningBizProcess = True Then
   		Exit Sub
   	End If 
	   
   	Call DisableToolBar(parent.TBC_QUERY)
		
   	If DBQuery = False Then
   		Call RestoreToolBar()
   		Exit Sub
   	End If
   End if    
End Sub

'========================================================================================
Function FncQuery()
   Dim IntRetCD

   FncQuery = False             '⊙: Processing is NG 

   Err.Clear               '☜: Protect system from crashing 

   '------ Check previous data area ------
   If lgBlnFlgChgValue = True Then
   IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")  '⊙: "Will you destory previous data" 
   	If IntRetCD = vbNo Then
   		Exit Function
   	End If
   End If

   '------ Erase contents area ------ 
   Call ggoOper.ClearField(Document, "2")         '⊙: Clear Contents  Field 
   Call InitVariables              '⊙: Initializes local global variables 

   '------ Check condition area ------ 
   If Not chkField(Document, "1") Then        '⊙: This function check indispensable field 
   	Exit Function
   End If

    '------ Query function call area ------ 
   If DbQuery = False Then Exit Function

   FncQuery = True               '⊙: Processing is OK 
End Function
 
'========================================================================================
Function FncNew()
   Dim IntRetCD 

   FncNew = False                                                          '⊙: Processing is NG  

   '------ Check previous data area ------ 
   If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
   	IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")  '⊙: "Will you destory previous data" 
   	If IntRetCD = vbNo Then
   		Exit Function
   	End If
   End If

   Call ggoOper.ClearField(Document, "A")       '⊙: Clear Contents  Field
   Call ggoOper.LockField(Document, "N")        '⊙: Lock  Suitable  Field
   Call InitVariables							'⊙: Initializes local global variables
	  
   Call SetDefaultVal

   FncNew = True               '⊙: Processing is OK

End Function

'========================================================================================
Function FncSave()
   Dim IntRetCD
	  
   FncSave = False                   '⊙: Processing is NG 
	  
   Err.Clear                    '☜: Protect system from crashing 
	  
   ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
   If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
   	IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
   	Exit Function
   End If
	    
   ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
   If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
   	Exit Function
   End If
	  
   '------ Save function call area ------ 
   If DbSave = False Then Exit Function 
	  
   FncSave = True                  '⊙: Processing is OK 
End Function

'========================================================================================================
Function FncCopy()
  
   lgCopyRow = true
	  
   frm1.vspdData.ReDraw = False

   If frm1.vspdData.Maxrows < 1 then exit function
	  
   ggoSpread.Source = frm1.vspdData 
   ggoSpread.CopyRow
	  
   SetSpreadColor  frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow 
	  
   frm1.vspdData.Row = frm1.vspdData.ActiveRow
   frm1.vspdData.Col = C_HsCd
   frm1.vspdData.Text = ""
	    
   frm1.vspdData.ReDraw = True
	  
   lgCopyRow = false
  
End Function

'========================================================================================================
Function FncCancel() 
    on error resume next
    if frm1.vspdData.Maxrows < 1 then exit function
 ggoSpread.Source = frm1.vspdData    
 ggoSpread.EditUndo              '☜: Protect system from crashing
  
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
   	.vspdData.focus
   	ggoSpread.Source = .vspdData

   	.vspdData.EditMode = True

   	.vspdData.ReDraw = False
   	ggoSpread.InsertRow, imRow
   	SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
   	lgBlnFlgChgValue = True
   	.vspdData.ReDraw = True

   End With
	
   Set gActiveElement = document.ActiveElement   
   If Err.number = 0 Then 
   	FncInsertRow = True                                                          '☜: Processing is OK
   End If
    
   frm1.vspdData.focus
End Function

'========================================================================================================
Function FncDeleteRow()
   on error resume next
   Dim lDelRows
   Dim iDelRowCnt, i
	
   If frm1.vspdData.Maxrows < 1 then exit function
	 
   With frm1.vspdData 
   	.focus
   	ggoSpread.Source = frm1.vspdData
   	lDelRows = ggoSpread.DeleteRow
   	lgBlnFlgChgValue = True
   End With
End Function

'========================================================================================================
Function FncPrint()  
   Call parent.FncPrint()
End Function

'============================================  5.1.10 FncPrev()  ========================================
Function FncPrev() 
  '------ Precheck area ------ 
 
   If lgIntFlgMode <> parent.OPMD_UMODE Then         'Check if there is retrived data 
   	Call DisplayMsgBox("900002","X","X","X")
   	Exit Function
   ElseIf lgPrevNo = "" Then         'Check if there is retrived data 
   	Call DisplayMsgBox("900011","X","X","X")
   End If
End Function

'============================================  5.1.11 FncNext()  ========================================
Function FncNext()
  '------ Precheck area ------ 
 
   If lgIntFlgMode <> parent.OPMD_UMODE Then         'Check if there is retrived data
   	Call DisplayMsgBox("900002","X","X","X")
   	Exit Function
   ElseIf lgNextNo = "" Then         'Check if there is retrived data 
   	Call DisplayMsgBox("900012","X","X","X")
   End If
End Function

'===========================================  5.1.12 FncExcel()  ========================================
Function FncExcel() 
   Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'===========================================  5.1.13 FncFind()  =========================================
Function FncFind() 
   Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function

'========================================================================================
Function FncExit()
   Dim IntRetCD

   FncExit = False

   If lgBlnFlgChgValue = True Then
   	IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")   <%'⊙: "Will you destory previous data"%>
   	If IntRetCD = vbNo Then
   		Exit Function
   	End If
   End If

   FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Err.Clear               '☜: Protect system from crashing

	DbQuery = False              '⊙: Processing is NG

	Dim strVal
	 
	If LayerShowHide(1) = False then
	   Exit Function 
	End if 
	 
	If lgIntFlgMode = parent.OPMD_UMODE Then
	  strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001    '☜: 비지니스 처리 ASP의 상태 
	  strVal = strVal & "&txtHsCd=" & Trim(frm1.txtHsCd.value)  '☆: 조회 조건 데이타 
	  strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	  strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	Else
	  strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001    '☜: 비지니스 처리 ASP의 상태 
	  strVal = strVal & "&txtHsCd=" & Trim(frm1.txtHsCd.value)  '☆: 조회 조건 데이타 
	  strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	  strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If

	Call RunMyBizASP(MyBizASP, strVal)         '☜: 비지니스 ASP 를 가동 
	 
	DbQuery = True    
  
End Function

Sub RemovedivTextArea()
	Dim ii
	
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'=============================================  5.2.2 DbSave()  =========================================
Function DbSave() 
	Dim lRow
	Dim lGrpCnt
	Dim strVal, strDel
	Dim intInsrtCnt
    Dim lColSep,lRowSep

	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size

	DbSave = False              <% '⊙: Processing is OK %>
	    
	  If LayerShowHide(1) = False Then
	     Exit Function 
	  End If

	With frm1
	 .txtMode.value = parent.UID_M0002
     lColSep = parent.gColSep
     lRowSep = parent.gRowSep
	 lGrpCnt = 1
	 strVal = ""
	 intInsrtCnt = 1
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferMaxCount = -1 
	iTmpDBufferMaxCount = -1 
	    
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

	 For lRow = 1 To .vspdData.MaxRows
	  .vspdData.Row = lRow
	  .vspdData.Col = 0

		Select Case .vspdData.Text
		Case ggoSpread.InsertFlag         '☜: 신규 
			strVal = "C" & lColSep & lRow & lColSep  '☜: C=Create, Row위치 정보 
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_HsCd,lRow,"X","X")) & lColSep
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_HsNm,lRow,"X","X")) & lColSep
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_HsSpec,lRow,"X","X")) & lColSep
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_HsUnit,lRow,"X","X")) & lRowSep

			lGrpCnt = lGrpCnt + 1
			intInsrtCnt = intInsrtCnt + 1

		Case ggoSpread.UpdateFlag         '☜: Update 
			strVal = "U" & lColSep & lRow & lColSep  '☜: U=Update, Row위치 정보 
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_HsCd,lRow,"X","X")) & lColSep
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_HsNm,lRow,"X","X")) & lColSep
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_HsSpec,lRow,"X","X")) & lColSep
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_HsUnit,lRow,"X","X")) & lRowSep

			lGrpCnt = lGrpCnt + 1

		Case ggoSpread.DeleteFlag         '☜: 삭제 
			strDel = "D" & lColSep & lRow & lColSep  '☜: D=Update, Row위치 정보 
			strDel = strDel & Trim(GetSpreadText(.vspdData,C_HsCd,lRow,"X","X")) & lColSep
			strDel = strDel & Trim(GetSpreadText(.vspdData,C_HsNm,lRow,"X","X")) & lColSep
			strDel = strDel & Trim(GetSpreadText(.vspdData,C_HsSpec,lRow,"X","X")) & lColSep
			strDel = strDel & Trim(GetSpreadText(.vspdData,C_HsUnit,lRow,"X","X")) & lRowSep

			lGrpCnt = lGrpCnt + 1
			      
		End Select

		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
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

	.txtMaxRows.value = lGrpCnt-1
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

	 Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)       '☜: 비지니스 ASP 를 가동 

	End With

	DbSave = True               '⊙: Processing is NG 
End Function

'========================================================================================
Function DbQueryOk()              '☆: 조회 성공후 실행로직 
	lgIntFlgMode = parent.OPMD_UMODE            '⊙: Indicates that current mode is Update mode 

	lgBlnFlgChgValue = False

	Call ggoOper.LockField(Document, "Q")         '⊙: This function lock the suitable field 
	  
	Call SetToolbar("1110111100111111")          '⊙: 버튼 툴바 제어 

	If frm1.vspdData.MaxRows > 0 Then
	 frm1.vspdData.Focus
	Else
	 frm1.txtHsCd.focus
	End If

	frm1.vspdData.ReDraw = False
	 
	lgQuery = False
	   
	frm1.vspdData.ReDraw = True
End Function

'=============================================  5.2.5 DbSaveOk()  =======================================
Function DbSaveOk()              '☆: 저장 성공후 실행 로직 
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call RemovedivTextArea
	Call MainQuery()
End Function
 

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<!--
'#########################################################################################################
'            6. Tag부 
'######################################################################################################### 
-->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
   <TR>
  <TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백 %></TD>
 </TR>
 <TR HEIGHT=23>
  <TD WIDTH="100%">
   <TABLE <%=LR_SPACE_TYPE_10%>>
    <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD CLASS="CLSMTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>HS부호</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=*>&nbsp;</TD>
    </TR>
    
   </TABLE>
  </TD>
 </TR>
 <TR HEIGHT=*>
  <TD WIDTH="100%" CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>>
    <TR>
     <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
    </TR>        
    <TR>
     <TD HEIGHT=20 WIDTH="100%">
      <FIELDSET CLASS="CLSFLD">
       <TABLE <%=LR_SPACE_TYPE_40%>>
        <TR>
         <TD CLASS="TD5" NOWRAP>HS부호</TD>
         <TD CLASS="TD6" NOWRAP><INPUT NAME="txtHsCd" MAXLENGTH="20" SIZE=20 ALT ="HS부호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnHsCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenHsCd()">&nbsp;
                 <INPUT NAME="txtHsNm" MAXLENGTH="50" SIZE=25 ALT  ="HS부호명" tag="14"></TD>
         <TD CLASS="TD6" NOWRAP>&nbsp;</TD>
         <TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
         <script language =javascript src='./js/b1b04ma1_vaSpread1_vspdData.js'></script>
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
    <tr HEIGHT="20">
  <td WIDTH="100%">
   <table <%=LR_SPACE_TYPE_30%>>
    <tr>
     <TD WIDTH=10>&nbsp;</TD>
     <TD WIDTH=10>&nbsp;</TD>
    </tr>
   </table>
  </td>
    </tr>
 <TR>
  <TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH="100%" SRC="<%=BIZ_PGM_QRY_ID%>" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
  </TD>
 </TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHHsCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


