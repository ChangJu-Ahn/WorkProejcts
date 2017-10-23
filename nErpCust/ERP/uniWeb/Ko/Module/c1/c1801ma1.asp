
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : c/c별 배부근거 등록 
'*  3. Program ID           : c1801ma1.asp
'*  4. Program Name         : c/c별 배부근거 등록 
'*  5. Program Desc         : Cost Center별 배부근거 등록 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2002/06/19
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : Cho Ig Sung / Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit		

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c1801mb1.asp"	                            'Biz Logic ASP 

'========================================================================================================
'=                       4.2 Constant variables 
'======================================================================================================== 
Dim C_CostCd  		
Dim C_CostCdPopup  															
Dim C_CostNm  
Dim C_ProdBasisQty  
Dim C_ProdBasisAmt  
Dim C_AdjRate  


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  
	 C_CostCd			= 1		
	 C_CostCdPopup		= 2															
	 C_CostNm			= 3
	 C_ProdBasisQty		= 4
	 C_ProdBasisAmt		= 5
	 C_AdjRate			= 6

End Sub

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE 
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0   
    
    lgStrPrevKey = ""  
    lgLngCurRows = 0  
    lgSortKey = 1
End Sub

Sub SetDefaultVal()

	Dim StartDate
	Dim EndDate
	
	StartDate	= "<%=GetSvrDate%>"
	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)

	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub


'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	With frm1.vspdData
	
    .MaxCols = C_AdjRate+1
    .Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
  
	.ReDraw = false
	
	Call GetSpreadColumnPos("A")
	
    ggoSpread.SSSetEdit C_CostCd, "코스트센타코드", 20,,,10,2
    ggoSpread.SSSetButton C_CostCdPopup
    ggoSpread.SSSetEdit C_CostNm, "코스트센타명", 35

    ggoSpread.SSSetFloat C_ProdBasisQty,"배부수량",20,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

    ggoSpread.SSSetFloat C_ProdBasisAmt,"배부금액",20,Parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

    ggoSpread.SSSetFloat C_AdjRate,"가중치(%)",20,Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

	call ggoSpread.MakePairsColumn(C_CostCd,C_CostCdPopup)
		
	Call ggoSpread.SSSetColHidden(C_AdjRate,C_AdjRate,True)

'    ggoSpread.SSSetSplit(C_CostNm)	                            
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_CostCd			, -1, C_CostCd
    ggoSpread.SpreadLock C_CostCdPopup		, -1, C_CostCdPopup
    ggoSpread.SpreadLock C_CostNm			, -1, C_CostNm
    ggoSpread.SpreadUnLock C_ProdBasisQty			, -1, C_ProdBasisQty
    ggoSpread.SpreadUnLock C_ProdBasisAmt			, -1, C_ProdBasisAmt
    ggoSpread.SSSetRequired	C_ProdBasisQty	, -1, -1
    ggoSpread.SSSetRequired	C_ProdBasisAmt	, -1, -1
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	C_CostCd		,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected	C_CostNm		,pvStartRow	,pvEndRow
    ggoSpread.SSSetRequired	C_ProdBasisQty	,pvStartRow	,pvEndRow
    ggoSpread.SSSetRequired	C_ProdBasisAmt	,pvStartRow	,pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub


'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_CostCd						= iCurColumnPos(1)
			C_CostCdPopup					= iCurColumnPos(2)
			C_CostNm						= iCurColumnPos(3)    
			C_ProdBasisQty					= iCurColumnPos(4)
			C_ProdBasisAmt					= iCurColumnPos(5)
			C_AdjRate						= iCurColumnPos(6)
    End Select    
End Sub


Function OpenCostCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "코스트센타팝업"
	arrParam(1) = "B_COST_CENTER a, B_MINOR b"	
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "a.COST_TYPE *= b.MINOR_CD and b.MAJOR_CD = " & FilterVar("c2203", "''", "S") & " "
	arrParam(5) = "코스트센타"			
	
    arrField(0) = "a.COST_CD"
    arrField(1) = "a.COST_NM"
    arrField(2) = "b.MINOR_NM"
    
    arrHeader(0) = "코스트센타코드"	
    arrHeader(1) = "코스트센타명"
    arrHeader(2) = "코스트센타구분"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtCostCd.focus
		Exit Function
	Else
		Call SetCostCd(arrRet, iWhere)
	End If	

End Function

Function SetCostCd(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 0 Then
    		 frm1.txtCostCd.focus
    		.txtCostCd.value = arrRet(0)
    		.txtCostNm.value = arrRet(1)
    	Else
    		.vspdData.Col = C_CostCd
    		.vspdData.Text = arrRet(0)
    		.vspdData.Col = C_CostNm
    		.vspdData.Text = arrRet(1)
            
    		Call vspdData_Change(.vspdData.Col, .vspdData.Row)	
    	End If
	
	End With
	
End Function

Function OpenDstbFctrCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "배부요소팝업"
	arrParam(1) = "C_DSTB_FCTR"	
	arrParam(2) = Trim(frm1.txtDstbFctrCd.value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "배부요소"			
	
    arrField(0) = "DSTB_FCTR_CD"
    arrField(1) = "DSTB_FCTR_NM"
    
    arrHeader(0) = "배부요소코드"
    arrHeader(1) = "배부요소명"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDstbFctrCd.focus
		Exit Function
	Else
		Call SetDstbFctrCd(arrRet)
	End If	

End Function

Function SetDstbFctrCd(Byval arrRet)
	
	With frm1
		 frm1.txtDstbFctrCd.focus
		.txtDstbFctrCd.value = arrRet(0)
		.txtDstbFctrNm.value = arrRet(1)
	End With
	
End Function

Sub Form_Load()

    Call LoadInfTB19029   
    
    Call ggoOper.LockField(Document, "N")               
                     
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
                                                                           
    Call InitSpreadSheet  
    Call InitVariables  
    
    Call SetDefaultVal
    Call SetToolbar("110011010010111")		
    frm1.txtYyyymm.focus
   	Set gActiveElement = document.activeElement			
    
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
   	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	
    
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp

	With frm1 
	
		ggoSpread.Source = frm1.vspdData
		   
		If Row > 0 And Col = C_CostCdPopUp Then
		    .vspdData.Col = Col
		    .vspdData.Row = Row
		        
			.vspdData.Col = C_CostCd
		    Call OpenCostCd(.vspdData.Text, 1)
		End If
    Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_CostNm Or NewCol <= C_CostNm Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
    	If lgStrPrevKey <> "" Then 
      	DbQuery
    	End If

    End if
    
End Sub

Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear   
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")	
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    	
    Call InitVariables  
    															
    if frm1.txtDstbFctrCd.value = "" then
		frm1.txtDstbFctrNm.value = ""
    end if

    if frm1.txtCostCd.value = "" then
		frm1.txtCostNm.value = ""
    end if

    If Not chkField(Document, "1") Then	
       Exit Function
    End If
    
    IF DbQuery	= False Then
		Exit function
	END IF
       
    FncQuery = True															
    
End Function

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear  
    
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If

    If Not chkField(Document, "1") Then 
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then  
       Exit Function
    End If
    
    IF DbSave = False Then
		Exit function
	END IF	
    
    FncSave = True                                                          
    
End Function


'========================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
    frm1.vspdData.Col = C_CostCd
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_CostNm
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
End Function

Function FncCancel() 

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  
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
		End If	
	End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function

Function FncDeleteRow() 
    Dim lDelRows

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   
End Function

Function FncPrev() 
    On Error Resume Next  
End Function

Function FncNext() 
    On Error Resume Next  
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)	
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)  
End Function


'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

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
'   Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub


Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")     
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 

    
    DbQuery = False
    
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
    
    Err.Clear 

	Dim strVal
	Dim stryyyymm
    Dim	strYear, strMonth, strDay
    
    With frm1

    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    
    stryyyymm = strYear & strMonth
    
	    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtYyyymm=" & .hYyyymm.value				
		strVal = strVal & "&txtDstbFctrCd=" & .hDstbFctrCd.value				
		strVal = strVal & "&txtCostCd=" & .hCostCd.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtYyyymm=" & stryyyymm
		strVal = strVal & "&txtDstbFctrCd=" & .txtDstbFctrCd.value				
		strVal = strVal & "&txtCostCd=" & .txtCostCd.value
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If

	Call RunMyBizASP(MyBizASP, strVal)
 
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()

    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")	
    
    IF Trim(frm1.hGenFlag.value) = "A" then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLock C_CostCd,-1,C_AdjRate,-1
		Call SetToolbar("110000000001111")	   
    ELSE
		Call SetSpreadLock 
		
		IF frm1.hDataExists.value = "Y" Then
			Call SetToolbar("110011110011111")	    
		Else
			Call SetToolbar("110011010011111")	    
		End IF
    END IF

	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   
	                 
End Function

Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	Dim	strYear, strMonth, strDay
    Dim iColSep 
    Dim iRowSep   
	
    DbSave = False                                                          
    
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	' 날짜 처리함수 관련하여 MA쪽에서 Extract하여 MB쪽으로 넘겨줌 
    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    
    frm1.hYYYYMM.value = strYear & strMonth  
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	  
	
	With frm1
		.txtMode.value = Parent.UID_M0002
    
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag	
				
				strVal = strVal & "C" & Parent.gColSep & lRow & iColSep 

                .vspdData.Col = C_CostCd	'1
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_ProdBasisQty
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep

                .vspdData.Col = C_ProdBasisAmt
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep

                 .vspdData.Col = C_AdjRate
                 strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
					
				strVal = strVal & "U" & iColSep & lRow & iColSep 

                .vspdData.Col = C_CostCd
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_ProdBasisQty
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep

                .vspdData.Col = C_ProdBasisAmt
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep

                .vspdData.Col = C_AdjRate
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag	

				strDel = strDel & "D" & iColSep & lRow & iColSep

                .vspdData.Col = C_CostCd
                strDel = strDel & Trim(.vspdData.Text) & iRowSep									
                
                lGrpCnt = lGrpCnt + 1
        End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()	
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	
    Call MainQuery()
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function ExeReflect()

    Dim IntRetCD
    Dim strVal
    Dim strYYYYMM
    Dim strDstbFctrCd
    Dim strYear,strMonth,strDay


    strVal = ""

	ExeReflect = False

    '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.txtYYYYMM.Text,frm1.txtYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM =   strYear & strMonth
    strDstbFctrCd = Trim(frm1.txtDstbFctrCd.value)

    Call CommonQueryRs("count(*)","c_dstb_basis_by_cc","yyyymm = " & FilterVar(strYYYYMM, "''", "S") & "and DSTB_FCTR_CD = " & FilterVar(strDstbFctrCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    if Trim(Replace(lgF0,Chr(11),"")) <> 0 then
	    IntRetCD = DisplayMsgBox("236071",Parent.VB_YES_NO,"배부요소","X")    '버젼정보가 이미 존재합니다.계속진행하시겠습니까?
    end if
    
   
    If IntRetCD = vbNo Then
		Exit Function
    End If

  
    if LayerShowHide(1) = false then
	    Exit Function
    end if


    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0004                 'copy
    strVal = strVal     & "&txtYYYYMM="			 & strYYYYMM
    strVal = strVal     & "&txtDstbFctrCd="		 & strDstbFctrCd					'배부요소 

	Call RunMyBizASP(MyBizASP, strVal)                                          '☜:  Run biz logic

    Call LayerShowHide(0)
    
    ExeReflect = True                                                           '⊙: Processing is NG
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>C/C별배부근거등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/c1801ma1_fpDateTime1_txtYyyymm.js'></script>
									</TD>								
									<TD CLASS="TD5" NOWRAP>배부요소</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDstbFctrCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="배부요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDstbFctrCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDstbFctrCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtDstbFctrNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>코스트센타</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="코스트센타"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCostCd frm1.txtCostCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=20 tag="14"></TD>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c1801ma1_I322049623_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>전월COPY</BUTTON>
 		</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYyyymm" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hDstbFctrCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hGenFlag" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hDataExists" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

