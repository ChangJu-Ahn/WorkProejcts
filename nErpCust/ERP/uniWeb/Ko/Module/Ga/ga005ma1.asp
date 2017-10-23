
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1 %>
<!-- '======================================================================================================
'*  1. Module Name          : 경영손익 
'*  2. Function Name        : 배부C/C정의 
'*  3. Program ID           : ga005ma1.asp
'*  4. Program Name         : 배부C/C정의 
'*  5. Program Desc         : 배부C/C정의 
'*  6. Modified date(First) : 2003/06/16
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park Joon-Won 
'*  9. Modifier (Last)      : 
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "ga005mb1.asp"	                          'Biz Logic ASP
 
Dim C_CostCd 
Dim C_Cost_Pop 	
Dim C_CostNm
Dim C_BizUnitCd
Dim C_BizUnit_Pop
Dim C_BizUnitNm

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          


'========================================================================================================
Sub initSpreadPosVariables()  
	 C_CostCd			= 1
	 C_Cost_Pop			= 2
	 C_CostNm			= 3															
	 C_BizUnitCd		= 4
	 C_BizUnitNm		= 5
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


'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
	With frm1.vspdData
	
    .MaxCols = C_BizUnitNm+1	
    .Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030616" ,,parent.gAllowDragDropSpread   

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false

    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit		C_CostCd		,"Cost Center"	,15 ,,,10,2
    ggoSpread.SSSetButton	C_Cost_Pop
    ggoSpread.SSSetEdit		C_CostNm		,"Cost Center명", 20,,,40
    ggoSpread.SSSetEdit		C_BizUnitCd		,"사업부"		,15 ,,,20,2
    ggoSpread.SSSetEdit		C_BizUnitNm		,"사업부명"		,20,,,35,2

	call ggoSpread.MakePairsColumn(C_CostCd,C_Cost_Pop)
	call ggoSpread.MakePairsColumn(C_BizUnitCd,C_BizUnit_Pop)

	
	.ReDraw = true

'   ggoSpread.SSSetSplit(C_MinorNm)	
    Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock	C_CostCd			, -1, C_CostCd
    ggoSpread.SpreadLock	C_Cost_Pop			, -1, C_Cost_Pop
    ggoSpread.SpreadLock	C_CostNm			, -1, C_CostNm
    ggoSpread.SpreadLock	C_BizUnitCd			, -1, C_BizUnitCd
    ggoSpread.SpreadLock	C_BizUnitNm			, -1, C_BizUnitNm
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                         'Col				Row				Row2
    ggoSpread.SSSetRequired		C_CostCd		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_CostNm		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_BizUnitCd		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_BizUnitNm		,pvStartRow		,pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub



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
			C_CostCd				= iCurColumnPos(1)
			C_Cost_Pop				= iCurColumnPos(2)
			C_CostNm				= iCurColumnPos(3)    
			C_BizUnitCd				= iCurColumnPos(4)
			C_BizUnitNm				= iCurColumnPos(5)
    End Select    
End Sub



'********************************************************************************************************* 
Function OpenPop(Byval strCode, Byval iWhere, ByVal Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_CostCd
	    	arrParam(1) = "b_cost_center"		' TABLE 명칭 
	    	arrParam(2) = Trim(strCode) 		' Code Condition
	    	arrParam(3) = "" 			' Name Cindition
	    	arrParam(4) = "cost_type <>" & FilterVar("C", "''", "S") & " " 		' Where Condition
	    	arrParam(5) = "Cost Center"		' TextBox 명칭 

	    	arrField(0) = "cost_cd"		 	    ' Field명(0)
	    	arrField(1) = "cost_nm"    		    ' Field명(1)%>

	    	arrHeader(0) = "Cost Center"	' Header명(0)%>
	    	arrHeader(1) = "Cost Center명"	' Header명(1)%>
	    	
	    Case C_BizUnitCd
	    	 
    		arrParam(0) = "사업부팝업"	
			arrParam(1) = "B_BIZ_UNIT"
			arrParam(2) = Trim(frm1.txtBizUnitCd.Value)
			arrParam(3) = ""											
			arrParam(4) = ""											
			arrParam(5) = "사업부"							
	
			arrField(0) = "biz_unit_cd"						
			arrField(1) = "biz_unit_nm"						
    
			arrHeader(0) = "사업부"				
    		arrHeader(1) = "사업부명"    	

    End Select


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet, iWhere)
	End If

End Function


'---------------------------------------------------------------------------------------------------------
Function SetCost(arrRet, iWhere)

	With frm1

		Select Case iWhere
		    Case C_CostCd
		        .vspdData.Col = C_CostCd
		    	.vspdData.text = arrRet(0) 
		        .vspdData.Col = C_CostNm
		    	.vspdData.text = arrRet(1)	
		    Case C_BizUnitCd
		        .vspdData.Col = C_BizUnitCd
		    	.vspdData.text = arrRet(0) 
		        .vspdData.Col = C_BizUnitNm
		    	.vspdData.text = arrRet(1)	
        End Select

	End With

End Function

Function OpenBizUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


			arrParam(0) = "사업부팝업"	
			arrParam(1) = "B_BIZ_UNIT"
			arrParam(2) = Trim(frm1.txtBizUnitCd.Value)
			arrParam(3) = ""											
			arrParam(4) = ""											
			arrParam(5) = "사업부"							
	
			arrField(0) = "biz_unit_cd"						
			arrField(1) = "biz_unit_nm"						
    
			arrHeader(0) = "사업부"				
    		arrHeader(1) = "사업부명"
    		
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizUnitCd.focus
		Exit Function
	Else
		frm1.txtBizUnitCd.focus
		frm1.txtBizUnitCd.Value = arrRet(0)
		frm1.txtBizUnitNm.value = arrRet(1)
	End If
		
End Function    		
    		


'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029     
    
    Call ggoOper.LockField(Document, "N")                      

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet 
    Call InitVariables
    
    Call SetDefaultVal

    Call SetToolbar("110011010010111")
    frm1.txtVerCd.focus 
    
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

    Dim iDx
    Dim IntRetCD,EFlag
    Dim grp_cd
    Dim C_CostCd
    Dim alloc_from

    EFlag = False

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col
	
'	currency_code = CStr(frm1.txtCurrencyCode.value)
	Select Case Col
		Case C_CostCd
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'=============================cost center 값 체크 시작 ==================================================
			alloc_from = Frm1.vspdData.Text
				If alloc_from <>"" Then
				    IntRetCD = CommonQueryRs("cost_nm","b_cost_center","cost_type <> " & FilterVar("C", "''", "S") & "  and cost_cd = " & FilterVar(alloc_from, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    If IntRetCD = False Then
					    Call DisplayMsgBox("124400","X","X","X")
					    Frm1.vspdData.Col = C_alloc_from
					    Frm1.vspdData.Text = ""
					    Frm1.vspdData.Col = C_cost_nm
					    Frm1.vspdData.Text = ""
					    Frm1.vspdData.Col = Col
					    Frm1.vspdData.Action = 0
					    Set gActiveElement = document.activeElement
					    EFlag = True
				    Else
					    Frm1.vspdData.Col = C_cost_nm
					    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
				    End If
				End If
	'=============================cost center 값 체크 끝 ==================================================
    End Select
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    '------ Developer Coding part (Start ) --------------------------------------------------------------
    '데이터 확인시 틀린데이터에 대해 undo 해준다.
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0

    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()
	End If
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub


'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"	'Split 상태코드 
        
    Set gActiveSpdSheet = frm1.vspdData

    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub



'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
	
End Sub


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
		   
		If Row > 0 And Col = C_Cost_Pop Then
		    .vspdData.Col = Col
		    .vspdData.Row = Row
		        
			.vspdData.Col = C_CostCd
		    Call OpenPop(.vspdData.Text,C_CostCd, Row)

		End If
		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")  
    End With
End Sub


'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_MinorNm Or NewCol <= C_MinorNm Then
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
    
    IF frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKey <> "" Then 
    	      	DbQuery
    	End If

    End if
    
End Sub


'========================================================================================================
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
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    IF DbQuery = False Then
		Exit Function
	END IF	
           
    FncQuery = True															
    
End Function


'========================================================================================================
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
		Exit Function
	END IF
    
    FncSave = True                                                          
    
End Function


'========================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow ,frm1.vspdData.ActiveRow
    
    frm1.vspdData.Col = C_CostCd
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_CostNm
    frm1.vspdData.Text = ""
    
    frm1.vspdData.col = C_BizUnitCd
	frm1.vspdData.Text = ""
	
	frm1.vspdData.col = C_BizUnitNm
	frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
End Function

'========================================================================================================
Function FncCancel()
    FncCancel = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCancel = True                                                             '☜: Processing is OK
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


'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function


'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   
End Function


'========================================================================================================
Function FncPrev() 
    On Error Resume Next
End Function


'========================================================================================================
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


'========================================================================================================
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


'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear 

	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001		
		strVal = strVal & "&txtBizUnitCd=" & .hBizUnitCd.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001		
		strVal = strVal & "&txtBizUnitCd=" & .txtBizUnitCd.value
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)	
        
    End With
    
    DbQuery = True
    
End Function


'========================================================================================================
Function DbQueryOk()	

    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")

	Call SetToolbar("110011110011111")
	
End Function


'========================================================================================================
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
	Dim strCostType
    Dim iColSep
    Dim iRowSep     

	
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    
	With frm1
		.txtMode.value = Parent.UID_M0002

    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
    iColSep = Parent.gColSep
	iRowSep = parent.gRowSep	
    
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag

				strVal = strVal & "C" & iColSep & lRow & iColSep 

				.vspdData.Col = C_CostCd

            Call CommonQueryRs("COST_TYPE","B_COST_CENTER","COST_CD =  " & FilterVar(.vspdData.Text, "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If Trim(Replace(lgF0,Chr(11),"")) = "X" then
				Call DisplayMsgBox("211210", vbInformation, "", "", I_MKSCRIPT)	
				'해당되는 Data가 없습니다.
				Exit Function
			Else
				strCostType = Trim(Replace(lgF0,Chr(11),""))
			End if

                strVal = strVal & Trim(.vspdData.Text) & iColSep
			    strVal = strVal & strCostType & iRowSep
                
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag

				strDel = strDel & "D" & iColSep & lRow & iColSep

                .vspdData.Col = C_CostCd	'1
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

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>배부C/C정의</font></td>
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
									<TD CLASS="TD5" NOWRAP>사업부</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtBizUnitCd"  SIZE=10  ALT ="사업부" tag="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizUnit() ">
													<INPUT NAME="txtBizUnitNm"  SIZE=25  ALT ="사업부명" tag="14X"></TD>
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
								<script language =javascript src='./js/ga005ma1_I299999934_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hBizUnitCd" tag="24" TABINDEX= "-1">

<INPUT  NAME="txtMaxRows3" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

