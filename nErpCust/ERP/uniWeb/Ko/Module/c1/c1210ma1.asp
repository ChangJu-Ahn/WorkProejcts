
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 간접비배부율등록 
'*  3. Program ID           : c1210ma1
'*  4. Program Name         : 간접비배부율등록 
'*  5. Program Desc         : 공장별 표준계산시 간접비에 대한 배부율을 등록한다 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2002/06/05
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : Cho Ig Sung / Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit				

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "c1210mb1.asp"							'Biz Logic ASP 
 
Dim C_IndElmt 
Dim C_IndElmtPop 
Dim C_IndElmtNm 
Dim C_DirElmt 
Dim C_DirElmtPop 
Dim C_DirElmtNm 
Dim C_OverheadRate 


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

Dim lgIndPrevKey
Dim lgDirPrevKey


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  
	 C_IndElmt		= 1
	 C_IndElmtPop	= 2	
	 C_IndElmtNm	= 3
	 C_DirElmt		= 4
	 C_DirElmtPop	= 5	
	 C_DirElmtNm	= 6
	 C_OverheadRate	= 7
End Sub


'========================================================================================================
sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE									'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False									'⊙: Indicates that current mode is Create mode	
    lgIntGrpCount = 0 
    
    lgIndPrevKey = ""											'⊙: initializes Previous Key	
    lgLngCurRows = 0   
	lgSortKey = 1												'⊙: initializes sort direction
	    
End Sub


'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	With frm1.vspdData
	
    .MaxCols = C_OverheadRate+1	
	.Col = .MaxCols						
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false
	
   
    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit C_IndElmt, "간접원가요소코드", 20,,,6,2
	ggoSpread.SSSetButton C_IndElmtPop
    ggoSpread.SSSetEdit C_IndElmtNm, "간접원가요소명", 20
    ggoSpread.SSSetEdit C_DirElmt, "직접원가요소코드", 20,,,6,2
	ggoSpread.SSSetButton C_DirElmtPop
    ggoSpread.SSSetEdit C_DirElmtNm, "직접원가요소명", 20
    ggoSpread.SSSetFloat C_OverheadRate,"간접비배부율(%)", 33,Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z",,999999999


	call ggoSpread.MakePairsColumn(C_IndElmt,C_IndElmtPop)
	call ggoSpread.MakePairsColumn(C_DirElmt,C_DirElmtPop)	


	.ReDraw = true

'    ggoSpread.SSSetSplit(C_IndElmtNm)	
    Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_IndElmt		, -1, C_IndElmt
	ggoSpread.SpreadLock C_IndElmtPop	, -1, C_IndElmtPop
	ggoSpread.SpreadLock C_IndElmtNm	, -1, C_IndElmtNm
	ggoSpread.SpreadLock C_DirElmt		, -1, C_DirElmt
	ggoSpread.SpreadLock C_DirElmtPop	, -1, C_DirElmtPop
	ggoSpread.SpreadLock C_DirElmtNm	, -1, C_DirElmtNm
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                         'Col               Row           Row2
    ggoSpread.SSSetRequired		C_IndElmt		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_IndElmtNm		,pvStartRow		,pvEndRow    
    ggoSpread.SSSetRequired		C_DirElmt		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_DirElmtNm		,pvStartRow		,pvEndRow    
    ggoSpread.SSSetRequired  	C_OverheadRate	,pvStartRow		,pvEndRow
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
			C_IndElmt				= iCurColumnPos(1)
			C_IndElmtPop			= iCurColumnPos(2)
			C_IndElmtNm				= iCurColumnPos(3)    
			C_DirElmt				= iCurColumnPos(4)
			C_DirElmtPop			= iCurColumnPos(5)
			C_DirElmtNm				= iCurColumnPos(6)
			C_OverheadRate			= iCurColumnPos(7)
    End Select    
End Sub



Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""			
	arrParam(4) = ""			
	arrParam(5) = "공장"		
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"		
    
    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If
		
End Function

Function OpenCostElmt(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "원가요소팝업"
	arrParam(1) = "C_COST_ELMT"	
	arrParam(2) = strCode
	arrParam(3) = ""		
	If iWhere = 1 Then
		arrParam(4) = "DI_FLAG= " & FilterVar("I", "''", "S") & " "			
	Else
		arrParam(4) = "DI_FLAG= " & FilterVar("D", "''", "S") & " "	
	End If
	arrParam(5) = "원가요소"   

	arrField(0) = "COST_ELMT_CD"			
	arrField(1) = "COST_ELMT_NM"		
    
	arrHeader(0) = "원가요소코드"	   		
	arrHeader(1) = "원가요소명"					
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCostElmt(arrRet, iWhere)
	End If	

End Function

Function SetPlant(byval arrRet)
	frm1.txtPlantCd.focus
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

Function SetCostElmt(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.vspdData.Col = C_IndElmt
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_IndElmtNm
				.vspdData.Text = arrRet(1)
				
				Call vspddata_Change(.vspddata.col, .vspddata.row)
				.vspdData.Col = C_IndElmt
				.vspdData.Action = 0
				
			Case 4
				.vspdData.Col = C_DirElmt
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DirElmtNm
				.vspdData.Text = arrRet(1)
			
				Call vspddata_Change(.vspddata.col, .vspddata.row)
				.vspdData.Col = C_DirElmt
				.vspdData.Action = 0
		End Select


	End With
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================

sub Form_Load()

    Call LoadInfTB19029 

    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
  
    Call InitSpreadSheet       
    Call InitVariables
   
    Call SetDefaultVal
    Call SetToolbar("110011010010111")			
    frm1.txtPlantCd.focus

   	Set gActiveElement = document.activeElement			
     
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
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

	'------ Developer Coding part (End   ) --------------------------------------------------------------         
	
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


sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub


sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_IndElmtPop
				.vspdData.Col = Col
				.vspdData.Row = Row

				.vspdData.Col = C_IndElmt
				Call OpenCostElmt(.vspdData.Text, 1)
			Case C_DirElmtPop        
				.vspdData.Col = Col
				.vspdData.Row = Row

				.vspdData.Col = C_DirElmt
				Call OpenCostElmt(.vspdData.Text, 4)
		End Select
			Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_IndElmtNm Or NewCol <= C_IndElmtNm Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	End If	

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgIndPrevKey <> "" Then        
      	DbQuery
    	End If

    End if
    
End Sub


'========================================================================================================
function FncQuery() 
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
    															
    if frm1.txtPlantCd.value = "" then
		frm1.txtPlantNm.value = ""
    end if
    
    If Not chkField(Document, "1") Then		
       Exit Function
    End If

    IF DbQuery = False Then
		Exit function	
    END If
       
    If Err.number = 0 Then	
       FncQuery = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False																  '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
function FncSave() 
    Dim IntRetCD 
    
    FncSave = False             
    
    Err.Clear 
    
    ggoSpread.Source = frm1.vspddata
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
    
	If DbSave = False Then
		Exit Function
	End If
    
    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                                         
    
End Function

'========================================================================================================
function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_IndElmt
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_IndElmtNm
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_DirElmt
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_DirElmtNm
    frm1.vspdData.Text = ""
    
	frm1.vspdData.ReDraw = True
End Function


'========================================================================================================
function FncCancel() 

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


'========================================================================================================
function FncDeleteRow() 
    Dim lDelRows
    
    if frm1.vspdData.maxrows < 1 then exit function 
	   
    
    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
End Function


'========================================================================================================
function FncPrint()
    Call parent.FncPrint()
End Function


'========================================================================================================
function FncPrev() 
	On Error Resume Next
End Function


'========================================================================================================
function FncNext() 
	On Error Resume Next  
End Function

function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)		
End Function

function FncFind() 
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
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub


'========================================================================================================
function FncExit()
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
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
function DbQuery() 
	Dim strVal

    DbQuery = False
    
	IF LayerShowHide(1) = False Then
		Exit Function                                        
	End If
	
    Err.Clear 

    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
 
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgIndPrevKey=" & lgIndPrevKey
			strVal = strVal & "&lgDirPrevKey=" & lgDirPrevKey
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)	 	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    	Else
	    	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
			strVal = strVal & "&lgIndPrevKey=" & lgIndPrevKey
			strVal = strVal & "&lgDirPrevKey=" & lgDirPrevKey
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
	
		Call RunMyBizASP(MyBizASP, strVal)	
        
    End With
    
    DbQuery = True

End Function


'========================================================================================================
function DbQueryOk()					
	
    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")	
	Call SetToolbar("110011110011111")	
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
function DbSave() 
    Dim lRow        
    Dim lGrpCnt
    Dim iColSep
    Dim iRowSep     
	Dim strVal, strDel
	
    DbSave = False        
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If	
	
	With frm1
		.txtMode.value = Parent.UID_M0002
		
		lGrpCnt = 1

		strVal = ""
	    strDel = ""
    
		iColSep = Parent.gColSep
		iRowSep = Parent.gRowSep	

		For lRow = 1 To .vspdData.MaxRows

			.vspdData.Row = lRow
			.vspdData.Col = 0
        
			Select Case .vspdData.Text

	            Case ggoSpread.InsertFlag		

					strVal = strVal & "C" & iColSep & lRow & iColSep

					.vspdData.Col = C_IndElmt	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_DirElmt
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_OverheadRate	
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep

					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag		

					strVal = strVal & "U" & iColSep & lRow & iColSep	

					.vspdData.Col = C_IndElmt	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_DirElmt	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_OverheadRate	
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep

					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag		

					strDel = strDel & "D" & iColSep & lRow & iColSep	

					.vspdData.Col = C_IndElmt		
					strDel = strDel & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_DirElmt	
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


'========================================================================================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>간접비배부율등록</font></td>
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
									<TD CLASS="TD656" COLSPAN=3><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
										 <INPUT TYPE=TEXT ID="txtPlantNm" NAME="txtPlantNm" SIZE=25 tag="14X">
									</TD>
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
								<script language =javascript src='./js/c1210ma1_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


