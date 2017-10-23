<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost Accounting
'*  2. Function Name        : Cost Work Version
'*  3. Program ID           : c1501mb1
'*  4. Program Name         : 실제원가계산 버젼 정보 등록 
'*  5. Program Desc         : 실제원가 계산시 현재 버젼 정보 
'*  6. Modified date(First) : 2000/11/08
'*  7. Modified date(Last)  : 2002/06/13
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : Cho Ig sung / Park Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================  -->

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

Const BIZ_PGM_ID = "C4018MB1.asp"	                         'Biz Logic ASP

Dim C_WorkStep	 
Dim C_WorkStepPop  
Dim C_WorkStepNm	 
Dim C_VerCd		 
Dim C_VerCdPop	 	
Dim C_TableId		 	
Dim C_ColumnId	 


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          


'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	 C_WorkStep			= 1
	 C_WorkStepPop		= 2
	 C_WorkStepNm		= 3
	 C_VerCd			= 4
	 C_VerCdPop			= 5	
	 C_TableId			= 6	
	 C_ColumnId			= 7
End Sub


'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode		= Parent.OPMD_CMODE 
    lgBlnFlgChgValue	= False
    lgIntGrpCount		= 0
    
    lgStrPrevKey		= ""
    lgLngCurRows		= 0
	lgSortKey			= 1
	    
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
sub SetDefaultVal()

	Dim StartDate
	Dim EndDate
	
	StartDate	= "<%=GetSvrDate%>"
	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)
		
	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
	With frm1.vspdData
	
    .MaxCols = C_ColumnId+1
	.Col = .MaxCols	
    .ColHidden = True
    
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread   

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	Call GetSpreadColumnPos("A")
    
    ggoSpread.SSSetEdit		C_WorkStep		, "작업단계코드"	, 36,,,2,2
	ggoSpread.SSSetButton	C_WorkStepPop    
    ggoSpread.SSSetEdit		C_WorkStepNm	, "작업단계명"		, 42
    ggoSpread.SSSetEdit		C_VerCd			, "버젼"			, 36,,,3,2
	ggoSpread.SSSetButton	C_VerCdPop    
    ggoSpread.SSSetEdit		C_TableId		, "테이블명"		, 30,,,20,2
    ggoSpread.SSSetEdit		C_ColumnId		, "컬럼명"			, 30,,,20,2

	call ggoSpread.MakePairsColumn(C_WorkStep,C_WorkStepPop)
	call ggoSpread.MakePairsColumn(C_VerCd,C_VerCdPop)  	

	Call ggoSpread.SSSetColHidden(C_TableId,C_TableId,True)
	Call ggoSpread.SSSetColHidden(C_ColumnId,C_ColumnId,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_WorkStep		, -1, C_WorkStep    
	ggoSpread.SpreadLock C_WorkStepPop	, -1, C_WorkStepPop    
    ggoSpread.SpreadLock C_WorkStepNm	, -1, C_WorkStepNm
	ggoSpread.SSSetRequired C_VerCd		, -1
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
										'Col             Row			 Row2
    ggoSpread.SSSetRequired		C_WorkStep		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected		C_WorkStepNm	,pvStartRow		,pvEndRow    
    ggoSpread.SSSetRequired  	C_VerCd			,pvStartRow		,pvEndRow
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

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
			C_WorkStep					= iCurColumnPos(1)
			C_WorkStepPop				= iCurColumnPos(2)
			C_WorkStepNm				= iCurColumnPos(3)    
			C_VerCd						= iCurColumnPos(4)
			C_VerCdPop					= iCurColumnPos(5)
			C_TableId					= iCurColumnPos(6)
			C_ColumnId					= iCurColumnPos(7)
    End Select    
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
Function OpenPopUp(Byval strCode, Byval strTbl,Byval strCol, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "작업단계팝업"  
			arrParam(1) = "B_MINOR"
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = "major_cd = 'C4018'"	
			arrParam(5) = "작업단계"

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"
    
			arrHeader(0) = "작업단계코드"
			arrHeader(1) = "작업단계명"
		Case 1
			arrParam(0) = "버젼팝업"
			arrParam(1) = strTbl
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "버젼"

			arrField(0) = strCol
    
			arrHeader(0) = "버젼"

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.vspdData.Col	= C_WorkStep
				.vspdData.Text	= arrRet(0)
				.vspdData.Col	= C_WorkStepNm
				.vspdData.Text	= arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			Case 1
				.vspdData.Col	= C_VerCd
				.vspdData.Text	= arrRet(0)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
		End Select

		lgBlnFlgChgValue = True
	End With
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029 
    
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet
    Call InitVariables

    Call SetDefaultVal

    Call SetToolbar("110011010010111")
    frm1.txtYYYYMM.focus
     
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

sub txtYYYYMM_DblClick(Button)
	If Button = 1 Then
		frm1.txtYYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYYYYMM.focus
	End If
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
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

sub vspdData_Change(ByVal Col, ByVal Row)
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
	
End Sub

sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strCode1
	Dim strCode2
	Dim strCode3
	Dim strWorkStep
	
	With frm1.vspdData 
		
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_WorkStepPop
				.Col = C_WorkStep
				.Row = Row
				
				strCode1 = Trim(.Text)

				Call OpenPopup(strCode1,,, 0)
			Case C_VerCdPop
				.Col = C_VerCd
				.Row = Row
				
				strCode1 = Trim(.Text)

				.Col = C_WorkStep
				strWorkStep = Trim(.Text)

				
				IF strWorkStep = "00" Then
					strCode2 = "C_MOVETYPE_CONFIGURATION_S"
				ELSEIf strWorkStep = "05" Then
					strCode2 = "c_mfc_dstb_rule_by_cc_s"
				ELSEIf strWorkStep = "06" Then
					strCode2 = "c_mfc_dstb_rule_by_wc_s"		
				ELSEIf strWorkStep = "07" Then
					strCode2 = "c_mfc_dstb_rule_by_order_s"									
				Else
					strCode2 = "C_DSTB_RULE_S"
				End If

				strCode3 = "VER_CD"

				Call OpenPopup(strCode1,strCode2,strCode3, 1)

		End Select
    End With
        Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")   	
    
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub



sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
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



Sub txtYYYYMM_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
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
    
'	Call InitSpreadSheet
    Call InitVariables
    
    If Not chkField(Document, "1") Then	
       Exit Function
    End If

    IF DbQuery = False Then
		Exit Function
	END IF
       
    FncQuery = True
    
End Function

function FncNew() 
    Dim IntRetCD 
    
    FncNew = False
    
    Err.Clear
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
			If IntRetCD = vbNo Then
				Exit Function
			End If
    End If
    
    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
     
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call SetDefaultVal
    
    FncNew = True

End Function

function FncSave() 
    Dim IntRetCD 
    
    FncSave = False
    
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False  Then 
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

    FncSave = True
    
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
    frm1.vspdData.Col = C_WorkStep
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_WorkStepNm
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
End Function
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


function FncCancel() 

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo 
    End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
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


function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    if frm1.vspdData.maxrows < 1 then exit function 
	   
    
    With frm1.vspdData 

    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow

    End With
End Function


function FncPrint()
    Call parent.FncPrint()
End Function


function FncPrev() 
End Function

function FncNext() 
End Function

function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub



'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub


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

function DbQuery() 
	Dim strVal
	Dim strYYYYMM
	Dim	strYear, strMonth, strDay
    
    DbQuery = False
    
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
    Err.Clear
    
    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    
    stryyyymm = strYear & strMonth
  
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtYYYYMM=" & strYYYYMM
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtYYYYMM=" & strYYYYMM
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    
		Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True

End Function

Function DbQueryOk()	

    lgIntFlgMode = Parent.OPMD_UMODE
    
    Call ggoOper.LockField(Document, "Q")	

	Call SetToolbar("110011110011111")
	
End Function


function DbSave() 
    Dim lRow        
    Dim lGrpCnt
	Dim strVal, strDel, strWorkStep
	Dim	strYear, strMonth, strDay
    Dim iColSep 
    Dim iRowSep   
    
    DbSave = False 
    
    IF LayerShowHide(1) = False Then
		Exit Function
	End If
    
    ' 날짜 처리함수 관련하여 MA쪽에서 Extract하여 MB쪽으로 넘겨줌 
    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    
    frm1.htxtYYYYMM.value = strYear & strMonth
      
	With frm1
		.txtMode.value = Parent.UID_M0002
		

		lGrpCnt = 1

		strVal = ""
		strDel = ""
		strWorkStep = ""
		
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	  
    

		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = 0
        
			Select Case .vspdData.Text

	            Case ggoSpread.InsertFlag
					strVal = strVal & "C" & iColSep & lRow & iColSep	
					.vspdData.Col = C_WorkStep	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					strWorkStep = Trim(.vspdData.Text)
					.vspdData.Col = C_VerCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					If strWorkStep = "05" Then
						strVal = strVal & "C_MFC_DSTB_RULE_BY_CC_S" & iColSep
					ElseIf strWorkStep = "06" Then
						strVal = strVal & "C_MFC_DSTB_RULE_BY_WC_S" & iColSep
					ElseIf strWorkStep = "07" Then
						strVal = strVal & "C_MFC_DSTB_RULE_BY_ORDER_S" & iColSep	
					ElseIf strWorkStep = "00" Then
						strVal = strVal & "C_MOVETYPE_CONFIGURATION_S" & iColSep						
					Else
						strVal = strVal & "C_DSTB_RULE_S" & iColSep
					End If
					strVal = strVal & "VER_CD" & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & iColSep & lRow & iColSep	
					.vspdData.Col = C_WorkStep	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					strWorkStep = Trim(.vspdData.Text)
					.vspdData.Col = C_VerCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					If strWorkStep = "05" Then
						strVal = strVal & "C_MFC_DSTB_RULE_BY_CC_S" & iColSep
					ElseIf strWorkStep = "06" Then
						strVal = strVal & "C_MFC_DSTB_RULE_BY_WC_S" & iColSep
					ElseIf strWorkStep = "07" Then
						strVal = strVal & "C_MFC_DSTB_RULE_BY_ORDER_S" & iColSep	
					ElseIf strWorkStep = "00" Then
						strVal = strVal & "C_MOVETYPE_CONFIGURATION_S" & iColSep						
					Else
						strVal = strVal & "C_DSTB_RULE_S" & iColSep
					End If
					strVal = strVal & "VER_CD" & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag	
					strDel = strDel & "D" & iColSep & lRow & iColSep	
					.vspdData.Col = C_WorkStep	
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
	frm1.vspddata.maxrows = 0
	Call MainQuery()
		
End Function


function DbDelete() 
End Function

Function ExeReflect()
    Dim IntRetCD
    Dim strVal

    Dim lRow
    Dim strYYYYMM
    Dim strYear,strMonth,strDay


    strVal = ""

	ExeReflect = False

    '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.txtYYYYMM.Text,frm1.txtYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM =   strYear & strMonth

    Call CommonQueryRs("count(*)","C_WORK_VERSION_S","yyyymm = " & FilterVar(strYYYYMM, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    if Trim(Replace(lgF0,Chr(11),"")) <> 0 then
	    IntRetCD = DisplayMsgBox("236071",Parent.VB_YES_NO,"버전","X")
    end if

    If IntRetCD = vbNo Then
		Exit Function
    End If

  
    if LayerShowHide(1) = false then
	    Exit Function
    end if


    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0003                     '☜: Query
    strVal = strVal     & "&txtYYYYMM="			 & strYYYYMM

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실제원가버젼관리</font></td>
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
									<TD CLASS="TD5">작업년월</TD>
									<TD CLASS="TD656" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYYYYMM" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="년월" id=txtYYYYMM> </OBJECT>');</SCRIPT>
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="htxtYyyyMm" tag="24" TABINDEX= "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

