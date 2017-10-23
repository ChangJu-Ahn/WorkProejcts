
<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<% Response.Expires = -1 %>
<!--
======================================================================================================
*  1. Module Name          : GA
*  2. Function Name        :
*  3. Program ID           : GA001MA1
*  4. Program Name         : 경영손익 option 등록 
*  5. Program Desc         : Single-Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/09
*  8. Modified date(Last)  : 2001/12/31
*  9. Modifier (First)     : song sang min
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
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

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit																	'☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "ga001MB1.asp"											'☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

Const TAB1 = 1																	'☜: Tab의 위치 

'Const C_SHEETMAXROWS_D   = 30													'☜: Fetch count at a time
Const C_SHEETMAXROWS     = 21
Const COOKIE_SPLIT       = 4877													'Cookie Split String


Dim C_OPTION_CODE
Dim C_OPTION_NAME
Dim C_MINOR_CD
Dim C_MINOR_PB
Dim C_MINOR_NM


'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop



'========================================================================================================
Sub InitVariables()
	
	lgIntFlgMode      = parent.OPMD_CMODE												'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False													'⊙: Indicates that no value changed
	lgIntGrpCount     = 0														'⊙: Initializes Group View Size
    lgStrPrevKey      = ""														'⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""														'⊙: initializes Previous Key Index
    lgSortKey         = 1														'⊙: initializes sort direction

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	lgStrPrevKey = ""															'initializes Previous Key
    lgLngCurRows = 0															'initializes Deleted Rows Count
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub



'========================================================================================================

Sub SetDefaultVal()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
		'------ Developer Coding part (End )   --------------------------------------------------------------
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   --------------------------------------------------------------
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
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
    Dim iDx
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'=
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Select Case Col
	    Case C_MINOR_PB
	        frm1.vspdData.Col = C_MINOR_CD
	        Call OpenMinor(frm1.vspdData.Text, 1, Row)
	End Select
			Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")   	
End Sub

'
'===========================================================================
Function OpenMinor(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim option_code
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	 Frm1.vspdData.Col = C_OPTION_CODE
	 Frm1.vspdData.Row = Row
     option_code       = Trim(Frm1.vspdData.Text)
     
     Frm1.vspdData.Col = C_MINOR_CD
  Select Case iWhere
	    Case 1
	    	arrParam(1) = "B_MINOR A, B_MAJOR B"								' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.vspdData.text)	 
	    	arrParam(3) = "" 													' Name Cindition
	    	arrParam(4) = "A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar(option_code, "''", "S") & "" ' Where Condition
	    	arrParam(5) = "선택항목"										' TextBox 명칭 

	    	arrField(0) = "A.MINOR_CD"											' Field명(0)
	    	arrField(1) = "A.MINOR_NM"    										' Field명(1)%>

	    	arrHeader(0) = "선택항목"										' Header명(0)%>
	    	arrHeader(1) = "선택항목명"										' Header명(1)%>
	End Select

    arrParam(3) = ""
  	arrParam(0) = arrParam(5)													' 팝업 명칭 

	   arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet, iWhere, Row)
	End If

End Function


'---------------------------------------------------------------------------------------------------------
Function SetCost(arrRet, iWhere, Row)


	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_MINOR_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_MINOR_NM
		    	.vspdData.text = arrRet(1)
		End Select

		lgBlnFlgChgValue = True

	End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Function


'========================================================================================================
Sub InitData()
End Sub


'========================================================================================================
Sub InitSpreadPosVariables()
	C_OPTION_CODE		=1
	C_OPTION_NAME		=2
	C_MINOR_CD			=3
	C_MINOR_PB			=4
	C_MINOR_NM			=5
End Sub


'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData

       .MaxCols   = C_MINOR_NM + 1												  ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols													  ' ☜:☜: Hide maxcols
       .ColHidden = True														  ' ☜:☜:

        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021121", ,parent.gAllowDragDropSpread
		ggoSpread.ClearSpreadData
				
       .ReDraw = false
        
        Call GetSpreadColumnPos("A")
        
        ggoSpread.SSSetEdit     C_OPTION_CODE      ,     "Code"  ,20,,,5,2
        ggoSpread.SSSetEdit     C_OPTION_NAME      ,     "Option",40,,,30,2
        ggoSpread.SSSetEdit     C_MINOR_CD         ,     "Value"  ,20,,,2,2
        ggoSpread.SSSetButton   C_MINOR_PB
        ggoSpread.SSSetEdit     C_MINOR_NM         ,     "선택항목명",30,,,50,2

		Call ggoSpread.MakePairsColumn(C_MINOR_CD,C_MINOR_PB)
		
	   .ReDraw = true

       Call SetSpreadLock
    End With

End Sub


'======================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	  ggoSpread.SpreadLock		C_OPTION_CODE, -1, C_OPTION_CODE
	  ggoSpread.SpreadLock		C_OPTION_NAME, -1, C_OPTION_NAME
	  ggoSpread.SpreadLock		C_MINOR_NM,	   -1, C_MINOR_NM
	  ggoSpread.SSSetRequired	C_MINOR_CD, -1, -1
	  ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True
    End With
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    .vspdData.ReDraw = False
     ggoSpread.SSSetProtected	C_OPTION_CODE, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_OPTION_NAME, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_MINOR_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_MINOR_CD, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub


'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0
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

			C_OPTION_CODE		= iCurColumnPos(1)
			C_OPTION_NAME		= iCurColumnPos(2)
			C_MINOR_CD			= iCurColumnPos(3)    
			C_MINOR_PB			= iCurColumnPos(4)
			C_MINOR_NM			= iCurColumnPos(5)
    End Select    
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub


'========================================================================================================
Sub Form_Load()
    Err.Clear																	'☜: Clear err status

	Call LoadInfTB19029															'☜: Load table , B_numeric_format

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
																				'Format Numeric Contents Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet														'Setup the Spread sheet
    Call InitVariables															'Initializes local global variables

    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100100100001111")
    Call CookiePage(0)

    '------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub


'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										  '⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
  ' Call SetDefaultVal
    Call InitVariables															  '⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									          '⊙: This function check indispensable field
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
 
   '------ Developer Coding part (End )   --------------------------------------------------------------
    If DbQuery = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncQuery = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "1")                                        '☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                        '☜: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Call SetToolbar("1100100100011111")
    Call SetDefaultVal
    Call InitVariables

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncNew = True															      '☜: Processing is OK
End Function


'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                 '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		                 '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbDelete = False Then                                                     '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncDelete = True                                                             '☜: Processing is OK
End Function


'========================================================================================================
Function FncSave()
    Dim IntRetCD

    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If


    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------

     Call MakeKeyStream("X")
	'------ Developer Coding part (End )   --------------------------------------------------------------

    If DbSave = False Then                                                       '☜: Query db data
    ' Call LayerShowHide(0)
       Exit Function
    End If
'    Set gActiveElement = document.ActiveElement
    FncSave = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	With Frm1

		If .vspdData.ActiveRow > 0 Then
			.vspdData.ReDraw = False

			ggoSpread.Source = frm1.vspdData
			ggoSpread.CopyRow
			SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	' Clear key field
	'----------------------------------------------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------

			.vspdData.ReDraw = True
			.vspdData.focus
		End If
	End With
    Set gActiveElement = document.ActiveElement
    FncCopy = True                                                                '☜: Processing is OK
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
Function FncInsertRow()
    
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next
    
    FncInsertRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	imRow = AskSpdSheetAddRowcount()
	
	If imRow = "" Then
		Exit function
	End If
		
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
       .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    
    IF Err.number = 0 Then
	    FncInsertRow = True                                                          '☜: Processing is OK
	End If
	
	
End Function


'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if
    With Frm1.vspdData
    	.focus
    	ggoSpread.Source = frm1.vspdData
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function


'========================================================================================================
Function FncPrint()
    FncPrint = False	                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True	                                                             '☜: Processing is OK
End Function


'========================================================================================================
Function FncPrev()
    Dim strVal
    Dim IntRetCD
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData


    Call SetDefaultVal
    Call InitVariables													         '⊙: Initializes local global variables

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
'   strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction


	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
    Set gActiveElement = document.ActiveElement
    FncPrev = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncNext()
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData


    Call SetDefaultVal
    Call InitVariables														     '⊙: Initializes local global variables

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
'   strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction


	'------ Developer Coding part (Start )   --------------------------------------------------------------
	'------ Developer Coding part (End   )   --------------------------------------------------------------

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
    Set gActiveElement = document.ActiveElement
    FncNext = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncExcel()
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function


'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub



'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			         '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    if LayerShowHide(1) = false then
	    Exit Function
	end if																		 '☜: Show Processing Message

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
'   strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
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
    Dim iColSep
    Dim iRowSep     

    Err.Clear                                                                    '☜: Clear err status

    DbSave = False                                                               '☜: Processing is NG
    if LayerShowHide(1) = false then
	    Exit Function
	end if																		 '☜: Show Processing Message

	'------ Developer Coding part (Start)  --------------------------------------------------------------
    With frm1
        .txtMode.value        = parent.UID_M0002                                        '☜: Delete
        .txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
    End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	

	With Frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

               Case ggoSpread.InsertFlag										'☜: Update

               Case ggoSpread.UpdateFlag										'☜: Update
                                                        strVal = strVal & "U" & iColSep
                                                        strVal = strVal & lRow & iColSep
                    .vspdData.Col = C_OPTION_CODE    : strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_MINOR_CD       : strVal = strVal & Trim(.vspdData.Text) & iRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag										'☜: Delete


           End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal

	End With

	'------ Developer Coding part (End )   --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
Function DbDelete()

    Err.Clear																		'☜: Clear err status
    DbDelete = False																'☜: Processing is NG
    'Call LayerShowHide(1)															'☜: Show Processing Message

	                            '☜: Delete
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------

    DbDelete = True																	'☜: Processing is OK
	'Call RunMyBizASP(MyBizASP, strVal)												'☜: Run Biz logic
End Function


'========================================================================================================
Sub DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE													'⊙: Indicates that current mode is Create mode
	'------ Developer Coding part (Start)  --------------------------------------------------------------



       Call SetToolbar("1100100100011111")											'버튼 툴바 제어 


	'------ Developer Coding part (End )   --------------------------------------------------------------
   ' Call InitData()
	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
End Sub


'========================================================================================================
Sub DbSaveOk()
	Call InitVariables
	'------ Developer Coding part (Start)  --------------------------------------------------------------
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
   Call SetToolbar("1100100100011111")
	'------ Developer Coding part (End )   --------------------------------------------------------------
    FncQuery()
    Set gActiveElement = document.ActiveElement
End Sub


'========================================================================================================
Sub DbDeleteOk()
  	'------ Developer Coding part (Start)  --------------------------------------------------------------
    FncQuery()
	Call SetToolbar("1100100100011111")
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call FncNew()
End Sub


'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD, Majorcd, Minorcd, EFlag

	EFlag = False

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Minorcd = Frm1.vspdData.Text

	Frm1.vspdData.Col = C_OPTION_CODE
	Majorcd = Frm1.vspdData.Text

	IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD =  " & FilterVar(Majorcd , "''", "S") & " And MINOR_CD =  " & FilterVar(Minorcd , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	IF Trim(Minorcd) <> "" Then
		If IntRetCD = False Then
			Call DisplayMsgBox("800054","X","X","X")
			Frm1.vspdData.Col = C_MINOR_CD
			Frm1.vspdData.Text = ""
			Frm1.vspdData.Col = C_MINOR_NM
			Frm1.vspdData.Text = ""
			Frm1.vspdData.Col = Col
			Frm1.vspdData.Action = 0
			Set gActiveElement = document.activeElement  
			EFlag = True
		Else
			Frm1.vspdData.Col = C_MINOR_NM
			Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
		End If
	END IF
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData,Col,Row)
	
	
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


'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	IF frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col,lgSortKey
            lgSortKey = 1
        End If
    End If

End Sub


'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
       If lgStrPrevKeyIndex <> "" Then
          lgCurrentSpd = "M"
          Call MakeKeyStream("X")
          DbQuery
       End If
    End if

End Sub


'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub


'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 	
End Sub


'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : cboYesNo_OnChange
'   Event Desc :
'========================================================================================================
Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
<BODY SCROLL="NO" TABINDDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<td nowrap <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<td nowrap WIDTH=10>&nbsp;</TD>
					<td nowrap CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td nowrap background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td nowrap background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>경영손익 Option 등록</font></td>
								<td nowrap background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<td nowrap WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<td nowrap WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<td nowrap WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<td nowrap HEIGHT="100%">
									<script language =javascript src='./js/ga001ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<td nowrap WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No  noresize framespacing=0 TABINDDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDDEX = "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

