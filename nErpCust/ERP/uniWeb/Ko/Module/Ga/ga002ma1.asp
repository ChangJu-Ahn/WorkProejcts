
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        :
*  3. Program ID           : GA002MA1
*  4. Program Name         : 경영손익 대분류 등록 
*  5. Program Desc         : 경영손익 대분류 등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/07
*  8. Modified date(Last)  : 2001/12/07
*  9. Modifier (First)     : Kwon Ki Soo
* 10. Modifier (Last)      : Lee Tae Soo
* 11. Comment                 :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/uni2kcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "ga002mb1.asp"						           '☆: Biz Logic ASP Name

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_OPTION_CODE      
Dim C_OPTION_NAME      
Dim C_MINOR_CD         
Dim C_MINOR_PB         
Dim C_MINOR_NM		 

'Const C_SHEETMAXROWS_D   = 30                                          '☜: Fetch count at a time
'Const C_SHEETMAXROWS     = 21
Const COOKIE_SPLIT       = 4877	                                      'Cookie Split String

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgIsOpenPop
Dim IsOpenPop


'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	lgStrPrevKey = ""                                           'initializes Previous Key
    lgLngCurRows = 0                                            'initializes Deleted Rows Count
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
   <% Call loadInfTB19029A("I", "G", "NOCOOKIE", "MA") %>  'batch= B , print = P , input = I
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
End Sub

'========================================================================================================
Sub MakeKeyStream(pOpt)
End Sub

'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
    Dim iDx
End Sub

'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Select Case Col
	    Case C_MINOR_PB
	        Call OpenMinor(Row)
	End Select
		Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")   	
End Sub

'===========================================================================
Function OpenMinor(Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function


	Frm1.vspdData.Col = C_MINOR_CD
	Frm1.vspdData.Row = Row

  
	IsOpenPop = True

    arrParam(1) = "B_MINOR"	' TABLE 명칭 
    arrParam(2) = Trim(frm1.vspdData.text) 	' Code Condition
    arrParam(3) = "" 				' Name Cindition
    arrParam(4) = "MAJOR_CD = " & FilterVar("G1005", "''", "S") & "" ' Where Condition
    arrParam(5) = "대분류"	' TextBox 명칭 

    arrField(0) = "MINOR_CD"		    ' Field명(0)
    arrField(1) = "MINOR_NM"    		' Field명(1)%>

    arrHeader(0) = "대분류"	' Header명(0)%>
    arrHeader(1) = "대분류명"	' Header명(1)%>

    arrParam(3) = ""
  	arrParam(0) = arrParam(5)								  ' 팝업 명칭 

	   arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet,Row)
	End If

End Function

'---------------------------------------------------------------------------------------------------------
Function SetCost(arrRet,Row)

	With frm1
        .vspdData.Row = Row
        .vspdData.Col = C_MINOR_CD
        .vspdData.text = arrRet(0)
        .vspdData.Col = C_MINOR_NM
        .vspdData.text = arrRet(1)

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

       .MaxCols   = C_MINOR_NM + 1                                                 ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021127", ,parent.gAllowDragDropSpread
		ggoSpread.ClearSpreadData
				
       .ReDraw = false

		Call GetSpreadColumnPos("A")       
  
        ggoSpread.SSSetEdit     C_OPTION_CODE      ,     "경영손익"  ,20,,,5,2
        ggoSpread.SSSetEdit     C_OPTION_NAME      ,     "경영손익명",40,,,30,2
        ggoSpread.SSSetEdit     C_MINOR_CD         ,     "대분류"  ,20,,,1,2
        ggoSpread.SSSetButton   C_MINOR_PB
        ggoSpread.SSSetEdit     C_MINOR_NM         ,     "대분류명",30,,,50,2
        
        Call ggoSpread.MakePairsColumn(C_MINOR_CD,C_MINOR_PB)

	   .ReDraw = true

       Call SetSpreadLock
    End With

End Sub

'======================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
      ggoSpread.SpreadLock      C_OPTION_CODE,	-1, C_OPTION_CODE
      ggoSpread.SpreadLock      C_OPTION_NAME,	-1, C_OPTION_NAME
      ggoSpread.SpreadLock      C_MINOR_NM,		-1, C_MINOR_NM
      ggoSpread.SSSetRequired	C_MINOR_CD,		-1, -1
      ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
Sub SetSpreadColor(lRow)
    With frm1
    .vspdData.ReDraw = False
     ggoSpread.SSSetProtected	C_OPTION_CODE, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_OPTION_NAME, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_MINOR_NM, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired	C_MINOR_CD, pvStartRow, pvEndRow
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
    Err.Clear                                                                        '☜: Clear err status

	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format

    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>
                                                                            <%'Format Numeric Contents Field%>
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>

    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100100100011111")
    Call CookiePage(0)
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
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
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "A")										  '☜: Clear Condition Field and Contents Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field

	Call SetToolbar("1100100100011111")
    Call SetDefaultVal
    Call InitVariables

    Set gActiveElement = document.ActiveElement
    FncNew = True															      '☜: Processing is OK
End Function

'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                 '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		                 '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

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

	imRow = AskSpdSheetAddRowCount()
	
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
    Set gActiveElement = document.ActiveElement
    
    IF Err.number = 0 Then
	    FncInsertRow = True                                                          '☜: Processing is OK
	End If
	                            '☜: Processing is OK
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
    Set gActiveElement = document.ActiveElement
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False	                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True	                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrev()
    Dim strVal
    Dim IntRetCD
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
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

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
'   strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction

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

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
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

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
'   strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
    Set gActiveElement = document.ActiveElement
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncExcel()
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(Parent.C_MULTI, True)

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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			         '⊙: Data is changed.  Do you want to exit?
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
	end if                                                       '☜: Show Processing Message

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
'   strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

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
	end if                                                     '☜: Show Processing Message

    With frm1
        .txtMode.value        = Parent.UID_M0002                                        '☜: Delete
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

               Case ggoSpread.InsertFlag                                      '☜: Update

               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                        strVal = strVal & "U" & iColSep
                                                        strVal = strVal & lRow & iColSep
                    .vspdData.Col = C_OPTION_CODE    : strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_MINOR_CD       : strVal = strVal & Trim(.vspdData.Text) & iRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete


           End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Sub DbQueryOk()

	lgIntFlgMode      = Parent.OPMD_UMODE                                                   '⊙: Indicates that current mode is Create mode
	'------ Developer Coding part (Start)  --------------------------------------------------------------



       Call SetToolbar("1100100100011111")									<%'버튼 툴바 제어 %>


	'------ Developer Coding part (End )   --------------------------------------------------------------
   ' Call InitData()
	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
Sub DbSaveOk()
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
   Call SetToolbar("1100100100011111")
    MainQuery()
    Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
Sub DbDeleteOk()
    MainQuery()
	Call SetToolbar("1100100100011111")
	Call MainNew()
End Sub

'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD, Majorcd, Minorcd, EFlag

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	EFlag = False

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Minorcd = Frm1.vspdData.Text

	Frm1.vspdData.Col = Col
	Majorcd = Frm1.vspdData.Text

	IF Trim(Minorcd) <> "" Then
		IntRetCD = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("G1005", "''", "S") & " AND MINOR_CD =  " & FilterVar(Minorcd , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If IntRetCD = False Then
			Call DisplayMsgBox("971001","X",Minorcd & " {에 해당하는 데이타}","X")
			Frm1.vspdData.Col = C_MINOR_CD
			Frm1.vspdData.Action = 0
			EFlag = True
		Else
			Frm1.vspdData.Col = C_MINOR_NM
			Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
		End If
	END IF

	Call CheckMinNumSpread(frm1.vspdData,Col,Row)
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    '데이터 확인시 틀린데이터에 대해 undo 해준다.
    If EFlag Then
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.EditUndo Row
		Set gActiveElement = document.ActiveElement
	End If
End Sub

'========================================================================================================
Sub vspdData_Click(Col, Row)
	 Call SetPopupMenuItemInf("0001111111")	
     gMouseClickStatus = "SPC"
     Set gActiveSpdSheet = frm1.vspdData

    If Row = 0 Then
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
      	   Call DisableToolBar(Parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
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
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
End Sub

</SCRIPT>
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<td nowrap background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>경영손익 대분류 등록</font></td>
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
									<script language =javascript src='./js/ga002ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<td nowrap WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No  noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDEX = "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

