
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        :
*  3. Program ID           : GC001BA1
*  4. Program Name         : 경영손익 Profit Center별 손익작업 
*  5. Program Desc         : 경영손익 Profit Center별 손익작업 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/04
*  8. Modified date(Last)  : 2001/12/04
*  9. Modifier (First)     : Kwon Ki Soo
* 10. Modifier (Last)      : Kwon Ki Soo
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!-- #Include file="../../inc/uni2kcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID    = "gc001bb1.asp"

Const COOKIE_SPLIT  =  4877	                                                        'Cookie Split String

Dim C_RUN_FG          
Dim C_PROGRESS_FG     
Dim C_JOB_CD          
Dim C_JOB_SP          
Dim C_JOB_NM          
Dim C_YYYYMM          
Dim C_ERR_CNT         
Dim C_ERR_PB          

'Const C_SHEETMAXROWS    =   100	                             '한 화면에 보여지는 최대갯수*1.5%>
'Const C_SHEETMAXROWS_D  =   100                               '☆: Server에서 한번에 fetch할 최대 데이타 건수 

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgIsOpenPop
Dim IsOpenPop

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction

	'------ Developer Coding part (Start ) --------------------------------------------------------------

    lgStrPrevKey = ""                                           'initializes Previous Key
    lgLngCurRows = 0                                            'initializes Deleted Rows Count
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Dim StartDate
	StartDate	= "<%=GetSvrDate%>"                                               'Get Server DB Date
	
	frm1.txtWork_dt.focus
	frm1.txtWork_dt.text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.txtWork_dt, Parent.gDateFormat, 2)
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	<% Call loadInfTB19029A("Q", "G", "NOCOOKIE", "BA") %>

	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) --------------------------------------------------------------
   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
Sub MakeKeyStream(pOpt)
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.txtWork_dt.Text,frm1.txtWork_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth
    lgKeyStream = strYYYYMM & Parent.gColSep       'You Must append one character(Parent.gColSep)

    '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	ggoSpread.source = frm1.vspdData
	ggoSpread.SetCombo "Y" & vbtab & "N" , C_PROGRESS_FG
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	 C_RUN_FG =         1
	 C_PROGRESS_FG =    2
	 C_JOB_CD  =        3
	 C_JOB_SP  =        4
	 C_JOB_NM  =        5
	 C_YYYYMM  =        6
	 C_ERR_CNT =        7
	 C_ERR_PB  =        8
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	With frm1.vspdData

       .MaxCols = C_ERR_PB + 1                                                      ' ☜:☜: Add 1 to Maxcols

	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

       .MaxRows = 0
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021202", ,parent.gAllowDragDropSpread

	   .ReDraw = false
	   
	    Call GetSpreadColumnPos("A") 

       ggoSpread.SSSetCheck    C_RUN_FG,       "실행구분",      16,2,,True
       ggoSpread.SSSetCombo    C_PROGRESS_FG,  "기작업여부",    16,2
       ggoSpread.SSSetEdit     C_JOB_CD,       "작업코드",      20,0,,10,2
       ggoSpread.SSSetEdit     C_JOB_SP,       "작업SP",        20,0,,40,2
       ggoSpread.SSSetEdit     C_JOB_NM,       "작업코드명",    40,0,,50,2
       ggoSpread.SSSetEdit     C_YYYYMM,       "작업년월",      6,0,,6,2
       ggoSpread.SSSetFloat     C_ERR_CNT,      "ERROR COUNT",   20,3,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
       ggoSpread.SSSetButton   C_ERR_PB
	
	   Call ggoSpread.MakePairsColumn(C_ERR_CNT,C_ERR_PB)
	   
	   Call ggoSpread.SSSetColHidden(C_JOB_SP,C_JOB_SP,True)			
	   Call ggoSpread.SSSetColHidden(C_YYYYMM,C_YYYYMM,True)

	   .ReDraw = true

       Call SetSpreadLock

    End With
End Sub

'======================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
      ggoSpread.SpreadLock      C_JOB_CD , -1, C_JOB_CD
      ggoSpread.SpreadLock      C_JOB_SP , -1, C_JOB_SP
      ggoSpread.SpreadLock      C_JOB_NM , -1, C_JOB_NM
      ggoSpread.SpreadLock      C_YYYYMM , -1, C_YYYYMM
      ggoSpread.SpreadLock      C_ERR_CNT , -1, C_ERR_CNT
      ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False
      ggoSpread.SSSetProtected    C_JOB_CD  , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_JOB_SP  , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_JOB_NM  , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_YYYYMM  , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_ERR_CNT , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_ERR_PB  , pvStartRow, pvEndRow
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
           frm1.vspdData.Col = iDx
           frm1.vspdData.Row = iRow
           If frm1.vspdData.ColHidden <> True And frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              frm1.vspdData.Col = iDx
              frm1.vspdData.Row = iRow
              frm1.vspdData.Action = 0 ' go to
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

			C_RUN_FG				= iCurColumnPos(1)
			C_PROGRESS_FG 			= iCurColumnPos(2)
			C_JOB_CD     			= iCurColumnPos(3)    
			C_JOB_SP        		= iCurColumnPos(4)
			C_JOB_NM        		= iCurColumnPos(5)
			C_YYYYMM        		= iCurColumnPos(6)
			C_ERR_CNT       		= iCurColumnPos(7)
			C_ERR_PB         		= iCurColumnPos(8)
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
'	Call InitData()
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                            '⊙: Lock  Suitable  Field

    Call InitVariables                                                               '⊙: Setup the Spread sheet
	Call SetDefaultVal

	Call SetToolbar("1100000000001111")                                              '☆: Developer must customize
    Call InitSpreadSheet
    Call InitComboBox
    Call BtnDisabled(1)
	'------ Developer Coding part (End )   --------------------------------------------------------------
'	Call CookiePage (0)                                                              '☜: Check Cookie

End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
Function FncQuery()
    Dim IntRetCD
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    Call InitVariables                                                           '⊙: Initializes local global variables
'    Call SetDefaultVal
    Call MakeKeyStream("X")

	'------ Developer Coding part (End )   --------------------------------------------------------------

'    Call SetSpreadLock                                   '자동입력때 풀어준 부분을 다시 조회할때 Lock시킴 

    Call BtnDisabled(1)
    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data

    Set gActiveElement = document.ActiveElement
    FncQuery = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncNew()
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
Function FncDelete()
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function FncSave()
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncCopy()
    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncCancel()
    Dim lRow
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
'    With frm1
'        for lRow=1 to .vspdData.MaxRows
'            ggoSpread.EditUndo lRow
'        Next
'    End With
    ggoSpread.EditUndo
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncDeleteRow()
    FncDeleteRow = False                                                         '☜: Processing is NG
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrev()
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncNext()
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncExcel()
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	Call Parent.FncExport(Parent.C_SINGLE)
    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	Call Parent.FncFind(Parent.C_SINGLE, True)
    FncFind = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncExit()
	FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function DbQuery()
    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

	if LayerShowHide(1) = false then
	    Exit Function
	end if

	Dim strVal
	'------ Developer Coding part (Start)  --------------------------------------------------------------

    With frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
'       strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)            '☜: Max fetched data at a time
    End With

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQuery = True

End Function

'========================================================================================================
Function DbSave()
    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    DbSave = True                                                                '☜: Processing is OK
End Function

'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Sub DbQueryOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	Call BtnDisabled(0)
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub DbSaveOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub vspdData_Click(Col, Row)
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
	Call SetPopupMenuItemInf("0000111111")
   
End Sub

'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row
	End With
End Sub

'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim yyyymm
    Dim job_cd

	frm1.vspdData.Row = Row
	Select Case Col
        Case C_ERR_PB
'            frm1.vspdData.Col = C_ERR_CNT
'            If CInt(frm1.vspdData.Text) > 0 Then
                frm1.vspdData.Col = C_JOB_CD
                job_cd = frm1.vspdData.Text
                frm1.vspdData.Col = C_YYYYMM
                yyyymm = frm1.vspdData.Text
                Call OpenErr(yyyymm,job_cd)
'            End If
    End Select
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub  

'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub  

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKeyIndex <> "" Then
      	   Call DisableToolBar(Parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
    	End If
    End if
End Sub

'=======================================================================================================
Function ExeReflect()
	Dim IntRetCD
    Dim lGrpCnt
    Dim strVal
    Dim strDel
    Dim lRow

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	ExeReflect = False                                                          '⊙: Processing is NG

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")

	If IntRetCD = vbNo Then
		Exit Function
	End If

    if SpreadWorkingChk = false then
        Exit Function      'spread check box 체크 유무 
    end if

	if LayerShowHide(1) = false then
	    Exit Function
	end if

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

    Call MakeKeyStream("X")
	With frm1
       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = C_RUN_FG

           if .vspdData.value = 1 then
                .vspdData.Col = C_JOB_SP
                strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                lGrpCnt = lGrpCnt + 1
            End If
       Next

       .txtMode.value        = Parent.UID_M0006
       .txtKeyStream.value        = lgKeyStream
	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal
	End With


	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	                                       '☜: 비지니스 ASP 를 가동 
	ExeReflect = True                                                           '⊙: Processing is NG
End Function

'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD

	IntRetCD =DisplayMsgBox("990000","X","X","X")
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data
	Call LayerShowHide(0)

End Function

'=======================================================================================================
Function ExeReflectNo()
	Dim IntRetCD
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data
    Call LayerShowHide(0)
End Function

Function SpreadWorkingChk()
    Dim iRows
    Dim ichkCnt
    Dim IntRetCD

    SpreadWorkingChk = False
    ichkCnt = 0

    with frm1.vspdData
	    For iRows = 1 to .MaxRows
	        .Col =  C_RUN_FG
	        .Row =  iRows

	        if .Value = 1 then
		        .Col = C_PROGRESS_FG
		        if .Text = "Y" then
		            IntRetCD = DisplayMsgBox("236020","X","X","X")  '기작업구분이 Y 인 작업은 실행할 수 없습니다.
		            Exit Function
		        end if
		        ichkCnt = ichkCnt + 1
	        end if
	    Next

	    if ichkCnt = 0 then
	        IntRetCD = DisplayMsgBox("236021","X","X","X")  '선택된 작업이 없습니다.
	        Exit Function
        end if
    End With

    SpreadWorkingChk = True
End Function

'=======================================================================================================
Function OpenErr(yyyymm, job_cd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "BATCH JOB ERROR"		    	<%' 팝업 명칭 %>
	arrParam(1) = "G_ERROR"           <%' TABLE 명칭 %>
	arrParam(2) = ""                        <%' Code Condition%>
	arrParam(3) = "" 		            	<%' Name Cindition%>
	arrParam(4) = " YYYYMM = " & FilterVar(yyyymm, "''", "S") & " AND JOB_CD = " & FilterVar(job_cd, "''", "S")                       <%' Where Condition%>
	arrParam(5) = "BATCH JOB ERROR"

    arrField(0) = "ED08" & parent.gColSep & "SEQ"	     			            <%' Field명(1)%>
    arrField(1) = "ED200" & parent.gColSep & "ERROR_CONTENTS"					<%' Field명(0)%>


    arrHeader(0) = "SEQ"			    	<%' Header명(0)%>
    arrHeader(1) = "ERROR_CONTENTS"				<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable:Yes; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
'		Call SetCode(arrRet)
	End If

End Function

'========================================================================================================
Sub txtWork_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtWork_dt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtWork_dt.focus
	End If
End Sub

'========================================================================================================
Sub txtWork_dt_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
	End If
End Sub
</SCRIPT>

</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%= LR_SPACE_TYPE_00 %>>
	<TR>
		<TD NOWRAP  <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD NOWRAP >
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD NOWRAP  WIDTH=10>&nbsp;</TD>
					<TD NOWRAP  CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>P/C별 손익작업</font></td>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
                    <TD NOWRAP  WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD NOWRAP  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD NOWRAP  <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD NOWRAP  HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    	<TR>
									<TD NOWRAP  CLASS=TD5 NOWRAP>작업년월</TD>
									<TD NOWRAP  CLASS=TD656 NOWRAP><script language =javascript src='./js/gc001ba1_fpDateTime_txtWork_dt.js'></script></TD>
								</TR>
    					</TABLE>
					</TD>
				</TR>
				<TR>
					<TD NOWRAP  <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD NOWRAP  WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD NOWRAP  HEIGHT="100%">
									<script language =javascript src='./js/gc001ba1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD NOWRAP  <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=20>
		<TD NOWRAP >
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD NOWRAP  WIDTH=10>&nbsp;</TD>
					<TD NOWRAP ><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>작업실행</BUTTON></TD>
                    <TD NOWRAP  WIDTH=*>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD NOWRAP  WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


