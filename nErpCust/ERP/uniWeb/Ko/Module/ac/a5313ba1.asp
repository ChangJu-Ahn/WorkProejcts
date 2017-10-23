<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 회계관리
'*  3. Program ID           : A5402ba1
'*  4. Program Name         : 환평가작업
'*  5. Program Desc         : 환평가작업
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2006/04/10
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
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

Const BIZ_PGM_ID    = "a5313bb1.asp"

Dim C_RUN_FG          
Dim C_PROGRESS_FG
Dim C_PROGRESS_NM
Dim C_PROGRESS_DT
Dim C_JOB_CD          
Dim C_JOB_NM          
Dim C_ERR_CNT         
Dim C_ERR_PB          

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
	'frm1.txtWork_dt.text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	frm1.fpDateTime.Text =  UNIMonthClientFormat(parent.gFiscEnd)
	Call ggoOper.FormatDate(frm1.txtWork_dt, Parent.gDateFormat, 2)
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) --------------------------------------------------------------
   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	C_RUN_FG				=1
	C_PROGRESS_FG     		=2
	C_PROGRESS_NM           =3
	C_PROGRESS_DT           =4
	C_JOB_CD          		=5
	C_JOB_NM          		=6
	C_ERR_CNT               =7
	C_ERR_PB                =8
End Sub

'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	With frm1.vspdData

		.MaxCols = C_ERR_PB + 1                                                      ' ☜:☜: Add 1 to Maxcols

		.Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021128", ,parent.gAllowDragDropSpread

		ggoSpread.ClearSpreadData

		.ReDraw = False

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck    C_RUN_FG,       "{{실행구분}}",      16,2,,True
		ggoSpread.SSSetEdit	    C_PROGRESS_FG,  "",					 10,2,,10,2
		ggoSpread.SSSetEdit	    C_PROGRESS_NM,  "{{작업상태}}",		 20,2,,20,2   
		ggoSpread.SSSetEdit	    C_PROGRESS_DT,  "{{작업일자}}",		 20,2,,20,2      
		ggoSpread.SSSetEdit     C_JOB_CD,       "{{작업모듈}}",      20,2,,10,2
		ggoSpread.SSSetEdit     C_JOB_NM,       "{{작업모듈명}}",    20,2,,50,2
		ggoSpread.SSSetEdit     C_ERR_CNT,      "{{ERROR COUNT}}",   12,1,,50,2
		ggoSpread.SSSetButton   C_ERR_PB

		Call ggoSpread.MakePairsColumn(C_ERR_CNT,C_ERR_PB)
'		Call ggoSpread.MakePairsColumn(C_PROGRESS_FG,C_PROGRESS_NM)	   
		Call ggoSpread.SSSetColHidden(C_PROGRESS_FG,C_PROGRESS_FG,True)	   
		Call ggoSpread.SSSetColHidden(C_ERR_CNT,C_ERR_PB,True)	   
	   
		.ReDraw = True

		Call SetSpreadLock
    End With
End Sub

'======================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		  ggoSpread.SpreadLock      C_PROGRESS_FG , -1, C_PROGRESS_FG
		  ggoSpread.SpreadLock      C_PROGRESS_DT , -1, C_PROGRESS_DT
		  ggoSpread.SpreadLock      C_PROGRESS_NM , -1, C_PROGRESS_NM		  
		  ggoSpread.SpreadLock      C_JOB_CD , -1, C_JOB_CD
		  ggoSpread.SpreadLock      C_JOB_NM , -1, C_JOB_NM
		  ggoSpread.SpreadLock      C_ERR_CNT , -1, C_ERR_CNT
		  ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		  ggoSpread.SSSetProtected    C_PROGRESS_FG  , pvStartRow, pvEndRow
		  ggoSpread.SSSetProtected    C_PROGRESS_DT  , pvStartRow, pvEndRow
		  ggoSpread.SSSetProtected    C_PROGRESS_NM  , pvStartRow, pvEndRow				  		
		  ggoSpread.SSSetProtected    C_JOB_CD  , pvStartRow, pvEndRow
		  ggoSpread.SSSetProtected    C_JOB_NM  , pvStartRow, pvEndRow
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

			C_RUN_FG			= iCurColumnPos(1)
			C_PROGRESS_FG		= iCurColumnPos(2)
			C_PROGRESS_NM		= iCurColumnPos(3)
			C_PROGRESS_DT		= iCurColumnPos(4)
			C_JOB_CD 			= iCurColumnPos(5)    
			C_JOB_NM			= iCurColumnPos(6)
			C_ERR_CNT			= iCurColumnPos(7)
			C_ERR_PB 			= iCurColumnPos(8)    
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
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                            '⊙: Lock  Suitable  Field

    Call InitVariables                                                               '⊙: Setup the Spread sheet
	Call SetDefaultVal

	Call SetToolbar("1100000000001111")                                              '☆: Developer must customize
    Call InitSpreadSheet
'    Call InitComboBox
    Call BtnDisabled(1)
    call Fncquery()
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

    If Not chkField(Document, "1") Then									         '☜: This function check required field
		Exit Function
    End If

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    Call InitVariables                                                           '⊙: Initializes local global variables

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
    ggoSpread.EditUndo
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
	FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function DbQuery()
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    Dim strVal
    Dim iRetCd,IntRetCD

    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

	If LayerShowHide(1) = False Then
	    Exit Function
	End If

'	iRetCd = ChkExistVersion
	
	If iRetCd = True Then
		IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If	

    Call ExtractDateFrom(frm1.txtWork_dt.Text,frm1.txtWork_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth

    With frm1
		strVal = BIZ_PGM_ID & "?txtMode="           & Parent.UID_M0001
        strVal = strVal     & "&txtYYYYMM="         & strYYYYMM							'☜: Query Key
        strVal = strVal     & "&txtMaxRows="        & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
    End With

	Call RunMyBizASP(MyBizASP, strVal)													'☜: Run Biz Logic

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
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize
	Call BtnDisabled(0)
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

Private Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row

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
            frm1.vspdData.Col = C_JOB_CD
            job_cd = frm1.vspdData.Text
            yyyymm = frm1.txthWork_dt.value

            Call OpenErr(yyyymm,job_cd)
    End Select
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
    Dim strYYYYMM
    Dim strYear,strMonth,strDay    
    Dim strAns

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	ExeReflect = False                                                          '⊙: Processing is NG

	

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		Exit Function
    End If
    strAns = "N"	
	With frm1
       For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = C_RUN_FG

           If .vspdData.value = 1 Then
                .vspdData.Col = C_JOB_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                 .vspdData.Col = C_PROGRESS_NM 
                 if Trim(.vspdData.Text) ="Y" then
                 .vspdData.Col = C_JOB_CD 
                 
					 IntRetCD = DisplayMsgBox("A12137",Parent.VB_YES_NO,"[" & .vspdData.Text & "]" ,"x")

					If IntRetCD = vbNo Then
						Exit Function
					else 
					 strAns = "Y"	
					End If
	
                 end if
                 
                .vspdData.Col = C_PROGRESS_FG
                strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                lGrpCnt = lGrpCnt + 1
           End If
       Next


	if lGrpCnt =1 then
		 call DisplayMsgBox("181216","x","","x") 
		exit function 
    end if
    if strAns<>"Y" then 
    
		IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    end if 
	If LayerShowHide(1) = False Then
	    Exit Function
	End If
       .txtMode.value        = Parent.UID_M0006
	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strVal
	End With



	Call ExtractDateFrom(frm1.txtWork_dt.Text,frm1.txtWork_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    frm1.txthWork_dt.value = strYear & strMonth

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	                                       '☜: 비지니스 ASP 를 가동

	ExeReflect = True                                                           '⊙: Processing is NG
End Function

'=======================================================================================================
Function ExeReflectOk()				            
	Dim IntRetCD

	IntRetCD =DisplayMsgBox("183114","X","X","X")

    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field

    If DbQuery = False Then
		Exit Function
    End If                                                                 '☜: Query db data
    Call LayerShowHide(0)
End Function

'=======================================================================================================
Function ExeReflectNo()
	Dim IntRetCD

    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field

    If DbQuery = False Then
		Exit Function
    End If                                                                 '☜: Query db data
    Call LayerShowHide(0)
End Function

'=======================================================================================================
Function OpenErr(yyyymm, job_cd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim strYYYYMM
    Dim strYear,strMonth,strDay

    Call ExtractDateFrom(frm1.txtWork_dt.Text,frm1.txtWork_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "{{BATCH JOB ERROR}}"											<%' 팝업 명칭%>
	arrParam(1) = "A_EXCHANGE_ERROR"											<%' TABLE 명칭%>
	arrParam(2) = ""															<%' Code Condition%>
	arrParam(3) = "" 		            										<%' Name Cindition%>
	arrParam(4) = " YYYYMM = " & FilterVar(strYYYYMM, "''", "S") & " AND MODULE_CD = " & FilterVar(job_cd, "''", "S")                       <%' Where Condition%>
	arrParam(5) = "{{BATCH JOB ERROR}}"

    arrField(0) = "ED08" & parent.gColSep & "SEQ"	     			            <%' Field명(1)%>
    arrField(1) = "ED200" & parent.gColSep & "ERROR_CONTENTS"					<%' Field명(0)%>

    arrHeader(0) = "{{SEQ}}"			    									<%' Header명(0)%>
    arrHeader(1) = "{{ERROR_CONTENTS}}"											<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=615px; dialogHeight=450px; center: Yes; help: No; resizable: Yes; status: No;")

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

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	IF frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If

    If Row = 0 Then
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
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5" NOWRAP>{{작업년월}}</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript>ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%>  NAME="txtWork_dt" CLASS="FPDTYYYYMM" TAG="14X1" ALT="{{작업년월}}" Title="FPDATETIME"  id=fpDateTime></OBJECT></TD>')</script>
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
					<TD WIDTH=100% HEIGHT=100% VALIGN=top COLSPAN=4>
						<script language =javascript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>')</script>
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
					<TD NOWRAP ><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>{{작업실행}}</BUTTON></TD>
                    <TD NOWRAP  WIDTH=*>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD NOWRAP  WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txthWork_dt" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txthVercd" tag="24" TABINDEX = "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
