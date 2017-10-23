<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        :
'*  3. Program ID           : A5955MA1
'*  4. Program Name         : a5955ma1
'*  5. Program Desc         : 월차결산 작업 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/09
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : 권기수 
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->


<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">           </SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'========================================================================================================

Const BIZ_PGM_ID    = "a5955mb1.asp"
Const BIZ_PGM_JUMP_ID   = "a5953ma1"
Const COOKIE_SPLIT  =  4877	                                                        'Cookie Split String
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================

Dim C_RUN_FG 
Dim C_CANCEL_FG 
Dim C_PROGRESS_FG
Dim C_PROGRESS_FG_CD
Dim C_BIZ_AREA_CD 
Dim C_JOB_CD  
Dim C_JOB_SP  
Dim C_JOB_NM  
Dim C_YYYYMM  
Dim C_ERR_CNT 
Dim C_FLAG 
Dim C_ERR_PB  


'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop
dim StartDate
dim LastDate


'========================================================================================================
Sub InitSpreadPosVariables()
	C_RUN_FG			= 1
	C_CANCEL_FG			= 2
	C_PROGRESS_FG		= 3
	C_PROGRESS_FG_CD	= 4
	C_BIZ_AREA_CD		= 5
	C_JOB_CD			= 6
	C_JOB_SP			= 7
	C_JOB_NM			= 8
	C_YYYYMM			= 9
	C_ERR_CNT			= 10
	C_FLAG				= 11
	C_ERR_PB			= 12
End Sub

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
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================

Sub SetDefaultVal()
	
'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim StartDate
	Dim LastDate			
	Dim strYear, strMonth, strDay
	
	StartDate	= "<%=GetSvrDate%>"
                                              'Get Server DB Date
	LastDate     = UNIGetLastDay(StartDate,Parent.gServerDateFormat)     
	
	Call ExtractDateFrom(StartDate,Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	frm1.txtWork_dt.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
'	frm1.txtWork_dt.Year	= strYear	
'	frm1.txtWork_dt.Month	= strMonth
'	frm1.txtWork_dt.Day		= strDay
'	frm1.txtWorkdate.text	= UNIDateClientFormat(LastDate)

	Call ExtractDateFrom(LastDate,Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)

	frm1.txtWorkdate.text	=  UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)

	Call ggoOper.FormatDate(frm1.txtWork_dt, Parent.gDateFormat, 2)		
	Call ggoOper.FormatDate(frm1.txtWorkdate, Parent.gDateFormat, 1)	
	
	'------ Developer Coding part (Start ) --------------------------------------------------------------
   
	frm1.txtWork_dt.focus

	
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("B", "A", "NOCOOKIE", "MA") %>
	
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value
'========================================================================================================
Sub CookiePage(Kubun)
'	Dim iRow,iCol
	Dim strYear,strMonth,strDay
	Dim TempFrDt
   '------ Developer Coding part (Start ) --------------------------------------------------------------       
	With Frm1                      	
	 Select Case Kubun		
		Case 0
			
			If ReadCookie("JumpFlag")	<>""	Then .txtJumpFlag.Value		= ReadCookie("JumpFlag")
			
			If UCase(Trim(.txtJumpFlag.Value)) = "A5955MA1" Then
				If ReadCookie("FrYYYYMM")	<>"" Then TempFrDt				= ReadCookie("FrYYYYMM")	
				Call ExtractDateFrom(TempFrDt, Parent.gDateFormat, Parent.gComDateType, strYear, strMonth, strDay)
				.txtWork_dt.Year	= strYear
				.txtWork_dt.Month	= strMonth
							
     		End If
			
			If ReadCookie("Unt_Code")	<>"" Then .txtBizAreaCd.value 			= ReadCookie("Unt_Code")
			
			If Trim(.txtWork_dt.Text) <> "" and Trim(.txtWorkDate.Text) <> "" and Trim(.txtBizAreaCd.value) <> ""  Then
				Call MainQuery()
      		End If		
      		
			WriteCookie "FrYYYYMM"		, ""
			WriteCookie "Unt_Code"      , ""
      		WriteCookie "JumpFlag"		, ""
					
		Case 1			
   		
		    WriteCookie "FrYYYYMM" , UniConvYYYYMMDDToDate(Parent.gDateFormat,Trim(.txtWork_dt.Year),Right("0" & Trim(.txtWork_dt.Month),2),"01")
		    WriteCookie "Unt_Code" , .txtBizAreaCd.value
		    WriteCookie "JumpFlag" , "a5953ma1"
	End Select
	
	End With
   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
'	Desc : 화면이동 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 계속하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

'   Call CookiePage(strPgmId)
   Call PgmJump(strPgmId)
End Function

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    Dim strYYYYMM, strYYYYMMDD 
    Dim strYear,strMonth,strDay
    Dim tstr
    '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.txtWork_dt.Text,frm1.txtWork_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strYYYYMM = strYear & strMonth
	
	Call ExtractDateFrom(frm1.txtWorkDate.Text,frm1.txtWorkDate.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strYYYYMMDD = strYear & strMonth & strDay

	
    lgKeyStream = strYYYYMM & Parent.gColSep
    lgKeyStream = lgKeyStream & Trim(frm1.txtBizAreaCd.value) & Parent.gColSep
    lgKeyStream = lgKeyStream & strYYYYMMDD & Parent.gColSep

   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
	ggoSpread.source = frm1.vspdData
	ggoSpread.SetCombo "Y" & vbTab & "N", C_PROGRESS_FG
	ggoSpread.SetCombo "Y" & vbTab & "N", C_PROGRESS_FG_CD
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_PROGRESS_FG_CD
			intIndex = .value
			.col = C_PROGRESS_FG
			.value = intindex						
		Next	
	End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021103",,Parent.gAllowDragDropSpread            '

	With frm1.vspdData

	   .ReDraw = false
	   
       .MaxCols   = C_ERR_PB + 1                                                    ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

	ggoSpread.Source= frm1.vspdData
    ggoSpread.ClearSpreadData
       
    call GetSpreadColumnPos("A")
              
		ggoSpread.SSSetCheck    C_RUN_FG,       "작업실행",      13,2,,True
		ggoSpread.SSSetCheck    C_CANCEL_FG,    "작업취소",      13,2,,True
		ggoSpread.SSSetCombo    C_PROGRESS_FG,  "기작업여부",    15,2
		ggoSpread.SSSetCombo    C_PROGRESS_FG_CD,  "기작업여부코드",    15,2
		ggoSpread.SSSetEdit     C_BIZ_AREA_CD,  "사업장",      10,0,,10,2
		ggoSpread.SSSetEdit     C_JOB_CD,       "작업코드",      13,0,,10,2
		ggoSpread.SSSetEdit     C_JOB_SP,       "작업SP",        20,0,,40,2
		ggoSpread.SSSetEdit     C_JOB_NM,       "월차결산 작업명",    38,0,,50,2
		ggoSpread.SSSetEdit     C_YYYYMM,       "작업년월",      6,0,,6,2
		ggoSpread.SSSetFloat    C_ERR_CNT,      "ERROR COUNT",   20,3,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit     C_FLAG,		   "기준작업여부",      6,0,,6,2
		ggoSpread.SSSetButton   C_ERR_PB
		
		Call ggoSpread.SSSetColHidden(C_BIZ_AREA_CD,C_BIZ_AREA_CD,True)
		Call ggoSpread.SSSetColHidden(C_JOB_SP,C_JOB_SP,True)
		Call ggoSpread.SSSetColHidden(C_YYYYMM,C_YYYYMM,True)
		Call ggoSpread.SSSetColHidden(C_FLAG,C_FLAG,True)
		Call ggoSpread.SSSetColHidden(C_PROGRESS_FG_CD,C_PROGRESS_FG_CD,True)
		Call ggoSpread.SSSetColHidden(C_ERR_CNT,C_ERR_CNT,True)
		Call ggoSpread.SSSetColHidden(C_ERR_PB,C_ERR_PB,True)
		
	   .ReDraw = true

       Call SetSpreadLock

    End With
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    Dim iRows
    
    With frm1.vspdData
     .ReDraw = False
		ggoSpread.SpreadLock      C_JOB_CD , -1, C_JOB_CD 
		ggoSpread.SpreadLock      C_JOB_SP , -1, C_JOB_SP 
		ggoSpread.SpreadLock      C_JOB_NM , -1, C_JOB_NM 
		ggoSpread.SpreadLock	.MaxCols, -1,.MaxCols
    .ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    .vspdData.ReDraw = False
		ggoSpread.SSSetProtected    C_JOB_CD  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_JOB_SP  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_JOB_NM  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_YYYYMM  , pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected    C_ERR_CNT , lRow, lRow
		ggoSpread.SSSetProtected    C_ERR_PB  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_PROGRESS_FG  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_PROGRESS_FG_CD  , pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadColorUser()  
	Dim lRow

    With frm1
     ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False
  For lRow = 1 To .vspdData.MaxRows
        
    .vspdData.Row = lRow
   	.vspdData.Col = C_FLAG
  
   	If Trim(.vspdData.Text) = "N" Then
		ggoSpread.SSSetProtected    C_RUN_FG  , lRow, lRow
		ggoSpread.SSSetProtected    C_CANCEL_FG  , lRow, lRow
		ggoSpread.SSSetProtected    C_BIZ_AREA_CD , lRow, lRow   
    End If
    
    ggoSpread.SSSetProtected    C_PROGRESS_FG  , lRow, lRow
	ggoSpread.SSSetProtected    C_PROGRESS_FG_CD  , lRow, lRow
    
    .vspdData.Col = C_PROGRESS_FG
    
	if Trim(.vspdData.Text) = "N" Then
		ggoSpread.SSSetProtected    C_CANCEL_FG , lRow, lRow
	end if	
	
	if Trim(.vspdData.Text) = "Y" Then
		ggoSpread.SSSetProtected    C_RUN_FG , lRow, lRow
	end if
    
  next  
   .vspdData.ReDraw = True
    End With
End Sub
'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
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

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_RUN_FG			=  iCurColumnPos(1)
			C_CANCEL_FG			=  iCurColumnPos(2)
			C_PROGRESS_FG		=  iCurColumnPos(3)
			C_PROGRESS_FG_CD	=  iCurColumnPos(4)
			C_BIZ_AREA_CD		=  iCurColumnPos(5)
			C_JOB_CD			=  iCurColumnPos(6)
			C_JOB_SP			=  iCurColumnPos(7)
			C_JOB_NM			=  iCurColumnPos(8)
			C_YYYYMM			=  iCurColumnPos(9)
			C_ERR_CNT			=  iCurColumnPos(10)
			C_FLAG				=  iCurColumnPos(11)
			C_ERR_PB			=  iCurColumnPos(12)
			
			
    End Select    
   
End Sub   


'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
    Call InitVariables
	Call SetDefaultVal
	Call InitData

	Call SetToolbar("1100000000001111")
    Call InitSpreadSheet
    Call InitComboBox
	Call CookiePage (0)                                                              '☜: Check Cookie
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
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
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
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
' Name : FncInsertRow
' Desc : developer describe this line Called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    FncDeleteRow = False                                                         '☜: Processing is NG
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
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
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev()
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncPrev = True                                                               '☜: Processing is OK
End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext()
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel()
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	Call Parent.FncExport(Parent.C_SINGLE)
    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	Call Parent.FncFind(Parent.C_SINGLE, True)
    FncFind = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	FncExit = True                                                               '☜: Processing is OK
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	Call SetSpreadColorUser()
End Sub

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
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
    End With

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQuery = True

End Function
'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    DbSave = True                                                                '☜: Processing is OK
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    DbDelete = True                                                              '☜: Processing is OK
End Function
'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize
	Call InitData()
	Call SetSpreadColorUser()
	Call BtnDisabled(0)
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
End Sub

'========================================================================================================
' Name : DbDeleteOk
' Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

Private Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
        Select Case Col
            Case C_PROGRESS_FG
                .Col = Col
                intIndex = .Value												'COMBO의 VALUE값 
				.Col = C_PROGRESS_FG_CD												'CODE값란으로 이동 
				.Value = intIndex												'CODE란의 값은 COMBO의 VALUE값이된다.
		End Select
	End With	

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생하는 콤보 박스 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
   
   Call SetPopupMenuItemInf("0000111111")  
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
    
    	frm1.vspdData.Row = Row

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevKeyIndex <> "" Then
      	   Call DisableToolBar(Parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
        End If
    End if
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
' Name : ExeReflect
' Desc :
'=======================================================================================================
Function ExeReflect()
	Dim IntRetCD
    Dim lGrpCnt
    Dim strVal
    Dim strDel
    Dim lRow
    Dim run_fg
    Dim cancel_fg

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
            run_fg = .vspdData.value
            .vspdData.Col = C_CANCEL_FG
            cancel_fg = .vspdData.value

            if run_fg = 1 then
                strVal = strVal & "R" & Parent.gColSep
				.vspdData.Col = C_JOB_SP
				strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				lGrpCnt = lGrpCnt + 1
            Elseif cancel_fg = 1 then
                strVal = strVal & "C" & Parent.gColSep
				.vspdData.Col = C_JOB_SP
				strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				lGrpCnt = lGrpCnt + 1
            End If

       Next

       .txtMode.value        = Parent.UID_M0006
       .txtKeyStream.value   = lgKeyStream
	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strVal
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	                                       '☜: 비지니스 ASP 를 가동 
	ExeReflect = True                                                           '⊙: Processing is NG
End Function

'======================================================================================================
' Name : ExeReflectOk
' Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD

	IntRetCD =DisplayMsgBox("990000","X","X","X")
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data
    Call LayerShowHide(0)
End Function

'======================================================================================================
' Name : ExeReflectNo
' Desc :
'=======================================================================================================
Function ExeReflectNo()
	Dim IntRetCD
    'Call DisplayMsgBox("800407","X","X","X") 				            '☆: 실행된 자료가 없습니다 
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data
    Call LayerShowHide(0)
End Function

Function SpreadWorkingChk()
    Dim iRows
    Dim ichkCnt
    Dim IntRetCD
    Dim run_fg
    Dim cancel_fg

    SpreadWorkingChk = False
    ichkCnt = 0

    with frm1.vspdData
	    For iRows = 1 to .MaxRows
            .Row =  iRows
	        .Col =  C_RUN_FG
	        run_fg = .Value
	        .Col =  C_CANCEL_FG
	        cancel_fg = .Value

	        if run_fg =1 or cancel_fg = 1 then
		        .Col = C_PROGRESS_FG
		     '   if .Text = "Y" then
		     '      IntRetCD = DisplayMsgBox("236020","X","X","X")  '기작업구분이 Y 인 작업은 실행할 수 없습니다.
		     '       Exit Function
		     '   end if
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

<%'======================================================================================================
'	Name : OpenBizArea)
'	Description : Major PopUp
'=======================================================================================================%>
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"		    	 <%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_AREA"					 <%' TABLE 명칭 %>
	arrParam(2) = frm1.txtBizAreaCd.value        <%' Code Condition%>
	arrParam(3) = "" 		            	<%' Name Cindition%>
	arrParam(4) = ""                        <%' Where Condition%>
	arrParam(5) = "사업장"

    arrField(0) = "BIZ_AREA_CD"	     			<%' Field명(1)%>
    arrField(1) = "BIZ_AREA_NM"					<%' Field명(0)%>


    arrHeader(0) = "사업장코드"			    	<%' Header명(0)%>
    arrHeader(1) = "사업장명"				<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetBizArea(arrRet)
	End If

End Function

'========================================================================================================
'	Name : SetBizArea()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetBizArea(Byval arrRet)
	With frm1
		.txtBizAreaCd.focus
		.txtBizAreaCd.value = arrRet(0)
		.txtBizArea.value	   = arrRet(1)
	End With
End Function





'========================================================================================================
Sub txtWork_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtWork_dt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtWork_dt.Focus
	End If
End Sub

Sub txtWorkDate_DblClick(Button)
	If Button = 1 Then
		frm1.txtWorkDate.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtWorkDate.Focus
	End If
End Sub

Sub txtWork_dt_Change()
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
	
'	If CheckDateFormat(frm1.txtWorkdate.Text,Parent.gDateFormat) Then 
	Call ExtractDateFrom(frm1.txtWork_dt.Text,frm1.txtWork_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	strYYYYMM = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)

	frm1.txtWorkdate.Text  = UNIGetLastDay (strYYYYMM,Parent.gDateFormat) 
'	End If	 
 
End Sub

'========================================================================================================
'   Event Name : txtWork_dt_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtWork_dt_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
	End If
End Sub

Sub txtWorkdate_Change()
    Dim strYYYYMM, strYYYYMMDD
    Dim strYear, strMonth, strDay
    Dim IntRetCD
 
	If CheckDateFormat(frm1.txtWorkdate.Text,Parent.gDateFormat) Then 
		Call ExtractDateFrom(frm1.txtWork_dt.Text,frm1.txtWork_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
		strYYYYMM = strYear & strMonth
		Call ExtractDateFrom(frm1.txtWorkdate.Text,frm1.txtWorkdate.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
		strYYYYMMDD = strYear & strMonth
	
		If Trim(strYYYYMM) <> "" Then
			If Trim(strYYYYMM) <> Trim(strYYYYMMDD) Then
				IntRetCD = DisplayMsgBox("am0027","x","x","x")      '결산일자는 결산년월과 같은 달이여야 합니다.
			End If
		End If	
	End If		

End Sub

Sub txtWorkdate_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
	End If
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
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
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월차결산작업</font></td>
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
									<TD NOWRAP  CLASS=TD5 NOWRAP>결산년월</TD>
									<TD NOWRAP  CLASS=TD6 NOWRAP><script language =javascript src='./js/a5955ma1_txtWork_dt_txtWork_dt.js'></script></TD>
                                    <TD NOWRAP  CLASS=TD5 NOWRAP>결산일자</TD>
									<TD NOWRAP  CLASS=TD6 NOWRAP><script language =javascript src='./js/a5955ma1_txtWorkDate_txtWorkDate.js'></script></TD>
								</TR>
                                <TR>
                                    <TD NOWRAP  CLASS=TD5 NOWRAP>사업장</TD>
									<TD NOWRAP  CLASS=TD6 NOWRAP>
									 <INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=20 tag="12XXXU" ALT="사업장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCode" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizArea()">
									 <INPUT TYPE=TEXT NAME="txtBizArea" SIZE=22 MAXLENGTH=50 tag="14" ALT="사업장명" >
									 </TD>
                                    <TD NOWRAP  CLASS=TD5 NOWRAP></TD>
									<TD NOWRAP  CLASS=TD6 NOWRAP></TD>
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
									<script language =javascript src='./js/a5955ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD NOWRAP >
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD NOWRAP  WIDTH=10>&nbsp;</TD>
					<TD NOWRAP ><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>작업실행</BUTTON></TD>
					<TD WIDTH=* ALIGN=RIGHT>
						<A HREF="VBSCRIPT:PgmJumpChk(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:Call CookiePage(1)">월차결산진행현황</A>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
            <!--        <TD NOWRAP  WIDTH=*>&nbsp;</TD>  -->
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD NOWRAP  WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtJumpFlag"	 TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
