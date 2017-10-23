
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 경영손익 
*  2. Function Name        :
*  3. Program ID           : GB002MA1
*  4. Program Name         : 회계Data 조회 
*  5. Program Desc         : Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/11/26
*  8. Modified date(Last)  : 2001/12/28
*  9. Modifier (First)     : park jai hong
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "GB002MB1.asp"                                      'Biz Logic ASP
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================

Dim C_COST_CD  
Dim C_COST_NM  
Dim C_COST_TYPE  
Dim C_ACCT_CD  
Dim C_ACCT_NM 
Dim C_ACCT_TYPE  
Dim C_CTRL_CD
Dim C_CTRL_NM
Dim C_CTRL_VAL
Dim C_CTRL_VAL_NM  
Dim C_MINOR_CD  
Dim C_MINOR_NM  
Dim C_AMOUNT  


Const COOKIE_SPLIT      = 4877	                                      '☆: Cookie Split String


'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
'Dim lgIsOpenPop
Dim IsOpenPop


'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
    lgPageNo		  = "0"

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


'========================================================================================================

Sub SetDefaultVal()
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Dim StartDate
	StartDate	= "<%=GetSvrDate%>"                                               'Get Server DB Date
	
	frm1.txtYyyymm.focus
	frm1.txtYyyymm.text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat, 2)
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	<% Call loadInfTB19029A("Q", "G","NOCOOKIE", "QA") %>

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
	Dim intRow
	Dim intIndex

	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow
		'	.Col = C_TYPECd
			intIndex = .value
		'	.col = C_TYPENm
			.value = intindex
		Next
	End With
End Sub


'========================================================================================================
Sub InitSpreadPosVariables()
 
	 C_COST_CD			= 1       
	 C_COST_NM			= 2 	 
	 C_COST_TYPE		= 3
	 C_ACCT_CD			= 4 
	 C_ACCT_NM			= 5 	 
	 C_ACCT_TYPE		= 6 
	 C_CTRL_CD			= 7   
	 C_CTRL_NM			= 8
	 C_CTRL_VAL			= 9
	 C_CTRL_VAL_NM		= 10
		C_MINOR_CD			= 11   
	 C_MINOR_NM			= 12 
     C_AMOUNT			= 13 
End Sub



'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData

       .MaxCols = C_AMOUNT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:


        ggoSpread.Source = Frm1.vspdData
        ggoSpread.Spreadinit "V20021128", ,parent.gAllowDragDropSpread
        
        ggoSpread.ClearSpreadData

	   .ReDraw = false

       Call AppendNumberPlace("6","3","0")
       
       Call GetSpreadColumnPos("A")


       ggoSpread.SSSetEdit  C_COST_CD , "C/C코드" ,10,   ,, 20
       ggoSpread.SSSetEdit  C_COST_NM , "C/C명" ,15,   ,, 20
       ggoSpread.SSSetEdit  C_COST_TYPE , "C/C Type" ,10,   ,, 20
       ggoSpread.SSSetEdit  C_ACCT_CD , "계정코드" ,10,   ,, 20
       ggoSpread.SSSetEdit  C_ACCT_NM , "계정명"   ,20,   ,, 30
       ggoSpread.SSSetEdit  C_ACCT_TYPE , "변동/고정/관세환급", 10 , 0 
       ggoSpread.SSSetEdit  C_CTRL_CD , "관리항목" ,10,   ,, 20
       ggoSpread.SSSetEdit  C_CTRL_NM , "관리항목명" ,10,   ,, 20
       ggoSpread.SSSetEdit  C_CTRL_VAL		, "관리항목Value" ,10,   ,, 20
       ggoSpread.SSSetEdit  C_CTRL_VAL_NM	, "관리항목Value명" ,20,   ,, 20
       ggoSpread.SSSetEdit  C_MINOR_CD		, "손익대분류" ,10,   ,, 20
       ggoSpread.SSSetEdit  C_MINOR_NM		, "손익항목" ,10,   ,, 20              
       ggoSpread.SSSetFloat  C_AMOUNT		,  "금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec       

	   .ReDraw = true

       Call SetSpreadLock

    End With

End Sub


'======================================================================================================
Sub SetSpreadLock()
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False

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



			C_COST_CD			= iCurColumnPos(1)       
			C_COST_NM			= iCurColumnPos(2) 	 
			C_COST_TYPE			= iCurColumnPos(3)   
			C_ACCT_CD			= iCurColumnPos(4) 	 
			C_ACCT_NM			= iCurColumnPos(5)   
			C_ACCT_TYPE			= iCurColumnPos(6) 
			C_CTRL_CD			= iCurColumnPos(7) 
			C_CTRL_NM			= iCurColumnPos(8) 
			C_CTRL_VAL 			= iCurColumnPos(9) 
			C_CTRL_VAL_NM		= iCurColumnPos(10) 
			C_MINOR_CD	 		= iCurColumnPos(11) 
			C_MINOR_NM 			= iCurColumnPos(12) 
			C_AMOUNT			= iCurColumnPos(13) 
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
	Call InitData()
End Sub


'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status

	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format

	'------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field

    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	Call SetToolbar("1100000000001111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call InitComboBox
    Call InitData()
    
	Call CookiePage (0)                                                              '☜: Check Cookie

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

    Call InitVariables															  '⊙: Initializes local global variables
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------


    '------ Developer Coding part (End )   --------------------------------------------------------------

    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncQuery = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncNew = True																 '☜: Processing is OK
End Function


'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncDelete = True                                                             '☜: Processing is OK
End Function


'========================================================================================================
Function FncSave()
    Dim IntRetCD

    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                      '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data.
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncSave = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow

            .ReDraw = True
		    .Focus
		 End If
	End With

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'----------------------------------------------------------------------------------------------------



	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCopy = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncCancel()
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCancel = False                                                            '☜: Processing is OK
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
	                          '☜: Processing is OK
End Function


'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
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
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
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
    Set gActiveElement = document.ActiveElement
    FncPrev = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncNext()
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncNext = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncExcel()
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function DbQuery()
	Dim strVal
	Dim strYear,strMonth,strDay
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    
    DbQuery = False                                                              '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
     '@Query_Hidden     
			strVal = BIZ_PGM_ID
			strVal = strVal & "?txtYyyymm="		& .hYYYYMM.value
			strVal = strVal & "&txtCostCd="		& .hCostCd.value				
			strVal = strVal & "&txtAcctGp="		& .hAcctGp.value
			strVal = strVal & "&txtCtrlCd="		& .hCtrlCd.value
		Else
      '@Query_Text     
			Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
			strVal = BIZ_PGM_ID
			strVal = strVal & "?txtYyyymm="		& strYear & strMonth
			strVal = strVal & "&txtCostCd="		& .txtCostCd.value				
			strVal = strVal & "&txtAcctGp="		& .txtAcctGp.value
			strVal = strVal & "&txtCtrlCd="		& .txtCtrlCd.value
		END IF
		
			
		strVal = strVal & "&lgPageNo="			& lgPageNo								'Next key tag
'		strVal = strVal & "&lgMaxCount="		& C_SHEETMAXROWS_D					'한번에 가져올수 있는 데이타 건수 
	    
    End With
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   End Function

'========================================================================================================
Function DbSave()
    Dim lRow
    Dim lGrpCnt
	Dim strVal, strDel

    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
'    Call LayerShowHide(1)                                                        '☜: Show Processing Message
	If	LayerShowHide(1) = False Then
		Exit Function
	End If

    '--------- Developer Coding Part (Start) ----------------------------------------------------------


	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
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

    lgIntFlgMode = Parent.OPMD_UMODE
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call InitData()

	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
End Sub


'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  --------------------------------------------------------------
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData


	Call SetToolbar("1111111111111111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   --------------------------------------------------------------
    DBQuery()
    Set gActiveElement = document.ActiveElement
End Sub


'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


'========================================================================================================
Function OpenCondAreaPopup(byVal iwhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	select case iwhere
		case 1
			arrParam(0) = "Cost Center"		    	<%' 팝업 명칭 %>
			arrParam(1) = "B_COST_CENTER"		<%' TABLE 명칭 %>
			arrParam(2) = frm1.txtCostCd.value	    <%' Code Condition%>
			arrParam(3) = "" 		            		<%' Name Cindition%>
			arrParam(4) = ""   <%' Where Condition%>
			arrParam(5) = "Cost Center"

			arrField(0) = "COST_CD"					<%' Field명(0)%>
			arrField(1) = "COST_NM"	     			<%' Field명(1)%>

			arrHeader(0) = "Cost Center"				<%' Header명(0)%>
			arrHeader(1) = "Cost Center명"				<%' Header명(1)%>
		
		case 2
			arrParam(0) = "계정그룹"		    	<%' 팝업 명칭 %>
			arrParam(1) = "A_ACCT_GP"		<%' TABLE 명칭 %>
			arrParam(2) = frm1.txtAcctGp.value	    <%' Code Condition%>
			arrParam(3) = "" 		            		<%' Name Cindition%>
			arrParam(4) = "gp_cd in (select distinct gp_cd from a_acct where temp_fg_3 LIKE " & FilterVar("G%", "''", "S") & ")"   <%' Where Condition%>
			arrParam(5) = "계정그룹"

			arrField(0) = "GP_CD"					<%' Field명(0)%>
			arrField(1) = "GP_NM"	     			<%' Field명(1)%>

			arrHeader(0) = "계정그룹"				<%' Header명(0)%>
			arrHeader(1) = "계정그룹명"				<%' Header명(1)%>
		case 3	
			arrParam(0) = "관리항목"		    	<%' 팝업 명칭 %>
			arrParam(1) = "A_CTRL_ITEM"		<%' TABLE 명칭 %>
			arrParam(2) = frm1.txtCtrlCd.value	    <%' Code Condition%>
			arrParam(3) = "" 		            		<%' Name Cindition%>
			arrParam(4) = "ctrl_cd in (" & FilterVar("MK", "''", "S") & "," & FilterVar("MG", "''", "S") & "," & FilterVar("BZ", "''", "S") & "," & FilterVar("CC", "''", "S") & "," & FilterVar("SO", "''", "S") & "," & FilterVar("SG", "''", "S") & "," & FilterVar("BP", "''", "S") & ")"   <%' Where Condition%>
			arrParam(5) = "계정그룹"

			arrField(0) = "ctrl_cd"					<%' Field명(0)%>
			arrField(1) = "ctrl_nm"	     			<%' Field명(1)%>

			arrHeader(0) = "관리항목"				<%' Header명(0)%>
			arrHeader(1) = "관리항목명"				<%' Header명(1)%>
		
	end select

	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	 Select case iwhere
	  Case 1
		frm1.txtCostCd.focus
	  Case 2
		frm1.txtAcctGp.focus
	  Case 3
		frm1.txtCtrlCd.focus
	 End Select
		Exit Function
	Else
		Call SetPopup(arrRet,iWhere)
	End If

End Function


'=======================================================================================================%>
Function SetPopup(Byval arrRet,ByVal iwhere )
	With frm1
		select case iWhere
			case 1
				.txtCostCd.focus
				.txtCostCd.value = arrRet(0)
				.txtCostNm.value = arrRet(1)
			case 2
				.txtAcctGp.focus
				.txtAcctGp.value = arrRet(0)
				.txtAcctGpNm.value = arrRet(1)
			case 3
				.txtCtrlCd.focus
				.txtCtrlCd.value = arrRet(0)
				.txtCtrlNm.value = arrRet(1)
		end select
	End With
End Function





'========================================================================================================
Sub vspdData_Change(Col , Row )
    Dim iDx

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
Sub vspdData_Click(Col, Row)
	
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


'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub



'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "0" Then
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
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub



'=======================================================================================================

Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub


'=======================================================================================================
Sub txtYyyymm_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
		End If
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>회계 Data 조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>대상년월</TD>
									<TD CLASS="TD6" NOWRAP ><script language =javascript src='./js/gb002ma1_txtYyyymm_txtYyyymm.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>Cost Center</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup(1)">
										<INPUT TYPE="Text" NAME="txtCostNm" SiZE=20 MAXLENGTH=20 tag="14XXXU" ALT="Cost Center명">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정그룹</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtAcctGp" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="계정그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup(2)">
										<INPUT TYPE="Text" NAME="txtAcctGpNm" SiZE=20 MAXLENGTH=50 tag="14XXXU" ALT="계정그룹명"></TD>
									</TD>
									<TD CLASS="TD5" NOWRAP>관리항목</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtCtrlCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="관리항목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup(3)">
										<INPUT TYPE="Text" NAME="txtCtrlNm" SiZE=20 MAXLENGTH=20 tag="14XXXU" ALT="관리항목명"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/gb002ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>

				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
							    	<TD CLASS=TD5 NOWRAP>비용합계&nbsp;&nbsp</TD>
							    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/gb002ma1_fpDoubleSingle1_txtExpenseSum.js'></script></TD>
							    	<TD CLASS=TD5 NOWRAP>수익합계&nbsp;&nbsp</TD>
							    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/gb002ma1_fpDoubleSingle2_txtProfitSum.js'></script></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMode"    tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM"     tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hCostCd"     tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hAcctGp"     tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hCtrlCd"     tag="24" TABINDEX = "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

