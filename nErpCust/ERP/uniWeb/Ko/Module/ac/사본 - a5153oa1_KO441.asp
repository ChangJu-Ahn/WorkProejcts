<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 회계관리 - 정비채권관리
*  2. Function Name        : 
*  3. Program ID           : A5152OA1_KO441
*  4. Program Name         : 계정별 잔액 증감명세서(관리항목)(S)
*  5. Program Desc         :  
*  6. Comproxy List        :
*  7. Modified date(First) : 2004/07/16
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : 유엔아이 신현호
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

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "A5153OB1_KO441.asp"						           '☆: Biz Logic ASP Name

Const C_SHEETMAXROWS_D    = 500
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------
'Dim C_ACCT
Dim C_CTRLVAL
Dim C_ACCTNM
Dim C_l_yr_bal
Dim C_l_prd_bal
Dim C_t_prd_inc
Dim C_t_prd_dec
Dim C_t_prd_bal

Dim gSpreadFlg
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim IsOpenPop
<% 
   BaseDate     = GetSvrDate                                                                  'Get DB Server Date
%>   
dim lastdate
dim firstdate
dim ExampleDate

LastDate    = UNIGetLastDay ("<%=BaseDate%>",parent.gServerDateFormat)                                  'Last  day of this month
FirstDate   = UNIGetFirstDay("<%=BaseDate%>",parent.gServerDateFormat)                                  'First day of this month
ExampleDate = UniDateAdd("m", -2, "<%=BaseDate%>",parent.gServerDateFormat)
ExampleDate = UNIConvDateAToB("<%=BaseDate%>" ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
    
		'C_ACCT			= 1
		C_CTRLVAL		= 1
		C_ACCTNM		= 2
		C_l_yr_bal		= 3
		C_l_prd_bal		= 4
		C_t_prd_inc		= 5
		C_t_prd_dec		= 6
		C_t_prd_bal		= 7
		
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	gSpreadFlg		  = 1
	
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()

	Dim StartDate, EndDate
	
	StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
	StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
	EndDate   = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtFromGlDt.Text	=  STARTDate
	frm1.txttoGlDt.Text		=  ENDDate
	Call ggoOper.FormatDate(frm1.txtFromGlDt, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txttoGlDt, parent.gDateFormat, 2)
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next


End Function

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pDirect)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
	lgKeyStream =               Trim(frm1.txtFromGlDt.Text)		& Parent.gColSep 
    lgKeyStream = lgKeyStream & Trim(frm1.txttoGlDt.Text)		& Parent.gColSep 
    lgKeyStream = lgKeyStream & Trim(frm1.txtAcctCd.value)		& Parent.gColSep 
    lgKeyStream = lgKeyStream & Trim(frm1.txtFrBizAreaCd.value)	& Parent.gColSep 
    lgKeyStream = lgKeyStream & Trim(frm1.txtToBizAreaCd.value)	& Parent.gColSep 
    lgKeyStream = lgKeyStream & Trim(frm1.txtAcctCd1.value)		& Parent.gColSep 
    	           
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

	'Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'XW014' ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	'Call SetCombo(frm1.txtEngineerCd,"","           ")
	'call setcombo2(frm1.txtEngineerCd,lgF0,lgF1,Chr(11))
	
End Sub

'========================================================================================================
' Function Name : InitSpreadComboBox
' Function Desc :
'========================================================================================================
Sub InitSpreadComboBox()
    
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
			.Col = C_ACCTNM
			If  Trim(.Value) = "총합계" THEN
			    
			    .Col = -1 
			    .Col2 = -1
			    .BackColor = RGB(255,230,255)
		   
			End If
    next
    End With    
	

End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021105",, parent.gAllowDragDropSpread    
		.ReDraw = false
		.MaxCols   = C_t_prd_bal + 1                                                  ' ☜:☜: Add 1 to Maxcols

		Call ggoSpread.ClearSpreadData()	
		Call AppendNumberPlace("6","15","0")
		Call GetSpreadColumnPos("A")
		
		'ggoSpread.SSSetEdit    C_ACCT			,"관리항목"			,10
		ggoSpread.SSSetEdit    C_CTRLVAL		,"계정코드"		,12
		ggoSpread.SSSetEdit    C_ACCTNM			,"계정명"		,20
		ggoSpread.SSSetFloat   C_l_yr_bal		,"전기이월"     ,15 ,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"       
		ggoSpread.SSSetFloat   C_l_prd_bal		,"전월이월"     ,15 ,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"       
		ggoSpread.SSSetFloat   C_t_prd_inc		,"증가"         ,15 ,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"       
		ggoSpread.SSSetFloat   C_t_prd_dec		,"감소"         ,15 ,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"       
		ggoSpread.SSSetFloat   C_t_prd_bal		,"당월잔액"     ,15 ,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"       
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		       		
		.ReDraw = true
	
    End With

End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of err
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
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)            

			'C_ACCT			= iCurColumnPos(1)
			C_CTRLVAL		= iCurColumnPos(1)
			C_ACCTNM		= iCurColumnPos(2)
			C_l_yr_bal		= iCurColumnPos(3)
			C_l_prd_bal		= iCurColumnPos(4)
			C_t_prd_inc		= iCurColumnPos(5)
			C_t_prd_dec		= iCurColumnPos(6)
			C_t_prd_bal		= iCurColumnPos(7)
			
    End Select    
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Group-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field

	Call InitVariables
    Call SetDefaultVal

    Call InitSpreadSheet                                                             'Setup the Spread sheet
	'Call ggoOper.FormatDate(frm1.fpdtCloseDt, Parent.gDateFormat, 2)

	Call SetToolbar("11000000000011")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Call InitComboBox
    'frm1.txtACCT.focus
End Sub
'========================================================================================================
Function CompareAcctCdByDB(ByVal FromCd , ByVal ToCd)
	Dim strSelect,strFrom,strWhere
	Dim iFlag,iRs

	CompareAcctCdByDB = False

    If FromCd.value <> "" And ToCd.value <> "" Then
        strSelect = ""
        strSelect = "  Case When  " & FilterVar(UCase(FromCd.value), "''", "S") & " "
        strSelect = strSelect & "  >  " & FilterVar(UCase(ToCd.value), "''", "S") & "  Then " & FilterVar("N", "''", "S") & "  "
        strSelect = strSelect & " When  " & FilterVar(UCase(FromCd.value), "''", "S") & " "
        strSelect = strSelect & "  <=  " & FilterVar(UCase(ToCd.value), "''", "S") & "  Then " & FilterVar("Y", "''", "S") & "  End "
        strFrom = ""
        strWhere = ""
        If CommonQueryRs2by2(strSelect, strFrom, strWhere, iRs) = True Then
            iFlag = Split(iRs, Chr(11))
            If Trim(iFlag(1)) = "N" Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    CompareAcctCdByDB = True
End Function
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
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False															  '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData
	
	If ValidDateCheck(frm1.txtFromGlDt, frm1.txtTOGlDt) = False Then Exit Function
	
	
	'alert frm1.txtFromGlDt
	'alert frm1.txtAcctCd1.value
	
	If CompareAcctCdByDB(frm1.txtAcctCd,frm1.txtAcctCd1) = False Then
        Call DisplayMsgBox("970025", "X", frm1.txtAcctCd.Alt, frm1.txtAcctCd1.Alt)
        frm1.txtAcctCd.focus
		Exit Function
	End If		

	
	
    'Call ggoOper.ClearField(Document, "2")										  '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    
    Call InitVariables															  '⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									          '⊙: This function check indispensable field
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    If DbQuery("Q") = False Then                                                       '☜: Query db data
       Exit Function
    End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    If Err.number = 0 Then
       FncQuery = True                                                            '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
   
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
  
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()

End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel()
 
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
  
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()

End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False	                                                          '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
 
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 

End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	'------ Developer Coding part (Start )   -------------------------------------------------------------- 
	Call Parent.FncExport(Parent.C_MULTI)
	'------ Developer Coding part (End   )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncExcel = True                                                            '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	'------ Developer Coding part (Start )   -------------------------------------------------------------- 
	Call Parent.FncFind(Parent.C_MULTI, True)
	'------ Developer Coding part (End   )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncFind = True                                                             '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

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
    
    Call ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.SaveSpreadColumnInf()

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG

	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			          '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    If Err.number = 0 Then
       FncExit = True                                                             '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Group-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect)

	Dim strVal
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
 
    DbQuery = False                                                               '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call MakeKeyStream(pDirect)

    strVal = BIZ_PGM_ID & "?txtMode="        & Parent.UID_M0001                          '☜: Query
    strVal = strVal     & "&txtKeyStream="   & lgKeyStream                        '☜: Query Key
    strVal = strVal     & "&txtPrevNext="    & pDirect                            '☜: Direction
    strVal = strVal     & "&lgStrPrevKey="   & lgStrPrevKey                       '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="     & Frm1.vspdData.MaxRows              '☜: Max fetched data
    strVal = strVal     & "&lgMaxCount="    & CStr(C_SHEETMAXROWS_D)            '☜: Max fetched data at a time
	
	'------ Developer Coding part (End )   --------------------------------------------------------------
	'alert  strVal
    Call RunMyBizASP(MyBizASP, strVal)                                            '☜:  Run biz logic

    If Err.number = 0 Then
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
 
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	Dim intRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	lgIntFlgMode      = Parent.OPMD_UMODE                                                '⊙: Indicates that current mode is Create mode

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    Frm1.vspdData.focus
	Call SetToolbar("11000000000111")                                           '☆: Developer must customize
    Call InitData()
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")

    Set gActiveElement = document.ActiveElement   

End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
	
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()

End Sub


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Supplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenACCT()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정 팝업"									' 팝업 명칭 
	arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtAcctCd.Value)						' Code Condition
	arrParam(3) = ""												' Name Cindition
	arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD"' AND a_acct.acct_type IN ('i0','i1','i2')"					' Where Condition
	arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

	arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
	arrField(1) = "A_ACCT.Acct_NM"	
	arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
	arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)

	arrHeader(0) = "계정코드"									' Header명(0)
	arrHeader(1) = "계정코드명"									' Header명(1)
	arrHeader(2) = "그룹코드"									' Header명(2)
	arrHeader(3) = "그룹명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtAcctCd.Value = arrRet(0)
		frm1.txtAcctNM.Value = arrRet(1)
	End If	
End Function

'========================================================================================================
'   Event Name : txtBpCd_Onchange
'   Event Desc :
'========================================================================================================

Function txtAcctCd_Onchange()
    Dim IntRetCd, strWhere
    
    IF Trim(frm1.txtAcctCd.value) = "" Then
        frm1.txtAcctCd.value        = ""
        frm1.txtAcctNM.value    = ""
        
    Else
        
        'strWhere =  " A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND a_acct.acct_type IN ('i0','i1','i2') AND A_ACCT.Acct_CD = " & FilterVar(Ucase(Trim(frm1.txtAcctCd.value)),"''","S")
        strWhere =  " A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.Acct_CD = " & FilterVar(Ucase(Trim(frm1.txtAcctCd.value)),"''","S")
        
        IntRetCd = CommonQueryRs( " A_ACCT.Acct_CD , A_ACCT.Acct_NM " ," A_Acct, A_ACCT_GP ", strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        If IntRetCd = false then
            Call DisplayMsgBox("800054", "X", "X", "X")
            frm1.txtAcctCd.value    = ""
            frm1.txtAcctNM.value    = ""
            frm1.txtAcctCd.focus
        ELSE    
            frm1.txtAcctCd.Value	= Trim(Replace(lgF0,Chr(11),""))
            frm1.txtAcctNM.Value	= Trim(Replace(lgF1,Chr(11),""))
        End If
    End If
    
End Function

Function txtAcctCd1_Onchange()
    Dim IntRetCd, strWhere
    
    IF Trim(frm1.txtAcctCd1.value) = "" Then
        frm1.txtAcctCd1.value        = ""
        frm1.txtAcctNM1.value    = ""
        
    Else
        
        'strWhere =  " A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND a_acct.acct_type IN ('i0','i1','i2') AND A_ACCT.Acct_CD = " & FilterVar(Ucase(Trim(frm1.txtAcctCd1.value)),"''","S")
        strWhere =  " A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.Acct_CD = " & FilterVar(Ucase(Trim(frm1.txtAcctCd1.value)),"''","S")
        
        IntRetCd = CommonQueryRs( " A_ACCT.Acct_CD , A_ACCT.Acct_NM " ," A_Acct, A_ACCT_GP ", strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        If IntRetCd = false then
            Call DisplayMsgBox("800054", "X", "X", "X")
            frm1.txtAcctCd1.value    = ""
            frm1.txtAcctNM1.value    = ""
            frm1.txtAcctCd1.focus
        ELSE    
            frm1.txtAcctCd1.Value	= Trim(Replace(lgF0,Chr(11),""))
            frm1.txtAcctNM1.Value	= Trim(Replace(lgF1,Chr(11),""))
        End If
    End If
    
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

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
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1111111111")    
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
	
	gSpreadFlg = 1
		
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If

        Exit Sub
    End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
'7.1. SpreadSheet의 이벤트명[DblClick]에 로직 추가
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery("R") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End if
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


Function OpenPopUp(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strSelect,strFrom,strWhere
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6	
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    Select Case iWhere
        Case 1     
	        arrParam(0) = "사업장 팝업"						' 팝업 명칭
	        arrParam(1) = "B_BIZ_AREA"					    ' TABLE 명칭
	        arrParam(2) = Trim(frm1.txtFrBizAreaCd.Value)	' Code Condition
	        arrParam(3) = ""							    ' Name Cindition
	        arrParam(4) = ""							    ' Where Condition
	        arrParam(5) = "사업장 코드"			
	
            arrField(0) = "BIZ_AREA_CD"						' Field명(0)
            arrField(1) = "BIZ_AREA_NM"						' Field명(1)
    
            arrHeader(0) = "사업장코드"						' Header명(0)
	        arrHeader(1) = "사업장명"				        ' Header명(1)	
	    Case 2
	        arrParam(0) = "계정코드 팝업"				
	        arrParam(1) = "A_ACCT"					
	        arrParam(2) = Trim(frm1.txtAcctCd.Value)
	        arrParam(3) = ""							
	        arrParam(4) = ""							
	        arrParam(5) = "계정 코드"			
	
            arrField(0) = "ACCT_CD"				
            arrField(1) = "ACCT_NM"				
    
            arrHeader(0) = "계정코드"				
	         arrHeader(1) = "계정코드명"					     
		Case 3
			If frm1.txtAcctCd.value = "" Then 	         
				msgbox "계정코드를 먼저 선택하십시오."
				Exit Function
	        End If
			
			arrParam(0) = "계정코드 팝업"				
	        arrParam(1) = "A_ACCT"					
	        arrParam(2) = Trim(frm1.txtAcctCd1.Value)
	        arrParam(3) = ""							
	        arrParam(4) = ""							
	        arrParam(5) = "계정 코드"			
	
            arrField(0) = "ACCT_CD"				
            arrField(1) = "ACCT_NM"				
    
            arrHeader(0) = "계정코드"				
	        arrHeader(1) = "계정코드명"
	        
	        'arrParam(0) = "관리항목 팝업"				
	       ' arrParam(1) = "A_ACCT_CTRL_ASSN, A_CTRL_ITEM"					
	       ' arrParam(2) = Trim(frm1.txtCtrlCd.Value)
	       '' arrParam(3) = ""							
	       ' arrParam(4) = "A_ACCT_CTRL_ASSN.Ctrl_CD = A_CTRL_ITEM.CTRL_CD AND ACCT_CD = '" & frm1.txtAcctCd.value & "' "							
	       ' arrParam(5) = "계정 코드"			

           ' arrField(0) = "A_ACCT_CTRL_ASSN.CTRL_CD"				
           ' arrField(1) = "CTRL_NM"				

            'arrHeader(0) = "관리항목"
	       ' arrHeader(1) = "관리항목명"
		
		Case 5
	        arrParam(0) = "사업장 팝업"						' 팝업 명칭
	        arrParam(1) = "B_BIZ_AREA"					    ' TABLE 명칭
	        arrParam(2) = Trim(frm1.txtToBizAreaCd.Value)	' Code Condition
	        arrParam(3) = ""							    ' Name Cindition
	        arrParam(4) = ""							    ' Where Condition
	        arrParam(5) = "사업장 코드"			
	
            arrField(0) = "BIZ_AREA_CD"						' Field명(0)
            arrField(1) = "BIZ_AREA_NM"						' Field명(1)
    
            arrHeader(0) = "사업장코드"						' Header명(0)
	        arrHeader(1) = "사업장명"				        ' Header명(1)	
	End Select         

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If	

End Function

'------------------------------------------  SetReturnVal()  ---------------------------------------------
'	Name : SetReturnVal()
'	Description : Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetReturnVal(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere 
		    Case 1
			  .txtFrBizAreaCd.value = arrRet(0)
			  .txtFrBizAreaNm.value = arrRet(1)
			Case 2
			  .txtAcctCd.value = arrRet(0)
			  .txtAcctNm.value = arrRet(1)	
			  '.txtCtrlCd.value = ""
             ' .txtCtrlNm.value = ""
			  '.txtCtrlVal.value = ""
			  '.txtCtrlValNm.value = ""
			Case 3
			  .txtAcctCd1.value = arrRet(0)
			  .txtAcctNm1.value = arrRet(1)
			  '.txtCtrlVal.value = ""
			  '.txtCtrlValNm.value = ""
			  			  
		    Case 5
			  .txtToBizAreaCd.value = arrRet(0)
			  .txtToBizAreaNm.value = arrRet(1)			  
		End Select
	End With
End Function

'========================================================================================================
'   Event Name : txtAcctCd_Onchange
'   Event Desc :
'========================================================================================================
'Function txtAcctCd_Onchange()
'	With frm1
'		If .txtAcctCd.value = "" Then
'			Exit Function
'		End If
 '   
'		Call CommonQueryRs("distinct A_ACCt.ACCT_CD, ACCT_NM ","A_ACCT, A_ACCT_CTRL_ASSN","A_ACCT.ACCT_CD = '" & .txtAcctCd.value & "' AND A_ACCT.acct_cd = a_acct_ctrl_assn.acct_cd" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 '   
'		If (lgF0 <> "X") And (Trim(lgF0) <> "") Then 
'			.txtAcctNm.value = Left(lgF1, Len(lgF1)-1)    
'			'.txtCtrlCd.value = ""
'		'	.txtCtrlNm.value = ""
'			'.txtCtrlVal.value = ""
'			'.txtCtrlValNm.value = ""
'		Else       
'			msgbox "해당 계정은 관리항목이 없습니다. 다시 선택하십시오"
'			.txtAcctNm.value = ""
'			'.txtCtrlCd.value = ""
'			'.txtCtrlNm.value = ""       
'			'.txtCtrlVal.value = ""
'			'.txtCtrlValNm.value = ""       
'			.txtAcctCd.focus       
'		End If   
'	End With
'	
'    txtAcctCd_OnChange = True
'End Function

'Function txtAcctCd1_Onchange()
'	With frm1
'		If .txtAcctCd1.value = "" Then
'			Exit Function
'		End If
 '   
'		Call CommonQueryRs("distinct A_ACCt.ACCT_CD, ACCT_NM ","A_ACCT, A_ACCT_CTRL_ASSN","A_ACCT.ACCT_CD = '" & .txtAcctCd1.value & "' AND A_ACCT.acct_cd = a_acct_ctrl_assn.acct_cd" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 '   
'		If (lgF0 <> "X") And (Trim(lgF0) <> "") Then 
'			.txtAcctNm1.value = Left(lgF1, Len(lgF1)-1)    
'			'.txtCtrlCd1.value = ""
'			'.txtCtrlNm1.value = ""
'			'.txtCtrlVal.value = ""
'			'.txtCtrlValNm.value = ""
'		Else       
'			msgbox "해당 계정은 관리항목이 없습니다. 다시 선택하십시오"
'			.txtAcctNm1.value = ""
'			''.txtCtrlCd1.value = ""
'			'.txtCtrlNm1.value = ""       
'			'.txtCtrlVal.value = ""
'			'.txtCtrlValNm.value = ""       
'			.txtAcctCd1.focus       
'		End If   
'	End With
'	
 '   txtAcctCd1_OnChange = True
'End Function

'========================================================================================================
'   Event Name : txtAcctCd_Onchange
'   Event Desc :
'========================================================================================================
'Function txtCtrlCd_Onchange()
'	With frm1
'		If (.txtCtrlCd.value = "") Or (.txtAcctCd.value = "") Then
''		   Exit Function
''		End if

'		Call CommonQueryRs("A_ACCT_CTRL_ASSN.CTRL_CD, CTRL_NM ", " A_ACCT_CTRL_ASSN, A_CTRL_ITEM", "a_acct_ctrl_assn.ctrl_cd = a_ctrl_item.ctrl_cd AND ACCT_CD = '" & frm1.txtAcctCd.value & "' and a_acct_ctrl_assn.ctrl_cd = '" & frm1.txtCtrlCd.value & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
'
''		If (lgF0 <> "X") And (trim(lgF0) <> "") Then
''			.txtCtrlNm.value = Left(lgF1, Len(lgF1)-1)
''		Else
''			.txtCtrlCd.focus
'		End If
'	End With

 '   txtCtrlCd_OnChange = True
'End Function

Function txtFrBizAreaCd_Onchange()
	With frm1
		If (.txtFrBizAreaCd.value = "") Then
		   Exit Function
		End if

		Call CommonQueryRs("BIZ_AREA_NM", " B_BIZ_AREA", "BIZ_AREA_CD = '" & frm1.txtFrBizAreaCd.value & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    

		If (lgF0 <> "X") And (trim(lgF0) <> "") Then
			.txtFrBizAreanm.value = Trim(Replace(lgF0,Chr(11),""))
		Else
			.txtFrBizAreaCd.value = ""
			.txtFrBizAreaNM.value = ""
			.txtFrBizAreaCd.focus
		End If
	End With

    
End Function

Function txtToBizAreaCd_Onchange()
	With frm1
		If (.txtToBizAreaCd.value = "") Then
		   Exit Function
		End if

		Call CommonQueryRs("BIZ_AREA_NM", " B_BIZ_AREA", "BIZ_AREA_CD = '" & frm1.txtToBizAreaCd.value & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    

		If (lgF0 <> "X") And (trim(lgF0) <> "") Then
			.txtToBizAreaNm.value = Trim(Replace(lgF0,Chr(11),""))
		Else
			.txtToBizAreaCd.value = ""
			.txtToBizAreaNM.value = ""
			.txtToBizAreaCd.focus
		End If
	End With

    
End Function
'=======================================================================================================
'   Event Name : 
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromGlDt.Action = 7
    End If
End Sub

Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
    End If
End Sub


'=======================================================================================================
'   Event Name : txtValidDt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행
'=======================================================================================================
Sub txtFromGLDt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtToGLDt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtFrBizAreaCd_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtToBizAreaCd_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtAcctCd_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtAcctCd1_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Function FncBtnPreView()

   Dim StrEbrFile, condvar
  
   If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
   End If
   
   If lgIntFlgMode = parent.OPMD_CMODE Then						'/조회여부 확인
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
   End If
  
   Call PrintCond(strEbrFile, condvar)
   
   ObjName = AskEBDocumentName(StrEBrFile, "ebr")


   Call FncEBRPreview(ObjName, condvar)    
    
End Function

'==========================================================================================
'   Event Name : FncBtnPrint()
'   Event Desc : 
'==========================================================================================
Function FncBtnPrint()

   Dim StrEbrFile, condvar
  
   If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
   End If
   
   If lgIntFlgMode = parent.OPMD_CMODE Then						'/조회여부 확인
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
   End If
  
   Call PrintCond(strEbrFile, condvar)
 
   ObjName = AskEBDocumentName(StrEBrFile, "ebr")
   
   Call FncEBRPrint(EBAction , ObjName , condvar)
    
End Function

'==========================================================================================
'   Event Name : PrintCond(strEbrFile, condvar)
'   Event Desc : 
'==========================================================================================
Sub PrintCond(strEbrFile, condvar)

   Dim StrFrEndYmd,StrToEndYmd,STRID,STRTYPE,STRCtrl,STRFrBiz,STRToBiz

   
   StrFrEndYmd = frm1.txtFromGlDt.text
   StrToEndYmd = frm1.txtTOGlDt.text
   STRID       = frm1.txtID.VALUE
   
   'STRCtrl     = frm1.txtCtrlCd.VALUE
   
   IF frm1.txtAcctCd.value <> "" THEN
   STRTYPE     = frm1.txtAcctCd.value
   ELSE 
   STRTYPE     = ""
   END IF
   
   IF frm1.txtAcctCd1.value <> "" THEN
   STRCtrl     = frm1.txtAcctCd1.value
   ELSE 
   STRCtrl     = ""
   END IF
   
   IF frm1.txtFrBizAreaCd.value <> "" THEN
   STRFrBiz     = frm1.txtFrBizAreaCd.value
   ELSE 
   STRFrBiz     = "%"
   END IF
   IF frm1.txtToBizAreaCd.value <> "" THEN
   STRToBiz     = frm1.txtToBizAreaCd.value
   ELSE 
   STRToBiz     = "ZZZZZ"
   END IF
          
    condvar = "FR_DT|" & StrFrEndYmd & "|TO_DT|" & StrToEndYmd & "|ID|" & STRID & "|TYPE|" & STRTYPE & "|CTRL|" & STRCtrl & "|FRBIZ|" & STRFrBiz & "|TOBIZ|" & STRTOBiz
    'msgbox condvar
    StrEbrFile = "A5153OA1_KO441"
   
End Sub

'==========================================================================================
'   Event Name : txtTOIN_YMD
'   Event Desc : Date OCX KeyPress
'==========================================================================================
Function  txtToEndYmd_KeyPress(KeyAscii)
	txtToEndYmd_KeyPress = false
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
	txtToEndYmd_KeyPress = true
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY SCROLL="NO" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정잔액증감명세서(S)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
                    <TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
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
                                <TD CLASS="TD5" NOWRAP>회계일자</TD>
								<TD CLASS="TD6" NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> NAME="txtFromGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="회계일자1" id=fpDateTime1></OBJECT>&nbsp;~&nbsp;
												       <OBJECT classid=<%=gCLSIDFPDT%> NAME="txtToGlDt"   CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="회계일자2" id=fpDateTime2></OBJECT></TD>
						        <TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFrBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="사업장"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(1)">
								                       <INPUT TYPE=TEXT NAME="txtFrBizAreaNm" SIZE=25 tag="14">&nbsp;~&nbsp;</TD>
						    </TR>	
						    <TR>
						        <TD CLASS="TD5" NOWRAP>계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="시작계정코드"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(2)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=25 tag="14"></TD>
						        <TD CLASS="TD5" NOWRAP></TD>								
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="사업장"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(5)">
								                       <INPUT TYPE=TEXT NAME="txtToBizAreaNm" SIZE=25 tag="14"></TD>
						    </TR>						
						    <TR>
								<TD CLASS="TD5" NOWRAP>계정코드</TD>
								<TD CLASS="TD656" colspan = 3 NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd1" SIZE=10 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="종료계정코드"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcctCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(3)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm1" SIZE=25 tag="14"></TD>
						   	</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vspdData>
										<PARAM NAME="MaxCols" VALUE="0">
										<PARAM NAME="MaxRows" VALUE="0">
									</OBJECT>
								</TD>
							</TR>
						</TABLE>
					</TD>			
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
			    <TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%>
					
					<TR>
					<TD>
						&nbsp;&nbsp;<BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreView()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
					</TD>
					</TR>
				    
					
					</TD>
					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtChgFlag"    tag="24">
<INPUT TYPE=HIDDEN NAME="txtID"       TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>
