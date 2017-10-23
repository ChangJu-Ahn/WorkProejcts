
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 회계 
*  2. Function Name        : 기초미결등록 
*  3. Program ID           : a5402ma
*  4. Program Name         : 기초미결등록 
*  5. Program Desc         : 기초미결등록,수정,삭제,조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/11/05
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : 문희정 
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

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">    </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "a5402mb1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================
'⊙: Grid Columns

'@Grid_Column
Dim C_CHOICE
Dim C_GLNO
Dim C_GLSEQ
Dim C_GL_DT
Dim C_GLDESC
Dim C_OPENAMT
Dim C_OPENLOCAMT
                        
Dim C_Name		
Dim C_SumOpenDocAmt 	
Dim C_SumOpenAmt 	

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
'<ADD>*******************************
Sub initSpreadPosVariables()
	C_CHOICE	= 1
	C_GLNO		= 2
	C_GLSEQ		= 3
	C_GL_DT		= 4
	C_GLDESC	= 5
	C_OPENAMT	= 6
	C_OPENLOCAMT= 7
End Sub
'***********************************
'========================================================================================================

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
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

	frm1.txtCurrency.value	= Parent.gCurrency
	EndDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(EndDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)

	frm1.txtGLDate.Text = StartDate
	frm1.txtGLDate1.Text = EndDate
	
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

	<%Call LoadInfTB19029A("I","*", "COOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "A","COOKIE","MA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(ByVal Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(ByVal pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   select case pOpt
		  case "MQ"
'				lgKeyStream = Frm1.txtAccountCd.text & Parent.gColSep							'You Must append one character(Parent.gColSep)
'				lgKeyStream = lgKeyStream & Frm1.txtCurrency.text & Parent.gColSep
		  case "MN"	
   end select
    
	 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call SetCombo(frm1.txtOpenAcctFg ,"N" ,"미연결")
	Call SetCombo(frm1.txtOpenAcctFg ,"Y" ,"연결")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()

End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	With frm1.vspdData
		
		'<ADD>************************
		Call initSpreadPosVariables()
		'*****************************
       .MaxCols   = C_OPENLOCAMT + 1                                                  ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True
                                                              ' ☜: Clear spreadsheet data 
        ggoSpread.Source = frm1.vspdData
        
		ggoSpread.Spreadinit "V20030116",,parent.gAllowDragDropSpread    
		Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 	        
		.ReDraw = false	

		Call GetSpreadColumnPos("A")
      ' Call AppendNumberPlace("6","4","2")
                             'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
       ggoSpread.SSSetCheck		C_CHOICE		,"선택"			,8	   ,2                  ,     ,15
       ggoSpread.SSSetEdit		C_GLNO			,"전표번호"     ,13    ,0                  ,     ,18     ,2
       ggoSpread.SSSetEdit		C_GLSEQ			,"순번"			,13    ,0                  ,     ,18     ,2
       ggoSpread.SSSetDate		C_GL_DT			,"전환일자"		,10    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetEdit		C_GLDESC		,"적요"			,30    ,0                  ,     ,128		,2
       ggoSpread.SSSetFloat		C_OPENAMT		,"금액"	,17    ,"2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,	,		,"Z"
       ggoSpread.SSSetFloat		C_OPENLOCAMT	,"금액(자국)"	,17    ,"2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,	,		,"Z"

	   .ReDraw = true
    
    End With

	Call SetSpreadLock 	

End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
		.vspdData.ReDraw = False 
		ggoSpread.Source = frm1.vspdData
		      
		                         'Col-1          Row-1       Col-2           Row-2   
		ggoSpread.SpreadLock       C_GLNO		, -1         , C_GLNO			, -1 
		ggoSpread.SpreadLock       C_GLSEQ		, -1         , C_GLSEQ			, -1 
		ggoSpread.SpreadLock       C_GLDESC		, -1         , C_GLDESC			, -1 
		ggoSpread.SpreadLock       C_GL_DT		, -1         , C_GL_DT			, -1 
		ggoSpread.SpreadLock       C_OPENAMT	, -1         , C_OPENAMT			, -1 
		ggoSpread.SpreadLock       C_OPENLOCAMT	, -1         , C_OPENLOCAMT			, -1 
				
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
                                'Col          Row   Row2
      ggoSpread.SSSetProtected   C_GLNO		, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_GLSEQ	, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_GLDESC	, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_GL_DT	, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_OPENAMT	, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_OPENLOCAMT, pvStartRow, pvEndRow
      
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

			C_CHOICE		= iCurColumnPos(1)
			C_GLNO			= iCurColumnPos(2)
			C_GLSEQ			= iCurColumnPos(3)
			C_GL_DT			= iCurColumnPos(4)
			C_GLDESC		= iCurColumnPos(5)
			C_OPENAMT		= iCurColumnPos(6)
			C_OPENLOCAMT	= iCurColumnPos(7)

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

	frm1.btnRun2.disabled = true
	frm1.btnRun3.disabled = true
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
      
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal 
    Call InitComboBox

	frm1.txtAccountCd.focus
'①메뉴탐색기 ②조회    ③신규    ④삭제        ⑤저장    ⑥행추가       ⑦행삭제 
'⑧취소       ⑨이전    ⑩이후    ⑪레코드복사  ⑫Export  ⑬인쇄         ⑭찾기	
	Call SetToolbar("11000000000111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	'Call CookiePage (0)   
                                                           '☜: Check Cookie
'call msgbox("20021106 Test중")
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
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False															  '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    

    Call ggoOper.ClearField(Document, "2")										  '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									          '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                            '⊙: Initializes local global variables

	If DbQuery("MQ") = False Then                                                 '☜: Query db data
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
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDelete = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'=======================================================================================================
' Function Name : FncSave
' Function	 Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                   '☜: Protect system from crashing
	'-----------------------
    'Precheck area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False  AND ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          'No data changed!!
        Exit Function
    End If    

	'-----------------------
    'Save function call area
    '----------------------- 

    IF  DbSave	= False Then
					                                                '☜: Save db data
		Exit Function
    End If
    
    
    FncSave = True                                                          
    
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
    Dim iDx



End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()


End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()


End Function
'========================================================================================================
' Name : fpdtFoundDt_ButtonHit
' Desc : developer describe this line
'========================================================================================================
Sub fpdtFoundDt_ButtonHit(Button, NewIndex)
	On Error Resume Next
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : fpdtCloseDt_ButtonHit
' Desc : developer describe this line
'========================================================================================================
Sub fpdtCloseDt_ButtonHit(Button, NewIndex)
	On Error Resume Next
    lgBlnFlgChgValue = True
End Sub


'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line 
'========================================================================================================
Function FncPrint() 

Dim StrEbrFile, ValAcctCd, ValDocCur
Dim StrUrl
Dim lngPos
Dim intCnt
Dim IntRetCd
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	StrEbrFile = "a5402ma1"
	
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If
		
	Call SetPrintCond(StrEbrFile, ValAcctCd, ValDocCur)

	StrUrl = StrUrl & "AcctCd|" & ValAcctCd
	StrUrl = StrUrl & "|DocCur|" & ValDocCur
	
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call FncEBRPrint(EBAction,StrEbrFile,StrUrl)
	
End Function


'========================================================================================================
' Name : FncPreview
' Desc : This function is related to Preview Button
'========================================================================================================
Function FncPreview()
 
Dim StrEbrFile, ValAcctCd, ValDocCur
Dim StrUrl
Dim lngPos
Dim intCnt
Dim IntRetCd
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If
		
	Call SetPrintCond(StrEbrFile, ValAcctCd, ValDocCur)

	StrUrl = StrUrl & "AcctCd|" & ValAcctCd
	StrUrl = StrUrl & "|DocCur|" & ValDocCur

	'------ Developer Coding part (End ) -------------------------------------------------------------- 
	
	Call FncEBRPreview(StrEbrFile,StrUrl)
			
End Function

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile, ValAcctCd, ValDocCur)
	
	StrEbrFile = "a5402ma1.ebr"
	
	With frm1

		ValAcctCd	= UCase(Trim(.txtAccountCd.value))
		ValDocCur	= UCase(Trim(.txtCurrency.value))
	End With
End Sub
	

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then	 
       FncNext = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

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

	Call Parent.FncFind(Parent.C_MULTI, True)

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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData

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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		              '⊙: Data is changed.  Do you want to exit? 
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
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If UCase(Trim(frm1.txtOpenAcctFg.value)) = "Y" THen
		frm1.btnRun2.disabled = true
    Else
		frm1.btnRun3.disabled = true
    ENd if

	Call MakeKeyStream(pDirect)
  
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="       & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream=" & lgKeyStream         '☜: Query Key
        strVal = strVal     & "&txtMaxRows="    & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey        '☜: Next key tag

    End With

	strVal = strVal & "&txtDateFr="				& frm1.txtGLDate.Text
	strVal = strVal & "&txtDateTo="				& frm1.txtGLDate1.Text
	strVal = strVal & "&txtAccountCd="			& UCase(Trim(frm1.txtAccountCd.value))
	strVal = strVal & "&txtOpenAcctFg="			& UCase(Trim(frm1.txtOpenAcctFg.value))
	strVal = strVal & "&txtCurrency="			& UCase(Trim(frm1.txtCurrency.value))

	If lgStrPrevKey = "" then	frm1.vspdData1.MaxRows = 0
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    If Err.number = 0 Then	 
       DbQuery = True                                                             '☜: Processing is OK
    End If

    If UCase(Trim(frm1.txtOpenAcctFg.value)) = "Y" THen
		frm1.btnRun3.disabled = false
    Else
		frm1.btnRun2.disabled = false
    ENd if

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : DbSave
' Desc : This sub is 
'========================================================================================================
Function DbSave()
		
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbSave = False                                                                '☜: Processing is NG

   ' Call DisableToolBar(Parent.TBC_SAVE)                                                 '☜: Disable Save Button Of ToolBar
   

    Call LayerShowHide(1) 
                                                            '☜: Show Processing Message
	With frm1
		.txtFlgMode.value			= lgIntFlgMode
		.txtUpdtUserId.value		= Parent.gUsrID
		.txtInsrtUserId.value		= Parent.gUsrID
		.txtMode.value				= Parent.UID_M0002
		.txtAuthorityFlag.value     = lgAuthorityFlag               '권한관리 추가 
				
		.hOrgChangeId.value = Parent.gChangeOrgId
	End With		
                                       '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    ggoSpread.Source = frm1.vspdData

    strVal = ""
    lGrpCnt = 1

	With Frm1
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C_CHOICE
			If .vspdData.Text = "1" Then
												  strVal = strVal & lRow							& Parent.gColSep                                                    
				.vspdData.Col = C_GLNO			: strVal = strVal & Trim(.vspdData.Text)			& Parent.gColSep
				.vspdData.Col = C_GLSEQ			: strVal = strVal & Trim(.vspdData.Text)			& Parent.gRowSep
				lGrpCnt = lGrpCnt + 1
            End If
		Next
	End With

	frm1.txtMaxRows.value     = lGrpCnt-1	
	frm1.txtSpread.value = strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002							'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtOpenAcctFg="		& frm1.txtOpenAcctFg.value 
	strVal = strVal & "&txtSpread="			& frm1.txtSpread.value

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
'   Event Name : OpenPopupOpeenAcct
'   Event Desc : 
'========================================================================================================
Function OpenPopupOpeenAcct()
	Dim strVal
	Dim lRow
	Dim lGrpCnt
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    OpenPopupOpeenAcctFg = False                                                              '☜: Processing is NG
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

	strVal = ""

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtDateFr="			& frm1.txtGLDate.Text
	strVal = strVal & "&txtDateTo="			& frm1.txtGLDate1.Text
	strVal = strVal & "&txtAccountCd="		& frm1.txtAccountCd.value 
	strVal = strVal & "&txtOpenAcctFg="		& frm1.txtOpenAcctFg.value 
	strVal = strVal & "&txtCurrency="		& frm1.txtCurrency.value 

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

	OpenPopupOpeenAcctFg = True
	Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
'   Event Name : OpenPopupOpeenAcctCancel
'   Event Desc : 
'========================================================================================================
Function OpenPopupOpeenAcctCancel()
	Dim strVal
	Dim lRow
	Dim lGrpCnt
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    OpenPopupOpeenAcctCancel = False                                                              '☜: Processing is NG
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

	strVal = ""
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtDateFr="			& frm1.txtGLDate.Text
	strVal = strVal & "&txtDateTo="			& frm1.txtGLDate1.Text
	strVal = strVal & "&txtAccountCd="		& frm1.txtAccountCd.value 
	strVal = strVal & "&txtOpenAcctFg="		& frm1.txtOpenAcctFg.value 
	strVal = strVal & "&txtCurrency="		& frm1.txtCurrency.value 

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

	OpenPopupOpeenAcctCancel = True
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()					'☆: 저장 성공후 실행 로직 

	lgBlnFlgChgValue = false
	
	'rm1.txtTempGlNo.value = UCase(Trim(TempGlNo))
    frm1.txtCommandMode.value = "UPDATE"
    
	Call ggoOper.ClearField(Document, "2")      '⊙: Condition field clear    
    Call InitVariables							'⊙: Initializes local global variables

	

	call DbQuery("MQ")
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbDelete = False                                                              '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	 
       DbDelete = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("11001000000111")                                              '☆: Developer must customize
	Frm1.vspdData.Focus
    Call InitData()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")
	
    Set gActiveElement = document.ActiveElement   

End Sub
	
'========================================================================================================
' Name : DbexeOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbexeOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery("MN") = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
    Set gActiveElement = document.ActiveElement   

End Sub

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

 '**********************  기초미결연결 Popup  ****************************************
'	기능: 기초미결연결 Popup창 오픈 
'   설명: 
'************************************************************************************** 
Function OpenPopupOpeenAccti()

	Dim arrRet
	Dim arrParam(2)	
	Dim arrRowVal,arrColVal
	Dim dr_sum,iDx
	Dim tmpFrDate,tmpToDate,FrYear,FrMonth,FrDay,ToYear,ToMonth,ToDay
	If IsOpenPop = True Then Exit Function
	
	
	
	arrParam(0) = Trim(frm1.txtAccountCd.value)
	arrParam(1) = Trim(frm1.txtAccountNm.value )
	arrParam(2) = Trim(frm1.txtCurrency.value)

'	If arrParam(0) = "전표번호" Then Exit Function		' 조회 이전 


   
	arrRet = window.showModalDialog("a5402ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=900px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		     '220
     

	if arrRet = "" then
		IsOpenPop = False
		
		Exit function
	end if
				
	IsOpenPop = False
	
	
End Function
 '**********************  회계전표 Popup  ****************************************
'	기능: 
'   설명: 
'************************************************************************************** 
Function OpenPopupGL()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5120ra1")


	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_gLno
			
		arrParam(0) = Trim(.Text)	'결의전표번호 
		arrParam(1) = ""			'Reference번호 
	End With

	
	If arrParam(0) = "전표번호" Then Exit Function		' 조회 이전 
	
	IsOpenPop = True   
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	
	IsOpenPop = False
	
End Function

 '**********************  POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd, strTempBankCd

	If IsOpenPop = True Then Exit Function

	
	Select Case iWhere
		Case 0		'계정코드 
			
			arrParam(0) = "계정 팝업"									' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD	and	MGNT_FG=" & FilterVar("Y", "''", "S") & " "	' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
			
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정코드명"									' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)

		Case 1		'거래통화 
			If frm1.txtCurrency.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If
			
			arrParam(0) = "통화코드 팝업"				' 팝업 명칭 
			arrParam(1) = "B_Currency"	    			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "통화코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "Currency"	    			' Field명(0)
			arrField(1) = "Currency_desc"	    		' Field명(1)
    
			arrHeader(0) = "통화코드"					' Header명(0)
			arrHeader(1) = "통화코드명"	
			
		    
		Case Else
			Exit Function
	End Select
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet, iWhere)
		'Call SetRefOpenAr(arrRet)
	End If	

End Function

 '**********************  SetReturnValue  ****************************************
'	기능: 기준 POP-UP에서 선택한 값을 Matching
'************************************************************************************** 

Function SetReturnVal(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 0	'계정코드 
				.txtAccountCD.value = arrRet(0)
				.txtAccountNM.value = arrRet(1)
'				.txtBankAcctNo.value = ""
			Case 1	'거래화폐 
				.txtCurrency.Value = arrret(0)
		End Select
	End With

End Function



'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

Dim lstxtAmtSum, lstxtLocAmtSum
	lstxtAmtSum = 0
	lstxtLocAmtSum = 0
	
    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 Then
			.Row = Row
'			.Col = C_CHOICE

				.Col = Col
				If ButtonDown = 1 Then
					ggoSpread.UpdateRow Row
					.col = C_OPENAMT
					lstxtAmtSum = UNIFormatNumber(UNICDbl(frm1.txtSumAmt.Text) + UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
					frm1.txtSumAmt.Text = lstxtAmtSum
					.col = C_OPENLOCAMT
					lstxtLocAmtSum = UNIFormatNumber(UNICDbl(frm1.txtSumLocAmt.Text) + UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
					frm1.txtSumLocAmt.Text = lstxtLocAmtSum
				Else
					ggoSpread.SSDeleteFlag Row,Row				
					.col = C_OPENAMT
					lstxtAmtSum = UNIFormatNumber(UNICDbl(frm1.txtSumAmt.Text) - UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
					frm1.txtSumAmt.Text = lstxtAmtSum
					.col = C_OPENLOCAMT
					lstxtLocAmtSum = UNIFormatNumber(UNICDbl(frm1.txtSumLocAmt.Text) - UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
					frm1.txtSumLocAmt.Text = lstxtLocAmtSum
				End If		
'			End If
		End If
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
       
'   	Frm1.vspdData.Row = Row
'   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
  
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
'    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
	
'	ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)

    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"	'Split 상태코드 

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
    Else
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_ProcurTypeNm Or NewCol <= C_ProcurTypeNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
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
    Dim iColumnName
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기초미결등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPopupGL()">회계전표</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAccountCd" SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAccountCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtAccountCd.value,0)">
														   <INPUT TYPE=TEXT NAME="txtAccountNm" SIZE=20 tag="14" ALT="계정명" STYLE="TEXT-ALIGN: Left">
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>연결상태</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="txtOpenAcctFg" STYLE="Width:150px;" tag="12XXXU" ALT="연결상태"></SELECT>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>전표일자</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5402ma1_fpDateTime1_txtGLDate.js'></script>&nbsp;~&nbsp;
												           <script language =javascript src='./js/a5402ma1_fpDateTime2_txtGLDate1.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>거래통화</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCurrency" ALT="거래통화" MAXLENGTH="5" SIZE=5 STYLE="TEXT-ALIGN: left" tag ="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCurrency.Value,1)">
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
								<TD HEIGHT="94%" NOWRAP>
									<script language =javascript src='./js/a5402ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100% >
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>		
								<TR>
									<TD CLASS=TD5 NOWRAP>합계</TD>
									<TD CLASS=TD6><script language =javascript src='./js/a5402ma1_txtSumAmt_txtSumAmt.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>합계(자국)</TD>
									<TD CLASS=TD6><script language =javascript src='./js/a5402ma1_txtSumLocAmt_txtSumLocAmt.js'></script></TD>
								</TR>
							</TABLE>
						</FIELDSET>
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
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><!--BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncPrint()" Flag=1>인쇄</BUTTON-->&nbsp;
					</TD>
					<TD WIDTH=* align=right>
					    <BUTTON NAME="btnRun2" CLASS="CLSMBTN" ONCLICK="vbscript:OpenPopupOpeenAcct()" >일괄연결</BUTTON>
					    <BUTTON NAME="btnRun3" CLASS="CLSMBTN" ONCLICK="vbscript:OpenPopupOpeenAcct()" >일괄연결취소</BUTTON>
					</TD>
					<TD WIDTH=10></TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		<!-- Debug -->
		<!--<TD WIDTH=100% HEIGHT=150><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=100 FRAMEBORDER=1 SCROLLING=Yes framespacing=0></IFRAME>-->
		</TD>
	</TR>	
</TABLE>
<script language =javascript src='./js/a5402ma1_vaSpread3_vspdData3.js'></script>
<TEXTAREA class=hidden name=txtSpread    tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3   tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24">
<INPUT TYPE=hidden NAME="htxtTempGlNo"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtCommandMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtAuthorityFlag"     tag="24" TABINDEX="-1"><!--권한관리추가 -->
<INPUT TYPE=HIDDEN NAME="hCongFg"        tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="hItemSeq"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows1" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
</BODY>
</HTML>

