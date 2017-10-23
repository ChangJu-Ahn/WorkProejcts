
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 경영손익 
*  2. Function Name        : 
*  3. Program ID           : gd004MA1
*  4. Program Name         : 영업그룹별 배부현황 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/06
*  8. Modified date(Last)  : 2002/01/04
*  9. Modifier (First)     : Kim Kyoung Ho
* 10. Modifier (Last)      : Kim Kyoung Ho
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "gd004mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------
Dim C_center                                                     'Column  for Spread Sheet 
Dim C_centerNm   															
Dim C_data       

Const C_SHEETMAXROWS_D  = 100                                          '☜: Fetch count at a time
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #2
'--------------------------------------------------------------------------------------------------------
Dim C_center1                                                    'Column  for Spread Sheet 
Dim C_centerNm1   															
Dim C_data1       


Const COOKIE_SPLIT      = 4877	                                      'Cookie Split String
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim lgIsOpenPop
Dim IsOpenPop
Dim lsStatusClick


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode       = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue   = False								    '⊙: Indicates that no value changed
	lgIntGrpCount      = 0										'⊙: Initializes Group View Size
    lgStrPrevKey       = ""                                     '⊙: initializes Previous Key
    lgStrPrevKeyIndex  = ""                                     '⊙: initializes Previous Key Index
    lgStrPrevKeyIndex1 = ""                                     '⊙: initializes Previous Key Index
    lgSortKey          = 1                                      '⊙: initializes sort direction
	
	frm1.hYYYYMM.value = ""	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	Dim StartDate
	StartDate	= "<%=GetSvrDate%>"                                               'Get Server DB Date

	frm1.fpdtWk_yymm.focus
	frm1.fpdtWk_yymm.text	= 	UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	
	Call ggoOper.FormatDate(frm1.fpdtWk_yymm, Parent.gDateFormat, 2)
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

	<% Call loadInfTB19029A("Q", "G", "NOCOOKIE", "QA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
Sub MakeKeyStream(pRow)

   Dim strYYYYMM

	strYYYYMM = frm1.hYYYYMM.Value     
 
	If lgCurrentSpd = "M" Then
 
       lgKeyStream       = strYYYYMM & Parent.gColSep
       


    Else
        frm1.vspdData.Row = pRow
	
		frm1.vspdData.Col = C_center
		frm1.txtGiveCostCd.value =  frm1.vspdData.Text
        
		lgKeyStream       = strYYYYMM                                        & Parent.gColSep                  '날짜                         'You Must append one character(Parent.gColSep)
		lgKeyStream       = lgKeyStream & frm1.txtGiveCostCd.value           & Parent.gColSep                  '품목    

		
    End If   

   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
'  SpreadSheet #1
	C_center     = 1                                                 
	C_centerNm   = 2															
	C_data       = 3
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables1()
'   SpreadSheet #2
	C_center1     = 1                                               
	C_centerNm1   = 2															
	C_data1       = 3

End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables() 
    '----------------------------------------------------------------------------------------------------
    ' Set SpreadSheet #1
    '----------------------------------------------------------------------------------------------------
	With Frm1.vspdData
	
       .MaxCols   = C_data + 1                                               '☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        '☜:☜: Hide maxcols
       .ColHidden = True                                                            '☜:☜:

       .OperationMode = 3                                                           '☜ 
        ggoSpread.Source = Frm1.vspdData
        ggoSpread.Spreadinit "V20021217",,parent.gAllowDragDropSpread 
        
        ggoSpread.ClearSpreadData()

	   .ReDraw = false
	
      Call GetSpreadColumnPos("A")
      
       ggoSpread.SSSetEdit  C_center     , "Profit Center"        ,15,   ,, 50,2
       ggoSpread.SSSetEdit  C_centerNm   , "Profit Center 명"     ,20,   ,, 30    
       ggoSpread.SSSetFloat C_data       , "금       액"          ,14,    Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	
	   .ReDraw = true
	
       lgActiveSpd = "M"
       Call SetSpreadLock(-1,-1) 
    
    End With
 End Sub
 
'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet1()    

	Call initSpreadPosVariables1() 
    '----------------------------------------------------------------------------------------------------
    ' Set SpreadSheet #2
    '----------------------------------------------------------------------------------------------------
	With Frm1.vspdData1
	
       .MaxCols   = C_data1 + 1                                               ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

        ggoSpread.Source = Frm1.vspdData1
        ggoSpread.Spreadinit "V20021217",,parent.gAllowDragDropSpread 

        ggoSpread.ClearSpreadData()

	   .ReDraw = false
	
       
       Call GetSpreadColumnPos("B")

       ggoSpread.SSSetEdit  C_center1     , "영업그룹"         ,15,   ,, 50,2
       ggoSpread.SSSetEdit  C_centerNm1   , "영업그룹명"       ,20,   ,, 80    
       ggoSpread.SSSetFloat C_data1       , "금       액"      ,14,    Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec       
       
	   .ReDraw = true
	
       lgActiveSpd = "S"
       Call SetSpreadLock(-1,-1) 
    
    End With

End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal lRow  , ByVal lRow2 )
    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If

    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
			  ggoSpread.Source = frm1.vspdData
			  ggoSpread.SpreadLockWithOddEvenRowColor()
  
        Case  "S"
              ggoSpread.Source = frm1.vspdData1
		      ggoSpread.SpreadLockWithOddEvenRowColor()
    End Select               
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
      With Frm1
             ggoSpread.Source = .vspdData
             .vspdData.ReDraw = False
                ggoSpread.SSSetProtected   C_center		,pvStartRow	,pvEndRow
                ggoSpread.SSSetProtected   C_centerNm	,pvStartRow	,pvEndRow
                ggoSpread.SSSetProtected   C_data		,pvStartRow	,pvEndRow
               
            .vspdData.ReDraw = True
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpread2Color()
	Dim strStartRow, strEndRow

    With frm1
    
		strStartRow = 1
		strEndRow	= .vspdData1.MaxRows

		ggoSpread.Source	= .vspdData1
		.vspdData1.ReDraw	= False
		
			ggoSpread.SSSetProtected   C_center1	,strStartRow, strEndRow
            ggoSpread.SSSetProtected   C_centerNm1	,strStartRow, strEndRow
            ggoSpread.SSSetProtected   C_data1		,strStartRow, strEndRow
		
		.vspdData1.ReDraw = True

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
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData1.Col    = iDx
              Frm1.vspdData1.Row    = iRow
              Frm1.vspdData1.Action = 0 
              Exit For
           End If
           
       Next
          
    End If   
End Sub



'======================================================================================================
'	Name : initMinor()
'	Description : 폼 로드시에 배부유형을 박아준다.
'=======================================================================================================%>
Function initMinor()

	Dim intRetCD   	  
	intRetCD = CommonQueryRs(" bm.minor_Cd, bm.minor_nm "," g_option go,b_minor bm","go.minor_Cd = bm.minor_cd and  go.major_cd = " & FilterVar("g1011", "''", "S") & " and  bm.major_cd = go.major_cd" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if intRetCd = False then
		Call CommonQueryRs(" bm.minor_Cd, bm.minor_nm ","b_minor bm"," bm.major_cd = " & FilterVar("g1011", "''", "S") & " and  bm.minor_cd = " & FilterVar("1", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		frm1.txtTypeCd.value= Trim(Replace(lgF0,Chr(11),""))
		frm1.txtTypeNm.value= Trim(Replace(lgF1,Chr(11),""))
	else
		frm1.txtTypeCd.value= Trim(Replace(lgF0,Chr(11),""))
		frm1.txtTypeNm.value= Trim(Replace(lgF1,Chr(11),""))
	End IF   	    
 	    
	   
End Function

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
			C_center 				= iCurColumnPos(1)
			C_centerNm   			= iCurColumnPos(2)
			C_data       			= iCurColumnPos(3)    
	   Case "B"
			ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_center1       		= iCurColumnPos(1)
			C_centerNm1     		= iCurColumnPos(2)
			C_data1         		= iCurColumnPos(3)  
    End Select    
End Sub


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
    Err.Clear                                                                       '☜: Clear err status
    
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
	Call initMinor()
	Call InitSpreadSheet1  
	Call InitVariables
    Call SetDefaultVal

	Call SetToolbar("1100000000001111")                                             '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
	Call CookiePage (0)                                                             '☜: Check Cookie
			
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
     Dim Row
   Dim strYear,strMonth,strDay 
       
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			         '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
    
    Call InitVariables															 '⊙: Initializes local global variables
'    Call SetDefaultVal	


    If Not chkField(Document, "1") Then									         '⊙: This function check indispensable field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
  if lgCurrentSpd = "M2" then
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
				
		Row = frm1.vspdData.ActiveRow
		Call MakeKeyStream(Row)
	
		If DbQuery2 = False Then
			Exit Function
		End If																'☜: Query db data
    else
		Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field  
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData		

   
		Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
		frm1.hYYYYMM.value = strYear & strMonth
			
		lgCurrentSpd = "M"   
		Call MakeKeyStream("X")
		If DbQuery = False Then
			Exit Function
		End If																'☜: Query db data  
    end if     
    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
'    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData1

    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = Frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call MakeKeyStream("S")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 




	'------ Developer Coding part (End )   -------------------------------------------------------------- 
     If Frm1.vspdData1.MaxRows < 1 Then
        Exit Function
     End If
    
     With Frm1
          If .vspdData1.ActiveRow > 0 Then
             .vspdData1.ReDraw = False
		
              ggoSpread.Source = .vspdData1	
              ggoSpread.CopyRow
              SetSpreadColor   .vspdData1.ActiveRow, .vspdData1.ActiveRow
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
              .vspdData1.Col  = 1
              .vspdData1.Text = ""
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
             .vspdData1.ReDraw = True
             .vspdData1.Focus
         End If
    End With

    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                           '☜: Processing is NG
    Err.Clear                                                                   '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    ggoSpread.Source = Frm1.vspdData1	
    ggoSpread.EditUndo  

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   
    FncCancel = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'====================================================================================================
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
    End If                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    If Frm1.vspdData1.MaxRows < 1 then
       Exit function
    End if	

    With Frm1.vspdData1 
              .Focus
              ggoSpread.Source = frm1.vspdData1 
              lDelRows = ggoSpread.DeleteRow
    End With
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                         '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                            '☜: Processing is NG
    Err.Clear                                                                   '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
    FncPrint = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
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
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
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
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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

	Dim indx
	Dim lngActRow1, lngActRow2
	Dim lngActCol1, lngActCol2

	lngActRow1 = frm1.vspdData.ActiveRow
	lngActCol1 = frm1.vspdData.ActiveCol
	lngActRow2 = frm1.vspdData1.ActiveRow
	lngActCol2 = frm1.vspdData1.ActiveCol
	
	If gActiveSpdSheet.Name <> "" Then
		For indx = 0 To frm1.vspdData.MaxRows
			frm1.vspdData.Row = indx
			frm1.vspdData.Col = 0
'			If frm1.vspdData.Text = ggoSpread.DeleteFlag Or _
'			   frm1.vspdData.Text = ggoSpread.UpdateFlag Then
'				Call FncUndoData(indx)
'			End If
			
			Select Case Trim(UCase(gActiveSpdSheet.Name))
				Case "VSPDDATA"
					frm1.vspdData.Row = lngActRow1 
					frm1.vspdData.Col = lngActCol1
					frm1.vspdData.Action = 0
					
			   		
				Case "VSPDDATA1"
					frm1.vspdData1.Row = lngActRow2
					frm1.vspdData1.Col = lngActCol2
					frm1.vspdData1.Action = 0
			End Select
		Next
	End If

	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
    
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call InitSpreadSheet()
'			Call InitSpreadComboBox()
			Call ggoSpread.ReOrderingSpreadData()
'			Call InitData()
			Call InitMinor()
			
		Case "VSPDDATA1"
'			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet1()			' 그리드2 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSprea2dColor()
	End Select

	If frm1.vspdData1.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If

End Sub
'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	ggoSpread.Source = Frm1.vspdData1
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True                                                               '☜: Processing is OK
End Function

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
    
    Dim strVal
    Dim Row
    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False                                                                  '☜: Processing is NG

    if LayerShowHide(1) = False then
	   Exit Function

	end if                                                         '☜: Show Processing Message

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
       strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex              '☜: Next key tag
'       strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)         '☜: Max fetched data at a time
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '☜: Max fetched data

    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
'    End If   

    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic

    DbQuery = True                                                                   '☜: Processing is NG

End Function

'========================================================================================================
' Name : DbQuery2
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery2()
    
    Dim strVal
    Dim Row
    Err.Clear                                                                        '☜: Clear err status

    DbQuery2 = False    
    
    
    Row = frm1.vspdData.ActiveRow

    If lsStatusClick = True Then
		if LayerShowHide(1) = False then
			Exit Function
		end if
	End If
    

    
    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
       strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex1             '☜: Next key tag
 '      strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D1)        '☜: Max fetched data at a time
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData1.MaxRows         '☜: Max fetched data
       strVal = strVal     & "&txtAmount="          & Frm1.txtAmount.Value
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic    

    DbQuery2 = True  
End Function


'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
    Err.Clear                                                                      '☜: Clear err status
    DbSave = False                                                                 '☜: Processing is NG
	
if LayerShowHide(1) = false then
 exit function
end if

		
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
  	With Frm1
		.txtMode.value      = Parent.UID_M0002                                            '☜: Delete
		.txtKeyStream.value = lgKeyStream
	End With

    strVal  = ""
    strDel  = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData1.MaxRows
    
           .vspdData1.Row = lRow
           .vspdData1.Col = 0
        
           Select Case .vspdData1.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                     strVal = strVal & "C" & Parent.gColSep                    '0
                                                     strVal = strVal & lRow & Parent.gColSep                    '1
                    .vspdData1.Col = C_MinorCd     : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep   '3
                    .vspdData1.Col = C_MinorNm     : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep   '4
                    .vspdData1.Col = C_MinorTypeCd : strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                     strVal = strVal & "U" & Parent.gColSep
                                                     strVal = strVal & lRow & Parent.gColSep
                    .vspdData1.Col = C_MinorCd     : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_MinorNm     : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_MinorTypeCd : strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                     strDel = strDel & "D" & Parent.gColSep
                                                     strDel = strDel & lRow & Parent.gColSep
                    .vspdData1.Col = C_MinorCd     : strDel = strDel & Trim(.vspdData1.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                                '☜: Processing is OK
    Set gActiveElement = document.ActiveElement   
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
    Call SetToolbar("1100000000011111")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
 	If lgCurrentSpd = "M" Then
		
		Call InitData()
        
        frm1.vspdData.row = 1
        frm1.vspdData.col = 1
        
		If frm1.vspdData.MaxRows >= 1 then
			lgCurrentSpd       = "M1"
		    Call MakeKeyStream(1)	    
		    Call Dbquery2()

        Else			
        End if

    Else
        lgCurrentSpd = "M"
    	Call InitData()
	End If

    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Name : DbQueryOk2
' Desc : Called by MB Area when query operation is successful
'========================================================================================================	
Function DbQueryOk2()

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call SetToolbar("1100000000011111")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

		Call InitData()

        lgCurrentSpd = "M"
	lsStatusClick = False
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

End Function	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
	Call InitVariables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    ggoSpread.Source = Frm1.vspdData1
    Frm1.vspdData1.MaxRows = 0
    lgCurrentSpd = "S"

	Call SetToolbar("110000000000000")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DBQuery()
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================


'=======================================================================================================
'   Event Name : txtYyyymm_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtYyyymm_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtDilig_dt_2_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub



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

	With Frm1.vspdData 
    
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
    Call SetPopupMenuItemInf("0000111111")
    
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
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
    
    If frm1.vspdData.MaxRows <= 0 Then
		Exit Sub
	End If
	'If Row = frm1.vspdData.Row Then	
	'	Exit Sub	
	'End If
	IF Row <> 0 Then
		ggoSpread.Source = frm1.vspdData

		frm1.vspdData1.MaxRows = 0
    
		lgCurrentSpd = "M2"
		lsStatusClick = True

		Call DBquery
    ENd IF
    
    
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



Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	gMouseClickStatus = "SPC"	'Split 상태코드    
	
    If Row <> NewRow And NewRow > 0 Then
		
		ggoSpread.Source       = Frm1.vspdData1		
		frm1.vspdData1.MaxRows = 0
	    lgCurrentSpd = "S"
		lgStrPrevKeyIndex1 = ""
		
	    Call MakeKeyStream(NewRow)

		
		If DBQuery = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If	    
	    
	End If    
	    

End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_OnFocus()
    lgActiveSpd      = "M"
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
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_Cost_Nm Or NewCol <= C_Cost_Nm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
    
End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_Cost_Nm Or NewCol <= C_Cost_Nm Then
     '   Cancel = True
     '   Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
    
End Sub


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )

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
'   Event Name : vspdData1_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_Change( Col ,  Row)

    Dim iDx
    Frm1.vspdData1.Row = Row
    Frm1.vspdData1.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
  
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
	Call CheckMinNumSpread(frm1.vpsdData1, Col, Row)
	
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row

End Sub
'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(Col, Row)
    Call SetPopupMenuItemInf("0000111111")
    
    gMouseClickStatus = "SP1C" 
    
    Set gActiveSpdSheet = frm1.vspdData1
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
	  
End Sub
'========================================================================================================
'   Event Name : vspdData1_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_OnFocus()
    lgActiveSpd      = "S"
End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
End Sub    
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if Frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	
       If lgStrPrevKeyIndex1 <> "" Then                         
          lgCurrentSpd = "S"
          DbQuery
       End If
    End if
End Sub


Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.fpdtWk_yymm.focus
	End If
End Sub


'========================================================================================================
' Name : BtnPreview
' Desc : This function is related to Preview Button
'========================================================================================================
Function FncBtnPreview() 
'On Error Resume Next                                                    '☜: Protect system from crashing
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
	dim condvar
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
	dim var1,var2,var3,var4
	'Dim strYear, strMonth, strDay
	Dim IntRetCD
	Dim lngPos
	Dim intCnt
	If frm1.vspdData.MaxRows = 0 Then
		IntRetCD = DisplayMsgBox("900002", "X", "X", "X")			'⊙: "Will you destory previous data"
		' 조회를 먼저 하십시오.
		Exit Function
	End if
	
	StrEbrFile = "gd004ma1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
	
    'Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	var1 = frm1.hYYYYMM.Value
	'--출력조건을 지정하는 부분 수정 - 끝 
	
	condvar = "YYYYMM|" & var1
	
	Call FncEBRPreview(ObjName,condvar)

End Function

'========================================================================================================
' Name : FncBtnPrint
' Desc : developer describe this line 
'========================================================================================================
Function FncBtnPrint() 
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
	Dim var1,var2, var3, var4
    Dim strYear, strMonth, strDay
    Dim IntRetCD
    
    If Not chkField(Document, "1") Then                                  '⊙: This function check indispensable field%>
       Exit Function
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
		IntRetCD = DisplayMsgBox("900002", "X", "X", "X")			'⊙: "Will you destory previous data"
			' 조회를 먼저 하십시오.
		Exit Function
	End if
	
    '--출력조건을 지정하는 부분 수정 
	
	StrEbrFile = "gd004ma1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
	
	Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	var1 = frm1.hYYYYMM.Value

    '--출력조건을 지정하는 부분 수정 
	condvar = "YYYYMM|" & var1
	
	Call FncEBRPrint(EBAction,ObjName,condvar)
		
End Function


'=======================================================================================================
'   Event Name : fpdtWk_yymm_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub fpdtWk_yymm_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub




</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY SCROLL="no" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>영업그룹별 배부현황</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
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
									<TD CLASS=TD5 NOWRAP>대상년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/gd004ma1_fpDateTime3_fpdtWk_yymm.js'></script>										
									</TD>									
									<TD CLASS=TD5 NOWRAP>배부유형</TD>                                    
									<TD CLASS=TD6 NOWRAP>                                    										
										<INPUT TYPE=TEXT NAME="txtTypeCd" SIZE=3 MAXLENGTH=1 tag="14XXU" ALT="배부유형코드">
										<INPUT TYPE=TEXT NAME="txtTypeNm" SIZE=22 MAXLENGTH=25 tag="14XXU"  ALT="배부유형">
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
					<TD HEIGHT=* WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD WIDTH="50%">
									<TABLE  <%=LR_SPACE_TYPE_20%> CLASS="BasicTB" CELLSPACING=0>
										<TR>
										<TD  COLSPAN=3>
											<script language =javascript src='./js/gd004ma1_vaSpread1_vspdData.js'></script>
										</TD>
										</TR>
										<TR HEIGHT=20>
											
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>											
											<TD CLASS=TD5 NOWRAP>합계</TD>
								            <TD CLASS=TD6 NOWRAP>
									            <script language =javascript src='./js/gd004ma1_txtDataAmt_txtDataAmt.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD WIDTH="50%">
									<TABLE  <%=LR_SPACE_TYPE_20%> CLASS="BasicTB" CELLSPACING=0>
										<TR>
											<TD COLSPAN=3>
												<script language =javascript src='./js/gd004ma1_vaSpread2_vspdData1.js'></script>
											</TD>
										</TR>
										<TR HEIGHT=20>
										    <TD CLASS=TD6 NOWRAP>&nbsp;</TD>										    
											<TD CLASS=TD5 NOWRAP>합계</TD>								            
											<TD CLASS="TD6" >
												<script language =javascript src='./js/gd004ma1_txtDataAmt1_txtDataAmt1.js'></script></TD>			    	    		            			    	    		            
										</TR>										
									</TABLE>
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
	
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
	
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"       TAG="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN       NAME="txtMode"         TAG="24" tabindex="-1">
<INPUT TYPE=HIDDEN       NAME="txtKeyStream"    TAG="24" tabindex="-1">
<INPUT TYPE=HIDDEN       NAME="txtFlgMode"      TAG="24" tabindex="-1">
<INPUT TYPE=HIDDEN       NAME="txtPrevNext"     TAG="24" tabindex="-1">
<INPUT TYPE=HIDDEN       NAME="txtMaxRows"      TAG="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtGiveCostCd" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtAmount" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="TotalAmount" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24" tabindex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>

