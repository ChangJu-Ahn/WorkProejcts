<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5954MA1
'*  4. Program Name         : 월차 환율 등록 
'*  5. Program Desc         : 회계관리 / 월차계산 / 웦차환율등록 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/15
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : Kim Kyoung Ho
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================
Const BIZ_PGM_ID = "a5954mb1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Dim C_DEALMON
Dim C_BTN	
Dim C_DEALRATE
Dim C_ExcCurDate	



Const COOKIE_SPLIT      = 4877	                                      '☆: Cookie Split String
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop

Dim lgIsOpenPop          

<%
Dim lsSvrDate
lsSvrDate = GetsvrDate
%>

'========================================================================================================
Sub InitSpreadPosVariables()

	 C_DEALMON		= 1                                                 'Column ant for Spread Sheet 
	 C_BTN			= 2
	 C_DEALRATE	= 3															
	 C_ExcCurDate	= 4

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE
	lgBlnFlgChgValue  = False
    lgStrPrevKey      = ""
    lgStrPrevKeyIndex = ""
    lgSortKey         = 1
		
End Sub

'========================================================================================================
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call ExtractDateFrom("<%=lsSvrDate%>",Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	
	frm1.fpdtWk_yymm.Year	= strYear
	frm1.fpdtWk_yymm.Month	= strMonth
	frm1.fpdtWk_yymm.Day	= strDay
	
	frm1.fpdtWk_yymmdd.Year	= strYear
	frm1.fpdtWk_yymmdd.Month= strMonth
	frm1.fpdtWk_yymmdd.Day	= strDay
	
	frm1.fpdtWk_yymm.focus
	
	Call ggoOper.FormatDate(frm1.fpdtWk_yymm, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.fpdtWk_yymmdd, Parent.gDateFormat, 1)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("I", "A", "NOCOOKIE", "MA" ) %>
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
' Desc : Make key stream of query or delete condition data
'========================================================================================================
Sub MakeKeyStream(pRow)
   Dim strYear,strMonth,strDay, strYYYYMM,strYYYYMM2, strYYYYMMDD
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call ExtractDateFrom(frm1.fpdtWk_yymm.text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM = strYear & strMonth
    
    Call ExtractDateFrom(frm1.fpdtWk_yymmdd.text,frm1.fpdtWk_yymmdd.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM2 = strYear & strMonth
	
'    Call ExtractDateFrom(frm1.fpdtWk_yymmdd.text,frm1.fpdtWk_yymmdd.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
'    strYYYYMMDD = strYear & strMonth & strDay
	 
	 lgKeyStream = strYYYYMM & Parent.gColSep       'You Must append one character(Parent.gColSep)
	 lgKeyStream = lgKeyStream & strYYYYMM2 & Parent.gColSep       'You Must append one character(Parent.gColSep)
'	 lgKeyStream = lgKeyStream & strYYYYMMDD & Parent.gColSep       'You Must append one character(Parent.gColSep)
	 lgKeyStream = lgKeyStream & UNIConvDate(Trim(frm1.fpdtWk_yymmdd.Text)) & Parent.gColSep      

   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()

End Sub

'========================================================================================================
Sub InitSpreadSheet()

Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021103",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	
	With frm1.vspdData
	
		.ReDraw = false
	
    	.MaxCols   = C_ExcCurDate + 1                                                  ' ☜:☜: Add 1 to Maxcols
	    .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
        .ColHidden = True           
       
		ggoSpread.Source= frm1.vspdData
		ggoSpread.ClearSpreadData

       call GetSpreadColumnPos("A")
       
     '  Call AppendNumberPlace("6","2","0")
       
		ggoSpread.SSSetEdit		C_DEALMON		,"거래통화"     ,40		,					,		,3		,2
		ggoSpread.SSSetButton	C_BTN
		ggoSpread.SSSetFloat C_DEALRATE,"평가환율",30,Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate		C_ExcCurDate	,"변동환율 적용기준일자"	,35		,2                  ,Parent.gDateFormat   ,-1														
		
       call ggoSpread.MakePairsColumn(C_DEALMON,C_BTN)
	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.Spreadlock      C_DEALMON		, -1, C_DEALMON     
      ggoSpread.Spreadlock      C_BTN			, -1, C_BTN   
      ggoSpread.SSSetRequired	C_DEALRATE		, -1, -1
      ggoSpread.Spreadlock      C_ExcCurDate	, -1, C_ExcCurDate 
      ggoSpread.SpreadLock	.vspdData.MaxCols, -1,.vspdData.MaxCols
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetRequired   C_DEALMON , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_DEALRATE , pvStartRow, pvEndRow
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

'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
            
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	Call SetToolbar("1100110100101111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
'	Call CookiePage (0)   

			
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
    Dim iRet
    Dim strYear,strMonth,strDay, strYYYYMM, strYYYYMMDD   
   

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
    
 	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call InitVariables                                                           '⊙: Initializes local global variables    
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call BtnDisabled(1)
    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
        
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
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
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
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
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()

    Dim IntRetCD, IntRetCD1, intRetCD2
    Dim DealMon, strYYYYMM
    Dim strYear, strMonth, strDay
    Dim lRow
    Dim lgCnt
    Dim RowFg
    
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
	Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM = strYear & strMonth
    
	With Frm1    
		lgCnt = 0
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0
			RowFg = .vspdData.Text
						
			.vspdData.Row = lRow
			.vspdData.Col = C_DEALMON
			
			DealMon = .vspdData.Text
			
			If RowFg <> "" Then
				Call CommonQueryRs("count(*)", "B_CURRENCY",  " CURRENCY = " & FilterVar(DealMon, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
				
				If lgF0 = 0 Then
					intRetCD2 = DisplayMsgBox("am0028", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"		
					Exit Function
				End If
				Call CommonQueryRs("count(*)", "A_EXCHANGE_RATE",  " yyyymm = " & FilterVar(strYYYYMM, "''", "S")  & " and doc_cur = " & FilterVar(DealMon, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
			
				lgCnt = lgCnt + lgF0				
			End If	
			
		Next
		
		If lgCnt <> 0 then
			IntRetCD1 = DisplayMsgBox("900007", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"		

			If IntRetCD1 = vbNo Then
				Exit Function
			End If
		End If
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                              '☜: Processing is OK
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
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 

'	With Frm1.VspdData
 '          .Col  = C_MAJORCD
'           .Row  = .ActiveRow
 '          .Text = ""
'    End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
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
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
	  Dim lRow, IntRetCD
	  Dim IntRetCD1,strYear,strMonth,strDay, strYYYYMM, strYYYYMMDD
    
    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error stat	
    
	FncInsertRow = False															'☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
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
        ggoSpread.InsertRow	.vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imrow -1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    Dim IntRetCD1,strYear,strMonth,strDay, strYYYYMM, strYYYYMMDD
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
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
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
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
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

    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

        if LayerShowHide(1) = false then
	    Exit Function
	end if

    
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag        
    End With
		
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
On Error Resume Next
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim strYear,strMonth,strDay, strYYYYMM, strYYYYMMDD
    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
        if LayerShowHide(1) = false then
	    Exit Function
	end if


    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM = strYear & strMonth
    
    'Call ExtractDateFrom(frm1.fpdtWk_yymmdd.Text,frm1.fpdtWk_yymmdd.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    'strYYYYMMDD = strYear & strMonth & strDay
        
    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0					
			
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Insert
													  strVal = strVal & "C" & Parent.gColSep
													  strVal = strVal & lRow & Parent.gColSep
													  strVal = strVal & strYYYYMM & Parent.gColSep
                    .vspdData.Col = C_DEALMON		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DEALRATE		: strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & Parent.gColSep
                    .vspdData.Col = C_ExcCurDate	: If Len(.vspdData.Text) > 0 Then
														Call ExtractDateFrom(.vspdData.Text,Parent.gDateFormat,Parent.gComDateType,strYear,strMonth,strDay)
														strYYYYMMDD = strYear & strMonth & strDay
													  End If	
													  strVal = strVal & Trim(strYYYYMMDD) & Parent.gRowSep  
													  
                    lGrpCnt = lGrpCnt + 1
					
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                                                  strVal = strVal & strYYYYMM & Parent.gColSep
                   .vspdData.Col = C_DEALMON 	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_DEALRATE	: strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & Parent.gColSep
                   .vspdData.Col = C_ExcCurDate	: strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep                                      
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & Parent.gColSep
                                                  strDel = strDel & lRow & Parent.gColSep
                                                  strDel = strDel & strYYYYMM & Parent.gColSep
				   .vspdData.Col = C_DEALMON    : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep									
                    lGrpCnt = lGrpCnt + 1

           End Select
       Next	
	
       .txtMode.value        = Parent.UID_M0002
      ' .txtUpdtUserId.value  = Parent.gUsrID
      ' .txtInsrtUserId.value = Parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal
		
	End With
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                               '☜: Processing is OK

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
    lgIntFlgMode = Parent.OPMD_UMODE    
	Call SetToolbar("1100111100111111")                                              '☆: Developer must customize
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables
    Call ggoOper.ClearField(Document, "2")
	Call SetToolbar("1100111100111111")
    MainQuery()
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
End Sub


'========================================================================================================
'   Event Name : fpdtWk_yymm_DblClick(Button)
'   Event Desc : 날짜 더블클릭시 발생 날짜 팝업 
'========================================================================================================

Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
		Call SetFocusToDocument("M")
		Frm1.fpdtWk_yymm.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : fpdtWk_yymmDD_DblClick(Button)
'   Event Desc : 날짜 더블클릭시 발생 날짜 팝업 
'========================================================================================================

Sub fpdtWk_yymmdd_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymmdd.Action = 7
		Call SetFocusToDocument("M")
		Frm1.fpdtWk_yymmdd.Focus
	End If
End Sub



'=======================================================================================================
'   Event Name : fpdtWk_yymm_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub fpdtWk_yymm_Keypress(Key)
    If Key = 13 Then
		frm1.fpdtWk_yymmdd.focus
        MainQuery()
    End If
End Sub

Sub fpdtWk_yymmdd_Keypress(Key)
    If Key = 13 Then
		frm1.fpdtWk_yymm.focus
        MainQuery()
    End If
End Sub




'=======================================================================================================
'   Event Name : ExeReflect(typel)
'   Event Desc : 고정환율, 변동환율버튼 에 따른 insert 
'=======================================================================================================
Function ExeReflect(typel)
	Dim IntRetCD,IntRetCD1 
    Dim lGrpCnt
    Dim strVal
    Dim strDel
    Dim lRow 
	Dim strYear,strMonth,strDay, strYYYYMM,strYYYYMM2, strYYYYMMDD
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	ExeReflect = False
  

   If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
		
    Call MakeKeyStream("X")
	
	
    Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM = strYear & strMonth   
    
    Call ExtractDateFrom(frm1.fpdtWk_yymmdd.Text,frm1.fpdtWk_yymmdd.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM2 = strYear & strMonth
    
'    Call ExtractDateFrom(frm1.fpdtWk_yymmdd.Text,frm1.fpdtWk_yymmdd.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
'    strYYYYMMDD = strYear & strMonth & strDay
   
   
	IF typel = 0 THEN
	
        Call CommonQueryRs("count(*)", "B_MONTHLY_EXCHANGE_RATE a, b_company b ",  " a.to_currency = b.loc_cur and  a.apprl_yrmnth = " & FilterVar(strYYYYMM2, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
          If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
          IntRetCD =  DisplayMsgBox("121602", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
			 If IntRetCD = vbNo Then
			     Exit Function
			 Else
				Call CommonQueryRs("count(*)", "A_EXCHANGE_RATE",  " yyyymm = " & FilterVar(strYYYYMM, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
				If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
					IntRetCD =  DisplayMsgBox("900007", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
					If IntRetCD = vbNo Then
						Exit Function
					End If
				End If	
			 End If	
		  Else
			IntRetCD =  DisplayMsgBox("121600", "X","X","X")			          '☜: "Will you destory previous data"
			Exit Function
       	  END IF
       		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0006                     '☜: Query  
    ELSE
       Call CommonQueryRs("count(*)", "B_DAILY_EXCHANGE_RATE a, b_company b ",  " a.to_currency = b.loc_cur and  a.APPRL_DT =  " & FilterVar(UNIConvDate(Trim(frm1.fpdtWk_yymmdd.Text)), "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
		IntRetCD =  DisplayMsgBox("121502", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
			If IntRetCD = vbNo Then
			    Exit Function
			Else
			  	Call CommonQueryRs("count(*)", "A_EXCHANGE_RATE",  " yyyymm = " & FilterVar(strYYYYMM, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
			  	If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
			  		IntRetCD =  DisplayMsgBox("900007", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
			  		If IntRetCD = vbNo Then
			  			Exit Function
			  		End If
			  	End If	
			 End If	
		 Else
			IntRetCD =  DisplayMsgBox("121500", "X","X","X")			          '☜: "Will you destory previous data"
			Exit Function
		 END IF
	      strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0005                     '☜: Query
    END IF    
        
	
	If LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data


	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

	ExeReflect = True                                                           '⊙: Processing is NG

End Function



'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Select Case Col				'추가부분을 위해..select로..
	    Case C_BTN              'Cost center
	        frm1.vspdData.Col = C_DEALMON
	        Call OpenCost(frm1.vspdData.Text, 1, Row)
			Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")
	End Select

End Sub

'===========================================================================
' Function Name : OpenCost
' Function Desc : OpenCode Reference Popup
'===========================================================================
Function OpenCost(strCode, iWhere, Row)

Dim IsOpenPop
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(0) = "통화 팝업"					<%' 팝업 명칭 %>
	    	arrParam(1) = " b_currency "					' TABLE 명칭 
	    	arrParam(2) = Trim(strCode)	                        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = ""							 		' Where Condition
	    	arrParam(5) = "통화"		   				    ' TextBox 명칭 

	    	arrField(0) = "currency "		                ' Field명(0)
	    	arrField(1) = "currency_desc"    						' Field명(1)

	    	arrHeader(0) = "통화코드"		        		' Header명(0)
	    	arrHeader(1) = "통화명"	      				' Header명(1)

	End Select

  
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet, iWhere, Row)
	End If

End Function

'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetCost()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCost(arrRet, iWhere, Row)


	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_DEALMON
		    	.vspdData.text = arrRet(0)		    	
		End Select

		lgBlnFlgChgValue = True
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row

	End With

End Function
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
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
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
    
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
'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
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

			C_DEALMON         =  iCurColumnPos(1)
			C_BTN         =  iCurColumnPos(2)
			C_DEALRATE  =  iCurColumnPos(3)
			C_ExcCurDate  =  iCurColumnPos(4)
		
    End Select    
   
End Sub


'******************************************  6.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

 '------------------------------------------ OpenLoanNo() -------------------------------------------------
'	Name : OpenPopupExc()
'	Description : Exchange reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupExc()

Dim arrRet
Dim arrParam(3)
Dim txtStdDt	
Dim txtStdYYMM
Dim iCalledAspName
Dim IntRetCD

	iCalledAspName = AskPRAspName("A5954RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5954RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
    
    txtStdYYMM = frm1.fpdtWk_yymm.Text
	txtStdDt = frm1.fpdtWk_yymmdd.Text
	
	arrRet = window.showModalDialog(iCalledAspName & "?txtStdYYMM=" & txtStdYYMM & "&txtStdDt=" & txtStdDt , Array(window.parent, arrParam), _
		     "dialogWidth=480px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	

	'arrRet = window.showModalDialog("a5954ra1.asp?txtStdYYMM=" & txtStdYYMM & "&txtStdDt=" & txtStdDt , Array(window.parent, arrParam), _
'		     "dialogWidth=480px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	
	If arrRet(0,0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpenExc(arrRet)
	
	End If
			
End Function

'------------------------------------------  SetRefOpenExc()  ---------------------------------------------
'	Name : SetRefOpenExc()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRefOpenExc(Byval arrRet)
	
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	
	TempRow = 0
	I = 0
	
	
	With frm1
	
		.vspddata.focus		
		ggoSpread.Source = .vspddata
		.vspddata.ReDraw = False	
	
		
		'.vspddata.MaxRows = .vspddata.MaxRows + (Ubound(arrRet, 1) + 1)			'☜: Reference Popup에서 선택되어진 Row만큼 추가		
		TempRow = .vspddata.MaxRows												'☜: 현재까지의 MaxRows
		
	'	msgbox ubound(arrRet,1)
	'	msgbox ubound(arrRet,2)
		For I = TempRow to TempRow + Ubound(arrRet, 1)
			
			.vspddata.MaxRows = .vspddata.MaxRows + 1
			.vspddata.Row = I + 1				
			.vspddata.Col = 0
			.vspddata.Text = ggoSpread.InsertFlag
			'FOR j = 0 to  C_BTN - 1
			'	.vspddata.Col = j + 1												'⊙: 첫번째 컬럼 
			'	.vspddata.text = arrRet(I - TempRow, j)				
			'Next
			.vspddata.Col = C_DEALMON 
			.vspddata.text = arrRet(I - TempRow, 0)
'			.vspddata.Col = C_BTN
'			.vspddata.text = arrRet(I - TempRow, 1)
			.vspddata.Col = C_DEALRATE
			.vspddata.text = arrRet(I - TempRow, 1)
								
			If Len(arrRet(I - TempRow, 2)) <> 6 Then
				.vspddata.Col = C_ExcCurDate
				.vspddata.text = arrRet(I - TempRow, 2)
			End If
			
		Next	
		
		
    End With
  
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월차평가환율</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupExc()">환율참조</A>&nbsp</TD>					
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
									<TD CLASS="TD5">결산년월</TD>
									<TD CLASS="TD6">										
										<script language =javascript src='./js/a5954ma1_fpDateTime3_fpdtWk_yymm.js'></script>
									</TD>
									<TD CLASS="TD5">결산일자</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/a5954ma1_fpDateTime3_fpdtWk_yymmdd.js'></script>
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
									<script language =javascript src='./js/a5954ma1_OBJECT1_vspdData.js'></script>
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
			
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT= <%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPayCd"     tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

