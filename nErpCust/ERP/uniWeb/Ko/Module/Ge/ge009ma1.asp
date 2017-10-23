
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        : 
*  3. Program ID           : GE009MA1
*  4. Program Name         : 품목그룹별 손익비교 
*  5. Program Desc         : 품목그룹별 손익비교 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/03/08
*  8. Modified date(Last)  : 2002/03/08
*  9. Modifier (First)     : Yoon Suck Kyu
* 10. Modifier (Last)      : Yoon Suck Kyu
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID		= "ge009mb1.asp"                                      'Biz Logic ASP 
Const BIZ_PGM_JUMP_ID	= "GE007MA1"                                      'Biz Logic ASP 

Dim C_ITEM_GRP	
Dim C_ITEM_GRPNM	
Dim C_SALE_AMT	
Dim C_SALE_COST	
Dim C_SALE_PROFIT 
Dim C_TOTAL_COST	
Dim C_CUR_PROFIT	
Dim C_NET_PROFIT	


'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop          

'========================================================================================================
Sub InitVariables()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    lgStrPrevKey      = ""										'⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""										'⊙: initializes Previous Key Index
    lgSortKey         = 1										'⊙: initializes sort direction

    lgIntFlgMode      = Parent.OPMD_CMODE								'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed    		
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub SetDefaultVal()
   Dim strYear ,strMonth ,strDay
   Dim BaseDate
   Dim FromDateOfDB
   
   BaseDate     = "<%=GetSvrDate%>"                                                          'Get DB Server Date    
   FromDateOfDB  = UniConvDateAToB(BaseDate ,Parent.gServerDateFormat,Parent.gDateFormat)               'Convert DB date type to Company

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 	
	
	frm1.txtFrYYYYMM.focus
	
	Call ggoOper.FormatDate(frm1.txtFrYYYYMM, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtToYYYYMM, Parent.gDateFormat, 2)
	
	Call ExtractDateFrom(FromDateOfDB, Parent.gDateFormat ,Parent.gComDateType ,strYear ,strMonth ,strDay) 
	
	frm1.txtFrYYYYMM.Year	= strYear
	frm1.txtFrYYYYMM.Month	= strMonth
	frm1.txtFrYYYYMM.Day	= strDay
	
	frm1.txtToYYYYMM.Year	= strYear
	frm1.txtToYYYYMM.Month	= strMonth
	frm1.txtToYYYYMM.Day	= strDay
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "G", "NOCOOKIE", "QA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
Dim iRow,iCol
Dim strYear,strMonth,strDay
Dim TempFrDt,TempToDt
   '------ Developer Coding part (Start ) --------------------------------------------------------------       
	With Frm1                      	
		Select Case Kubun		
		Case 0
			
			If ReadCookie("JumpFlag")	<>""	Then .txtJumpFlag.Value		= ReadCookie("JumpFlag")
			
			If UCase(Trim(.txtJumpFlag.Value)) = "GE007MA1" Then
				If ReadCookie("FrYYYYMM")	<>"" Then 
					TempFrDt				= ReadCookie("FrYYYYMM")				
					Call ExtractDateFrom(TempFrDt, Parent.gDateFormat, Parent.gComDateType, strYear, strMonth, strDay)
					.txtFrYYYYMM.Year	= strYear
					.txtFrYYYYMM.Month	= strMonth
					.txtFrYYYYMM.Day	= strDay
				End If
				
				If ReadCookie("ToYYYYMM")	<>"" Then 
					TempToDt				= ReadCookie("ToYYYYMM")
					Call ExtractDateFrom(TempToDt, Parent.gDateFormat, Parent.gComDateType, strYear, strMonth, strDay)
					.txtToYYYYMM.Year	= strYear
					.txtToYYYYMM.Month	= strMonth
					.txtToYYYYMM.Day	= strDay
				End If
			
				WriteCookie "FrYYYYMM"		, ""
				WriteCookie "ToYYYYMM"		, ""
				
				If Trim(.txtFrYYYYMM.Text) <> "" and Trim(.txtToYYYYMM.Text) <> ""  Then
					Call MainQuery()
      			End If
      		End If
      		
      		WriteCookie "JumpFlag"		, ""
      		
		Case 1			
		
		    WriteCookie "FrYYYYMM" , UniConvYYYYMMDDToDate(Parent.gDateFormat,Trim(.txtFrYYYYMM.Year),Right("0" & Trim(.txtFrYYYYMM.Month),2),"01")
		    WriteCookie "ToYYYYMM" , UniConvYYYYMMDDToDate(Parent.gDateFormat,Trim(.txtToYYYYMM.Year),Right("0" & Trim(.txtToYYYYMM.Month),2),"01")
		    
		    If .vspdData.MaxRows > 0 Then
				iRow = .vspdData.ActiveRow			
		    
				.vspdData.Row = iRow
				.vspdData.Col = C_ITEM_GRP			
				WriteCookie "ItemGrp" , .vspdData.Text		    
			End If
			
			WriteCookie "JumpFlag" , "GE009MA1"
		    		    
		End Select
	End With
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Function PgmJumpCheck()         
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
	   
    Call PgmJump(BIZ_PGM_JUMP_ID)
	    
End Function	  

'========================================================================================================
Sub MakeKeyStream()
 
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   With Frm1
		lgKeyStream = .txtFrYYYYMM.Year & Right("0" & .txtFrYYYYMM.Month,2)					& Parent.gColSep       'You Must append one character(Parent.gColSep)
		lgKeyStream = lgKeyStream & .txtToYYYYMM.Year & Right("0" & .txtToYYYYMM.Month,2)	& Parent.gColSep
		lgKeyStream = lgKeyStream & .txtBizUnitcd.value										& Parent.gColSep
		lgKeyStream = lgKeyStream & .txtCostcd.value										& Parent.gColSep
		lgKeyStream = lgKeyStream & .txtSalesOrg.value										& Parent.gColSep
		lgKeyStream = lgKeyStream & .txtSalesGrp.value										& Parent.gColSep
		lgKeyStream = lgKeyStream & .cboSort.value											& Parent.gColSep
	End With
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'========================================================================================================
Sub InitComboBox()
Dim strName, strCode   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	strName = "품목그룹코드" & Chr(11) & "품목그룹명" & Chr(11) & "매출액" & Chr(11) & "매출원가" & Chr(11)	
	strName = strName & "매출이익" & Chr(11) & "총원가" & Chr(11) & "경상이익" & Chr(11) & "순이익" & Chr(11)
	
	strCode = C_ITEM_GRP & Chr(11) & C_ITEM_GRPNM & Chr(11) & C_SALE_AMT & Chr(11) & C_SALE_COST & Chr(11)	
	strCode = strCode & C_SALE_PROFIT & Chr(11) & C_TOTAL_COST & Chr(11) & C_CUR_PROFIT & Chr(11) & C_NET_PROFIT & Chr(11)
	
    Call SetCombo2(frm1.cboSort ,strCode ,strName ,Chr(11))    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub InitData()

End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	C_ITEM_GRP			= 1	
	C_ITEM_GRPNM		= 2	
	C_SALE_AMT			= 3	
	C_SALE_COST			= 4	
	C_SALE_PROFIT		= 5 
	C_TOTAL_COST		= 6	
	C_CUR_PROFIT		= 7	
	C_NET_PROFIT		= 8	
End Sub

'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	With frm1.vspdData
	
       .MaxCols   = C_NET_PROFIT + 1                                                  ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True

        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021218", ,parent.gAllowDragDropSpread
		
		ggoSpread.ClearSpreadData                                            ' ☜: Clear spreadsheet data

	   .ReDraw = false
	
		Call GetSpreadColumnPos("A")	
      
								'ColumnPosition		Header			Width	Align(0:L,1:R,2:C)	Row						Length					CharCase(0:L,1:N,2:U)
	   ggoSpread.SSSetEdit		C_ITEM_GRP		,"품목그룹코드"	,12     ,					,						,10						,2
       ggoSpread.SSSetEdit		C_ITEM_GRPNM	,"품목그룹명"	,18     ,					,						,40						,2
       
								'ColumnPosition     Header			Width	Grp					IntegeralPart			DeciPointpart			Align			Sep				PZ   Min       Max 
       ggoSpread.SSSetFloat		C_SALE_AMT		,"매출액"		,14     ,Parent.ggAmtOfMoneyNo		,ggStrIntegeralPart		,ggStrDeciPointPart		,Parent.gComNum1000	,Parent.gComNumDec
       ggoSpread.SSSetFloat		C_SALE_COST		,"매출원가"		,14     ,Parent.ggAmtOfMoneyNo		,ggStrIntegeralPart		,ggStrDeciPointPart		,Parent.gComNum1000	,Parent.gComNumDec
       ggoSpread.SSSetFloat		C_SALE_PROFIT	,"매출이익"		,14     ,Parent.ggAmtOfMoneyNo		,ggStrIntegeralPart		,ggStrDeciPointPart		,Parent.gComNum1000	,Parent.gComNumDec
       ggoSpread.SSSetFloat		C_TOTAL_COST	,"총원가"		,14     ,Parent.ggAmtOfMoneyNo		,ggStrIntegeralPart		,ggStrDeciPointPart		,Parent.gComNum1000	,Parent.gComNumDec
       ggoSpread.SSSetFloat		C_CUR_PROFIT	,"경상이익"		,14     ,Parent.ggAmtOfMoneyNo		,ggStrIntegeralPart		,ggStrDeciPointPart		,Parent.gComNum1000	,Parent.gComNumDec
       ggoSpread.SSSetFloat		C_NET_PROFIT	,"순이익"		,14     ,Parent.ggAmtOfMoneyNo		,ggStrIntegeralPart		,ggStrDeciPointPart		,Parent.gComNum1000	,Parent.gComNumDec
       
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

			C_ITEM_GRP                    	= iCurColumnPos(1)
			C_ITEM_GRPNM               		= iCurColumnPos(2)    
			C_SALE_AMT						= iCurColumnPos(3)
			C_SALE_COST						= iCurColumnPos(4)
			C_SALE_PROFIT					= iCurColumnPos(5)
			C_TOTAL_COST                    = iCurColumnPos(6)
			C_CUR_PROFIT                    = iCurColumnPos(7)
			C_NET_PROFIT                    = iCurColumnPos(8)
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
'	Call initMinor()
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) --------------------------------------------------------------                          
        
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
            
	Call InitVariables()	
    Call SetDefaultVal

    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call SetToolbar("1100000000001111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
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

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    Call ggoOper.ClearField(Document, "3")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If CompareDateByFormat(frm1.txtFrYYYYMM.Text,frm1.txtToYYYYMM.Text,frm1.txtFrYYYYMM.Alt,frm1.txtToYYYYMM.Alt,"970023",frm1.txtFrYYYYMM.UserDefinedFormat,Parent.gComDateType,True) = False Then
        frm1.txtFrYYYYMM.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End If
    Call InitVariables()                                                           '⊙: Initializes local global variables    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If DbQuery() = False Then                                                      '☜: Query db data
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
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                             '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                            '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")                         '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbDelete = False Then                                                      '☜: Query db data
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    '

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncCancel() 
    Dim iDx
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
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
    End If   
    
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
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")	                 '⊙: Data is changed.  Do you want to exit? 
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
    On Error Resume Next
    
    DbQuery = False                                                              '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    Call MakeKeyStream()

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
           strVal = BIZ_PGM_ID  & "?txtMode="            & Parent.UID_M0001						         
           strVal = strVal      & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
           strVal = strVal      & "&txtMaxRows="         & .vspdData.MaxRows
           strVal = strVal      & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
'           strVal = strVal      & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)    '☜: Max fetched data at a time
    End With
    '--------- Developer Coding Part (End) ------------------------------------------------------------
   
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbSave()

    On Error Resume Next
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG

	Call LayerShowHide(1)
		
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID1)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	Call LayerShowHide(1)

    Call MakeKeyStream()
		
    strVal = BIZ_PGM_ID1 & "?txtMode="          & Parent.UID_M0003                       '☜: Query
    strVal = strVal      & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function

'========================================================================================================
Sub DbQueryOk()
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Frm1.vspdData.Focus                                    
'	Call ggoOper.LockField(Document, "Q")	    
	If Frm1.vspdData.MaxRows > 0 Then Call SetToolbar("1100000000011111")                        '☆: Developer must customize
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
Sub DbSaveOk()

    Call InitVariables()												     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "3")				                           '⊙: Clear Contents  Field
End Sub

'========================================================================================================
Function OpenCd(Kubun)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)		
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case Trim(UCase(Kubun))
		Case "BIZ"
			arrParam(0) = "사업부"											' Popup Name
			arrParam(1) = RetPopupTable(Kubun)									' Table Name
			arrParam(2) = Trim(frm1.txtBizUnitcd.Value)							' Code Condition
			arrParam(3) = ""													' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "사업부"
				
			arrField(0) = " Z.BIZ_UNIT_CD "								' Field명(0)
			arrField(1) = " Z.BIZ_UNIT_NM "								' Field명(1)
			    
			arrHeader(0) = "사업부코드"										' Header명(0)
			arrHeader(1) = "사업부명"										' Header명(1)
		Case "COST"
			arrParam(0) = "Profit Center"									' Popup Name
			arrParam(1) = RetPopupTable(Kubun)									' Table Name
			arrParam(2) = Trim(frm1.txtCostcd.Value)							' Code Condition
			arrParam(3) = ""													' Name Cindition			
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "Profit Center"
				
			arrField(0) = " Z.COST_CD "								' Field명(0)
			arrField(1) = " Z.COST_NM "								' Field명(1)
			    
			arrHeader(0) = "Profit Center코드"								' Header명(0)
			arrHeader(1) = "Profit Center명"								' Header명(1)			 
		Case "SALESORG"
			arrParam(0) = "영업조직"										' Popup Name
			arrParam(1) = RetPopupTable(Kubun)									' Table Name
			arrParam(2) = Trim(frm1.txtSalesOrg.Value)							' Code Condition
			arrParam(3) = ""													' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "영업조직"
				
			arrField(0) = " Z.SALES_ORG "								' Field명(0)
			arrField(1) = " Z.SALES_ORG_NM "								' Field명(1)
			    
			arrHeader(0) = "영업조직코드"									' Header명(0)
			arrHeader(1) = "영업조직명"										' Header명(1)
		Case "SALESGRP"
			arrParam(0) = "영업그룹"										' Popup Name
			arrParam(1) = RetPopupTable(Kubun)									' Table Name
			arrParam(2) = Trim(frm1.txtSalesGrp.Value)							' Code Condition
			arrParam(3) = ""													' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "영업그룹"
				
			arrField(0) = " Z.SALES_GRP "								' Field명(0)
			arrField(1) = " Z.SALES_GRP_NM "								' Field명(1)
			    
			arrHeader(0) = "영업그룹코드"									' Header명(0)
			arrHeader(1) = "영업그룹명"										' Header명(1)
		
	End Select
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	  Select Case Trim(UCase(Kubun))
		Case "BIZ"
			frm1.txtBizUnitcd.focus
		Case "COST"
			frm1.txtCostcd.focus
		Case "SALESORG"
			frm1.txtSalesOrg.focus
		Case "SALESGRP"
			frm1.txtSalesGrp.focus
		Case Else
		
	  End Select				
		Exit Function
	Else
		Call SubSetRet(arrRet,Kubun)
	End If	
	
End Function

'=======================================================================================================
Sub SubSetRet(arrRet,Kubun)
	With Frm1
		Select Case Trim(UCase(Kubun))
			Case "BIZ"
				.txtBizUnitcd.value = arrRet(0)
				.txtBizUnitnm.value = arrRet(1)
				.txtBizUnitcd.focus 
			Case "COST"
				.txtCostcd.value = arrRet(0)
				.txtCostnm.value = arrRet(1)
				.txtCostcd.focus 
			Case "SALESORG"
				.txtSalesOrg.value = arrRet(0)
				.txtSalesOrgnm.value = arrRet(1)
				.txtSalesOrg.focus 
			Case "SALESGRP"
				.txtSalesGrp.value = arrRet(0)
				.txtSalesGrpnm.value = arrRet(1)
				.txtSalesGrp.focus 
		End Select
	End With
End Sub

'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col

			End Select
		End If
    
	End With
	Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")
End Sub

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
Sub vspdData_Click(Col, Row)
    Dim IntRetCD
    
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
    Else
    	If frm1.vspdData.MaxRows = 0 Then                                      'If there is no data.
    	   Exit Sub
    	End If
    	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
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
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
           If DbQuery() = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub

'=======================================================================================================
Sub txtFrYYYYMM_DblClick(Button)
    If Button = 1 Then
       frm1.txtFrYYYYMM.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtFrYYYYMM.Focus
    End If
'    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtToYYYYMM_DblClick(Button)
    If Button = 1 Then
       frm1.txtToYYYYMM.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtToYYYYMM.Focus
    End If
'    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtFrYYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
'    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtToYYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
'    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub cboSort_onChange()
	IF frm1.vspdData.Maxrows > 0 Then
		ggoSpread.Source = frm1.vspdData
		If UNICDbl(Trim(frm1.cboSort.value)) = C_ITEM_GRP Or UNICDbl(Trim(frm1.cboSort.value)) = C_ITEM_GRPNM Then
			ggoSpread.SSSort UNICDbl(Trim(frm1.cboSort.value)), 1 '1:오름차순 ASC, 2:내림차순 DESC
		Else
			ggoSpread.SSSort UNICDbl(Trim(frm1.cboSort.value)), 2 '1:오름차순 ASC, 2:내림차순 DESC
		End If
	End If
'	lgBlnFlgChgValue = True	
End Sub

'========================================================================================================
Function RetPopupTable(Kubun)

Dim TempTable, TempCon
Dim StrField, StrTable, StrCon
Dim BIZFlag, COSTFlag, ORGFlag, GRPFlag

	TempTable = ""	: TempCon = ""		: StrField = ""		: StrTable = ""		: StrCon = ""
	BIZFlag = False : COSTFlag = False	: ORGFlag = False	: GRPFlag = False

	With Frm1
		
		StrTable = " FROM B_BIZ_UNIT A, B_COST_CENTER B, B_SALES_ORG C, B_SALES_GRP D "
		StrCon = " WHERE A.BIZ_UNIT_CD = B.BIZ_UNIT_CD AND B.COST_CD = D.COST_CD AND D.SALES_ORG = C.SALES_ORG AND B.COST_TYPE=" & FilterVar("S", "''", "S") & "  "
		
		If Trim(.txtBizUnitcd.value) <> "" And Trim(UCase(Kubun)) <> "BIZ" Then
			TempCon = " AND A.BIZ_UNIT_CD =  " & FilterVar(.txtBizUnitcd.value, "''", "S") & " "
			BIZFlag = True
		End If
		If Trim(.txtCostcd.value) <> "" And Trim(UCase(Kubun)) <> "COST" Then
			TempCon = TempCon & " AND B.COST_CD =  " & FilterVar(.txtCostcd.value, "''", "S") & " "
			COSTFlag = True
		End If
		If Trim(.txtSalesOrg.value) <> "" And Trim(UCase(Kubun)) <> "SALESORG" Then
			TempCon = TempCon & " AND C.SALES_ORG =  " & FilterVar(.txtSalesOrg.value, "''", "S") & " "
			ORGFlag = True
		End If
		If Trim(.txtSalesGrp.value) <> "" And Trim(UCase(Kubun)) <> "SALESGRP" Then
			TempCon = TempCon & " AND D.SALES_GRP =  " & FilterVar(.txtSalesGrp.value, "''", "S") & " "
			GRPFlag = True
		End If
		
		StrCon = StrCon & TempCon
		
		Select Case Trim(UCase(Kubun))
			Case "BIZ"
				StrField = " SELECT DISTINCT A.BIZ_UNIT_CD,A.BIZ_UNIT_NM "
				If Trim(TempCon) = "" Then				
					StrTable = " FROM B_BIZ_UNIT A "
					StrCon = ""
				End If
			Case "COST"
				StrField = " SELECT DISTINCT B.COST_CD,B.COST_NM "
				If Trim(TempCon) = "" Then				
					StrTable = " FROM B_COST_CENTER B "
					StrCon = ""
				End If
			Case "SALESORG"
				StrField = " SELECT DISTINCT C.SALES_ORG,C.SALES_ORG_NM "
				If Trim(TempCon) = "" Then
					StrTable = " FROM B_SALES_ORG C "
					StrCon = ""
				End If
			Case "SALESGRP"
				StrField = " SELECT DISTINCT D.SALES_GRP,D.SALES_GRP_NM "
				If Trim(TempCon) = "" Then
					StrTable = " FROM B_SALES_GRP D "
					StrCon = ""
				End If			
		End Select
	End With
	
	RetPopupTable = " (" & StrField & StrTable & StrCon & ") Z "
	
End Function
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목그룹별손익비교</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
									<TD CLASS=TD5 NOWRAP>작업년월</TD>
                                    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/ge009ma1_txtFrYYYYMM_txtFrYYYYMM.js'></script> ~
                                                         <script language =javascript src='./js/ge009ma1_txtToYYYYMM_txtToYYYYMM.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>사업부</TD>
                                    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBizUnitcd"  SIZE=10 MAXLENGTH=10  TAG="11xxxU" ALT="사업부"><IMG SRC="../../../CShared/image/btnPopup.gif" ALIGN=top TYPE="BUTTON" OnClick="vbscript:Call OpenCd('BIZ')">
                                                         <INPUT TYPE=TEXT NAME="txtBizUnitnm"  SIZE=20 MAXLENGTH=50  TAG="14X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>Profit Center</TD>
                                    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCostcd"  SIZE=10 MAXLENGTH=10  TAG="11xxxU" ALT="Profit Center"><IMG SRC="../../../CShared/image/btnPopup.gif" ALIGN=top TYPE="BUTTON" OnClick="vbscript:Call OpenCd('COST')">
                                                         <INPUT TYPE=TEXT NAME="txtCostnm"  SIZE=20 MAXLENGTH=50  TAG="14X"></TD>
									<TD CLASS=TD5 NOWRAP>영업조직</TD>
                                    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesOrg"  SIZE=10 MAXLENGTH=4  TAG="11xxxU" ALT="영업조직"><IMG SRC="../../../CShared/image/btnPopup.gif" ALIGN=top TYPE="BUTTON" OnClick="vbscript:Call OpenCd('SALESORG')">
                                                         <INPUT TYPE=TEXT NAME="txtSalesOrgnm"  SIZE=20 MAXLENGTH=50  TAG="14X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
                                    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp"  SIZE=10 MAXLENGTH=4  TAG="11xxxU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" ALIGN=top TYPE="BUTTON" OnClick="vbscript:Call OpenCd('SALESGRP')">
                                                         <INPUT TYPE=TEXT NAME="txtSalesGrpnm"  SIZE=20 MAXLENGTH=50  TAG="14X"></TD>
									<TD CLASS=TD5 NOWRAP>Sort</TD>
                                    <TD CLASS=TD6 NOWRAP><SELECT NAME="cboSort"  CLASS=cboNormal TAG="11" ALT="Sort"><OPTION VALUE=""></OPTION></SELECT></TD>
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
								<TD HEIGHT="80%" COLSPAN=4>
									<script language =javascript src='./js/ge009ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
							<TR>								
								<TD CLASS=TD5 NOWRAP>매출액합계</TD>
								<TD CLASS=TD6 ><script language =javascript src='./js/ge009ma1_totSalesAmt_totSalesAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>매출원가합계</TD>
								<TD CLASS=TD6 ><script language =javascript src='./js/ge009ma1_totCostAmt_totCostAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>매출이익합계</TD>
								<TD CLASS=TD6 ><script language =javascript src='./js/ge009ma1_totPorfitAmt_totPorfitAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>총원가합계</TD>
								<TD CLASS=TD6 ><script language =javascript src='./js/ge009ma1_totTotCostAmt_totTotCostAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>경상이익합계</TD>
								<TD CLASS=TD6 ><script language =javascript src='./js/ge009ma1_totCurProfitAmt_totCurProfitAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>순이익합계</TD>
								<TD CLASS=TD6 ><script language =javascript src='./js/ge009ma1_totNetProfitAmt_totNetProfitAmt.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
				    <TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:PgmJumpCheck()" ONCLICK="VBSCRIPT:Call CookiePage(1)">품목그룹별 손익현황</a></TD>
					<TD WIDTH=10>&nbsp;</TD>

				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"   TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"        TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"           TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"        TAG="X4" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtJumpFlag"		TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

