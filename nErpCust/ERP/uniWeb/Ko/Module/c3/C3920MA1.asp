
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : COSTING
*  2. Function Name        : Actual Cost Reflection
*  3. Program ID           : C3920MA1.asp
*  4. Program Name         : Actual Cost Reflection
*  5. Program Desc         : Actual Cost Reflection BIZ Logic
*  6. Comproxy List        : +
*  7. Modified date(First) : 2002/09/25
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Lee Tae Soo 
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================
-->
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
Const BIZ_PGM_ID = "C3920MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'@Grid_Column
Dim C_ChkFlag 
Dim C_ItemCd 										'Spread Sheet의 Column별 상수 
Dim C_ItemNm 
Dim C_Basicunit 
Dim C_ItemSpec 
Dim C_RcptPrc 
Dim C_TotAvgPrc 
Dim C_StockStdPrc 


'Const C_SHEETMAXROWS	= 100	                                      '☜: Visble row


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop 


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
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call SetToolbar("11000000000000")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim StartDate
	
	StartDate	= "<%=GetSvrDate%>"
	
	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)

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

	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "BA" ) %>

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
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
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
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	 C_ChkFlag		= 1
	 C_ItemCd		= 2										'Spread Sheet의 Column별 상수 
	 C_ItemNm		= 3
	 C_Basicunit	= 4
	 C_ItemSpec		= 5
	 C_RcptPrc		= 6
	 C_TotAvgPrc	= 7
	 C_StockStdPrc	= 8
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	With frm1.vspdData
	
       .MaxCols = C_StockStdPrc + 1                                                 ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
    
      
        ggoSpread.Source = Frm1.vspdData
        ggoSpread.Spreadinit "V20030102",,parent.gAllowDragDropSpread  
        
        Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	   .ReDraw = false
	   
	   Call GetSpreadColumnPos("A")

 		ggoSpread.SSSetCheck C_ChkFlag, "실행구분", 10, ,"",true    
		ggoSpread.SSSetEdit C_ItemCd, "품목코드",15, 0
		ggoSpread.SSSetEdit C_ItemNm, "품목명", 33, 0
		ggoSpread.SSSetEdit C_BasicUnit, "기준단위",10,0
		ggoSpread.SSSetEdit	C_ItemSpec, "품목규격",20,0
		ggoSpread.SSSetFloat C_RcptPrc,"입고단가",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_TotAvgPrc,"평가단가",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_StockStdPrc,"재고표준단가",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

		if frm1.rdoRcpt.value = "Y" then
			Call ggoSpread.SSSetColHidden(C_TotAvgPrc ,C_TotAvgPrc	,True)	
			frm1.hRadio.value = "RCPT"
		Else
			Call ggoSpread.SSSetColHidden(C_RcptPrc ,C_RcptPrc	,True)
			frm1.hRadio.value = "TOTAVG"
		end if 

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
		ggoSpread.SpreadLock C_ItemCd, -1, C_ItemCd
		ggoSpread.SpreadLock C_ItemNm, -1, C_ItemNm
		ggoSpread.SpreadLock C_BasicUnit , -1, C_BasicUnit
		ggoSpread.SpreadLock C_ItemSpec , -1, C_ItemSpec      
		ggoSpread.SpreadLock C_RcptPrc , -1, C_RcptPrc
		ggoSpread.SpreadLock C_TotAvgPrc , -1, C_TotAvgPrc
		ggoSpread.SpreadLock C_StockStdPrc , -1, C_StockStdPrc  
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_ItemCd		,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_ItemNm		,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_BasicUnit	,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_ItemSpec		,pvStartRow	,pvEndRow      
		ggoSpread.SSSetProtected C_RcptPrc		,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_TotAvgPrc	,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_StockStdPrc	,pvStartRow	,pvEndRow
    .vspdData.ReDraw = True
    
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
			C_ChkFlag			= iCurColumnPos(1)
			C_ItemCd		    = iCurColumnPos(2)
			C_ItemNm		    = iCurColumnPos(3)    
			C_Basicunit		    = iCurColumnPos(4)
			C_ItemSpec		    = iCurColumnPos(5)
			C_RcptPrc			= iCurColumnPos(6)
			C_TotAvgPrc		    = iCurColumnPos(7)
			C_StockStdPrc		= iCurColumnPos(8)
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
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
            
    Call InitSpreadSheet                                                             'Setup the Spread sheet
	Call SetDefaultVal
	Call InitVariables
	Call SetToolbar("11000000000000")                                              '☆: Developer must customize
	
	frm1.txtYYYYMM.focus
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
	Call CookiePage (0) 
	Set gActiveElement = document.activeElement	                                                             '☜: Check Cookie
			
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
    
    if frm1.rdoRcpt.checked = True Then
		frm1.hRadio.value = "RCPT"
    else
    	frm1.hRadio.value = "TOTAVG"
    end if
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
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
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
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
	Call ggoSpread.ReOrderingSpreadData()
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
	Dim strYear,strMonth,strDay
	
    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG
	
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF                                                       '☜: Show Processing Message
    
    With Frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
     '@Query_Hidden     
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'Hidden의 검색조건으로 Query
			strVal = strVal & "&txtYYYYMM=" & .hYYYYMM.value
			strVal = strVal & "&txtPlantCd=" &  .hPlantCd.value				
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtItemAccntCd=" & .hItemAccntCd.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
'			strVal = strVal & "&lgMaxCount=" & C_SHEETMAXROWS
			strVal = strVal & "&txtItemCd=" & .hItemCd.value
			strVal = strVal & "&txtFlag=" & .hRadio.value
		Else
      '@Query_Text     
			Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'현재 검색조건으로 Query
			strVal = strVal & "&txtYYYYMM=" & strYear & strMonth
			strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value				
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtItemAccntCd=" & .txtItemAccntCd.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
'			strVal = strVal & "&lgMaxCount=" & C_SHEETMAXROWS
			strVal = strVal & "&txtItemCd=" & .txtItemCd.value
			strVal = strVal & "&txtFlag=" & .hRadio.value
			
			.hYYYYMM.value = strYear & strMonth
			
		END IF	    
    End With
   
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
	Call SetToolbar("11000000000111")
    DbQuery = True                                                               '☜: Processing is OK
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
Function DbQueryOk()

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
                                '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
 	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
   
End Function
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables	   

    '------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadUnLock C_ChkFlag, -1, C_ChkFlag	
	
	frm1.vspdData.MaxRows = 0
    
	Call Dbquery()
   '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    On Error Resume Next
    
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
    
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
  

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
      	   DbQuery
    	End If
    End if
End Sub

'=======================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtYyyymm_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYyyymm.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
    End If
End Sub


Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub
'======================================================================================================
'	Name : OpenPlant()
'	Description : Plant PopUp
'=======================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function

'======================================================================================================
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetPlant(byval arrRet)
	
	frm1.txtPlantCd.Value = arrRet(0)
	frm1.txtPlantNM.value = arrRet(1)
End Function

'===========================================================================
' Function Name : OpenItemAccnt()
' Function Desc : OpenItemAccnt(품목계정) Reference Popup
'===========================================================================
Function OpenItemAccnt()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function

	
	lgIsOpenPop = True
	
	arrParam(0) = "품목계정팝업"				' 팝업 명칭 
	arrParam(1) = "B_MINOR a,b_item_acct_inf b"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtItemAccntCd.value)		' Code Condition
	arrParam(3) = ""	' Name Cindition
	arrParam(4) = "MAJOR_CD=" & FilterVar("P1001", "''", "S")  & "  and A.MINOR_CD = B.ITEM_ACCT AND B.ITEM_ACCT_GROUP <> " & FilterVar("6MRO","''","S") 
	arrParam(5) = "품목계정"						' TextBox 명칭 
		
    arrField(0) = "MINOR_CD"						' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)
    
    arrHeader(0) = "품목계정코드"					' Header명(0)
    arrHeader(1) = "품목계정명"						' Header명(1)

   
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemAccntCd.focus
		Exit Function
	Else
		Call SetItemAccnt(arrRet)
	End If	
	
End Function

'------------------------------------------  SetMinor()  --------------------------------------------------
'	Name : SetItemAccnt()
'	Description : Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemAccnt(Byval arrRet)

If arrRet(0) <> "" Then 
	frm1.txtItemAccntCd.focus								' 업태 
	frm1.txtItemAccntCd.value = arrRet(0)
	frm1.txtItemAccntNm.value = arrRet(1)

End If

End Function

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item Code PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("189220","x","x","x") '공장을 먼저 입력하세요 
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	If Trim(frm1.txtItemAccntCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("990003","x","x","x") '품목구분을 먼저 입력하세요 
		frm1.txtPlantCd.focus
		Exit Function
	End If

	lgIsOpenPop = True
	
	' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 

	If Trim(frm1.txtItemAccntCd.value) <> "" Then
		arrParam(2) = Mid(CStr(Trim(frm1.txtItemAccntCd.value)),1,1) & Mid(CStr(Trim(frm1.txtItemAccntCd.value)),1,1)
	ELSE
		arrParam(2) = "15"
	END IF
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(3) = ""							' Default Value
	

	arrField(0) = 1 								' Field명(0) :"ITEM_CD"
	arrField(1) = 2									' Field명(1) :"ITEM_NM"

	arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent,arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If	

End Function


'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetItemCd(Byval arrRet)
	With frm1
		.txtItemCd.focus
		.TxtItemCd.Value = arrRet(0)
		.TxtItemNm.Value = arrRet(1)

		lgBlnFlgChgValue = True
		
	End With
	
End Function

'======================================================================================================
' Function Name : FncBtnTotalExe
' Function Desc : This function is related to BtnTotalExe(일괄반영)
'=======================================================================================================
Function FncBtnTotalExe() 
	
	Dim IntRetCD 
	Dim lRow
	Dim lGrpCnt
	Dim strVal
	
	FncBtnTotalExe = False                                                  		       '⊙: Processing is NG

	Err.Clear                                                            	 		  '☜: Protect system from crashing
	
	On Error Resume Next                                           	       '☜: Protect system from crashing

	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	If Not chkField(Document, "1")  Then  '⊙: Check contents area
		Exit Function
	End If
    	
	IF lgIntFlgMode <> Parent.OPMD_UMODE  Then
		Call DisplayMsgBox("800167","x","x","x") '조회를 먼저하세요 
		Exit Function
	END IF


	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	

	With frm1
		.txtMode.value = Parent.UID_M0002
		.hChecked.value = "A"
			
		.txtMaxRows.value = lGrpCnt 
		
	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					'☜: 비지니스 ASP 를 가동 
	End With
	
	
	FncBtnTotalExe = True                                      	                    '⊙: Processing is OK
End Function


'======================================================================================================
' Function Name : FncBtnSelectedExe
' Function Desc : This function is related to BtnSelectedExe(선택반영)
'=======================================================================================================
Function FncBtnSelectedExe() 
	Dim IntRetCD 
	Dim lRow
	Dim lGrpCnt
	Dim strVal
		
	FncBtnSelectedExe = False                                                  		       '⊙: Processing is NG

	Err.Clear                                                            	 		  '☜: Protect system from crashing
	
	On Error Resume Next                                           	       '☜: Protect system from crashing


	
	'-----------------------
	'Check content area
	'-----------------------
	'ggoSpread.Source = frm1.vspdData
	'ggoSpread.SpreadLock C_ChkFlag, -1, C_ChkFlag
	
	If Not chkField(Document, "1")  Then  '⊙: Check contents area
		Exit Function
	End If
    	
	IF lgIntFlgMode <> Parent.OPMD_UMODE  Then
		Call DisplayMsgBox("800167","x","x","x") '조회를 먼저하세요 
		Exit Function
	END IF

	if SpreadWorkingChk = false then  Exit Function							'spread check box 체크 유무 

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If


	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	

 	With frm1
		.txtMode.value = Parent.UID_M0002
		.hChecked.value = "S"
			
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 0
    	strVal = ""
    		
    	'-----------------------
		'Data manipulate area
		'-----------------------
		
		
		For lRow = 1 To .vspdData.MaxRows
	    	.vspdData.Row = lRow
			.vspdData.Col = C_ChkFlag
			
			if .vspdData.value = 1  then
					
					.vspdData.Col = C_ItemCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep & Parent.gRowSep
					
					lGrpCnt = lGrpCnt + 1
			End if
		Next

		.txtMaxRows.value = lGrpCnt
		.txtSpread.value =  strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					'☜: 비지니스 ASP 를 가동 
	End With
	
	
	FncBtnSelectedExe = True                                      	                    '⊙: Processing is OK

End Function



'======================================================================================================
' Function Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

Function SpreadWorkingChk()
   Dim iRows
   Dim ichkCnt
   Dim IntRetCD

   SpreadWorkingChk = False
   ichkCnt = 0

   with frm1.vspdData
	For iRows = 1 to .MaxRows
	    .Col =  C_ChkFlag
	    .Row =  iRows
	    
	    if .Value = 1 then 
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

Function rdoRcpt_OnClick()
	With frm1	
	if .rdoTotAvg.checked = True Then
		.vspdData.Col = C_RcptPrc
		.vspdData.ColHidden = False		
		.vspdData.Col = C_TotAvgPrc
		.vspdData.ColHidden = True
	
		.rdoTotAvg.checked = False
		
		.vspdData.MaxRows = 0
		.hRadio.value = "RCPT"
		
		lgIntFlgMode = Parent.OPMD_CMODE
	END IF
	
	END WITH
	
	
End Function

Function rdoTotAvg_OnClick()
	With Frm1
	if .rdoRcpt.checked = True Then

		.vspdData.Col = C_RcptPrc
		.vspdData.ColHidden = True	
		.vspdData.Col = C_TotAvgPrc
		.vspdData.ColHidden = False

		.rdoRcpt.checked = False
	
		.hRadio.value = "TOTAVG"
		.vspdData.MaxRows = 0
		lgIntFlgMode = Parent.OPMD_CMODE
	END IF
	End With	
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
	
'======================================================================================================= -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>실제원가반영</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD>					
					<TD>&nbsp;</TD>					
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
									<TD CLASS="TD6"><script language =javascript src='./js/c3920ma1_fpDateTime1_txtYyyymm.js'></script></TD>

								    <TD CLASS=TD5 NOWRAP></TD>
								    <TD CLASS=TD6 NOWRAP></TD>

								</TR>
								<TR>	
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD6"><INPUT NAME="txtPlantCD" MAXLENGTH="4" SIZE=10  ALT ="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
													<INPUT NAME="txtPlantNM" MAXLENGTH="30" SIZE=25  ALT ="공장명" tag="14X"></TD>

									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemAccntCd" MAXLENGTH="2" SIZE=10  ALT ="품목계정" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenItemAccnt()">
													<INPUT NAME="txtItemAccntNM" MAXLENGTH="30" SIZE=20  ALT ="품목명" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenItemCd()">
														<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14"></TD>
								
								    <TD CLASS=TD5 NOWRAP>단가구분</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=radio CLASS="RADIO" NAME="rdoRcpt" id="rdoRcpt" VALUE="Y" tag = "11" CHECKED>
											<LABEL FOR="rdoRcpt">입고단가</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
										<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoTotAvg" id="rdoTotAvg" VALUE="N" tag = "11">
											<LABEL FOR="rdoTotAvg">평가단가</LABEL></TD>
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
						<script language =javascript src='./js/c3920ma1_vaSpread1_vspdData.js'></script>
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
				<TD><BUTTON NAME="btnTotalExe" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnTotalExe()" Flag=1>일괄반영</BUTTON>&nbsp;
					<BUTTON NAME="btnSelectedExe" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnSelectedExe()" Flag=1>선택반영</BUTTON>&nbsp;
				<TD>&nbsp</TD>
				<TD>&nbsp</TD>				
			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="2x" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAccntCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hChecked" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hRadio" tag="2x" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

