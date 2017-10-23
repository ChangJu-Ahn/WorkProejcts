
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 수불유형별 Variance 조회 
'*  3. Program ID           : c3950ma1.asp
'*  4. Program Name         : 수불유형별 Variance 조회 
'*  5. Program Desc         : 원가요소등록 
'*  6. Modified date(First) : 2003/03/25
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Tae Soo
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->
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
<!-- #Include file="../../inc/incSvrHTML.inc" -->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "C3950MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim C_PlantCd                                  '☆: Spread Sheet의 Column별 상수 
Dim C_PlantPriority
Dim C_ItemAcctNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_TrnsType
Dim C_TrnsTypeNm
Dim C_MoveType
Dim C_MoveTypeNm
Dim C_CostCenter
Dim C_CostCenterNm
Dim C_TrnsPlantCd
Dim C_TrnsPlantPriority
Dim C_TrnsSlCd
Dim	C_TrnsItemAcctNm
Dim	C_TrnsItemCd
Dim C_TrnsItemNm
Dim	C_Qty
Dim C_Amt
Dim C_DiffAmt
Dim C_IssuePrice



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
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_PlantCd			= 1                              
	C_PlantPriority		= 2
	C_ItemAcctNm		= 3
	C_ItemCd			= 4
	C_ItemNm			= 5
	C_TrnsType			= 6
	C_TrnsTypeNm		= 7
	C_MoveType			= 8
	C_MoveTypeNm		= 9
	C_CostCenter		=10
	C_CostCenterNm		=11
	C_TrnsPlantCd		=12
	C_TrnsPlantPriority	=13
	C_TrnsSlCd			=14
	C_TrnsItemAcctNm	=15
	C_TrnsItemCd		=16
	C_TrnsItemNm		=17
	C_Qty				=18
	C_Amt				=19
	C_DiffAmt			=20
	C_IssuePrice		=21


End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgSortKey         = 1                                       '⊙: initializes sort direction
    lgPageNo         = "0"
		
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

	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "MA") %>

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
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	    
	.MaxCols = C_IssuePrice + 1						
 	
    .Col = .MaxCols							
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	Call GetSpreadColumnPos("A")


	' ColumnPosition Header
    ggoSpread.SSSetEdit		C_PlantCd			,"공장"			,10,,,,2
    ggoSpread.SSSetEdit		C_PlantPriority		,"우선순위"		,10,,,,2
    ggoSpread.SSSetEdit		C_ItemAcctNm		,"품목계정명"	,10
    ggoSpread.SSSetEdit		C_ItemCd			,"품목"			,15,,,,2
    ggoSpread.SSSetEdit		C_ItemNm			,"품목명"		,20
	ggoSpread.SSSetEdit	    C_TrnsType			,"수불구분"		,10,,,,2
	ggoSpread.SSSetEdit		C_TrnsTypeNm		,"수불구분명"	,15
	ggoSpread.SSSetEdit	    C_MoveType			,"수불유형"		,10,,,,2
	ggoSpread.SSSetEdit		C_MoveTypeNm		,"수불유형명"	,15
	ggoSpread.SSSetEdit	    C_CostCenter		,"C/C"			,10,,,,2
	ggoSpread.SSSetEdit	    C_CostCenterNm		,"C/C명"		,15
	ggoSpread.SSSetEdit		C_TrnsPlantCd		,"이동공장"		,10,,,,2
	ggoSpread.SSSetEdit		C_TrnsPlantPriority	,"우선순위"		,10,,,,2
	ggoSpread.SSSetEdit		C_TrnsSlCd			,"이동창고"		,10,,,,2
	ggoSpread.SSSetEdit		C_TrnsItemAcctNm	,"이동품목계정"	,10,,,,2
	ggoSpread.SSSetEdit		C_TrnsItemCd		,"이동품목"		,15,,,,2
	ggoSpread.SSSetEdit		C_TrnsItemNm		,"이동품목명"	,20,,,,2
	ggoSpread.SSSetFloat	C_Qty				,"수량"			,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat	C_Amt				,"금액"			,15,Parent.ggAmtofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat	C_DiffAmt			,"차이금액"		,15,Parent.ggAmtofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat	C_IssuePrice		,"단가"			,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	
	
	

	if frm1.rdoItemAcct.value = "Y" then
		Call ggoSpread.SSSetColHidden(C_ItemCd ,C_ItemCd	,True)
		Call ggoSpread.SSSetColHidden(C_ItemNm ,C_ItemNm	,True)		
		Call ggoSpread.SSSetColHidden(C_TrnsPlantCd ,C_TrnsPlantCd	,True)		
		Call ggoSpread.SSSetColHidden(C_TrnsPlantPriority ,C_TrnsPlantPriority	,True)		
		Call ggoSpread.SSSetColHidden(C_TrnsSlCd ,C_TrnsSlCd	,True)
		Call ggoSpread.SSSetColHidden(C_TrnsItemAcctNm ,C_TrnsItemAcctNm	,True)	
		Call ggoSpread.SSSetColHidden(C_TrnsItemCd ,C_TrnsItemCd	,True)	
		Call ggoSpread.SSSetColHidden(C_TrnsItemNm ,C_TrnsItemNm	,True)	
		Call ggoSpread.SSSetColHidden(C_IssuePrice ,C_IssuePrice	,True)	
		frm1.hRadio.value = "ITEMACCT"
	Else
		frm1.hRadio.value = "ITEM"
	end if 
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
     
    .vspdData.ReDraw = False
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
              Frm1.vspdData.Action = 0 
              Exit For
           End If
       Next
          
    End If   
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
            
			C_PlantCd			= iCurColumnPos(1)                             
			C_PlantPriority		= iCurColumnPos(2)
			C_ItemAcctNm		= iCurColumnPos(3)
			C_ItemCd			= iCurColumnPos(4)
			C_ItemNm			= iCurColumnPos(5)
			C_TrnsType			= iCurColumnPos(6)
			C_TrnsTypeNm		= iCurColumnPos(7)
			C_MoveType			= iCurColumnPos(8)
			C_MoveTypeNm		= iCurColumnPos(9)
			C_CostCenter		= iCurColumnPos(10)
			C_CostCenterNm		= iCurColumnPos(11)
			C_TrnsPlantCd		= iCurColumnPos(12)
			C_TrnsPlantPriority	= iCurColumnPos(13)
			C_TrnsSlCd			= iCurColumnPos(14)
			C_TrnsItemAcctNm	= iCurColumnPos(15)
			C_TrnsItemCd		= iCurColumnPos(16)
			C_TrnsItemNm		= iCurColumnPos(17)
			C_Qty				= iCurColumnPos(18)
			C_Amt				= iCurColumnPos(19)
			C_DiffAmt			= iCurColumnPos(20)
			C_IssuePrice		= iCurColumnPos(21)
			            
    End Select    
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

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
    'Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	
    Call SetToolbar("1100000000001111")
    frm1.txtYyyymm.focus	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables
    
    if frm1.rdoItemAcct.checked = True Then
		frm1.hRadio.value = "ITEMACCT"
    else
    	frm1.hRadio.value = "ITEM"
    end if                                                           '⊙: Initializes local global variables
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If DbQuery() = False Then                                                      '☜: Query db data
       Exit Function
    End If
	
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
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
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
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDelete = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
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

   If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
			'SetSpreadColor2 .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    On Error Resume Next
    
    Dim iDx
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
	Call parent.FncPrint()                                                       '☜: Protect system from crashing
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

	Call parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(Parent.C_MULTI, True)

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
    On Error Resume Next
    
    DbQuery = False                                                              '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
  
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
     '@Query_Hidden     
			strVal = BIZ_PGM_ID
			strVal = strVal & "?txtYYYYMM="		& .hYYYYMM.value
			strVal = strVal & "&txtPlantCd="	& .hPlantCd.value				
			strVal = strVal & "&txtCostCd="		& .hCostCd.value				
			strVal = strVal & "&txtItemAcctCd=" & .hItemAcctCd.value
			strVal = strVal & "&txtItemCd="		& .hItemCd.value
			strVal = strVal & "&txtTrnsTypeCd="	& .hTrnsTypeCd.value
			strVal = strVal & "&txtMovTypeCd="	& .hMovTypeCd.value
			strVal = strVal & "&txtFlag="		& .hRadio.value
		Else
      '@Query_Text     
			Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
			strVal = BIZ_PGM_ID 
			strVal = strVal & "?txtYYYYMM="		& strYear & strMonth
			strVal = strVal & "&txtPlantCd="	& .txtPlantCd.value				
			strVal = strVal & "&txtCostCd="		& .txtCostCd.value				
			strVal = strVal & "&txtItemAcctCd=" & .txtItemAcctCd.value
			strVal = strVal & "&txtItemCd="		& .txtItemCd.value
			strVal = strVal & "&txtTrnsTypeCd="	& .txtTrnsTypeCd.value
			strVal = strVal & "&txtMovTypeCd="	& .txtMovTypeCd.value
			strVal = strVal & "&txtFlag="		& .hRadio.value
		END IF
		strVal = strVal & "&lgPageNo="			& lgPageNo								'Next key tag
'		strVal = strVal & "&lgMaxCount="		& C_SHEETMAXROWS_D					'한번에 가져올수 있는 데이타 건수 
	    
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
		
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel

    On Error Resume Next
    DbSave = False                                                               '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call DisableToolBar(Parent.TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
		
    Frm1.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

   If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

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
     If Err.number = 0 Then	 
       DbDelete = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	Dim iRow
	
    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'Call SetToolbar("110000000001111")	
	Frm1.vspdData.Focus
	

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery() = False Then
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
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "1")										     '⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	Call DisplayMsgBox("800154", "x","x","x")					 '☜: 작업이 완료되었습니다 
	frm1.txtVersion.focus()
	
	Set gActiveElement = document.ActiveElement   
 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : OpenPlant
' Desc : 공장 팝업 
'========================================================================================================
Function OpenPopup(ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strYear,strMonth,strDay,strYyyymm

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strYyyymm = strYear & strMonth
	
	Select case iWhere
		case 1
			arrParam(0) = "공장팝업"	
			arrParam(1) = "B_PLANT"
			arrParam(2) = Trim(frm1.txtPlantCd.Value)
			arrParam(3) = ""											
			arrParam(4) = ""											
			arrParam(5) = "공장"							
	
			arrField(0) = "plant_cd"						
			arrField(1) = "plant_nm"						
    
			arrHeader(0) = "공장"				
    		arrHeader(1) = "공장명"				
		case 2
			arrParam(0) = "코스트센터팝업"	
			arrParam(1) = "B_COST_CENTER "
			arrParam(2) = Trim(frm1.txtCostCd.Value)
			arrParam(3) = ""											
			arrParam(4) = ""
			arrParam(5) = "코스트센터"							
	
			arrField(0) = "cost_cd"						
			arrField(1) = "cost_nm"						
    
			arrHeader(0) = "코스트센터"				
    		arrHeader(1) = "코스트센터명"	
    	case 3
			arrParam(0) = "수불구분팝업"	
			arrParam(1) = "B_MINOR"
			arrParam(2) = Trim(frm1.txtTrnsTypeCd.Value)
			arrParam(3) = ""											
			arrParam(4) = "major_cd = " & FilterVar("I0002", "''", "S") & " "											
			arrParam(5) = "수불구분"							
	
			arrField(0) = "minor_cd"						
			arrField(1) = "minor_nm"						
    
			arrHeader(0) = "수불구분"				
    		arrHeader(1) = "수불구분명"
    	case 4
			arrParam(0) = "수불유형팝업"	
			arrParam(1) = "B_MINOR"
			arrParam(2) = Trim(frm1.txtMovTypeCd.Value)
			arrParam(3) = ""
			IF Trim(frm1.txtTrnsTypeCd.value) <> "" Then
				arrParam(4) = "major_cd = " & FilterVar("I0001", "''", "S") & "  and minor_cd in (select mov_type from i_movetype_configuration where trns_type = " & FilterVar(frm1.txtTrnsTypeCd.value, "''", "S") & ")"																						
			ELSE
				arrParam(4) = "major_cd = " & FilterVar("I0001", "''", "S") & " "
			END IF
			arrParam(5) = "코스트센터"							
	
			arrField(0) = "minor_cd"						
			arrField(1) = "minor_nm"						
    
			arrHeader(0) = "수불유형"				
    		arrHeader(1) = "수불유형명"
    	 case 5
			arrParam(0) = "품목계정팝업"	
			arrParam(1) = "B_MINOR a,b_item_acct_inf b"
			arrParam(2) = Trim(frm1.txtItemAcctCd.Value)
			arrParam(3) = ""											
			arrParam(4) = "major_cd = " & FilterVar("P1001", "''", "S") & " and a.minor_cd = b.item_acct and b.item_acct_group <> "	& FilterVar("6MRO","''","S")										
			arrParam(5) = "품목계정"							
	
			arrField(0) = "minor_cd"						
			arrField(1) = "minor_nm"						
    
			arrHeader(0) = "품목계정"				
    		arrHeader(1) = "품목계정명"
    	case 6
			arrParam(0) = "품목팝업"	
			IF Trim(frm1.txtPlantCd.value) <> "" Then
				arrParam(1) = "B_ITEM a,B_ITEM_BY_PLANT b"
				arrParam(2) = Trim(frm1.txtItemCd.Value)
				arrParam(3) = ""
				IF Trim(frm1.txtItemAcctCd.value) <> "" Then
					arrParam(4) = " b.plant_cd = " & FilterVar(frm1.txtPlantCD.value, "''", "S") & " and a.item_cd = b.item_cd and b.item_acct = " & FilterVar(frm1.txtItemAcctCd.value, "''", "S") 
				ELSE
					arrParam(4) = " b.plant_cd = " & FilterVar(frm1.txtPlantCD.value, "''", "S") & " and a.item_cd = b.item_cd "
				END IF
			ELSE
				arrParam(1) = "B_ITEM a"
				arrParam(2) = Trim(frm1.txtItemCd.Value)
				arrParam(3) = ""
				IF Trim(frm1.txtItemAcctCd.value) <> "" Then
					arrParam(4) = " a.item_acct = " & FilterVar(frm1.txtItemAcctCd.value, "''", "S") 
				ELSE
					arrParam(4) = ""
				END IF
			END IF
			
			arrParam(5) = "품목"							
	
			arrField(0) = "a.item_cd"						
			arrField(1) = "a.item_nm"						
    
			arrHeader(0) = "품목"				
    		arrHeader(1) = "품목명"    		    	 	    				 
	
	end select

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	  Select case iWhere
		case 1
			frm1.txtPlantCD.focus		
		case 2
			frm1.txtCostCd.focus		
		case 3
			frm1.txtTrnsTypeCd.focus		
		case 4
			frm1.txtMovTypeCd.focus		
		case 5
			frm1.txtItemAcctCd.focus		
		case 6
			frm1.txtItemCd.focus
	  End Select		
		Exit Function
	Else
		Call SetPopUp(iWhere,arrRet)
	End If
		
End Function


Function SetPopUp(byval iWhere,byval arrRet)
	select case iWhere
		case 1
			frm1.txtPlantCD.focus	
			frm1.txtPlantCd.Value = arrRet(0)
			frm1.txtPlantNM.value = arrRet(1)
		case 2
			frm1.txtCostCd.focus
			frm1.txtCostCd.Value = arrRet(0)
			frm1.txtCostNM.value = arrRet(1)
		case 3
			frm1.txtTrnsTypeCd.focus	
			frm1.txtTrnsTypeCd.Value = arrRet(0)
			frm1.txtTrnsTypeNm.value = arrRet(1)
		case 4
			frm1.txtMovTypeCd.focus	
			frm1.txtMovTypeCd.Value = arrRet(0)
			frm1.txtMovTypeNm.value = arrRet(1)
		case 5
			frm1.txtItemAcctCd.focus
			frm1.txtItemAcctCd.Value = arrRet(0)
			frm1.txtItemAcctNm.value = arrRet(1)
		case 6
			frm1.txtItemCd.focus
			frm1.txtItemCd.Value = arrRet(0)
			frm1.txtItemNm.value = arrRet(1)
	end select
						
End Function
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

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
Sub vspdData_Click(Col, Row)
    
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

	
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort	Col			'Sort in ascending
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col	,lgSortKey	'Sort in descending
            lgSortKey = 1
        End If
    Else
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_CE_NM Or NewCol <= C_CE_NM Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
    	If lgPageNo <> "0" Then                         
           If DbQuery() = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    'If Row > 0 And Col = C Then
    '    .Col = Col
    '    .Row = Row
    '    
	'End If
    
    End With
End Sub

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


Function rdoItemAcct_OnClick()
	With frm1	
		if .rdoItem.checked = True Then
			.vspdData.Col = C_ItemCd
			.vspdData.ColHidden = True
			.vspdData.Col = C_ItemNm
			.vspdData.ColHidden = True
			.vspdData.Col = C_TrnsPlantCd
			.vspdData.ColHidden = True
			.vspdData.Col = C_TrnsPlantPriority
			.vspdData.ColHidden = True
			.vspdData.Col = C_TrnsSlCd
			.vspdData.ColHidden = True
			.vspdData.Col = C_TrnsItemAcctNm
			.vspdData.ColHidden = True
			.vspdData.Col = C_TrnsItemCd
			.vspdData.ColHidden = True
			.vspdData.Col = C_TrnsItemNm
			.vspdData.ColHidden = True
			.vspdData.Col = C_IssuePrice
			.vspdData.ColHidden = True
						
							
			.rdoItem.checked = False
			
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
			
			.txtQtySum.text = "0"
			.txtAmtSum.text = "0"	
			.txtDiffSum.text = "0"
					
			.hRadio.value = "ITEMACCT"
			
			lgIntFlgMode = Parent.OPMD_CMODE
		END IF
	
	END WITH
	
	
End Function

Function rdoItem_OnClick()
	With Frm1
		if .rdoItemAcct.checked = True Then
			.vspdData.Col = C_ItemCd
			.vspdData.ColHidden = False
			.vspdData.Col = C_ItemNm
			.vspdData.ColHidden = False
			.vspdData.Col = C_TrnsPlantCd
			.vspdData.ColHidden = False
			.vspdData.Col = C_TrnsPlantPriority
			.vspdData.ColHidden = False
			.vspdData.Col = C_TrnsSlCd
			.vspdData.ColHidden = False
			.vspdData.Col = C_TrnsItemAcctNm
			.vspdData.ColHidden = False
			.vspdData.Col = C_TrnsItemCd
			.vspdData.ColHidden = False
			.vspdData.Col = C_TrnsItemNm
			.vspdData.ColHidden = False
			.vspdData.Col = C_IssuePrice
			.vspdData.ColHidden = False
			
			.rdoItemAcct.checked = False
			
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

			.txtQtySum.text = "0"
			.txtAmtSum.text = "0"	
			.txtDiffSum.text = "0"
							
			.hRadio.value = "ITEM"
			
			lgIntFlgMode = Parent.OPMD_CMODE
		END IF
	End With	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" bgColor=White text=Black>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수불유형별 Variance조회</font></td>
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
									<TD CLASS="TD6">
										<script language =javascript src='./js/c3950ma1_OBJECT1_txtYyyymm.js'></script>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlantCD"  SIZE=10  ALT ="공장" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(1)">
													<INPUT NAME="txtPlantNM"  SIZE=30  ALT ="공장명" tag="14X"></TD>
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>Cost Center</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCostCd"  SIZE=10  ALT ="코스트센터" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(2)">
													<INPUT NAME="txtCostNm"  SIZE=30  ALT ="코스트센터명" tag="14X"></TD>

									<TD CLASS="TD5" NOWRAP>수불구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTrnsTypeCd"  SIZE=10  ALT ="수불구분" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTnsTypeCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(3)">
													<INPUT NAME="txtTrnsTypeNm" MAXLENGTH="30" SIZE=30  ALT ="수불구분명" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>수불유형</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtMovTypeCd" SIZE=10  tag="11XXXU" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovTypeCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(4)">
														<INPUT TYPE=TEXT NAME="txtMovTypeNm" SIZE=30 tag="14"></TD>

									<TD CLASS="TD5" NOWRAP>품목계정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemAcctCd" SIZE=10  tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(5)">
														<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 tag="14"></TD>

								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemCd" SIZE=10  tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(6)">
														<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14"></TD>
								    <TD CLASS=TD5 NOWRAP>조회구분</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=radio CLASS="RADIO" NAME="rdoItemAcct" id="rdoItemAcct" VALUE="Y" tag = "11" CHECKED>
											<LABEL FOR="rdoItemAcct">품목계정</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
										<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoItem" id="rdoItem" VALUE="N" tag = "11">
											<LABEL FOR="rdoItem">품목</LABEL></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c3950ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>수량합계</TD>
									<TD CLASS=TD5 NOWRAP>											
										<script language =javascript src='./js/c3950ma1_fpDoubleSingle2_txtQtySum.js'></script>&nbsp;
	                                </TD>
									<TD CLASS="TD5" NOWRAP>금액합계</TD>
									<TD CLASS=TD5 NOWRAP>
										<script language =javascript src='./js/c3950ma1_fpDoubleSingle2_txtAmtSum.js'></script>&nbsp;
									</TD>
									<TD CLASS="TD5" NOWRAP>차이금액합계</TD>
									<TD CLASS=TD5 NOWRAP>
										<script language =javascript src='./js/c3950ma1_fpDoubleSingle2_txtDiffSum.js'></script>&nbsp;
									</TD>									
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hTrnsTypeCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hMovTypeCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAcctCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hRadio" tag="2x" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

