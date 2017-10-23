<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 이동유형별 Variance 조회 
'*  3. Program ID           : c3950ma1.asp
'*  4. Program Name         : 이동유형별 Variance 조회 
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
Const BIZ_PGM_ID = "GB003MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim C_BizUnitCd                                  '☆: Spread Sheet의 Column별 상수 
Dim C_CostCd
Dim C_CostNm
Dim C_SalesOrg
Dim C_SalesGrp
Dim C_SalesGrpNm
Dim C_BpCd
Dim C_BpNm
Dim C_SoType
Dim C_SoTypeNm
Dim C_ItemGroupNm
Dim C_ItemCd
Dim C_ItemNm
Dim	C_Qty
Dim C_SalesAmt
Dim C_Gain
Dim C_GainNm
Dim C_CostAmt


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          


'========================================================================================================
Sub initSpreadPosVariables() 
	C_BizUnitCd     =1                             '☆: Spread Sheet의 Column별 상수 
	C_CostCd		=2
	C_CostNm=3
	C_SalesOrg=4
	C_SalesGrp=5
	C_SalesGrpNm=6
	C_BpCd=7
	C_BpNm=8
	C_SoType=9
	C_SoTypeNm=10
	C_ItemGroupNm=11
	C_ItemCd=12
	C_ItemNm=13
	C_Qty=14
	C_SalesAmt=15
	C_Gain=16
	C_GainNm=17
	C_CostAmt=18


End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgSortKey         = 1                                       '⊙: initializes sort direction
    lgPageNo         = "0"
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	With frm1
		.hYYYYMM.value = ""
		.hBizUnitCd.value = ""				
		.hCostCd.value	 = ""			
		.hSalesOrg.value = ""
		.hSalesGrp.value = ""
		.hBpCd.value = ""
		.hSoTypeCd.value = ""
		.hItemAcct.value = ""
		.hItemGroupCd.value = ""
		.txtPrevCostCd.value = ""
		.txtPrevSalesGrp.value = ""
		.txtPrevBpCd.value = ""
		.txtPrevSoType.value= ""
		.txtPrevItemCd.value= ""		
	end with

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

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
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        



'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	

	With frm1.vspdData
	    
	.MaxCols = C_CostAmt + 1						
 	
    .Col = .MaxCols							
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	Call GetSpreadColumnPos("A")


	' ColumnPosition Header
    ggoSpread.SSSetEdit		C_BizUnitCd			,"사업부"			,10,,,,2
    ggoSpread.SSSetEdit		C_CostCd		,"C/C"	,10,,,,2
    ggoSpread.SSSetEdit		C_CostNm			,"C/C명"			,15
    ggoSpread.SSSetEdit		C_SalesOrg			,"영업조직"			,10,,,,2
	ggoSpread.SSSetEdit	    C_SalesGrp			,"영업그룹"			,10,,,,2
	ggoSpread.SSSetEdit		C_SalesGrpNm		,"영업그룹명"	,	15
	ggoSpread.SSSetEdit	    C_BpCd			,"거래처"		,10,,,,2
	ggoSpread.SSSetEdit		C_BpNm			,"거래처명"	,15
	ggoSpread.SSSetEdit	    C_SoType		,"판매유형"			,10,,,,2
	ggoSpread.SSSetEdit	    C_SoTypeNm		,"판매유형명"		,20
	ggoSpread.SSSetEdit		C_ItemGroupNm		,"품목그룹"		,20
	ggoSpread.SSSetEdit		C_ItemCd		,"품목"		,15,,,,2
	ggoSpread.SSSetEdit		C_ItemNm		,"품목명"	,20
	ggoSpread.SSSetFloat	C_Qty			,"수량"		,10,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat	C_SalesAmt		,"매출액"	,15,Parent.ggAmtofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetEdit		C_Gain			,"매출원가항목"		,15,,,,2
	ggoSpread.SSSetEdit		C_GainNm		,"매출원가항목명"		,15
	ggoSpread.SSSetFloat	C_CostAmt		,"매출원가"	,15,Parent.ggAmtofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	

	.ReDraw = true
	
    Call SetSpreadLock() 
    
    End With
    
End Sub


'======================================================================================================
Sub SetSpreadLock()
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
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
              Frm1.vspdData.Action = 0 
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

			C_BizUnitCd	= iCurColumnPos(1)
			C_CostCd	= iCurColumnPos(2)
			C_CostNm	= iCurColumnPos(3)
			C_SalesOrg	= iCurColumnPos(4)
			C_SalesGrp	= iCurColumnPos(5)
			C_SalesGrpNm	= iCurColumnPos(6)
			C_BpCd	= iCurColumnPos(7)
			C_BpNm	= iCurColumnPos(8)
			C_SoType	= iCurColumnPos(9)
			C_SoTypeNm	= iCurColumnPos(10)
			C_ItemGroupNm	= iCurColumnPos(11)
			C_ItemCd	= iCurColumnPos(12)
			C_ItemNm	= iCurColumnPos(13)
			C_Qty	= iCurColumnPos(14)
			C_SalesAmt	= iCurColumnPos(15)
			C_Gain	= iCurColumnPos(16)
			C_GainNm	= iCurColumnPos(17)
			C_CostAmt	= iCurColumnPos(18)

    End Select    
End Sub


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
    
  	Set gActiveElement = document.activeElement			
			
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
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables
    
                                                       '⊙: Initializes local global variables
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
	Call parent.FncPrint()                                                       '☜: Protect system from crashing
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

	Call parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(Parent.C_MULTI, True)

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
    Call ggoSpread.ReOrderingSpreadData()
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
			strVal = strVal & "&txtBizUnitCd="	& .hBizUnitCd.value				
			strVal = strVal & "&txtCostCd="		& .hCostCd.value				
			strVal = strVal & "&txtSalesOrg="	& .hSalesOrg.value
			strVal = strVal & "&txtSalesGrp="	& .hSalesGrp.value
			strVal = strVal & "&txtBpCd="		& .hBpCd.value
			strVal = strVal & "&txtSoType="		& .hSoTypeCd.value
			strVal = strVal & "&txtItemAcct="	& .hItemAcct.value
			strVal = strVal & "&txtItemGroupCd="	& .txtItemGroupCd.value
			strVal = strVal & "&txtPrevCostCd="		& .txtPrevCostCd.value
			strVal = strVal & "&txtPrevSalesGrp="	& .txtPrevSalesGrp.value
			strVal = strVal & "&txtPrevBpCd="		& .txtPrevBpCd.value
			strVal = strVal & "&txtPrevSoType="		& .txtPrevSoType.value
			strVal = strVal & "&txtPrevItemCd="		& .txtPrevItemCd.value
		Else
      '@Query_Text     
			Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
			strVal = BIZ_PGM_ID
			strVal = strVal & "?txtYyyymm="		& strYear & strMonth
			strVal = strVal & "&txtBizUnitCd="	& .txtBizUnitCd.value				
			strVal = strVal & "&txtCostCd="		& .txtCostCd.value				
			strVal = strVal & "&txtSalesOrg="	& .txtSalesOrg.value
			strVal = strVal & "&txtSalesGrp="	& .txtSalesGrp.value
			strVal = strVal & "&txtBpCd="		& .txtBpCd.value
			strVal = strVal & "&txtSoType="		& .txtSoType.value
			strVal = strVal & "&txtItemAcct="	& .txtItemAcct.value
			strVal = strVal & "&txtItemGroupCd="	& .txtItemGroupCd.value
			strVal = strVal & "&txtPrevCostCd="		& .txtPrevCostCd.value
			strVal = strVal & "&txtPrevSalesGrp="	& .txtPrevSalesGrp.value
			strVal = strVal & "&txtPrevBpCd="		& .txtPrevBpCd.value
			strVal = strVal & "&txtPrevSoType="		& .txtPrevSoType.value
			strVal = strVal & "&txtPrevItemCd="		& .txtPrevItemCd.value
			
			
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
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "1")										     '⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	Call DisplayMsgBox("800154", "x","x","x")					 '☜: 작업이 완료되었습니다 
	
	Set gActiveElement = document.ActiveElement   
 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


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
			arrParam(0) = "사업부팝업"	
			arrParam(1) = "B_BIZ_UNIT"
			arrParam(2) = Trim(frm1.txtBizUnitCd.Value)
			arrParam(3) = ""											
			arrParam(4) = ""											
			arrParam(5) = "사업부"							
	
			arrField(0) = "biz_unit_cd"						
			arrField(1) = "biz_unit_nm"						
    
			arrHeader(0) = "사업부"				
    		arrHeader(1) = "사업부명"				
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
			arrParam(0) = "영업조직팝업"	
			arrParam(1) = "B_SALES_ORG"
			arrParam(2) = Trim(frm1.txtSalesOrg.Value)
			arrParam(3) = ""											
			arrParam(4) = ""											
			arrParam(5) = "영업조직"							
	
			arrField(0) = "sales_org"						
			arrField(1) = "sales_org_nm"						
    
			arrHeader(0) = "영업조직"				
    		arrHeader(1) = "영업조직명"
    	case 4
			arrParam(0) = "영업그룹팝업"	
			arrParam(1) = "B_SALES_GRP"
			arrParam(2) = Trim(frm1.txtSalesGrp.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "영업그룹"							
	
			arrField(0) = "sales_grp"						
			arrField(1) = "sales_grp_nm"						
    
			arrHeader(0) = "영업그룹"				
    		arrHeader(1) = "영업그룹명"
    	 case 5
			arrParam(0) = "거래처"	
			arrParam(1) = "B_BIZ_PARTNER"
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""											
			arrParam(4) = ""											
			arrParam(5) = "거래처"							
	
			arrField(0) = "bp_Cd"						
			arrField(1) = "bp_nm"						
    
			arrHeader(0) = "거래처"				
    		arrHeader(1) = "거래처명"
    	case 6
			arrParam(0) = "판매유형"	
			arrParam(1) = "(select so_type,so_type_nm from s_so_type_Config union all select minor_cd as so_type,minor_nm as so_type_nm from b_minor where major_cd = " & FilterVar("G1025", "''", "S") & "  ) a"
			arrParam(2) = Trim(frm1.txtSoType.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "판매유형"
	
			arrField(0) = "a.so_type"						
			arrField(1) = "a.so_type_nm"						
    
			arrHeader(0) = "판매유형"
			arrHeader(1) = "판매유형명"    		    	 	    				 
    	case 7
			arrParam(0) = "품목계정"	
			arrParam(1) = "B_MINOR"
			arrParam(2) = Trim(frm1.txtItemAcct.Value)
			arrParam(3) = ""
			arrParam(4) = "major_cd = " & FilterVar("P1001", "''", "S") & ""
			arrParam(5) = "품목계정"
	
			arrField(0) = "minor_cd"
			arrField(1) = "minor_nm"
    
			arrHeader(0) = "품목계정"
			arrHeader(1) = "품목계정명"    		    	 	    				 
    	case 8
			arrParam(0) = "품목그룹"	
			arrParam(1) = "B_ITEM_GROUP"
			arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "품목그룹"
	
			arrField(0) = "item_group_Cd"						
			arrField(1) = "item_group_nm"						
    
			arrHeader(0) = "품목그룹"
			arrHeader(1) = "품목그룹명"    		    	 	    				 
	
	
	end select

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	 Select case iWhere
		case 1
			frm1.txtBizUnitCd.focus
		case 2
			frm1.txtCostCd.focus
		case 3
			frm1.txtSalesOrg.focus
		case 4
			frm1.txtSalesGrp.focus
		case 5
			frm1.txtBpCd.focus
		case 6
			frm1.txtSoType.focus
		case 7
			frm1.txtItemAcct.focus
		case 8
			frm1.txtItemGroupCd.focus
	 End Select	
		Exit Function
	Else
		Call SetPopUp(iWhere,arrRet)
	End If
		
End Function


Function SetPopUp(byval iWhere,byval arrRet)

	select case iWhere
		case 1
			frm1.txtBizUnitCd.focus
			frm1.txtBizUnitCd.Value = arrRet(0)
			frm1.txtBizUnitNm.value = arrRet(1)
		case 2
			frm1.txtCostCd.focus
			frm1.txtCostCd.Value = arrRet(0)
			frm1.txtCostNM.value = arrRet(1)
		case 3
			frm1.txtSalesOrg.focus
			frm1.txtSalesOrg.Value = arrRet(0)
			frm1.txtSalesOrgNm.value = arrRet(1)
		case 4
			frm1.txtSalesGrp.focus
			frm1.txtSalesGrp.Value = arrRet(0)
			frm1.txtSalesGrpNm.value = arrRet(1)
		case 5
			frm1.txtBpCd.focus
			frm1.txtBpCd.Value = arrRet(0)
			frm1.txtBpNm.value = arrRet(1)
		case 6
			frm1.txtSoType.focus
			frm1.txtSoType.Value = arrRet(0)
			frm1.txtSoTypeNm.value = arrRet(1)
		case 7
			frm1.txtItemAcct.focus
			frm1.txtItemAcct.Value = arrRet(0)
			frm1.txtItemAcctNm.value = arrRet(1)
		case 8
			frm1.txtItemGroupCd.focus
			frm1.txtItemGroupCd.Value = arrRet(0)
			frm1.txtItemGroupNm.value = arrRet(1)						
	end select
						
End Function


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
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    


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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출/매출원가조회</font></td>
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
										<script language =javascript src='./js/gb003ma1_OBJECT1_txtYyyymm.js'></script>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>										
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>사업부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBizUnitCd"  SIZE=10  ALT ="사업부" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(1)">
													<INPUT NAME="txtBizUnitNm"  SIZE=25  ALT ="사업부명" tag="14X"></TD>


									<TD CLASS="TD5" NOWRAP>Cost Center</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCostCd"  SIZE=10  ALT ="코스트센터" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(2)">
													<INPUT NAME="txtCostNm"  SIZE=25  ALT ="코스트센터명" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>영업조직</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtSalesOrg"  SIZE=10  ALT ="영업조직" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTnsTypeCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(3)">
													<INPUT NAME="txtSalesOrgNm" MAXLENGTH="25" SIZE=25  ALT ="영업조직명" tag="14X"></TD>
								
								
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtSalesGrp" SIZE=10  tag="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovTypeCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(4)">
														<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=25 tag="14" ALT="영업그룹명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtBpCd" SIZE=10  tag="11XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(5)">
														<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14"></TD>

									<TD CLASS="TD5" NOWRAP>판매유형</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtSoType" SIZE=10  tag="11XXXU" ALT="판매유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(6)">
														<INPUT TYPE=TEXT NAME="txtSoTypeNm" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목계정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemAcct" SIZE=10  tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(7)">
														<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=25 tag="14"></TD>

									<TD CLASS="TD5" NOWRAP>품목그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemGroupCd" SIZE=10  tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(8)">
														<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 tag="14"></TD>

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
								<script language =javascript src='./js/gb003ma1_OBJECT1_vspdData.js'></script>
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
									<TD CLASS="TD5" NOWRAP>매출액합계</TD>
									<TD CLASS=TD5 NOWRAP>											
										<script language =javascript src='./js/gb003ma1_fpDoubleSingle2_txtSalesAmtSum.js'></script>&nbsp;
	                                </TD>
									<TD CLASS="TD5" NOWRAP>매출원가합계</TD>
									<TD CLASS=TD5 NOWRAP>
										<script language =javascript src='./js/gb003ma1_fpDoubleSingle2_txtCostAmtSum.js'></script>&nbsp;
									</TD>
									<TD CLASS="TD5" NOWRAP>매출이익합계</TD>
									<TD CLASS=TD5 NOWRAP>
										<script language =javascript src='./js/gb003ma1_fpDoubleSingle2_txtProfitSum.js'></script>&nbsp;
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
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hBizUnitCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hSalesOrg" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hSalesGrp" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hBpCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hSoTypeCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevCostCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevSalesGrp" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevBpCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevSoType" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevItemCd" tag="2x" TABINDEX= "-1">


</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

	
