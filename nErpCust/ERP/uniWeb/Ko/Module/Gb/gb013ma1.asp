
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : 경영손익 
'*  2. Function Name        : 매출/매출원가 시점차이내역조회 
'*  3. Program ID           : GB013MA1
'*  4. Program Name         : 매출/매출원가 시점차이내역조회 
'*  5. Program Desc         : 매출/매출원가 시점차이내역조회 
'*  6. Modified date(First) : 2003/06/16
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park Joon-Won
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
Const BIZ_PGM_ID = "GB013MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim C_BizUnitCd                                  '☆: Spread Sheet의 Column별 상수 
Dim	C_CostCd		
Dim	C_CostNm		
Dim	C_BpCd			
Dim	C_BpNm			
Dim	C_SalesOrg		
Dim	C_SalesGrp		
Dim	C_SalesGrpNm	
Dim	C_ItemGrp       
Dim	C_ItemGrpNm     
Dim	C_ItemCd		
Dim	C_ItemNm		
Dim	C_SoType		
Dim	C_SoTypeNm		
Dim	C_ItemDocNo		
Dim	C_ShipMentNo		
Dim	C_ShipMentSeq	
Dim	C_BillingNo		
Dim	C_BillingSeq	
Dim	C_BillPostFlag	
Dim	C_TaxNo			
Dim	C_TaxSeq		
Dim	C_TaxPostFlag	
Dim	C_IssueDt		
Dim	C_SaleBondDt		
Dim	C_TaxDt			
Dim	C_StockUnit		
Dim	C_IssueQty		
Dim	C_IssueAmt		
Dim	C_BillingQty		
Dim	C_BillingAmt	
Dim	C_TaxQty		
Dim	C_TaxAmt		
	
Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 

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
	C_BizUnitCd     =1                             '☆: Spread Sheet의 Column별 상수 
	C_CostCd		=2
	C_CostNm		=3
	C_BpCd			=4
	C_BpNm			=5
	C_SalesOrg		=6
	C_SalesGrp		=7
	C_SalesGrpNm	=8
	C_ItemGrp       =9
	C_ItemGrpNm     =10
	C_ItemCd		=11
	C_ItemNm		=12
	C_SoType		=13
	C_SoTypeNm		=14
	C_ItemDocNo		=15
	C_ShipMentNo	=16	
	C_ShipMentSeq	=17
	C_BillingNo		=18
	C_BillingSeq	=19
	C_BillPostFlag	=20
	C_TaxNo			=21
	C_TaxSeq		=22
	C_TaxPostFlag	=23
	C_IssueDt		=24
	C_SaleBondDt	=25	
	C_TaxDt			=26
	C_StockUnit		=27
	C_IssueQty		=28
	C_IssueAmt		=29
	C_BillingQty	=30	
	C_BillingAmt	=31
	C_TaxQty		=32
	C_TaxAmt		=33

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
		.hFromSaleDt.value = ""
		.hToSaleDt.value = ""
		.hFromTaxDt.value = ""
		.hToTaxDt.value = ""				
	end with

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
	
'	Call ggoOper.ClearField(Document, "2")	 
	
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

	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "QA") %>

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

	.MaxCols = C_TaxAmt + 1						
 	
    .Col = .MaxCols							
    .ColHidden = True

    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030616",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	Call GetSpreadColumnPos("A")


	' ColumnPosition Header
    ggoSpread.SSSetEdit		C_BizUnitCd		,"사업부"		,10,,,,2
    ggoSpread.SSSetEdit		C_CostCd		,"C/C"			,10,,,,2
    ggoSpread.SSSetEdit		C_CostNm		,"C/C명"		,15
	ggoSpread.SSSetEdit	    C_BpCd			,"거래처"		,10,,,,2
    ggoSpread.SSSetEdit		C_BpNm			,"거래처명"		,15
    ggoSpread.SSSetEdit		C_SalesOrg		,"영업조직"		,10,,,,2
	ggoSpread.SSSetEdit	    C_SalesGrp		,"영업그룹"		,10,,,,2
	ggoSpread.SSSetEdit		C_SalesGrpNm	,"영업그룹명"	,15
	ggoSpread.SSSetEdit		C_ItemGrp		,"품목그룹"		,15,,,,2
	ggoSpread.SSSetEdit		C_ItemGrpNm		,"품목그룹명"	,20
	ggoSpread.SSSetEdit		C_ItemCd		,"품목"			,15,,,,2
	ggoSpread.SSSetEdit		C_ItemNm		,"품목명"		,20
	ggoSpread.SSSetEdit	    C_SoType		,"판매유형"		,10,,,,2
	ggoSpread.SSSetEdit	    C_SoTypeNm		,"판매유형명"	,20
	ggoSpread.SSSetEdit		C_ItemDocNo		,"수불번호"		,15
	ggoSpread.SSSetEdit		C_ShipMentNo	,"출하번호"		,20
	ggoSpread.SSSetFloat	C_ShipMentSeq	,"출하SEQ"		,8, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
	ggoSpread.SSSetEdit		C_BillingNo		,"매출채권번호"	,15	
	ggoSpread.SSSetFloat	C_BillingSeq	,"매출채권SEQ"	,12, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
	ggoSpread.SSSetEdit		C_BillPostFlag	,"POSTING여부"	,12,,,,2
	ggoSpread.SSSetEdit		C_TaxNo			,"세금계산서번호"	,15	
	ggoSpread.SSSetFloat	C_TaxSeq		,"세금계산서SEQ",12, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
	ggoSpread.SSSetEdit		C_TaxPostFlag	,"POSTING여부"	,12,,,,2
	ggoSpread.SSSetDate		C_IssueDt		,"출고일"		,12, 2, Parent.gDateFormat
	ggoSpread.SSSetDate		C_SaleBondDt	,"매출채권일"	,12, 2, Parent.gDateFormat
	ggoSpread.SSSetDate		C_TaxDt			,"세금계산서일"	,12, 2, Parent.gDateFormat
	ggoSpread.SSSetEdit		C_StockUnit		,"재고단위"		,10,0
	ggoSpread.SSSetFloat	C_IssueQty		,"출고수량"		,15, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat	C_IssueAmt		,"출고금액"		,15, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
	ggoSpread.SSSetFloat	C_BillingQty	,"매출채권수량"	,15, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat	C_BillingAmt	,"매출채권금액"	,15, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
	ggoSpread.SSSetFloat	C_TaxQty		,"세금계산서수량",15, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat	C_TaxAmt		,"세금계산서금액",15, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec		
	
    Call ggoSpread.SSSetColHidden(C_SoType,C_SoType,True)

	
	.ReDraw = true
	
	
    Call SetSpreadLock() 
    
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
			C_BizUnitCd     =iCurColumnPos(1)                             '☆: Spread Sheet의 Column별 상수 
			C_CostCd		=iCurColumnPos(2)
			C_CostNm		=iCurColumnPos(3)
			C_BpCd			=iCurColumnPos(4)
			C_BpNm			=iCurColumnPos(5)
			C_SalesOrg		=iCurColumnPos(6)
			C_SalesGrp		=iCurColumnPos(7)
			C_SalesGrpNm	=iCurColumnPos(8)
			C_ItemGrp       =iCurColumnPos(9)
			C_ItemGrpNm     =iCurColumnPos(10)
			C_ItemCd		=iCurColumnPos(11)
			C_ItemNm		=iCurColumnPos(12)
			C_SoType		=iCurColumnPos(13)
			C_SoTypeNm		=iCurColumnPos(14)
			C_ItemDocNo		=iCurColumnPos(15)
			C_ShipMentNo	=iCurColumnPos(16)	
			C_ShipMentSeq	=iCurColumnPos(17)
			C_BillingNo		=iCurColumnPos(18)
			C_BillingSeq	=iCurColumnPos(19)
			C_BillPostFlag	=iCurColumnPos(20)
			C_TaxNo			=iCurColumnPos(21)
			C_TaxSeq		=iCurColumnPos(22)
			C_TaxPostFlag	=iCurColumnPos(23)
			C_IssueDt		=iCurColumnPos(24)
			C_SaleBondDt	=iCurColumnPos(25)	
			C_TaxDt			=iCurColumnPos(26)
			C_StockUnit		=iCurColumnPos(27)
			C_IssueQty		=iCurColumnPos(28)
			C_IssueAmt		=iCurColumnPos(29)
			C_BillingQty	=iCurColumnPos(30)	
			C_BillingAmt	=iCurColumnPos(31)
			C_TaxQty		=iCurColumnPos(32)
			C_TaxAmt		=iCurColumnPos(33)

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
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
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
 
   If UniConvDateToYYYYMMDD(frm1.txtFromSaleDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToSaleDt.text, Parent.gDateFormat,"")Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If

   If UniConvDateToYYYYMMDD(frm1.txtFromTaxDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToTaxDt.text, Parent.gDateFormat,"")Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If
 
    
'    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
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
			strVal = strVal & "&txtFromSaleDt="     & Trim(.hFromSaleDt.Text)
			strVal = strVal & "&txtToSaleDt="       & Trim(.hToSaleDt.Text)
			strVal = strVal & "&txtFromTaxDt="      & Trim(.hFromTaxDt.Text)
			strVal = strVal & "&txtToTaxDt="        & Trim(.hToTaxDt.Text)
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
			strVal = strVal & "&txtFromSaleDt="     & Trim(.txtFromSaleDt.Text)
			strVal = strVal & "&txtToSaleDt="       & Trim(.txtToSaleDt.Text)
			strVal = strVal & "&txtFromTaxDt="      & Trim(.txtFromTaxDt.Text)
			strVal = strVal & "&txtToTaxDt="        & Trim(.txtToTaxDt.Text)
		END IF
			
		strVal = strVal & "&lgPageNo="			& lgPageNo								'Next key tag
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


Sub txtFromSaleDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromSaleDt.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtFromSaleDt.focus
    End If
End Sub

Sub txtToSaleDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToSaleDt.Action = 7
         Call SetFocusToDocument("M")
		frm1.txtToSaleDt.focus
    End If
End Sub

Sub txtFromTaxDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromTaxDt.Action = 7
         Call SetFocusToDocument("M")
		frm1.txtFromTaxDt.focus
    End If
End Sub

Sub txtToTaxDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToTaxDt.Action = 7
         Call SetFocusToDocument("M")
		frm1.txtToTaxDt.focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출/매출원가시점차이내역조회</font></td>
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
										<script language =javascript src='./js/gb013ma1_OBJECT1_txtYyyymm.js'></script>
									<TD CLASS="TD5" NOWRAP>사업부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBizUnitCd"  SIZE=10  ALT ="사업부" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(1)">
													<INPUT NAME="txtBizUnitNm"  SIZE=25  ALT ="사업부명" tag="14X"></TD>
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>Cost Center</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCostCd"  SIZE=10  ALT ="코스트센터" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(2)">
													<INPUT NAME="txtCostNm"  SIZE=25  ALT ="코스트센터명" tag="14X"></TD>


									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtBpCd" SIZE=10  tag="11XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(5)">
														<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14"></TD>
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
									<TD CLASS="TD5" NOWRAP>품목그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemGroupCd" SIZE=10  tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(8)">
														<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 tag="14"></TD>

									<TD CLASS="TD5" NOWRAP>판매유형</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtSoType" SIZE=10  tag="11XXXU" ALT="판매유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(6)">
														<INPUT TYPE=TEXT NAME="txtSoTypeNm" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>매출채권일</TD>																							    
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/gb013ma1_fpDateTime1_txtFromSaleDt.js'></script>&nbsp;~&nbsp;
                                        <script language =javascript src='./js/gb013ma1_fpDateTime2_txtToSaleDt.js'></script>
									</TD>

									<TD CLASS="TD5" NOWRAP>세금계산서일</TD>																							    
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/gb013ma1_fpDateTime3_txtFromTaxDt.js'></script>&nbsp;~&nbsp;
                                        <script language =javascript src='./js/gb013ma1_fpDateTime4_txtToTaxDt.js'></script>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/gb013ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
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
<INPUT TYPE=HIDDEN NAME="hFromSaleDt" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hToSaleDt" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hFromTaxDt" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hToTaxDt" tag="2x" TABINDEX= "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

	
