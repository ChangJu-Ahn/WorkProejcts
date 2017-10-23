<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :품목/오더별 실제원가조회 
'*  3. Program ID           : c4204ma1.asp
'*  4. Program Name         : 품목/오더별 실제원가 조회 
'*  5. Program Desc         : 품목/오더별 실제원가 조회 
'*  6. Modified date(First) : 2005-10-04
'*  7. Modified date(Last)  : 2005-10-04
'*  8. Modifier (First)     : HJO
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						3.    External File Include Part
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4204mb1.asp"                               'Biz Logic ASP
Const BIZ_PGM_ID2 = "c4204mb2.asp"                               'Biz Logic ASP

Dim iDBSYSDate
Dim iStrFromDt
Dim iStrToDt

iDBSYSDate = "<%=GetSvrDate%>"
iStrFromDt = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)	
iStrToDt= UNIDateAdd("m", -1,iStrFromDt, parent.gServerDateFormat)
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          
Dim lgCurrGrid
Dim lgCopyVersion
Dim lgErrRow, lgErrCol
Dim lgStrPrevKey2
Dim lgSTime		' -- 디버깅 타임체크 
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================
'--spread A
Dim C_PlantCd		
Dim C_CCCd			
Dim C_CCNm		
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_ProType
Dim C_Unit
Dim C_RcptQty
Dim C_TotCost
Dim C_MatCost
Dim C_LaborCost
Dim C_ExpCost
Dim C_SubCost
Dim C_UnitTotCost
Dim C_UnitMatCost
Dim C_UnitLaborCost
Dim C_UnitExpCost
Dim C_UnitSubCost
Dim C_SemiCost
Dim C_UnitSemiCost

'--spread B
Dim C_OrderNo2
Dim C_OderSeq2
Dim C_PoGu2
Dim C_CloseFlag2

Dim C_OrderQty2
Dim C_TotRcptQty2
Dim C_WipQty2
Dim C_WipAmt2
Dim C_RcptQty2

Dim C_TotCost2
Dim C_MatCost2
Dim C_LaborCost2
Dim C_ExpCost2
Dim C_SubCost2
Dim C_UnitTotCost2
Dim C_UnitMatCost2
Dim C_UnitLaborCost2
Dim C_UnitExpCost2
Dim C_UnitSubCost2
Dim C_SemiCost2
Dim C_UnitSemiCost2
'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(byVal pvSpd)
	If frm1.rdoTYPE1.checked Then
		If pvSpd="" or pvSpd="A" Then
			C_PlantCd		=1
			C_CCCd			=2
			C_CCNm			=3
			C_ItemAcct		=4
			C_ItemAcctNm	=5
			C_ItemCd		=6
			C_ItemNm		=7
			C_ProType		=8
			C_Unit			=9
			C_RcptQty		=10
			C_TotCost		=11
			C_MatCost		=12		
			C_LaborCost		=13
			C_ExpCost		=14
			C_SubCost		=15
			C_UnitTotCost	=16
			C_UnitMatCost	=17
			C_UnitLaborCost	=18
			C_UnitExpCost	=19
			C_UnitSubCost	=20
	
		End If
	
		If pvSpd="" or pvSpd="B" Then	
			C_OrderNo2		=1
			C_OderSeq2		=2
			C_PoGu2			=3
			C_CloseFlag2	=4

			C_OrderQty2		=5
			C_TotRcptQty2	=6
			C_WipQty2		=7
			C_WipAmt2		=8
			C_RcptQty2		=9

			C_TotCost2		=10
			C_MatCost2		=11
			C_LaborCost2	=12
			C_ExpCost2		=13
			C_SubCost2		=14
			C_UnitTotCost2	=15
			C_UnitMatCost2	=16
			C_UnitLaborCost2	=17
			C_UnitExpCost2	=18
			C_UnitSubCost2	=19
		End If
	Else
			If pvSpd="" or pvSpd="A" Then
			C_PlantCd		=1
			C_CCCd			=2
			C_CCNm			=3
			C_ItemAcct		=4
			C_ItemAcctNm	=5
			C_ItemCd		=6
			C_ItemNm		=7
			C_ProType		=8
			C_Unit			=9
			C_RcptQty		=10
			C_TotCost		=11
			C_MatCost		=12	
			C_SemiCost		=13
			C_LaborCost		=14
			C_ExpCost		=15
			C_SubCost		=16
			C_UnitTotCost	=17
			C_UnitMatCost	=18
			C_UnitSemiCost	=19
			C_UnitLaborCost	=20
			C_UnitExpCost	=21
			C_UnitSubCost	=22
	
		End If
	
		If pvSpd="" or pvSpd="B" Then	
			C_OrderNo2		=1
			C_OderSeq2		=2
			C_PoGu2			=3
			C_CloseFlag2	=4

			C_OrderQty2		=5
			C_TotRcptQty2	=6
			C_WipQty2		=7
			C_WipAmt2		=8
			C_RcptQty2		=9

			C_TotCost2		=10
			C_MatCost2		=11
			C_SemiCost2		=12
			C_LaborCost2	=13
			C_ExpCost2		=14
			C_SubCost2		=15
			C_UnitTotCost2	=16
			C_UnitMatCost2	=17
			C_UnitSemiCost2	=18
			C_UnitLaborCost2=19
			C_UnitExpCost2	=20
			C_UnitSubCost2	=21
		End If
	End If
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    
    lgStrPrevKey = ""
    lgStrPrevKey2 = ""		

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	frm1.txtSTART_DT.Text =UniConvDateAToB(iStrToDt, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtEND_DT.Text = UniConvDateAToB(iStrFromDt, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtTmp.value=""
	
	Call ggoOper.FormatDate(frm1.txtSTART_DT, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtEND_DT, parent.gDateFormat, 2)
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
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "MA") %>
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
Sub InitSpreadSheet(byVal pvSpd)
	Dim i, ret
	
	Call InitSpreadPosVariables(pvSpd)
    'Call AppendNumberPlace("6","3","0")
    
    If pvSpd = "" or pvSpd ="A" Then 
		With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		'ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		ggoSpread.Spreadinit "V20021106", , ""
			
		.MaxCols = C_UnitSubCost +1
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		.ReDraw = False
		
		ggoSpread.SSSetEdit		C_PlantCd,	"공장",				10,,,20,1	
		ggoSpread.SSSetEdit		C_CCCd,		"작업지시C/C",		10,,,20,1	
		ggoSpread.SSSetEdit		C_CCNm	,	"작업지시C/C명",	20
		ggoSpread.SSSetEdit		C_ItemAcct,	"품목계정",			10,,,20,1	
		ggoSpread.SSSetEdit		C_ItemAcctNm,	"품목계정명",	10
		ggoSpread.SSSetEdit		C_ItemCd,	"품목",				10,,,20,1	
		ggoSpread.SSSetEdit		C_ItemNm,	"품목명",			25
		ggoSpread.SSSetEdit		C_ProType,	"조달구분",			10
		ggoSpread.SSSetEdit		C_Unit,		"재고단위",			10		
		ggoSpread.SSSetFloat	C_RcptQty,	"입고수량",			10,		Parent.ggQtyNo,				ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_TotCost,	"제조원가",			10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		
		If frm1.rdoTYPE1.checked Then		
			ggoSpread.SSSetFloat	C_MatCost,		"재료비",			10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_LaborCost,	"노무비",			10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_ExpCost,		"경비",				10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_SubCost,		"외주가공비",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitTotCost,	"제조단가",			10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitMatCost,	"재료비단가",		10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitLaborCost,"노무비단가",		10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitExpCost,	"경비단가",			10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitSubCost,	"외주가공비단가",	10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Else
			ggoSpread.SSSetFloat	C_MatCost,		"원부재료비",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_SemiCost,		"(반)제품비",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_LaborCost,	"노무비",			10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_ExpCost,		"경비",				10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_SubCost,		"외주가공비",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitTotCost,	"제조단가",			10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitMatCost,	"원부재료비단가",	10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitSemiCost,	"(반)제품비단가",	10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitLaborCost,"노무비단가",		10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitExpCost,	"경비단가",			10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitSubCost,	"외주가공비단가",	10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		End If
			
	
		Call ggoSpread.SSSetColHidden(C_ItemAcct,C_ItemAcct,True)
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		'ggoSpread.SSSetSplit2(C_ItemNm) 
		.ReDraw = True		
		End With
		Call SetSpreadLock("A")
	End If
	
	If pvSpd = "" or pvSpd ="B" Then 
		With frm1.vspdData2
		
		ggoSpread.Source = frm1.vspdData2
'		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		ggoSpread.Spreadinit "V20021106", , ""
		
		.MaxCols = C_UnitSubCost2+1
		.MaxRows = 0
		
		Call GetSpreadColumnPos("B")
		.ReDraw = False
	
		ggoSpread.SSSetEdit		C_OrderNo2,		"오더번호",			15,,,20,1	
		ggoSpread.SSSetEdit		C_OderSeq2,		"오더SEQ",			10,,,20,1	
		ggoSpread.SSSetEdit		C_PoGu2	,		"사내/외주구분",	10
		ggoSpread.SSSetEdit		C_CloseFlag2,	"마감여부",			6
		
		ggoSpread.SSSetFloat	C_OrderQty2,	"오더수량",			10,		Parent.ggQtyNo,			ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_TotRcptQty2,	"누적입고수량",		10,		Parent.ggQtyNo,			ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_WipQty2,		"재공수량",			10,		Parent.ggQtyNo,			ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_WipAmt2,		"재공금액",			10,		Parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_RcptQty2,		"입고수량",			10,		Parent.ggQtyNo,			ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		If frm1.rdoTYPE1.checked Then
			ggoSpread.SSSetFloat	C_TotCost2,		"제조원가",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_MatCost2,		"재료비",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_LaborCost2,	"노무비",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_ExpCost2,		"경비",			10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_SubCost2,		"외주가공비",	10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitTotCost2,	"제조단가",		10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitMatCost2,	"재료비단가",	10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitLaborCost2,"노무비단가",	10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitExpCost2,	"경비단가",		10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitSubCost2,	"외주가공비단가", 10,	Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Else
			ggoSpread.SSSetFloat	C_TotCost2,		"제조원가",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_MatCost2,		"원부재료비",	10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_SemiCost2,	"(반)제품비",	10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			
			ggoSpread.SSSetFloat	C_LaborCost2,	"노무비",		10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_ExpCost2,		"경비",			10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_SubCost2,		"외주가공비",	10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitTotCost2,	"제조단가",		10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitMatCost2,	"재료비단가",	10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitSemiCost2,"(반)제품비단가",10,	Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			
			ggoSpread.SSSetFloat	C_UnitLaborCost2,"노무비단가",	10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitExpCost2,	"경비단가",		10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_UnitSubCost2,	"외주가공비단가",10,	Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		End If
	

		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		ggoSpread.SSSetSplit2(C_CloseFlag2) 
		.ReDraw = True
		End With
		Call SetSpreadLock("B")
	End If

	'ggoSpread.SpreadLockWithOddEvenRowColor()		

End Sub


'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock(byVal pvSpd)
	If pvSpd="A" Then
		ggoSpread.Source = frm1.vspdData    
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
	If pvSpd="B" Then
		ggoSpread.Source = frm1.vspdData2    
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False

	
    .vspdData.ReDraw = True
    
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	If frm1.rdoTYPE1.checked Then
		Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData 
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PlantCd		= iCurColumnPos(1)	
			C_CCCd			= iCurColumnPos(2)	
			C_CCNm			= iCurColumnPos(3)
			C_ItemAcct		= iCurColumnPos(4)
			C_ItemAcctNm	= iCurColumnPos(5)
			C_ItemCd		= iCurColumnPos(6)
			C_ItemNm		= iCurColumnPos(7)
			C_ProType		= iCurColumnPos(8)
			C_Unit			= iCurColumnPos(9)
			C_RcptQty		= iCurColumnPos(10)
			C_TotCost		= iCurColumnPos(11)
			C_MatCost		= iCurColumnPos(12)
			C_LaborCost		= iCurColumnPos(13)
			C_ExpCost		= iCurColumnPos(14)
			C_SubCost		= iCurColumnPos(15)
			C_UnitTotCost	= iCurColumnPos(16)
			C_UnitMatCost	= iCurColumnPos(17)
			C_UnitLaborCost	= iCurColumnPos(18)
			C_UnitExpCost	= iCurColumnPos(19)
			C_UnitSubCost	= iCurColumnPos(20)

		Case "B"
			ggoSpread.Source = frm1.vspdData2		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		
			
			C_OrderNo2		= iCurColumnPos(1)
			C_OderSeq2		= iCurColumnPos(2)
			C_PoGu2			= iCurColumnPos(3)
			C_CloseFlag2	= iCurColumnPos(4)

			C_OrderQty2		= iCurColumnPos(5)
			C_TotRcptQty2	= iCurColumnPos(6)
			C_WipQty2		= iCurColumnPos(7)
			C_WipAmt2		= iCurColumnPos(8)
			C_RcptQty2		= iCurColumnPos(9)

			C_TotCost2		= iCurColumnPos(10)
			C_MatCost2		= iCurColumnPos(11)
			C_LaborCost2	= iCurColumnPos(12)
			C_ExpCost2		= iCurColumnPos(13)
			C_SubCost2		= iCurColumnPos(14)
			C_UnitTotCost2	= iCurColumnPos(15)
			C_UnitMatCost2	= iCurColumnPos(16)
			C_UnitLaborCost2= iCurColumnPos(17)
			C_UnitExpCost2	= iCurColumnPos(18)				
		End Select	
	Else
			Select Case UCase(pvSpdNo)
			Case "A"
			ggoSpread.Source = frm1.vspdData 
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PlantCd		= iCurColumnPos(1)	
			C_CCCd			= iCurColumnPos(2)	
			C_CCNm			= iCurColumnPos(3)
			C_ItemAcct		= iCurColumnPos(4)
			C_ItemAcctNm	= iCurColumnPos(5)
			C_ItemCd		= iCurColumnPos(6)
			C_ItemNm		= iCurColumnPos(7)
			C_ProType		= iCurColumnPos(8)
			C_Unit			= iCurColumnPos(9)
			C_RcptQty		= iCurColumnPos(10)
			C_TotCost		= iCurColumnPos(11)
			C_MatCost		= iCurColumnPos(12)
			C_SemiCost		= iCurColumnPos(13)
			C_LaborCost		= iCurColumnPos(14)
			C_ExpCost		= iCurColumnPos(15)
			C_SubCost		= iCurColumnPos(16)
			C_UnitTotCost	= iCurColumnPos(17)
			C_UnitMatCost	= iCurColumnPos(18)
			C_UnitSemiCost	= iCurColumnPos(19)
			C_UnitLaborCost	= iCurColumnPos(20)
			C_UnitExpCost	= iCurColumnPos(21)
			C_UnitSubCost	= iCurColumnPos(22)

		Case "B"
			ggoSpread.Source	= frm1.vspdData2		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		
			
			C_OrderNo2		= iCurColumnPos(1)
			C_OderSeq2		= iCurColumnPos(2)
			C_PoGu2			= iCurColumnPos(3)
			C_CloseFlag2	= iCurColumnPos(4)

			C_OrderQty2		= iCurColumnPos(5)
			C_TotRcptQty2	= iCurColumnPos(6)
			C_WipQty2		= iCurColumnPos(7)
			C_WipAmt2		= iCurColumnPos(8)
			C_RcptQty2		= iCurColumnPos(9)

			C_TotCost2		= iCurColumnPos(10)
			C_MatCost2		= iCurColumnPos(11)
			C_SemiCost2		= iCurColumnPos(12)
			C_LaborCost2	= iCurColumnPos(13)
			C_ExpCost2		= iCurColumnPos(14)
			C_SubCost2		= iCurColumnPos(15)
			C_UnitTotCost2	= iCurColumnPos(16)
			C_UnitMatCost2	= iCurColumnPos(17)
			C_UnitSemiCost2	= iCurColumnPos(18)
			C_UnitLaborCost2= iCurColumnPos(19)
			C_UnitExpCost2	= iCurColumnPos(20)				
		End Select 
	End IF
End Sub
'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 ' -- 그리드1에서 팝업 클릭시 
Function OpenPopUp(Byval iWhere)
	Dim arrRet, sTmp
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	Select Case iWhere
		Case 0
			arrParam(0) = "공장 팝업"
			arrParam(1) = "dbo.B_PLANT"	
			arrParam(2) = Trim(.txtPLANT_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장" 

			arrField(0) = "PLANT_CD"	
			arrField(1) = "PLANT_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "공장"	
			arrHeader(1) = "공장명"
			arrHeader(2) = ""
			
		Case 1
			arrParam(0) = "품목계정 팝업"
			arrParam(1) = "dbo.B_MINOR"	
			arrParam(2) = Trim(.txtITEM_ACCT.value)
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1001", "''", "S")
			arrParam(5) = "품목계정" 

			arrField(0) ="ED10" & Parent.gColSep & "MINOR_CD"
			arrField(1) ="ED30" & Parent.gColSep & "MINOR_NM"		
			arrField(2) = ""	
    
			arrHeader(0) = "품목계정"
			arrHeader(1) = "품목계정명"
			arrHeader(2) = "C/C Level"	

		Case 2
			arrParam(0) = "품목 팝업"
			arrParam(1) = "dbo.B_ITEM"	
			arrParam(2) = Trim(.txtITEM_CD.value)
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "품목" 

			arrField(0) = "ED20" & Parent.gColSep &"ITEM_CD"	
			arrField(1) = "ED30" & Parent.gColSep &"ITEM_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "품목"	
			arrHeader(1) = "품목명"
			arrHeader(2) = ""
		Case 3
			arrParam(0) = "작업지시C/C 팝업"
			arrParam(1) = "b_cost_center"	
			arrParam(2) = Trim(.txtCC_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "작업지시C/C" 

			arrField(0) ="ED10" & Parent.gColSep &  "COST_CD"					' Field명(0)
			arrField(1) = "ED30" & Parent.gColSep & "COST_NM"					' Field명(1)
    
			arrHeader(0) = "작업지시C/C"
			arrHeader(1) = "작업지시C/C명"
		Case 4
			arrParam(0) = "작업단계 팝업"
			arrParam(1) = " ( select minor_cd,minor_nm from b_minor where major_cd = 'C4001') A"	
			arrParam(2) = Trim(.txtWORK_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "작업단계" 

			arrField(0) ="ED10" & Parent.gColSep &  "MINOR_CD"					' Field명(0)
			arrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"					' Field명(1)
    
			arrHeader(0) = "작업단계"
			arrHeader(1) = "작업단계명"

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

	End With
End Function


Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim sTmp
	
	With frm1
		Select Case iWhere		
			Case 0
				.txtPLANT_CD.value		= arrRet(0)
				.txtPLANT_NM.value		= arrRet(1)
				
			Case 1
				.txtITEM_ACCT.value		= arrRet(0)
				.txtITEM_ACCT_NM.value	= arrRet(1)
				
			Case 2
				.txtITEM_CD.value		= arrRet(0)
				.txtITEM_NM.value		= arrRet(1)
			Case 3
				.txtCC_CD.value			= arrRet(0)
				.txtCC_NM.value			= arrRet(1)
			Case 4
				.txtWORK_CD.value		= arrRet(0)
				.txtWORK_NM.value		= arrRet(1)				
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    
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
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtSTART_DT, parent.gDateFormat,2)
	Call ggoOper.FormatDate(frm1.txtEND_DT, parent.gDateFormat,2)

    Call InitSpreadSheet("")

    Call InitVariables

    Call SetDefaultVal
    Call SetToolbar("110000000001111")	
     If parent.gPlant <> "" Then
		frm1.txtPlant_Cd.value = UCase(parent.gPlant)
		frm1.txtPlant_Nm.value = parent.gPlantNm
		frm1.txtCC_Cd.focus		
	Else
		frm1.txtPlant_Cd.focus		
	End If
    
	Set gActiveElement = document.activeElement			
    
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
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================

'==========================================================================================
'   Event Desc : 배부규칙 설정확인 버튼 클릭시 
'==========================================================================================
Function BtnPrint(byval strPrintType)
	Dim StrUrl  , i,  StrUrl1, StrUrl2
	Dim strPlantCd, strCostCd, strItemCd	
	Dim sYear,sMon,sDay, sEndDt, sStartDt

	Dim intCnt,IntRetCD

	
    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
       Exit Function
    End If
    
	with frm1
	
		Call parent.ExtractDateFromSuper(.txtSTART_DT.Text, parent.gDateFormat,sYear,sMon,sDay)		
		sStartDt= sYear & sMon 
		Call parent.ExtractDateFromSuper(.txtEnd_DT.Text, parent.gDateFormat,sYear,sMon,sDay)
		sEndDt=  sYear & sMon 
	
	
	StrUrl = StrUrl & "Start_Dt|"			& sStartDt
	StrUrl = StrUrl & "|End_Dt|"			& sEndDt
	StrUrl = StrUrl & "|Work_Cd|"			& .txtwORK_CD.value
	
	StrUrl1 = StrUrl & "|Plant_Cd|"			& .txtPLANT_CD.value
	StrUrl1 = StrUrl1 & "|Cost_Cd|"			& .txtCC_CD.value
	StrUrl1 = StrUrl1 & "|Item_Acct|"		& .txtITEM_ACCT.value
	StrUrl1 = StrUrl1 & "|Item_Cd|"			& .txtITEM_CD.value 
	
	.vspdData.Row = .vspdData.ActiveRow : .vspdData.Col = C_PlantCd : strPlantCd = .vspdData.Text
	.vspdData.Row = .vspdData.ActiveRow : .vspdData.Col = C_CCCD : strCostCd = .vspdData.Text
	.vspdData.Row = .vspdData.ActiveRow : .vspdData.Col = C_ItemCd : strItemCd = .vspdData.Text

	StrUrl2 = StrUrl & "|Plant_Cd|"			& strPlantCd
	StrUrl2 = StrUrl2 & "|Cost_Cd|"			& strCostCd
	
	StrUrl2 = StrUrl2 & "|Item_Cd|"			& strItemCD
	
	If .rdoTYPE1.checked then
			If  strPrintType = "VIEW1" then
				ObjName = AskEBDocumentName("C4204MA1A", "ebr")
				Call FncEBRPreview(ObjName, StrUrl1)
			ElseIf strPrintType = "VIEW2" then
				ObjName = AskEBDocumentName("C4204MA1B", "ebr")
				Call FncEBRPreview(ObjName, StrUrl2)
			else
				'Call FncEBRPrint(EBAction,ObjName,StrUrl)
			End if	
		Else
			If  strPrintType = "VIEW1" then
				ObjName = AskEBDocumentName("C4204MA2A", "ebr")
				Call FncEBRPreview(ObjName, StrUrl1)
			ElseIf strPrintType = "VIEW2" then
				ObjName = AskEBDocumentName("C4204MA2B", "ebr")
				Call FncEBRPreview(ObjName, StrUrl2)
			else
				'Call FncEBRPrint(EBAction,ObjName,StrUrl)
			End if			
		End IF     
	End with
	
  
     
End Function 
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub txtSTART_DT_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtSTART_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtSTART_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtSTART_DT.Focus
    End If
End Sub

Sub txtEND_DT_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtEND_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtEND_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtEND_DT.Focus
    End If
End Sub

Sub txtPLANT_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtITEM_ACCT_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtITEM_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub rdoTYPE1_onClick()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	Call InitSpreadSheet("")
		'ggoSpread.Source = frm1.vspdData
'	ggoSpread.ClearSpreadData
End	Sub

Sub rdoTYPE2_onClick()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	Call InitSpreadSheet("")
End	Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    'ggoSpread.Source = frm1.vspdData
    'Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemCd
           .vspdData2.MaxRows = 0
        End With

        frm1.vspddata.Col = 0
		lgStrPrevKey2=""

		Call DbDtlQuery(NewRow)
    End If
End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정		
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData         

	lgStrPrevKey2=""
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	'msgbox lgStrPrevKey & "," & frm1.vspdData.MaxRows
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
	'If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey =frm1.vspdData.activerow Then
		'If CheckRunningBizProcess = True Then Exit Sub
	
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub


'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD , sStartDt, sEndDt
    
    FncQuery = False
    
    Err.Clear
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
	If CompareDateByFormat(frm1.txtSTART_DT.text,frm1.txtEND_DT.text,frm1.txtSTART_DT.Alt,frm1.txtEND_DT.Alt, _
		"970024",frm1.txtSTART_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
	frm1.txtSTART_DT.focus
	Exit Function
	End If
    
    If ChkKeyField=false then Exit Function 
    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
	Call InitVariables	

    IF DbQuery = False Then
		Exit Function
	END IF
       
    FncQuery = True		
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
  
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave() 
    
    FncSave = True      
    
End Function


'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 


End Function


Function FncCancel() 
    Dim lDelRows

	lgBlnFlgChgValue = True
End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD, iSeqNo, iSubSeqNo
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

End Function


Function FncDeleteRow() 
    Dim lDelRows
	
	lgBlnFlgChgValue = True
End Function
Function FncPrint()
    Call parent.FncPrint() 
End Function

Function FncPrev() 
End Function

Function FncNext() 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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
    Call InitSpreadSheet(gActiveSpdSheet.id)      
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
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
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

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	
    
    Dim sStartDt, sEndDt, sYear, sMon, sDay
    
    With frm1
		Call parent.ExtractDateFromSuper(.txtSTART_DT.Text, parent.gDateFormat,sYear,sMon,sDay)		
		sStartDt= sYear & sMon & sDay
		Call parent.ExtractDateFromSuper(.txtEnd_DT.Text, parent.gDateFormat,sYear,sMon,sDay)
		sEndDt= sYear & sMon & sDay

		
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtStartDt=" & sStartDt
		strVal = strVal & "&txtEndDt=" & sEndDt	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPLANT_CD=" & Trim(.txtPLANT_CD.value)
		strVal = strVal & "&txtITEM_ACCT=" & Trim(.txtITEM_ACCT.value)
		strVal = strVal & "&txtITEM_CD=" & Trim(.txtITEM_CD.value)
		strVal = strVal & "&txtCC_CD=" & Trim(.txtCC_CD.value)
		strVal = strVal & "&txtWork_CD=" & Trim(.txtWork_CD.value)
		
		If .rdoTYPE1.checked then
			strVal = strVal & "&rdoTYPE="	& .rdoTYPE1.value		
		Else
			strVal = strVal & "&rdoTYPE="	& .rdoTYPE2.value
		End If
		
		If lgStrPrevKey = "" Then Call InitSpreadSheet("")
	
	'	lgSTime = Time	' -- 디버깅용 
		Call RunMyBizASP(MyBizASP, strVal)
   
    End With
    
    DbQuery = True

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
	'Call DbDtlQuery(1)
	IF lgStrPrevKey<>"" Then
	call vspdData_ScriptLeaveCell( frm1.vspdData.ActiveCol,  lgStrPrevKey,  frm1.vspdData.ActiveCol,  frm1.vspdData.ActiveRow, "")
	else
	
	End If
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   

End Function
'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(byVal strRow) 

Dim strVal
Dim boolExist
Dim lngRows
Dim strPlantCd
Dim strCostCd
Dim strItemCd
Dim strGubun,sStartDt,sEndDt,i
Dim sYear,sMon,sDay

Dim tmpRow, tmpRow1,tmpRow2

    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

	tmpRow = frm1.txtTmp.value	
	tmpRow1 = split(tmpRow,parent.gColSep)
	tmpRow2=cdbl(strRow)
	
	For i=0 To ubound(tmpRow1)-1
		If cDbl(tmpRow1(i)) =tmpRow2 Then Exit Function
	Next

    With frm1

		.vspdData2.MaxRows = 0

	.vspdData.Row =strRow
	.vspdData.Col = C_PlantCd		:	strPlantCd = .vspdData.Text
	.vspdData.Col = C_CCCd		:	strCostCd = .vspdData.Text    
	.vspdData.Col = C_ItemCd		:	strItemCd = .vspdData.Text    
	
		If frm1.rdoType1.checked Then
			strGubun = frm1.rdoType1.value
		Else
			strGubun = frm1.rdoType2.value
		End If
		
		Call parent.ExtractDateFromSuper(.txtSTART_DT.Text, parent.gDateFormat,sYear,sMon,sDay)		
		sStartDt= sYear & sMon & sDay
		Call parent.ExtractDateFromSuper(.txtEnd_DT.Text, parent.gDateFormat,sYear,sMon,sDay)
		sEndDt= sYear & sMon & sDay
		
		DbDtlQuery = False   
    
		.vspdData.Row = strRow

		If LayerShowHide(1) = False Then Exit Function

		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001						'☜:				
		strVal = strVal & "&txtStartDt=" & sStartDt
		strVal = strVal & "&txtEndDt=" & sEndDt	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey2
		strVal = strVal & "&txtPLANT_CD=" & Trim(strPlantCd)
		strVal = strVal & "&txtITEM_CD=" & Trim(strItemCd)
		strVal = strVal & "&txtCC_CD=" & Trim(strCostCd)
		strVal = strVal & "&txtWork_CD=" & Trim(.txtWork_CD.value)
		strVal = strVal & "&rdoTYPE="	&strGubun

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With
    DbDtlQuery = True
End Function

Function DbDtlQueryOk()												'☆: 조회 성공후 실행로직 
	
	'-----------------------
    'Reset variables area
    '-----------------------
     ggoSpread.Source = frm1.vspdData2   
	lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	
End Function
'========================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : 소계 및 총계 색상변경 
'========================================================================================
Sub SetQuerySpreadColor(byVal arrStr)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt, strRowI

	With frm1.vspdData
	.ReDraw = False
	
	arrRow = Split(arrStr, Parent.gRowSep)
	
	iLoopCnt = UBound(arrRow, 1)

	For i = 0 to iLoopCnt -1
		arrCol = Split(arrRow(i), Parent.gColSep)
	
		.Col = -1
		.Row = CDbl(arrCol(2))	' -- 행 
	
		Select Case arrCol(0)
			Case "%1"
				iRow = .Row	: .Row2=.Row
				.Col = arrCol(1)+1  : .Col2=.MaxCols
				.BlockMode = True

				ret = .AddCellSpan(.Col,iRow, 9,1)
				.BackColor = RGB(250,250,210) 
				.ForeColor = vbBlack
				.BlockMode = False
			Case "%2"
				iRow = .Row :.Row2=.Row
				.Col = arrCol(1)+1 :.Col2=.MaxCols
				.BlockMode = True

				ret = .AddCellSpan(.Col,iRow, 8,1)
				.BackColor = RGB(204,255,153) 
				.ForeColor = vbBlack
				.BlockMode = False
			Case "%3"
				iRow = .Row : .Row2=.Row
				.Col = arrCol(1)  : .Col2 =.MaxCols
				.BlockMode = True

				ret = .AddCellSpan(.Col,iRow, 5,1)
				.BackColor = RGB(204,255,255) 
				.ForeColor = vbBlack
				.BlockMode = False
			Case "%4"  
				iRow = .Row :.Row2=.Row
				.Col = arrCol(1)+1 :.Col2=.MaxCols
				.BlockMode = True

				ret = .AddCellSpan(.Col,iRow, 4,1)
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
				.BlockMode =False
		End Select
		
		strRowI = strRowI & CDbl(arrCol(2)) & Parent.gColSep
	Next

	.Col = 1: .Row = -1: .ColMerge = 1
	.Col = 2: .Row = -1: .ColMerge = 1
	.Col = 3: .Row = -1: .ColMerge = 1
	.Col = 4: .Row = -1: .ColMerge = 1
	.Col = 5: .Row = -1: .ColMerge = 1

	frm1.txtTmp.value=frm1.txtTmp.value & strRowI
	.ReDraw = True
	End With

End Sub
'========================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : 소계 및 총계 색상변경 
'========================================================================================
Sub SetQuerySpreadColor2(Byval pGrpRow)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt

	With frm1.vspdData2	
	.ReDraw = False
	
	.Row=pGrpRow :	.Col = C_OrderNo2

	If .Text="합계" Then	
		.Col=-1
		.BackColor = RGB(250,250,153) 
		.ForeColor = vbBlack
	End If
	.ReDraw = True
	End With

End Sub

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave() 
    DbSave = True        
End Function

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()	
   
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		
'check plant
	If Trim(frm1.txtPLANT_CD.value) <> "" Then
		strWhere = " plant_cd= " & FilterVar(frm1.txtPLANT_CD.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm ","	b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPLANT_CD.alt,"X")
			frm1.txtPLANT_CD.focus 
			frm1.txtPLANT_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPLANT_NM.value = strDataNm(0)
	Else
		frm1.txtPLANT_NM.value=""
	End If
'check item_acct
	If Trim(frm1.txtITEM_ACCT.value) <> "" Then
		strWhere = " minor_cd  = " & FilterVar(frm1.txtITEM_ACCT.value, "''", "S") & " "		
		strWhere = strWhere & "		and major_cd=" & filterVar("P1001","","S")
		
		Call CommonQueryRs(" minor_nm  ","	b_minor ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtITEM_ACCT.alt,"X")
			frm1.txtITEM_ACCT.focus 
			frm1.txtITEM_ACCT_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtITEM_ACCT_NM.value = strDataNm(0)
	ELSE
		frm1.txtITEM_ACCT_NM.value=""
	End If
'check item
	If Trim(frm1.txtITEM_CD.value) <> "" Then
		strFrom = " B_ITEM "
		strWhere = " item_cd  = " & FilterVar(frm1.txtITEM_CD.value, "''", "S") & " "		
		
		Call CommonQueryRs(" item_nm  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtITEM_CD.alt,"X")
			frm1.txtITEM_CD.focus 
			frm1.txtITEM_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtITEM_NM.value = strDataNm(0)
	ELSE
		frm1.txtITEM_NM.value=""
	End If

'check CC
	If Trim(frm1.txtCC_CD.value) <> "" Then
		strFrom = " b_cost_center  "
		strWhere = " COST_CD  = " & FilterVar(frm1.txtCC_CD.value, "''", "S") & " "		
		
		Call CommonQueryRs(" COST_NM  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtCC_CD.alt,"X")
			frm1.txtCC_CD.focus 
			frm1.txtCC_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtCC_NM.value = strDataNm(0)
	ELSE
		frm1.txtCC_NM.value=""
	End If
	'check WC
	If Trim(frm1.txtWORK_CD.value) <> "" Then
		strFrom = " ( select minor_cd,minor_nm from b_minor where major_cd = 'C4001') A"	

		strWhere = " MINOR_CD  = " & FilterVar(frm1.txtWORK_CD.value, "''", "S") & " "		
		
		Call CommonQueryRs(" MINOR_NM  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtWORK_CD.alt,"X")
			frm1.txtWORK_CD.focus 
			frm1.txtWORK_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtWORK_NM.value = strDataNm(0)
	ELSE
		frm1.txtWORK_NM.value=""
	End If
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;&nbsp;</TD>
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
									<TD CLASS="TD5">작업년월</TD>
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSTART_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 작업년월" tag="12" id=txtSTART_DT></OBJECT>');</SCRIPT>&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtEND_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="종료 작업년월" tag="12" id=txtEND_DT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtPLANT_CD" TYPE="Text" MAXLENGTH="4" tag="15XXXU" size="10" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtPLANT_NM" TYPE="TEXT"  tag="14XXX" size="25">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5" NOWRAP>작업지시C/C</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtCC_CD" TYPE="Text" MAXLENGTH="10" tag="15XXXU" size="20" ALT="작업지시C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCC" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(3)">
									<input NAME="txtCC_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_ACCT" TYPE="Text" MAXLENGTH="2" tag="15XXXU" size="15" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(1)">
									<input NAME="txtITEM_ACCT_NM" TYPE="TEXT"  tag="14XXX" size="25">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_CD" TYPE="Text" MAXLENGTH="18" tag="15XXXU" size="20" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(2)">
									<input NAME="txtITEM_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5" NOWRAP>작업단계</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtWORK_CD" TYPE="Text" MAXLENGTH="2" tag="15XXXU" size="15" ALT="작업단계"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(4)">
									<input NAME="txtWORK_NM" TYPE="TEXT"  tag="14XXX" size="25">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">반제품원가분리여부</TD>
									<TD CLASS="TD6" valign=top><INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE1 tag="15XXX" VALUE="A" checked><LABEL FOR="rdoTYPE1">요소통합</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE2 tag="15XXX" VALUE="B"><LABEL FOR="rdoTYPE2">요소분리</LABEL>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<TD HEIGHT="60%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="*" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=10></TD>
					<TD WIDTH=10>&nbsp;<BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('VIEW1')" Flag=1 style="width: 150" disabled>미리보기(집계)</BUTTON></TD>
					<TD Width =10 > &nbsp;<BUTTON NAME="bttnPreview1"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('VIEW2')" Flag=1 style="width: 150" disabled>미리보기(오더별상세)</BUTTON></TD>
					<TD WIDTH=*></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtTmp" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

