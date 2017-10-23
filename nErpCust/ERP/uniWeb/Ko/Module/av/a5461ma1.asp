<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Accounting
*  2. Function Name        : 
*  3. Program ID           : a5432ma1
*  4. Program Name         : Verify AP
*  5. Program Desc         : 
*  6. Comproxy List        : None (coding with ADO )
*  7. Modified date(First) : 2003/06/13
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     :
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "A5461MB1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "A5461MB2.asp"	
'Const BIZ_PGM_ID2     = "A5442MB3.asp"	

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
'Const TAB3 = 3

'========================================================================================================
'                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------
Dim C_VatCd1
Dim C_VatNm1
Dim C_ItemLocAmt1
Dim C_NetLocAmt1
Dim C_GlInputCd1
Dim C_GlInputNm1
Dim C_IssuedDt1
Dim C_BpCd1
Dim C_BpNm1
Dim C_BizAreaCd1
Dim C_BizAreaNm1

'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #2
'--------------------------------------------------------------------------------------------------------
Dim C_VatCd2
Dim C_VatNm2
Dim C_ItemLocAmt2
Dim C_NetLocAmt2
Dim C_GlInputCd2
Dim C_GlInputNm2
Dim C_IssuedDt2
Dim C_BpCd2
Dim C_BpNm2
Dim C_BizAreaCd2
Dim C_BizAreaNm2

'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #3
'--------------------------------------------------------------------------------------------------------
Dim C_BpCd3
Dim C_BpNm3
Dim	C_VatCd3 
Dim	C_VatNm3		
Dim C_ItemLocAmt3
Dim C_NetLocAmt3
Dim C_GlInputCd3
Dim C_GlInputNm3
Dim C_GlDt3
Dim C_IssuedDt3
Dim C_BizAreaCd3
Dim C_BizAreaNm3
Dim C_ReBizAreaCd3
Dim C_ReBizAreaNm3
Dim C_VatBizAreaCd3
Dim C_VatBizAreaNm3


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  IsOpenPop          
Dim  lgRetFlag
Dim  gSelframeFlg

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			C_VatCd1		= 1
			C_VatNm1		= 2
			C_ItemLocAmt1	= 3
			C_NetLocAmt1	= 4
			C_GlInputCd1	= 5
			C_GlInputNm1	= 6
			C_IssuedDt1		= 7
			C_BpCd1			= 8
			C_BpNm1			= 9
			C_BizAreaCd1	= 10
			C_BizAreaNm1	= 11

		Case "B"
			C_VatCd2		= 1
			C_VatNm2		= 2
			C_ItemLocAmt2	= 3
			C_NetLocAmt2	= 4
			C_GlInputCd2	= 5
			C_GlInputNm2	= 6
			C_IssuedDt2		= 7
			C_BpCd2			= 8
			C_BpNm2			= 9
			C_BizAreaCd2	= 10
			C_BizAreaNm2	= 11

		Case "C"
			C_BpCd3			= 1
			C_BpNm3			= 2	
			C_VatCd3 		= 3
			C_VatNm3		= 4	
			C_ItemLocAmt3	= 5
			C_NetLocAmt3	= 6
			C_GlInputCd3	= 7
			C_GlInputNm3	= 8
			C_GlDt3			= 9
			C_IssuedDt3		= 10
			C_BizAreaCd3	= 11
			C_BizAreaNm3	= 12
			C_ReBizAreaCd3	= 13
			C_ReBizAreaNm3	= 14
			C_VatBizAreaCd3 = 15
			C_VatBizAreaNm3 = 16
	End Select 			
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False											'⊙: Indicates that no value changed
    lgStrPrevKey      = ""												'⊙: initializes Previous Key
    lgSortKey         = 1												'⊙: initializes sort direction
    lgPageNo		  = ""
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim StartDate, FirstDate, LastDate

	StartDate	= "<%=GetSvrDate%>"
	FirstDate	= UNIGetFirstDay(UNIDateAdd("m", -1, StartDate, parent.gServerDateFormat),Parent.gServerDateFormat)
	LastDate	= UNIGetLastDay(FirstDate , Parent.gServerDateFormat)
	frm1.txtFrDt.Text  = UniConvDateAToB(FirstDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtToDt.Text  = UniConvDateAToB(LastDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtFrDt2.Text  = UniConvDateAToB(FirstDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtToDt2.Text  = UniConvDateAToB(LastDate, Parent.gServerDateFormat, Parent.gDateFormat)

	frm1.txtShowDt.value = "N"
	frm1.txtShowBp.value = "N"
	frm1.txtShowBiz.value = "N"

	frm1.txtFrDt.focus 	
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q","*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Call initSpreadPosVariables(pvSpdNo)
	
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData1
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.Spreadinit "V20021227",, Parent.gAllowDragDropSpread
				.ReDraw = False
				.MaxCols   = C_BizAreaNm1 + 1                                                  ' ☜:☜: Add 1 to Maxcols
				.Col =.MaxCols
				.ColHidden = True
			   
				Call ggoSpread.ClearSpreadData()	
				Call GetSpreadColumnPos("A")

				                      'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
				ggoSpread.SSSetEdit    C_VatCd1			,"계산서유형"           ,9    ,,,20     ,2
				ggoSpread.SSSetEdit    C_VatNm1			,"계산서유형명"         ,13    ,3
				                      'ColumnPosition     Header            Width   Grp                    IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 
				ggoSpread.SSSetFloat   C_ItemLocAmt1	,"세액(자국)"			,14     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True
				ggoSpread.SSSetFloat   C_NetLocAmt1		,"공급가액(자국)"		,14     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True
				ggoSpread.SSSetEdit    C_GlInputCd1		,"입력경로"             ,8    ,,,10     ,2
				ggoSpread.SSSetEdit    C_GlInputNm1		,"입력경로명"           ,12    ,3
				ggoSpread.SSSetDate    C_IssuedDt1		,"계산서발행일자"		,12    ,2        ,Parent.gDateFormat   ,-1 
				ggoSpread.SSSetEdit    C_BpCd1			,"거래처"				,12    ,,,20     ,2
				ggoSpread.SSSetEdit    C_BpNm1			,"거래처명"				,15    ,3
				ggoSpread.SSSetEdit    C_BizAreaCd1		,"세금신고사업장"		,12    ,,,20     ,2
				ggoSpread.SSSetEdit    C_BizAreaNm1		,"세금신고사업장명"		,13    ,3

				call ggoSpread.MakePairsColumn(C_VatCd1,C_VatNm1)
				call ggoSpread.MakePairsColumn(C_GlInputCd1,C_GlInputNm1)
				call ggoSpread.MakePairsColumn(C_BpCd1,C_BpNm1)
				call ggoSpread.MakePairsColumn(C_BizAreaCd1,C_BizAreaNm1)
			   
				.ReDraw = True
	
				Call SetSpreadLock("A")
			End With
		Case "B"
			With frm1.vspdData2
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.Spreadinit "V20021227",, Parent.gAllowDragDropSpread
				.ReDraw = False
				.MaxCols   = C_BizAreaNm2 + 1                                                  ' ☜:☜: Add 1 to Maxcols
				.Col =.MaxCols
				.ColHidden = True
			   
				Call ggoSpread.ClearSpreadData()	
				Call GetSpreadColumnPos("B")

			                      'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
				ggoSpread.SSSetEdit    C_VatCd2			,"계산서유형"           ,9    ,,,20     ,2
				ggoSpread.SSSetEdit    C_VatNm2			,"계산서유형명"         ,13    ,3
				                      'ColumnPosition     Header            Width   Grp                    IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 
				ggoSpread.SSSetFloat   C_ItemLocAmt2	,"세액(자국)"			,14     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True
				ggoSpread.SSSetFloat   C_NetLocAmt2		,"공급가액(자국)"		,14     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True
				ggoSpread.SSSetEdit    C_GlInputCd2		,"입력경로"             ,8    ,,,10     ,2
				ggoSpread.SSSetEdit    C_GlInputNm2		,"입력경로명"           ,12    ,3
				ggoSpread.SSSetDate    C_IssuedDt2		,"계산서발행일자"		,12    ,2        ,Parent.gDateFormat   ,-1 
				ggoSpread.SSSetEdit    C_BpCd2			,"거래처"				,12    ,,,20     ,2
				ggoSpread.SSSetEdit    C_BpNm2			,"거래처명"				,15    ,3
				ggoSpread.SSSetEdit    C_BizAreaCd2		,"세금신고사업장"		,12    ,,,20     ,2
				ggoSpread.SSSetEdit    C_BizAreaNm2		,"세금신고사업장명"		,13    ,3

				call ggoSpread.MakePairsColumn(C_VatCd2,C_VatNm2)
				call ggoSpread.MakePairsColumn(C_GlInputCd2,C_GlInputNm2)
				call ggoSpread.MakePairsColumn(C_BpCd2,C_BpNm2)
				call ggoSpread.MakePairsColumn(C_BizAreaCd2,C_BizAreaNm2)
			   
				.ReDraw = True
	
				Call SetSpreadLock("B")
			End With			
		Case "C"
			With frm1.vspdData3
				ggoSpread.Source = frm1.vspdData3
				ggoSpread.Spreadinit "V20021227",, Parent.gAllowDragDropSpread
				.ReDraw = False
				.MaxCols = C_VatBizAreaNm3 + 1                                                  ' ☜:☜: Add 1 to Maxcols
				.Col =.MaxCols
				.ColHidden = True

				Call ggoSpread.ClearSpreadData()	
				Call GetSpreadColumnPos("C")
					                      'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
				ggoSpread.SSSetEdit    C_BpCd3			,"거래처"				,12    ,,,20     ,2
				ggoSpread.SSSetEdit    C_BpNm3			,"거래처명"				,15    ,3
				ggoSpread.SSSetEdit    C_VatCd3			,"계산서유형"           ,9    ,,,20     ,2
				ggoSpread.SSSetEdit    C_VatNm3			,"계산서유형명"         ,13    ,3
				                      'ColumnPosition     Header            Width   Grp                    IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 
				ggoSpread.SSSetFloat   C_ItemLocAmt3	,"세액(자국)"			,14     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True
				ggoSpread.SSSetFloat   C_NetLocAmt3		,"공급가액(자국)"		,14     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True
				ggoSpread.SSSetEdit    C_GlInputCd3		,"입력경로"             ,8    ,,,10     ,2
				ggoSpread.SSSetEdit    C_GlInputNm3		,"입력경로명"           ,12    ,3
				ggoSpread.SSSetDate    C_GlDt3			,"전표일자"				,12    ,2        ,Parent.gDateFormat   ,-1 
				ggoSpread.SSSetDate    C_IssuedDt3		,"계산서발행일자"		,12    ,2        ,Parent.gDateFormat   ,-1 
				ggoSpread.SSSetEdit    C_BizAreaCd3		,"사업장"				,13    ,,,10     ,2
				ggoSpread.SSSetEdit    C_BizAreaNm3		,"사업장명"				,15    ,3
				ggoSpread.SSSetEdit    C_ReBizAreaCd3	,"세금신고사업장"		,13    ,,,10     ,2
				ggoSpread.SSSetEdit    C_ReBizAreaNm3	,"세금신고사업장명"		,15    ,3
				ggoSpread.SSSetEdit    C_VatBizAreaCd3	,"부가세세금신고사업장"		,13    ,,,10     ,2
				ggoSpread.SSSetEdit    C_VatBizAreaNm3	,"부가세세금신고사업장명"		,15    ,3

				call ggoSpread.MakePairsColumn(C_VatCd3,C_VatNm3)
				call ggoSpread.MakePairsColumn(C_GlInputCd3,C_GlInputNm3)
				call ggoSpread.MakePairsColumn(C_BizAreaCd3,C_BizAreaNm3)
				call ggoSpread.MakePairsColumn(C_ReBizAreaCd3,C_ReBizAreaNm3)
				call ggoSpread.MakePairsColumn(C_VatBizAreaCd3,C_VatBizAreaNm3)
				call ggoSpread.MakePairsColumn(C_BpCd3,C_BpNm3)
				
				.ReDraw = True
	
				Call SetSpreadLock("C")
			End With			    
	End Select			
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData1
				.ReDraw = False 
					ggoSpread.Source = frm1.vspdData1
					ggoSpread.SpreadLockWithOddEvenRowColor()
				.ReDraw = True
			End With
		Case "B"
			With frm1.vspdData2
				.ReDraw = False 
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.SpreadLockWithOddEvenRowColor()
				.ReDraw = True
			End With
		Case "C"
			With frm1.vspdData3
				.ReDraw = False 
					ggoSpread.Source = frm1.vspdData3
					ggoSpread.SpreadLockWithOddEvenRowColor()
				.ReDraw = True
			End With								
	End Select
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_VatCd1		= iCurColumnPos(1)
			C_VatNm1		= iCurColumnPos(2)
			C_ItemLocAmt1	= iCurColumnPos(3)
			C_NetLocAmt1	= iCurColumnPos(4)
			C_GlInputCd1	= iCurColumnPos(5)
			C_GlInputNm1	= iCurColumnPos(6)
			C_IssuedDt1		= iCurColumnPos(7)
			C_BpCd1			= iCurColumnPos(8)
			C_BpNm1			= iCurColumnPos(9)
			C_BizAreaNm1	= iCurColumnPos(10)
			C_BizAreaNm1	= iCurColumnPos(11)

		Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_VatCd2		= iCurColumnPos(1)
			C_VatNm2		= iCurColumnPos(2)
			C_ItemLocAmt2	= iCurColumnPos(3)
			C_NetLocAmt2	= iCurColumnPos(4)
			C_GlInputCd2	= iCurColumnPos(5)
			C_GlInputNm2	= iCurColumnPos(6)
			C_IssuedDt1		= iCurColumnPos(7)
			C_BpCd2			= iCurColumnPos(8)
			C_BpNm2			= iCurColumnPos(9)
			C_BizAreaNm2	= iCurColumnPos(10)
			C_BizAreaNm2	= iCurColumnPos(11)

		Case "C"
            ggoSpread.Source = frm1.vspdData3
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_BpCd3			= iCurColumnPos(1)
			C_BpNm3			= iCurColumnPos(2)
			C_VatCd3 		= iCurColumnPos(3)
			C_VatNm3		= iCurColumnPos(4)	
			C_ItemLocAmt3	= iCurColumnPos(5)
			C_NetLocAmt3	= iCurColumnPos(6)
			C_GlInputCd3	= iCurColumnPos(7)
			C_GlInputNm3	= iCurColumnPos(8)
			C_GlDt3			= iCurColumnPos(9)
			C_IssuedDt3		= iCurColumnPos(10)
			C_BizAreaCd3	= iCurColumnPos(11)
			C_BizAreaNm3	= iCurColumnPos(12)
			C_ReBizAreaCd3	= iCurColumnPos(13)
			C_ReBizAreaNm3	= iCurColumnPos(14)
			C_VatBizAreaCd3 = iCurColumnPos(15)
			C_VatBizAreaNm3 = iCurColumnPos(16)
    End Select    
End Sub

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function 
 
	Select Case iWhere
		Case 0
			If frm1.txtTaxBizAreaCd.className = parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "사업장팝업"								' 팝업 명칭 
			arrParam(1) = "B_TAX_BIZ_AREA"									' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(4) = ""											' Where Condition
			arrParam(5) = "사업장"   
 
			arrField(0) = "TAX_Biz_AREA_CD"									' Field명(0)
			arrField(1) = "TAX_Biz_AREA_NM"									' Field명(1)    
			 
			arrHeader(0) = "사업장"									' Header명(0)
			arrHeader(1) = "사업장명"
		Case 1
			If frm1.txtBizAreaCd.className = parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "사업장팝업"								' 팝업 명칭 
			arrParam(1) = "B_Biz_AREA"									' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(4) = ""											' Where Condition
			arrParam(5) = "사업장"   
 
			arrField(0) = "Biz_AREA_CD"									' Field명(0)
			arrField(1) = "Biz_AREA_NM"									' Field명(1)    
			 
			arrHeader(0) = "사업장"									' Header명(0)
			arrHeader(1) = "사업장명"
		Case 2
			arrParam(0) = "거래처팝업"								' 팝업 명칭 
			arrParam(1) = "b_biz_partner"								' TABLE 명칭 
			arrParam(2) = strCode						 				' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(5) = "거래처"			
	
			arrField(0) = "BP_CD"										' Field명(0)
			arrField(1) = "BP_NM"										' Field명(1)
    
			arrHeader(0) = "거래처"									' Header명(0)
			arrHeader(1) = "거래처"									' Header명(1)
		Case 3
			arrParam(0) = "전표입력경로팝업"						' 팝업 명칭 
			arrParam(1) = "(SELECT GL_INPUT_TYPE FROM AV_GL_ITEM_VAT UNION SELECT GL_INPUT_TYPE FROM AV_VAT_INPUT) A " & vbcr & _
							"INNER JOIN B_MINOR B ON (B.MINOR_CD = A.GL_INPUT_TYPE AND B.MAJOR_CD=" & FilterVar("A1001", "''", "S") & " ) "
			arrParam(2) = strCode						 				' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = ""											' Where Condition
			arrParam(5) = "전표입력경로"			
			arrField(0) = "A.GL_INPUT_TYPE"									' Field명(0)
			arrField(1) = "B.MINOR_NM"									' Field명(1)
    
			arrHeader(0) = "전표입력경로"							' Header명(0)
			arrHeader(1) = "전표입력경로명"							' Header명(1)			
		Case 4
			arrParam(0) = "매입매출구분"						' 팝업 명칭 
			arrParam(1) = " b_minor b "					' TABLE 명칭 
			arrParam(2) = strCode						 				' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = " b.major_cd=" & FilterVar("a1003", "''", "S") & "  " ' Where Condition
			arrParam(5) = "매입매출코드"			
			arrField(0) = "b.minor_cd"									' Field명(0)
			arrField(1) = "b.minor_nm"									' Field명(1)
    
			arrHeader(0) = "매입매출코드"							' Header명(0)
			arrHeader(1) = "매입매출"							' Header명(1)			
		Case 5
			arrParam(0) = "계산서유형"						' 팝업 명칭 
			arrParam(1) = " b_minor b "					' TABLE 명칭 
			arrParam(2) = strCode						 				' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = " b.major_cd=" & FilterVar("B9001", "''", "S") & "  " ' Where Condition
			arrParam(5) = "계산서유형"			
			arrField(0) = "b.minor_cd"									' Field명(0)
			arrField(1) = "b.minor_nm"									' Field명(1)
    
			arrHeader(0) = "계산서유형"							' Header명(0)
			arrHeader(1) = "계산서유형명"							' Header명(1)			
	End Select    

	IsOpenPop = True
	 
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")   

	IsOpenPop = False
 
	If arrRet(0) = "" Then     
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function			

'======================================================================================================
'	Name : EscPopup()
'	Description : Dept Popup에서 Return되는 값 setting
'======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
		    Case "0"
				.txtTaxBizAreaCd.focus
		    Case "1"
				.txtBizAreaCd.focus
			Case "2"
				.txtBpCd.focus
			Case "3"
				.txtGlInputCd.focus
			Case "4"
				.txtVatIoFg.focus
			Case "5"
				.txtVatTypeCd.focus
	    End Select
	End With
End Function

'======================================================================================================
'	Name : SetPopup()
'	Description : Dept Popup에서 Return되는 값 setting
'======================================================================================================
Function SetPopup(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case "0"
				.txtTaxBizAreaCd.value     = arrRet(0)
				.txtTaxBizAreaNM.value     = arrRet(1)
				.txtTaxBizAreaCd.focus
			Case "1"
				.txtBizAreaCd.value     = arrRet(0)
				.txtBizAreaNM.value     = arrRet(1)
				.txtBizAreaCd.focus
			Case "2"
				.txtBpCd.value      = arrRet(0)
				.txtBpNm.value      = arrRet(1)
				.txtBpCd.focus
			Case "3"
				.txtGlInputCd.value   = arrRet(0)
				.txtGlInputNm.value = arrRet(1)
				.txtGlInputCd.focus				
			Case "4"
				.txtVatIoFg.value     = arrRet(0)
				.txtVatIoNm.value     = arrRet(1)
				.txtVatIoFg.focus
			Case "5"
				.txtVatTypeCd.value     = arrRet(0)
				.txtVatTypeNm.value     = arrRet(1)
				.txtVatTypeCd.focus
	    End Select
	End With
End Function     


'================================================================
'상환번호 참조 팝업 
'================================================================
Function OpenPopupVat()

	Dim arrRet
	Dim arrParam(0)
	Dim iCalledAspName
	Dim iColSep
	iColSep = Parent.gColSep

	iCalledAspName = AskPRAspName("A5461RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5461RA1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function
	If frm1.vspdData1.Maxrows <= 0 and frm1.vspdData2.Maxrows <= 0 Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtFrDt.text) & iColSep
	arrParam(0) = arrParam(0) & Trim(frm1.txtToDt.text) & iColSep
	arrParam(0) = arrParam(0) & Trim(frm1.txtVatIoFg.value) & iColSep
	arrParam(0) = arrParam(0) & Trim(frm1.txtVatIoNm.value) & iColSep

	If gMouseClickStatus <> "SPC" and gMouseClickStatus <> "SP1C" Then		Exit Function
	
	if gMouseClickStatus <> "SPC" and frm1.vspdData1.Maxrows > 0 then     'JJ
			With frm1.vspdData1
				.Row = .ActiveRow
				.Col = C_VatCd1		:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_VatNm1		:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_GlInputCd1	:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_GlInputNm1	:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_IssuedDt1	:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_BpCd1		:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_BpNm1		:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_BizAreaCd1	:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_BizAreaNm1	:	arrParam(0) = arrParam(0) & Trim(.Text)
			End With
	ElseIf frm1.vspdData2.Maxrows > 0 then
			With frm1.vspdData2
				.Row = .ActiveRow
				.Col = C_VatCd2		:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_VatNm2		:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_GlInputCd2	:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_GlInputNm2	:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_IssuedDt2	:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_BpCd2		:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_BpNm2		:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_BizAreaCd2	:	arrParam(0) = arrParam(0) & Trim(.Text) & iColSep
				.Col = C_BizAreaNm2	:	arrParam(0) = arrParam(0) & Trim(.Text)
			end with

	End IF
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=900px; dialogHeight=500px; center: Yes; help: No; resizable: YES; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0,1) = ""  Then
		frm1.txtFrDt.focus
		Exit Function
	Else
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		frm1.txtFrDt.focus
	End If

End Function


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
    Err.Clear																						'☜: Clear err status
    
	Call LoadInfTB19029																				'☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")															'⊙: Lock Field
             
	Call InitVariables
    Call SetDefaultVal
    
    Call InitSpreadSheet("A")																		'Setup the Spread sheet
	Call InitSpreadSheet("B")
	Call InitSpreadSheet("C")
	
	Call CurFormatNumericOCX()

	Call ggoOper.SetReqAttr(frm1.txtbpCd, "Q")
	Call ggoOper.SetReqAttr(frm1.txtTaxBizAreaCd, "Q")

	Call ClickTab1()
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    On Error Resume Next																			'☜: If process fails
    Err.Clear																						'☜: Clear error status

    FncQuery = False	
    Call InitVariables																				'☜: Processing is NG

    If Not chkField(Document, "1") Then																'⊙: This function check indispensable field
		Exit Function
	End If

	Select Case gSelframeFlg
		Case TAB1
			ggoSpread.Source = Frm1.vspdData1
			Call ggoSpread.ClearSpreadData()
			ggoSpread.Source = Frm1.vspdData2
			Call ggoSpread.ClearSpreadData()

			If DbQuery("A") = False Then															'☜: Query db data
				Exit Function
			End If

		Case TAB2
			ggoSpread.Source = Frm1.vspdData3
			Call ggoSpread.ClearSpreadData()
    
			If DbQuery("C") = False Then															'☜: Query db data
				Exit Function
			End If
	End Select

    If Err.number = 0 Then
		FncQuery = True																				'☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   
End Function
	
'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False	                                                          '☜: Processing is NG

	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '☜: Processing is OK
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
    
    Call ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next																	'☜: If process fails
    Err.Clear																				'☜: Clear error status

    FncExit = False																			'☜: Processing is NG

    If Err.number = 0 Then
       FncExit = True																		'☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   
End Function

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
    
	Select Case UCase(gActiveSpdSheet.Name)
		Case "VSPDDATA1"
			Call InitSpreadSheet("A")      
		Case "VSPDDATA2"
			Call InitSpreadSheet("B")      		
		Case "VSPDDATA3"
			Call InitSpreadSheet("C")      		
	End Select	

	Call ggoSpread.ReOrderingSpreadData()
End Sub

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
	Dim strVal, txtRdoType
	
    On Error Resume Next																		'☜: If process fails
    Err.Clear																					'☜: Clear error status
 
    DbQuery = False																				'☜: Processing is NG

    Call DisableToolBar(Parent.TBC_QUERY)														'☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)																		'☜: Show Processing Message
	
	Select Case pDirect
		Case "A"
			strVal = BIZ_PGM_ID		& "?txtMode="		& Parent.UID_M0001                      '☜: Query
		    strVal = strVal			& "&txtTaxBizAreaCd="	& frm1.txtTaxBizAreaCd.value							'☜:
			strVal = strVal			& "&txtFrDt="		& frm1.txtFrDt.text								'☜:
			strVal = strVal			& "&txtToDt="		& frm1.txtToDt.text								'☜:
			strVal = strVal			& "&txtMaxRows1="	& Frm1.vspdData1.MaxRows				'☜: Max fetched data
			strVal = strVal			& "&txtMaxRows2="	& Frm1.vspdData2.MaxRows				'☜: Max fetched data
		Case "C"
			strVal = BIZ_PGM_ID1	& "?txtMode="		& Parent.UID_M0001						'☜: Query
			strVal = strVal			& "&DispMeth="		& Trim(frm1.RdoDiff.checked)
		    strVal = strVal			& "&txtBizAreaCd="	& frm1.txtBizAreaCd.value							'☜:
		    strVal = strVal			& "&txtTaxBizAreaCd="	& frm1.txtTaxBizAreaCd.value							'☜:
			strVal = strVal			& "&txtFrDt="		& frm1.txtFrDt2.text								'☜:
			strVal = strVal			& "&txtToDt="		& frm1.txtToDt2.text								'☜:
			strVal = strVal			& "&txtGlFrDt="		& frm1.txtGlFrDt.text								'☜:
			strVal = strVal			& "&txtGlToDt="		& frm1.txtGlToDt.text								'☜:
			strVal = strVal			& "&txtMaxRows="	& Frm1.vspdData3.MaxRows				'☜: Max fetched data
			strVal = strVal			& "&lgPageNo="		& lgPageNo        
	End Select 
	
	strVal = strVal		& "&txtGlInputCd="	& frm1.txtGlInputCd.value
    strVal = strVal		& "&txtShowDt="		& frm1.txtShowDt.value							'☜:
    strVal = strVal		& "&txtVatIoFg="	& frm1.txtVatIoFg.value							'☜:
    strVal = strVal		& "&txtVatTypeCd="	& frm1.txtVatTypeCd.value							'☜:
    strVal = strVal     & "&txtBpCd="		& frm1.txtbpCd.value                           '☜:
    strVal = strVal     & "&txtShowBp="		& frm1.txtShowbp.value                             '☜:
    strVal = strVal     & "&txtShowBiz="	& frm1.txtShowBiz.value

    Call RunMyBizASP(MyBizASP, strVal)															'☜:  Run biz logic

    If Err.number = 0 Then
       DbQuery = True																			'☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
    On Error Resume Next																		'☜: If process fails
    Err.Clear																					'☜: Clear error status

	lgIntFlgMode      = Parent.OPMD_UMODE														'⊙: Indicates that current mode is Create mode

    Select Case gSelframeFlg
		Case TAB1 
			Frm1.vspdData1.focus
		Case TAB2
			Frm1.vspdData3.focus
	End Select
	
	Call CurFormatNumericOCX()

	Call SetToolbar("1100000000001111")															'☆: Developer must customize
    
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
' Name : CurFormatNumericOCX
' Desc : 
'========================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtVatLocAmt1, parent.gCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtGlLocAmt1,  parent.gCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtGlLocAmt2,  parent.gCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
'========================================================================================================
'   Event Name : txtBizAreaCd_onChange
'   Event Desc : 
'========================================================================================================
Sub txtBizAreaCd_onChange()
	Dim IntRetCD
	Dim arrVal

	If frm1.txtBizAreaCd.value = "" Then frm1.txtBizAreaNm.value = "":	Exit Sub

	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(frm1.txtBizAreaCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBizAreaNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtBizAreaCd.alt,"X")  	
		frm1.txtBizAreaCd.value = ""
		frm1.txtBizAreaNm.value = ""
		frm1.txtBizAreaCd.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtBizAreaCd_onChange
'   Event Desc : 
'========================================================================================================
Sub txtTaxBizAreaCd_onChange()
	Dim IntRetCD
	Dim arrVal

	If frm1.txtTaxBizAreaCd.value = "" Then frm1.txtTaxBizAreaNm.value = "" : Exit Sub

	If CommonQueryRs("TAX_BIZ_AREA_NM", "B_TAX_BIZ_AREA ", " TAX_BIZ_AREA_CD=  " & FilterVar(frm1.txtTaxBizAreaCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtTaxBizAreaNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtTaxBizAreaCd.alt,"X")  	
		frm1.txtTaxBizAreaCd.value = ""
		frm1.txtTaxBizAreaNm.value = ""
		frm1.txtTaxBizAreaCd.focus
	End If
End Sub
'========================================================================================================
'   Event Name : txtVatTypeCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtVatTypeCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtVatTypeCd.value = "" Then frm1.txtVatTypeNm.value = "" : Exit Sub
	
	If CommonQueryRs("MINOR_NM", "B_MINOR ", "MAJOR_CD = " & FilterVar("B9001", "''", "S") & "  AND MINOR_CD=  " & FilterVar(frm1.txtVatTypeCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtVatTypeNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtVatTypeCd.alt,"X")  	
		frm1.txtVatTypeCd.value = ""
		frm1.txtVatTypeNm.value = ""
		frm1.txtVatTypeCd.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtVatIoFg_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtVatIoFg_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtVatIoFg.value = "" Then frm1.txtVatIoNm.value = "": Exit Sub
	
	If CommonQueryRs("MINOR_NM", "B_MINOR ", "MAJOR_CD = " & FilterVar("A1003", "''", "S") & "  AND MINOR_CD=  " & FilterVar(frm1.txtVatIoFg.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtVatIoNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtVatIoFg.alt,"X")  	
		frm1.txtVatIoFg.value = ""
		frm1.txtVatIoNm.value = ""
		frm1.txtVatIoFg.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtGlInputCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtGlInputCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtGlInputCd.value = "" Then frm1.txtGlInputNm.value = "" :Exit Sub
	
	If CommonQueryRs("MINOR_NM", "B_MINOR ", "MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  AND MINOR_CD=  " & FilterVar(frm1.txtGlInputCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtGlInputNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtGlInputCd.alt,"X")  	
		frm1.txtGlInputCd.value = ""
		frm1.txtGlInputNm.value = ""
		frm1.txtGlInputCd.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtbpCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtBpCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtBpCd.value = "" Then	frm1.txtBpNm.value = "" : Exit Sub	
		
	If CommonQueryRs("BP_NM", "B_BIZ_PARTNER ", " BP_CD=  " & FilterVar(frm1.txtBpCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBpNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("126100","X","X","X")  	
		frm1.txtBpCd.value = ""
		frm1.txtBpNm.value = ""
		frm1.txtBpCd.focus
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1111111111")    
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1111111111")    
    gMouseClickStatus = "SP1C"   
    Set gActiveSpdSheet = frm1.vspdData2

    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData3_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData3_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1111111111")    
    gMouseClickStatus = "SP2C"   
    Set gActiveSpdSheet = frm1.vspdData3

    If frm1.vspdData3.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData3
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub


'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData3_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then

    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then

    End If
End Sub

'========================================================================================================
'   Event Name : vspdData3_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData3_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then

    End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = Frm1.vspdData1
End Sub

'========================================================================================================
'   Event Name : vspdData2_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
End Sub

'========================================================================================================
'   Event Name : vspdData3_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData3_GotFocus()
    ggoSpread.Source = Frm1.vspdData3
End Sub

'========================================================================================================
'   Event Name : vspdData1_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
'   Event Name : vspdData2_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
End Sub    

'========================================================================================================
'   Event Name : vspdData3_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData3_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
End Sub    

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
  
'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub  

'========================================================================================================
'   Event Name : vspdData3_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData3_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("C")
End Sub  
  
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData3_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
    
    If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then	           
    	If lgPageNo <> "" Then                         
           If DbQuery("C") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub

'========================================================================================================
' Name : txtFrDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtFrDt_DblClick(Button)
    If Button = 1 Then
		frm1.txtFrDt.Action = 7                                    ' 7 : Popup Calendar ocx
		Call SetFocusToDocument("M")	
		frm1.txtFrDt.Focus
    End If
End Sub

'========================================================================================================
' Name : txtToDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
		frm1.txtToDt.Action = 7                                    ' 7 : Popup Calendar ocx
		Call SetFocusToDocument("M")	
		frm1.txtToDt.Focus
    End If
End Sub

'========================================================================================================
' Name : txtFrDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtFrDt2_DblClick(Button)
    If Button = 1 Then
		frm1.txtFrDt2.Action = 7                                    ' 7 : Popup Calendar ocx
		Call SetFocusToDocument("M")	
		frm1.txtFrDt2.Focus
    End If
End Sub

'========================================================================================================
' Name : txtToDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtToDt2_DblClick(Button)
    If Button = 1 Then
		frm1.txtToDt2.Action = 7                                    ' 7 : Popup Calendar ocx
		Call SetFocusToDocument("M")	
		frm1.txtToDt2.Focus
    End If
End Sub
'========================================================================================================
' Name : txtGlFrDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtGlFrDt_DblClick(Button)
    If Button = 1 Then
		frm1.txtGlFrDt.Action = 7                                    ' 7 : Popup Calendar ocx
		Call SetFocusToDocument("M")	
		frm1.txtGlFrDt.Focus
    End If
End Sub

'========================================================================================================
' Name : txtToDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtGlToDt_DblClick(Button)
    If Button = 1 Then
		frm1.txtGlToDt.Action = 7                                    ' 7 : Popup Calendar ocx
		Call SetFocusToDocument("M")	
		frm1.txtGlToDt.Focus
    End If
End Sub



Sub txtFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtFrDt.Focus
	   Call MainQuery
	End If   
End Sub

Sub txtToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToDt.Focus
	   Call MainQuery
	End If   
End Sub

Sub txtFrDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtFrDt2.Focus
	   Call MainQuery
	End If   
End Sub

Sub txtToDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToDt2.Focus
	   Call MainQuery
	End If   
End Sub


Sub txtGlFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtGlFrDt.Focus
	   Call MainQuery
	End If   
End Sub

Sub txtGlToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtGlToDt.Focus
	   Call MainQuery
	End If   
End Sub
'========================================================================================================
' Name : chkShowBp_onchange
' Desc : 
'========================================================================================================
Sub chkShowBp_onchange()
	If frm1.chkShowBp.checked = True Then
		frm1.txtShowBp.value = "Y"
		Call ggoOper.SetReqAttr(frm1.txtbpCd, "D")
	Else
		frm1.txtShowBp.value = "N"	
		frm1.txtbpCd.value = ""
		Call ggoOper.SetReqAttr(frm1.txtbpCd, "Q")		
	End If
End Sub

Sub chkShowBiz_onchange()
	If frm1.chkShowBiz.checked = True Then
		frm1.txtShowBiz.value = "Y"
		Call ggoOper.SetReqAttr(frm1.txtTaxBizAreaCd, "D")
	Else
		frm1.txtShowBiz.value = "N"	
		frm1.txtTaxBizAreaNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtTaxBizAreaCd, "Q")		
	End If
End Sub


'========================================================================================================
' Name : chkShowBp_onchange
' Desc : 
'========================================================================================================
Sub chkShowDt_onchange()
	If frm1.chkShowDt.checked = True Then
		frm1.txtShowDt.value = "Y"
	Else
		frm1.txtShowDt.value = "N"	
	End If
End Sub

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	gMouseClickStatus = "SPC"

	frm1.chkShowBp.style.display = ""
	frm1.chkShowBiz.style.display = ""

	txtDate1.style.display = ""
	txtDate2.style.display = "NONE"
	DifTotal.style.DISPLAY = "NONE"
	BizArea1.style.DISPLAY = "NONE"
	VATREF1.style.display = ""
	VATREF2.style.display = "NONE"
	Call ggoOper.SetReqAttr(frm1.txtFrDt, "N")
	Call ggoOper.SetReqAttr(frm1.txtToDt, "N")
	If frm1.chkShowBp.checked = false THen
		frm1.txtBpCd.value = ""
		frm1.txtBpNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtBpCd, "Q")
	End If
	If frm1.chkShowBiz.checked = false THen
		frm1.txtTaxBizAreaCd.value = ""
		frm1.txtTaxBizAreaNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtTaxBizAreaCd, "Q")
	End If

	Call SetToolbar("1100000000001111") 				 
End Function

Function ClickTab2()
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2

	frm1.chkShowBp.style.display = "NONE"
	frm1.chkShowBiz.style.display = "NONE"

	txtDate1.style.display = "NONE"
	txtDate2.style.display = ""
	DifTotal.style.DISPLAY = ""
	BizArea1.style.DISPLAY = ""
	VATREF1.style.display = "NONE"
	VATREF2.style.display = ""
	VATREF2.disabled = TRUE
	Call ggoOper.SetReqAttr(frm1.txtFrDt, "D")
	Call ggoOper.SetReqAttr(frm1.txtToDt, "D")
	Call ggoOper.SetReqAttr(frm1.txtBpCd, "D")
	Call ggoOper.SetReqAttr(frm1.txtTaxBizAreaCd, "D")
	
	Call SetToolbar("1100000000001111") 
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><IMG height=23 src="../../image/table/seltab_up_left.gif" width=9></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>부가세현황조회</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><IMG height=23 src="../../image/table/tab_up_left.gif" width=9></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>전표부가세확인</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
		    		<TD WIDTH=* align=right ID=VATREF1><A HREF="VBSCRIPT:OpenPopupVat()">내역참조</a></TD>
		    		<TD WIDTH=* align=right ID=VATREF2>내역참조</TD>
					<TD WIDTH=10></TD>
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
								<TD CLASS="TD5" NOWRAP>계산서발행일자</TD>
								<TD CLASS="TD6" NOWRAP ID=txtDate1>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpFrDt name=txtFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="시작일자" ALIGN=TOP></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpToDt name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="종료일자" ALIGN=TOP></OBJECT>');</SCRIPT> &nbsp; &nbsp;
							        <INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkShowDt ID=chkShowDt tag="11" onclick=chkShowDt_onchange()>
								</TD>
								<TD CLASS="TD6" NOWRAP ID=txtDate2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpFrDt2 name=txtFrDt2 CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="시작일자" ALIGN=TOP></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpToDt2 name=txtToDt2 CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="종료일자" ALIGN=TOP></OBJECT>');</SCRIPT> &nbsp; &nbsp;
								</TD>
								<TD CLASS=TD5 NOWRAP>매입매출구분</TD>
								<TD CLASS=TD6 nowrap>
									<INPUT TYPE=TEXT NAME=txtVatIoFg ALT="매입매출구분" SIZE=10 MAXLENGTH=20 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnType ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtVatIoFg.value, 4)">
									<INPUT TYPE=TEXT NAME=txtVatIoNm ALT="매입매출구분" SIZE=18  tag="14" >
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=TEXT NAME=txtGlInputCd ALT="전표입력경로" SIZE=10 MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtGlInputCd.value, 3)">
									<INPUT TYPE=TEXT NAME=txtGlInputNm ALT="전표입력경로명" SIZE="18"  tag="14" >
								</TD>
								<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
								<TD CLASS=TD6 nowrap>
									<INPUT TYPE=TEXT NAME=txtTaxBizAreaCd ALT="부가세세금신고사업장" SIZE=10 MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtTaxBizAreaCd.value, 0)">
									<INPUT TYPE=TEXT NAME=txtTaxBizAreaNm ALT="부가세세금신고사업장명" SIZE="18"  tag="14" >
									<INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkShowBiz ID=chkShowBiz  tag="11" onclick=chkShowBiz_onchange()>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계산서유형</TD>
								<TD CLASS=TD6 nowrap>
									<INPUT TYPE=TEXT NAME=txtVatTypeCd ALT="계산서유형" SIZE=10 MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtVatTypeCd.value, 5)">
									<INPUT TYPE=TEXT NAME=txtVatTypeNm ALT="계산서유형명" SIZE="18" tag="14" >
								</TD>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 nowrap>
									<INPUT TYPE=TEXT NAME=txtBpCd ALT="거래처코드" SIZE=10 MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif"  NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtBpCd.value, 2)">
									<INPUT TYPE=TEXT NAME=txtBpNm ALT="거래처명" SIZE="18"  tag="14">
									<INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkShowBp ID=chkShowBp  tag="11" onclick=chkShowBp_onchange()>
								</TD>
							</TR>
							<TR ID=DifTotal>
								<TD CLASS="TD5" NOWRAP>전표일자</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpGLFrDt name=txtGLFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="시작일자" ALIGN=TOP></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpGLToDt name=txtGLToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="종료일자" ALIGN=TOP></OBJECT>');</SCRIPT>							</TD>
								<TD CLASS=TD5 NOWRAP>사업장</TD>
								<TD CLASS=TD6 nowrap>
									<INPUT TYPE=TEXT NAME=txtBizAreaCd ALT="사업장코드" SIZE=10 MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtBizAreaCd.value, 1)">
									<INPUT TYPE=TEXT NAME=txtBizAreaNm ALT="사업장명" SIZE="18"  tag="14" >
								</TD>
							</TR>
							<TR ID=BizArea1>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 nowrap></TD>
								<TD CLASS=TD5 NOWRAP>표시구분</TD>
								<TD CLASS=TD6 nowrap><INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoType" ID="RdoDiff" VALUE="D" TAG="11" ><LABEL FOR="RdoDiff" Id="RdoDiff">차이분</LABEL>&nbsp;&nbsp
									<INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoType" ID="RdoTotal" VALUE="T" TAG="11" Checked><LABEL FOR="RdoTotal" Id="RdoTotal">전체</LABEL></TD>
							</TR>
						</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=*>
					<TD WIDTH=100% VALIGN=TOP>
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD colspan="2">
										<FIELDSET CLASS="CLSFLD">
											<TABLE <%=LR_SPACE_TYPE_20%>>
												<TR>
													<TD CLASS="TD5" NOWRAP>부가세합계금액</TD>
													<TD CLASS=TD6 NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtVatLocAmt1" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="부가세합계금액" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
													<TD CLASS="TD5" NOWRAP>전표합계금액</TD>
													<TD CLASS=TD6 NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtGlLocAmt1" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="전표합계금액" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
												</TR>
											</TABLE>
										</FIELDSET>
									</TD>
								</TR>
								<TR>
									<TD HEIGHT="100%" WIDTH="50%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
									<TD HEIGHT="100%" WIDTH="50%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</DIV>		
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD>
										<FIELDSET CLASS="CLSFLD">
											<TABLE <%=LR_SPACE_TYPE_20%>>
												<TR>
													<TD CLASS="TD5" NOWRAP>전표합계금액</TD>
													<TD CLASS=TD6 NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtGlLocAmt2" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="전표합계금액" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
													<TD CLASS="TD5" NOWRAP></TD>
													<TD CLASS=TD6 NOWRAP></TD>
												</TR>
											</TABLE>
										</FIELDSET>
									</TD>
								</TR>
								<TR>
									<TD WIDTH=100% HEIGHT=100%>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</DIV>		
						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
	<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"		  tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtShowDt" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtShowBp" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtShowBiz" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
