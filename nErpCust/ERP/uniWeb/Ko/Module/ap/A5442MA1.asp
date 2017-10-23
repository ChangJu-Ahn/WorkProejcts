<%@ LANGUAGE="VBSCRIPT" %>
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

Const BIZ_PGM_ID      = "A5442MB1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "A5442MB2.asp"	
Const BIZ_PGM_ID2     = "A5442MB3.asp"	

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

'========================================================================================================
'                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------
Dim C_AcctCd1
Dim C_AcctNm1
Dim C_ApLocAmt1
Dim C_GlLocAmt1
Dim C_DiffAmt1  
Dim C_TempGlLocAmt1
Dim C_BatchLocAmt1
Dim C_BizAreaCd1
Dim C_BizAreaNm1
Dim C_DealBpCd1
Dim C_DealBpNm1

'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #2
'--------------------------------------------------------------------------------------------------------
Dim C_AcctCd2
Dim C_AcctNm2
Dim C_ApLocAmt2
Dim C_GlLocAmt2
Dim C_DiffAmt2  
Dim C_TempGlLocAmt2
Dim C_BatchLocAmt2
Dim C_GLInPutType
Dim C_GLInPutNm   
Dim C_BizAreaCd2
Dim C_BizAreaNm2
Dim C_DealBpCd2
Dim C_DealBpNm2

'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #3
'--------------------------------------------------------------------------------------------------------
Dim	C_AcctCd3 
Dim	C_AcctNm3		
Dim	C_ApNo			
Dim	C_ApDt			
Dim	C_GlDt			
Dim	C_ApLocAmt3		
Dim	C_GlLocAmt3		
Dim C_DiffAmt3 
Dim	C_TempGlLocAmt3	
Dim	C_BatchLocAmt3	
Dim C_GLNo			
Dim	C_TempGlNo		
Dim	C_BatchNO		
Dim	C_TempGlDt		
Dim	C_ApType3       
Dim	C_ApTypeNm3     
Dim	C_BizAreaCd3	
Dim	C_BizAreaNm3	
Dim	C_DealBpCd3		
Dim	C_DealBpNm3		

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
Dim  lgGlInputType 
Dim  lgGlInputTypeNm

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
			C_AcctCd1		= 1
			C_AcctNm1		= 2
			C_ApLocAmt1		= 3
			C_GlLocAmt1		= 4
			C_DiffAmt1      = 5
			C_TempGlLocAmT1	= 6
			C_BatchLocAmt1	= 7
			C_BizAreaCd1	= 8
			C_BizAreaNm1	= 9
			C_DealBpCd1		= 10
			C_DealBpNm1		= 11
		Case "B"
			C_AcctCd2		= 1
			C_AcctNm2		= 2
			C_ApLocAmt2		= 3
			C_GlLocAmt2		= 4
			C_DiffAmt2      = 5
			C_TempGlLocAmt2 = 6
			C_BatchLocAmt2	= 7
			C_GLInPutType	= 8  
			C_GLInPutNm		= 9
			C_BizAreaCd2	= 10
			C_BizAreaNm2	= 11
			C_DealBpCd2		= 12
			C_DealBpNm2		= 13
		Case "C"
			C_AcctCd3		= 1
			C_AcctNm3		= 2
			C_ApNo			= 3
			C_ApDt			= 4
			C_GlDt			= 5			
			C_ApLocAmt3		= 6
			C_GlLocAmt3		= 7
			C_DiffAmt3      = 8
			C_TempGlLocAmt3	= 9
			C_BatchLocAmt3	= 10
			C_GLNo			= 11
			C_TempGlNo		= 12
			C_BatchNO		= 13
			C_TempGlDt		= 14
			C_ApType3       = 15
			C_ApTypeNm3     = 16
			C_BizAreaCd3	= 17
			C_BizAreaNm3	= 18
			C_DealBpCd3		= 19
			C_DealBpNm3		= 20			
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
	frm1.txtApFrDt.Text  = UniConvDateAToB(FirstDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtApToDt.Text  = UniConvDateAToB(LastDate, Parent.gServerDateFormat, Parent.gDateFormat)

	frm1.txtShowBiz.value = "N"
	frm1.txtShowBp.value = "N"

	frm1.txtApFrDt.focus 	



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
				.MaxCols   = C_DealBpNm1 + 1                                                  ' ☜:☜: Add 1 to Maxcols
				.Col =.MaxCols
				.ColHidden = True
			   
				Call ggoSpread.ClearSpreadData()	
				Call GetSpreadColumnPos("A")
				                      'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
				ggoSpread.SSSetEdit    C_AcctCd1           ,"계정코드"           ,10    ,,,20     ,2
				ggoSpread.SSSetEdit    C_AcctNm1           ,"계정코드명"         ,15    ,3
				                      'ColumnPosition     Header            Width   Grp                    IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 
				ggoSpread.SSSetFloat   C_ApLocAmt1         ,"채무금액(자국)"     ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True 
				ggoSpread.SSSetFloat   C_GlLocAmt1         ,"회계전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True 
				ggoSpread.SSSetFloat   C_DiffAmt1			,"차이금액"           ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_TempGlLocAmT1     ,"결의전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True 
				ggoSpread.SSSetFloat   C_BatchLocAmt1      ,"Batch금액(자국)"    ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True 
				ggoSpread.SSSetEdit    C_BizAreaCd1        ,"사업장"             ,15    ,,,10     ,2
				ggoSpread.SSSetEdit    C_BizAreaNm1        ,"사업장명"           ,15    ,3
				ggoSpread.SSSetEdit    C_DealBpCd1         ,"거래처"             ,15    ,,,20     ,2
				ggoSpread.SSSetEdit    C_DealBpNm1         ,"거래처명"           ,15    ,3

				call ggoSpread.MakePairsColumn(C_BizAreaCd1,C_BizAreaNm1)
				call ggoSpread.MakePairsColumn(C_AcctCd1,C_AcctNm1)
				call ggoSpread.MakePairsColumn(C_DealBpCd1,C_DealBpNm1)
			   	Call ggoSpread.SSSetColHidden(C_BizAreaCd1,C_BizAreaNm1,True)
				Call ggoSpread.SSSetColHidden(C_DealBpCd1,C_DealBpNm1,True)
				Call ggoSpread.SSSetColHidden(C_BatchLocAmt1,C_BatchLocAmt1,True)
				.ReDraw = True
	
				Call SetSpreadLock("A")
			End With
		Case "B"
			With frm1.vspdData2
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.Spreadinit "V20021227",, Parent.gAllowDragDropSpread
				.ReDraw = False
				.MaxCols   = C_DealBpNm2 + 1                                                  ' ☜:☜: Add 1 to Maxcols
				.Col =.MaxCols
				.ColHidden = True
			   
				Call ggoSpread.ClearSpreadData()	
				Call GetSpreadColumnPos("B")
				                      'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
				ggoSpread.SSSetEdit    C_AcctCd2           ,"계정코드"           ,10    ,         ,    ,20       ,2
				ggoSpread.SSSetEdit    C_AcctNm2           ,"계정코드명"         ,15    ,3
				                      'ColumnPosition     Header            Width   Grp                    IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 
				ggoSpread.SSSetFloat   C_ApLocAmt2         ,"채무금액(자국)"     ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_GlLocAmt2         ,"회계전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_DiffAmt2			,"차이금액"           ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_TempGlLocAmT2     ,"결의전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_BatchLocAmt2      ,"Batch금액(자국)"    ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetEdit    C_GLInPutType       ,"전표입력경로"       ,15    ,         ,    ,20       ,2
				ggoSpread.SSSetEdit    C_GLInPutNm         ,"전표입력경로명"     ,15    ,3
				ggoSpread.SSSetEdit    C_BizAreaCd2        ,"사업장"             ,15    ,         ,    ,10       ,2
				ggoSpread.SSSetEdit    C_BizAreaNm2        ,"사업장명"           ,15    ,3
				ggoSpread.SSSetEdit    C_DealBpCd2         ,"거래처"             ,15    ,         ,    ,20       ,2
				ggoSpread.SSSetEdit    C_DealBpNm2         ,"거래처명"           ,15    ,3


				call ggoSpread.MakePairsColumn(C_BizAreaCd2,C_BizAreaNm2)
				call ggoSpread.MakePairsColumn(C_AcctCd2,C_AcctNm2)
				call ggoSpread.MakePairsColumn(C_DealBpCd2,C_DealBpNm2)
			  	Call ggoSpread.SSSetColHidden(C_BizAreaCd2,C_BizAreaNm2,True)
				Call ggoSpread.SSSetColHidden(C_DealBpCd2,C_DealBpNm2,True) 
				Call ggoSpread.SSSetColHidden(C_BatchLocAmt2,C_BatchLocAmt2,True)
				Call ggoSpread.SSSetColHidden(C_GLInPutType,C_GLInPutType,True)
				.ReDraw = True
	
				Call SetSpreadLock("B")
			End With			
		Case "C"
			With frm1.vspdData3
				ggoSpread.Source = frm1.vspdData3
				ggoSpread.Spreadinit "V20021227",, Parent.gAllowDragDropSpread
				.ReDraw = False
				.MaxCols = C_DealBpNm3 + 1                                                  ' ☜:☜: Add 1 to Maxcols
				.Col =.MaxCols
				.ColHidden = True

				Call ggoSpread.ClearSpreadData()	
				Call GetSpreadColumnPos("C")
				                      'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
				ggoSpread.SSSetEdit    C_AcctCd3       ,"계정코드"           ,10    ,         ,    ,20       ,2
				ggoSpread.SSSetEdit    C_AcctNm3       ,"계정코드명"         ,15    ,3
				ggoSpread.SSSetEdit    C_ApNo          ,"채무번호"           ,15    ,3        ,     ,15     ,2
				ggoSpread.SSSetDate    C_ApDt          ,"채무일자"           ,10    ,2        ,Parent.gDateFormat   ,-1 
				ggoSpread.SSSetDate    C_GlDt          ,"회계전표일자"       ,10    ,2        ,Parent.gDateFormat   ,-1 
				                      'ColumnPosition     Header            Width   Grp                    IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 
				
				ggoSpread.SSSetFloat   C_ApLocAmt3     ,"채무금액(자국)"     ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_GlLocAmt3     ,"회계전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_DiffAmt3      ,"차이금액"           ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_TempGlLocAmT3 ,"결의전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetFloat   C_BatchLocAmt3  ,"Batch금액(자국)"    ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
				ggoSpread.SSSetEdit    C_GLNo          ,"회계전표번호"       ,15    ,3        ,     ,18     ,2
				ggoSpread.SSSetEdit    C_TempGlNo      ,"결의전표번호"       ,15    ,3        ,     ,18     ,2				
				ggoSpread.SSSetEdit    C_BatchNO       ,"Batch 번호"         ,15    ,2        ,     ,18     ,2
				ggoSpread.SSSetDate    C_TempGlDt      ,"결의전표일자"       ,10    ,2        ,Parent.gDateFormat   ,-1 
				ggoSpread.SSSetEdit    C_ApType3       ,"전표입력경로"       , 5    ,         ,    ,10       ,2
				ggoSpread.SSSetEdit    C_ApTypeNm3     ,"전표입력경로명"     ,15    ,3				
				ggoSpread.SSSetEdit    C_BizAreaCd3    ,"사업장"             ,15    ,         ,    ,10       ,2
				ggoSpread.SSSetEdit    C_BizAreaNm3    ,"사업장명"           ,15    ,3
				ggoSpread.SSSetEdit    C_DealBpCd3     ,"거래처"             ,15    ,         ,    ,20       ,2
				ggoSpread.SSSetEdit    C_DealBpNm3     ,"거래처명"           ,15    ,3	
				
				call ggoSpread.MakePairsColumn(C_BizAreaCd3,C_BizAreaNm3)
				call ggoSpread.MakePairsColumn(C_AcctCd3,C_AcctNm3)
				call ggoSpread.MakePairsColumn(C_DealBpCd3,C_DealBpNm3)
				Call ggoSpread.SSSetColHidden(C_BizAreaCd3,C_BizAreaNm3,True)
				Call ggoSpread.SSSetColHidden(C_DealBpCd3,C_DealBpNm3,True)
				Call ggoSpread.SSSetColHidden(C_BatchLocAmt3,C_BatchLocAmt3,True)
				Call ggoSpread.SSSetColHidden(C_BatchNO,C_BatchNO,True)
				Call ggoSpread.SSSetColHidden(C_ApType3,C_ApType3,True)
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
			C_AcctCd1			= iCurColumnPos(1)    
			C_AcctNm1			= iCurColumnPos(2)
			C_ApLocAmt1			= iCurColumnPos(3)
			C_GlLocAmt1			= iCurColumnPos(4)
			C_DiffAmt3          = iCurColumnPos(5)
			C_TempGlLocAmT1     = iCurColumnPos(6)
			C_BatchLocAmt1		= iCurColumnPos(7)			
			C_BizAreaCd1        = iCurColumnPos(8)
			C_BizAreaNm1        = iCurColumnPos(9)
			C_DealBpCd1			= iCurColumnPos(10)
			C_DealBpNm1			= iCurColumnPos(11)
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)            
			C_AcctCd2			= iCurColumnPos(1)
			C_AcctNm2			= iCurColumnPos(2)
			C_ApLocAmt2			= iCurColumnPos(3)
			C_GlLocAmt2			= iCurColumnPos(4)	
			C_DiffAmt2          = iCurColumnPos(5)		
			C_TempGlLocAmt2     = iCurColumnPos(6)		
			C_BatchLocAmt2		= iCurColumnPos(7)
			C_GLInPutType    	= iCurColumnPos(8)
			C_GLInPutNm			= iCurColumnPos(9)
			C_BizAreaCd2		= iCurColumnPos(10)    
			C_BizAreaNm2        = iCurColumnPos(11)
			C_DealBpCd2			= iCurColumnPos(12)
			C_DealBpNm2			= iCurColumnPos(13)
		Case "C"
            ggoSpread.Source = frm1.vspdData3
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)            
			C_AcctCd3		    = iCurColumnPos(1)
			C_AcctNm3		    = iCurColumnPos(2)
			C_ApNo			    = iCurColumnPos(3)
			C_ApDt			    = iCurColumnPos(4)
			C_GlDt			    = iCurColumnPos(5)
			
			C_ApLocAmt3		    = iCurColumnPos(6)
			C_GlLocAmt3		    = iCurColumnPos(7)
			C_DiffAmt3          = iCurColumnPos(8)
			C_TempGlLocAmt3	    = iCurColumnPos(9)
			C_BatchLocAmt3	    = iCurColumnPos(10)
			C_GLNo			    = iCurColumnPos(11)
			C_TempGlNo		    = iCurColumnPos(12)
			C_BatchNO		    = iCurColumnPos(13)
			C_TempGlDt			= iCurColumnPos(14)
			C_ApType3			= iCurColumnPos(15)
			C_ApTypeNm3			= iCurColumnPos(16)
			C_BizAreaCd3		= iCurColumnPos(17)
			C_BizAreaNm3		= iCurColumnPos(18)
			C_DealBpCd3			= iCurColumnPos(19)
			C_DealBpNm3			= iCurColumnPos(20)			
    End Select    
End Sub

'========================================== OpenPopupTempGl() ============================================
'	Name : OpenPopuptempGL()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'=========================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	With frm1.vspdData3
		.Row = .ActiveRow
		.Col = C_TempGlNo
		arrParam(0) = Trim(.Text)							        '결의전표번호 
	    arrParam(1) = ""											'Reference번호	
	End With


	If arrParam(0) = "" Then
		IntRetCD = DisplayMsgBox("970000","X" , "결의전표", "X") 	
		IsOpenPop = False
		Exit Function
	End If	
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'========================================== OpenPopupGL()  =============================================
'	Name : OpenPopupGL()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData3
		.Row = .ActiveRow
		.Col = C_GLNo
		arrParam(0) = Trim(.Text)							        '회계전표번호 
	    arrParam(1) = ""											'Reference번호	
	End With
	
	If arrParam(0) = "" Then
		IntRetCD = DisplayMsgBox("970000","X" , "회계전표", "X") 	
		IsOpenPop = False
		Exit Function
	End If

	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'========================================== 2.4.2 OpenPopupBatch()  =============================================
'	Name : OpenPopupBatch()
'	Description : Ref 화면을 call한다. 
'================================================================================================================ 
Function OpenPopupBatch()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD

	iCalledAspName = AskPRAspName("a5140ra1")
'	If Trim(iCalledAspName) = "" Then
'		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5140ra1", "X")
'		IsOpenPop = False
'		Exit Function
'	End If

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData3
		.Row = .ActiveRow
		.Col = C_BatchNO
		arrParam(0) = Trim(.Text)							        '배치번호 
	    arrParam(1) = ""											'Reference번호	
	End With



	IsOpenPop = True
'	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _	
	arrRet = window.showModalDialog("a5140ra1.asp", Array(window.parent,arrParam), _	
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

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
			If frm1.txtBizAreaCd.className = "protected" Then Exit Function
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
		Case 1
			arrParam(0)  = "계정코드 팝업"							' 팝업 명칭 
			arrParam(1)  = "A_ACCT "									' TABLE 명칭 
			arrParam(2)  = strCode			       						' Code Condition
			arrParam(3)  = ""											' Name Cindition
			arrParam(4)  = "ACCT_TYPE LIKE " & FilterVar("%J%", "''", "S") & "  "						' Where Condition
			arrParam(5)  = "계정코드"			
	
			arrField(0)  = "ACCT_CD"									' Field명(0)
			arrField(1)  = "ACCT_NM"									' Field명(1)
    
			arrHeader(0) = "계정코드"								' Header명(0)
			arrHeader(1) = "계정코드명"								' Header명(3)
		Case 2
			If frm1.txtdealbpCd.className = "protected" Then Exit Function
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
			arrParam(1) = " a_open_ap a , b_minor b "					' TABLE 명칭 
			arrParam(2) = strCode						 				' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = " b.major_cd=" & FilterVar("A1001", "''", "S") & "  and a.ap_type=b.minor_cd " ' Where Condition
			arrParam(5) = "전표입력경로"			
			arrField(0) = "a.ap_type"									' Field명(0)
			arrField(1) = "b.minor_nm"									' Field명(1)
    
			arrHeader(0) = "전표입력경로"							' Header명(0)
			arrHeader(1) = "전표입력경로명"							' Header명(1)			
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
				.txtBizAreaCd.focus
			Case "1"
				.txtAcctCd.focus
			Case "2"
				.txtdealbpCd.focus
			Case "3"
				.txtGlInputType.focus				
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
		     	.txtBizAreaCd.value     = arrRet(0)
				.txtBizAreaNM.value     = arrRet(1)
				.txtBizAreaCd.focus
			 Case "1"
				.txtAcctCd.value        = arrRet(0)
				.txtAcctNM.value        = arrRet(1)
				.txtAcctCd.focus
			Case "2"
				.txtdealbpCd.value      = arrRet(0)
				.txtdealbpNM.value      = arrRet(1)
				.txtdealbpCd.focus
			Case "3"
				.txtGlInputType.value   = arrRet(0)
				.txtGlInputTypeNM.value = arrRet(1)
				lgGlInputType = .txtGlInputType.value
				lgGlInputTypeNm =  .txtGlInputTypeNm.value 	 				
				.txtGlInputType.focus				
	    End Select
	End With
End Function     

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1

	Call MoveJmpClick() 
	Call SetToolbar("1100000000001111") 				 
End Function

Function ClickTab2()
	Call changeTabs(TAB2)	 
	gSelframeFlg = TAB2
	Call MoveJmpClick()
	Call SetToolbar("1100000000001111") 
End Function

Function ClickTab3()
	Call changeTabs(TAB3)	 
	gSelframeFlg = TAB3
	Call MoveJmpClick()
	Call SetToolbar("1100000000001111")
End Function

'======================================================================================================
'	기능: 
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function MoveJmpClick()
	Dim innerString
	
	Select Case gSelframeFlg
		Case TAB1
		
			RefView.innerHTML =  "<font color=""#777777"">회계전표&nbsp;|&nbsp;결의전표</font>"
			inputTypeViewtemp1.style.DISPLAY=""
			inputTypeViewtemp2.style.DISPLAY=""
			inputTypeView1.style.DISPLAY="NONE"
			inputTypeView2.style.DISPLAY="NONE"


		Case TAB2			
			RefView.innerHTML =  "<font color=""#777777"">회계전표&nbsp;|&nbsp;결의전표</font>"

			inputTypeViewtemp1.style.DISPLAY="NONE"
			inputTypeViewtemp2.style.DISPLAY="NONE"
			inputTypeView1.style.DISPLAY=""
			inputTypeView2.style.DISPLAY=""
			frm1.txtGlInputType.value = lgGlInputType
			frm1.txtGlInputTypeNm.value = lgGlInputTypeNm


		Case TAB3
			RefView.innerHTML =  "<A href='vbscript:OpenPopupGL()'>회계전표</A>&nbsp;|&nbsp;<A href='vbscript:OpenPopupTempGL()'>결의전표</A>"

			inputTypeViewtemp1.style.DISPLAY="NONE"
			inputTypeViewtemp2.style.DISPLAY="NONE"
			inputTypeView1.style.DISPLAY=""
			inputTypeView2.style.DISPLAY=""
			frm1.txtGlInputType.value = lgGlInputType
			frm1.txtGlInputTypeNm.value = lgGlInputTypeNm
			

	End Select    
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

	Call ggoOper.SetReqAttr(frm1.txtBizAreaCd, "Q")		
	Call ggoOper.SetReqAttr(frm1.txtdealbpCd, "Q")		
	
	Call ClickTab1()																				'☜: Check Cookie
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
	
	if frm1.txtAcctCd.value = "" then
		frm1.txtAcctNm.value = ""
    end if
    
   if frm1.txtBizAreaCd.value = "" then
		frm1.txtBizAreaNm.value = ""
    end if

    if frm1.txtdealbpCd.value = "" then
		frm1.txtdealbpNm.value = ""
    end if
    
    if frm1.txtGlInputType.value = "" then
		frm1.txtGlInputTypeNm.value = ""
    end if
	

	Select Case gSelframeFlg
		Case TAB1
			ggoSpread.Source = Frm1.vspdData1

			Call ggoSpread.ClearSpreadData()

			If DbQuery("A") = False Then															'☜: Query db data
				Exit Function
			End If
		Case TAB2
			ggoSpread.Source = Frm1.vspdData2

			Call ggoSpread.ClearSpreadData()
    
			If DbQuery("B") = False Then															'☜: Query db data
				Exit Function
			End If
		Case TAB3
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
	Dim strVal
	
    On Error Resume Next																		'☜: If process fails
    Err.Clear																					'☜: Clear error status
 
    DbQuery = False																				'☜: Processing is NG

    Call DisableToolBar(Parent.TBC_QUERY)														'☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)																		'☜: Show Processing Message
	
	Select Case pDirect
		Case "A"
			strVal = BIZ_PGM_ID		& "?txtMode="        & Parent.UID_M0001                      '☜: Query
			strVal = strVal			& "&txtMaxRows=" 	 & Frm1.vspdData1.MaxRows				'☜: Max fetched data
		Case "B"
			strVal = BIZ_PGM_ID1	& "?txtMode="        & Parent.UID_M0001						'☜: Query
			strVal = strVal			& "&txtGlInputType=" & frm1.txtGlInputType.value
			strVal = strVal			& "&txtMaxRows="	 & Frm1.vspdData2.MaxRows				'☜: Max fetched data
		Case "C"
			strVal = BIZ_PGM_ID2	& "?txtMode="        & Parent.UID_M0001						'☜: Query
			strVal = strVal			& "&txtGlInputType=" & frm1.txtGlInputType.value
			strVal = strVal			& "&txtMaxRows="	 & Frm1.vspdData3.MaxRows				'☜: Max fetched data
			strVal = strVal			& "&lgStrPrevKey="	 & lgStrPrevKey              
	End Select 
	
    strVal = strVal		& "&txtApFrDt="      & frm1.txtApFrDt.text								'☜:
    strVal = strVal		& "&txtApToDt="      & frm1.txtApToDt.text								'☜:
    strVal = strVal		& "&txtShowBiz="     & frm1.txtShowBiz.value							'☜:
    strVal = strVal		& "&txtBizAreaCd="   & frm1.txtBizAreaCd.value							'☜:
    strVal = strVal     & "&txtAcctCd="		 & frm1.txtAcctCd.value								'☜:
    strVal = strVal     & "&txtShowBp="		 & frm1.txtShowbp.value                             '☜:
    strVal = strVal     & "&txtDealBpCd="	 & frm1.txtdealbpCd.value                           '☜:
    strVal = strVal     & "&DispMeth="		 & Trim(frm1.RdoDiff.checked )                      '☜:

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
			Frm1.vspdData2.focus
		Case TAB3
			Frm1.vspdData3.focus
	End Select
	Call DOSUM()
	Call SetToolbar("1100000000001111")															'☆: Developer must customize
    
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
' Name : OpenReference1
' Desc : developer describe this line 
'========================================================================================================
'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'=======================================================================================================
Sub DoSum()
	Dim dbTotApLocAmt
	Dim dbTotGlLocAmt
	DIm dbTotDiffLocAmt
	Dim dbTotTempGlLocAmt
	
	Select Case gSelframeFlg
		Case TAB1
			With frm1.vspdData1
				dbTotApLocAmt		= FncSumSheet(frm1.vspdData1,C_ApLocAmt1, 1, .MaxRows, False, -1, -1, "V")
				dbTotGlLocAmt		= FncSumSheet(frm1.vspdData1,C_GlLocAmt1, 1, .MaxRows, False, -1, -1, "V")
				dbTotDiffLocAmt		= FncSumSheet(frm1.vspdData1,C_DiffAmt1, 1, .MaxRows, False, -1, -1, "V")
				dbTotTempGlLocAmt	= FncSumSheet(frm1.vspdData1,C_TempGlLocAmt1, 1, .MaxRows, False, -1, -1, "V")
	
				frm1.txtTotApLocAmt1.text = UNIConvNumPCToCompanyByCurrency(dbTotApLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				frm1.txtTotGlLocAmt1.text = UNIConvNumPCToCompanyByCurrency(dbTotGlLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    
				frm1.txtTotDiffLocAmt1.text = UNIConvNumPCToCompanyByCurrency(dbTotDiffLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    
				frm1.txtTotTempGlLocAmt1.text = UNIConvNumPCToCompanyByCurrency(dbTotTempGlLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")        
			End With 
		Case TAB2
			With frm1.vspdData2
				dbTotApLocAmt		= FncSumSheet(frm1.vspdData2,C_ApLocAmt2, 1, .MaxRows, False, -1, -1, "V")
				dbTotGlLocAmt		= FncSumSheet(frm1.vspdData2,C_GlLocAmt2, 1, .MaxRows, False, -1, -1, "V")
				dbTotDiffLocAmt		= FncSumSheet(frm1.vspdData2,C_DiffAmt2, 1, .MaxRows, False, -1, -1, "V")
				dbTotTempGlLocAmt	= FncSumSheet(frm1.vspdData2,C_TempGlLocAmt2, 1, .MaxRows, False, -1, -1, "V")
				
				frm1.txtTotApLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dbTotApLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				frm1.txtTotGlLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dbTotGlLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    
				frm1.txtTotDiffLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dbTotDiffLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    
				frm1.txtTotTempGlLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dbTotTempGlLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")        
			End With 
	End Select 
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

	If frm1.txtBizAreaCd.value = "" Then Exit Sub

	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(frm1.txtBizAreaCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBizAreaNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("124200","X","X","X")  	
		frm1.txtBizAreaCd.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtAcctCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtAcctCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtAcctCd.value = "" Then Exit Sub
	
	If CommonQueryRs("ACCT_NM", "A_ACCT ", " ACCT_CD=  " & FilterVar(frm1.txtAcctCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtAcctNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("110100","X","X","X")  	
		frm1.txtAcctCd.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtdealbpCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtdealbpCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtdealbpCd.value = "" Then Exit Sub	
		
	If CommonQueryRs("BP_NM", "B_BIZ_PARTNER ", " BP_CD=  " & FilterVar(frm1.txtdealbpCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtdealbpNm.value= Trim(arrVal(0)) 

	Else
		IntRetCD = DisplayMsgBox("126100","X","X","X")  	
		frm1.txtdealbpCd.focus
	End If

End Sub
'========================================================================================================
'   Event Name : txtGlInputType_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtGlInputType_onChange()
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtGlInputType.value = "" Then 
		lgGlInputType = frm1.txtGlInputType.value
		lgGlInputTypeNm = ""
		Exit Sub	
	End if
		
	If CommonQueryRs("MINOR_NM", "B_MINOR ", " MAJOR_CD=" & FilterVar("A1001", "''", "S") & " AND MINOR_CD=  " & FilterVar(frm1.txtGlInputType.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtGlInputTypeNM.value= Trim(arrVal(0)) 
		lgGlInputType = frm1.txtGlInputType.value
		lgGlInputTypeNm = frm1.txtGlInputTypeNm.value 	
	Else
		IntRetCD = DisplayMsgBox("800506","X","X","X")  	
		frm1.txtGlInputType.focus
	End If
	
End Sub

'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0001111111")   
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
    Call SetPopupMenuItemInf("0001111111")     
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
    Call SetPopupMenuItemInf("0001111111")    
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
    	If lgStrPrevKey <> "" Then                         
           If DbQuery("C") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub

'========================================================================================================
' Name : txtApFrDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtApFrDt_DblClick(Button)
    If Button = 1 Then
		frm1.txtApFrDt.Action = 7                                    ' 7 : Popup Calendar ocx
		Call SetFocusToDocument("M")	
		frm1.txtApFrDt.Focus
    End If
End Sub

Sub txtApFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtApFrDt.focus
		Call FncQuery
	end if
End Sub
'========================================================================================================
' Name : txtApToDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtApToDt_DblClick(Button)
    If Button = 1 Then
		frm1.txtApToDt.Action = 7                                    ' 7 : Popup Calendar ocx
		Call SetFocusToDocument("M")	
		frm1.txtApToDt.Focus
    End If
End Sub

Sub txtApToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtApToDt.focus
		Call FncQuery
	end if
End Sub

'========================================================================================================
' Name : chkShowBiz_onchange
' Desc : 
'========================================================================================================
Sub chkShowBiz_onchange()
	If frm1.chkShowBiz.checked = True Then
		frm1.txtShowBiz.value = "Y"
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCd, "D")
		ggoSpread.Source = frm1.vspdData1
		Call ggoSpread.SSSetColHidden(C_BizAreaCd1,C_BizAreaNm1,FALSE)

		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.SSSetColHidden(C_BizAreaCd2,C_BizAreaNm2,FALSE)

		ggoSpread.Source = frm1.vspdData3
		Call ggoSpread.SSSetColHidden(C_BizAreaCd3,C_BizAreaNm3,FALSE)
	Else
		frm1.txtShowBiz.value = "N"	
		frm1.txtBizAreaCd.value = ""
		frm1.txtBizAreaNm.value=""
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCd, "Q")		

		ggoSpread.Source = frm1.vspdData1
		Call ggoSpread.SSSetColHidden(C_BizAreaCd1,C_BizAreaNm1,True)

		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.SSSetColHidden(C_BizAreaCd2,C_BizAreaNm2,True)

		ggoSpread.Source = frm1.vspdData3
		Call ggoSpread.SSSetColHidden(C_BizAreaCd3,C_BizAreaNm3,True)	
	End If
End Sub

'========================================================================================================
' Name : chkShowBp_onchange
' Desc : 
'========================================================================================================
Sub chkShowBp_onchange()
	If frm1.chkShowBp.checked = True Then
		frm1.txtShowBp.value = "Y"
		Call ggoOper.SetReqAttr(frm1.txtdealbpCd, "D")

		ggoSpread.Source = frm1.vspdData1
		Call ggoSpread.SSSetColHidden(C_DealBpCd1,C_DealBpNm1,False)

		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.SSSetColHidden(C_DealBpCd2,C_DealBpNm2,False)
	
		ggoSpread.Source = frm1.vspdData3
		Call ggoSpread.SSSetColHidden(C_DealBpCd3,C_DealBpNm3,False)			

	Else
		frm1.txtShowBp.value = "N"	
		frm1.txtdealbpCd.value = ""
		frm1.txtdealbpNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtdealbpCd, "Q")	
		
		ggoSpread.Source = frm1.vspdData1
		Call ggoSpread.SSSetColHidden(C_DealBpCd1,C_DealBpNm1,TRUE)

		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.SSSetColHidden(C_DealBpCd2,C_DealBpNm2,TRUE)
	
		ggoSpread.Source = frm1.vspdData3
		Call ggoSpread.SSSetColHidden(C_DealBpCd3,C_DealBpNm3,TRUE)		
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY SCROLL="NO" TABINDEX="-1">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><IMG height=23 src="../../image/table/seltab_up_left.gif" width=9></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>계정코드별합계</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><IMG height=23 src="../../image/table/tab_up_left.gif" width=9></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>전표입력경로별합계</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><IMG height=23 src="../../image/table/tab_up_left.gif" width=9></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>채무별발생금액</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
		    		<TD WIDTH=* align=right><span id="RefView"><A href="vbscript:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<A href="vbscript:OpenPopupTempGL()">결의전표</A> </SPAN></TD>					
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
									<TD CLASS="TD5" NOWRAP>발생일자</TD>
									<TD CLASS="TD6" NOWRAP>
									    <script language =javascript src='./js/a5442ma1_fpDateTime1_txtApFrDt.js'></script>
									    &nbsp;~&nbsp;
									    <script language =javascript src='./js/a5442ma1_fpDateTime2_txtApToDt.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>계정코드</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME=txtAcctCd ALT="계정코드" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtACctCd.value, 1)">
														 <INPUT TYPE=TEXT NAME=txtAcctNm ALT="계정코드명" SIZE="18" style="HEIGHT: 20px; " tag="14" ></TD>
								</TR>
								<TR>		
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME=txtBizAreaCd ALT="사업장코드" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtBizAreaCd.value, 0)">
														 <INPUT TYPE=TEXT NAME=txtBizAreaNm ALT="사업장명" SIZE="18" style="HEIGHT: 20px; " tag="14" >
														 <INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkShowBiz ID=chkShowBiz tag="1" onclick=chkShowBiz_onchange()></TD>
									<TD CLASS=TD5 NOWRAP>거래처</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME=txtdealbpCd ALT="거래처" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtdealbpCd.value, 2)">
														 <INPUT TYPE=TEXT NAME=txtdealbpNm ALT="거래처명" SIZE="18" style="HEIGHT: 20px; " tag="14" >
														 <INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkShowBp ID=chkShowBp tag="" onclick=chkShowBp_onchange()></TD>
								</TR>
								<TR>									
									<TD CLASS=TD5 ID=inputTypeViewtemp1>&nbsp;</TD>								
									<TD CLASS=TD6 ID=inputTypeViewtemp2>&nbsp;</TD>
									<TD CLASS=TD5 ID=inputTypeView1>전표입력경로</TD>								
									<TD CLASS=TD6 ID=inputTypeView2><INPUT TYPE=TEXT NAME=txtGlInputType ALT="전표입력경로" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnGlInputTypee ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtGlInputType.value, 3)" >
												   <INPUT TYPE=TEXT NAME=txtGlInputTypeNm ALT="전표입력경로명" SIZE="18" style="HEIGHT: 20px; " tag="14" ></TD>
									<TD CLASS=TD5 NOWRAP>조회방식</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoType" ID="RdoDiff" VALUE="S" TAG="11" ><LABEL FOR="rdoReport1">차이금액</LABEL>&nbsp;&nbsp
														 <INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoType" ID="RdoTotal" VALUE="D" TAG="11" Checked><LABEL FOR="rdoReport2">전체금액</LABEL></TD>		
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
										<script language =javascript src='./js/a5442ma1_OBJECT1_vspdData1.js'></script>
									</TD>
								</TR>
								<TR>
									<TD HEIGHT=40 WIDTH=25%>
										<FIELDSET CLASS="CLSFLD">
											<TABLE  <%=LR_SPACE_TYPE_20%>>
												<TR>
													<TD CLASS="TD5" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;채무금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotApLocAmt1.js'></script>
													</TD>
													<TD class=TD5 STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TD5" NOWRAP>차이금액합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotDiffLocAmt1.js'></script>
													</TD>

												</TR>
												<TR>
													<TD CLASS="TD5" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;회계전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotGlLocAmt1.js'></script>
													</TD>
													<TD class=TD5 STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TD5" NOWRAP>결의전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotTempGlLocAmt1.js'></script>
													</TD>
												</TR>
											</TABLE>
										</FIELDSET>
									</TD>
								</TR>
							</TABLE>
						</DIV>		
						
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
										<script language =javascript src='./js/a5442ma1_OBJECT1_vspdData2.js'></script>
									</TD>
								</TR>
								<TR>
									<TD HEIGHT=40 WIDTH=25%>
										<FIELDSET CLASS="CLSFLD">
											<TABLE  <%=LR_SPACE_TYPE_20%>>
												<TR>
													<TD CLASS="TD5" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;채무금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotApLocAmt2.js'></script>
													</TD>
													<TD class=TD5 STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TD5" NOWRAP>차이금액합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotDiffLocAmt2.js'></script>
													</TD>
												</TR>
												<TR>
													<TD CLASS="TD5" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;회계전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotGlLocAmt2.js'></script>
													</TD>
													<TD class=TD5 STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TD5" NOWRAP>결의전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotTempGlLocAmt2.js'></script>
													</TD>
												</TR>
											</TABLE>
										</FIELDSET>
									</TD>
								</TR>
							</TABLE>
						</DIV>		
						
						
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
										<script language =javascript src='./js/a5442ma1_OBJECT1_vspdData3.js'></script>
									</TD>
								</TR>
								<TR>
									<TD HEIGHT=40 WIDTH=25%>
										<FIELDSET CLASS="CLSFLD">
											<TABLE  <%=LR_SPACE_TYPE_20%>>
												<TR>
													<TD CLASS="TD5" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;채무금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotArLocAmt3.js'></script>
													</TD>
													<TD class=TD5 STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TD5" NOWRAP>차이금액합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotDiffLocAmt3.js'></script>
													</TD>
												</TR>
												<TR>
													<TD CLASS="TD5" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;회계전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotGlLocAmt3.js'></script>
													</TD>
													<TD class=TD5 STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TD5" NOWRAP>결의전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5442ma1_OBJECT1_txtTotTempGlLocAmt3.js'></script>
													</TD>
												</TR>
											</TABLE>
										</FIELDSET>
									</TD>
								</TR>
							</TABLE>
						</DIV>		
									
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtShowBiz"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtShowBp"       TAG="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
