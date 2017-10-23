<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : A5456ma1
*  4. Program Name         : 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2003/06/24
*  8. Modified date(Last)  : 2003/06/
*  9. Modifier (First)     : Oh, Soo Min 
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
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "A5456MB1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "A5456MB2.asp"	

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #A
'--------------------------------------------------------------------------------------------------------
Dim C_AcctCd1
Dim C_AcctNm1
Dim C_SttlLocAmt1
Dim C_GlLocAmt1
Dim C_DiffLocAmt1
Dim C_TempGlLocAmt1
Dim C_BatchLocAmt1
Dim C_GLInPutCd1
Dim C_GLInPutNm1
Dim C_BizAreaCd1
Dim C_BizAreaNm1
Dim C_BpCd1
Dim C_BpNm1

'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #B
'--------------------------------------------------------------------------------------------------------
Dim C_AcctCd2
Dim C_AcctNm2
Dim C_NoteNo2
Dim C_SttlDt2
Dim C_GlDt2
Dim C_SttlLocAmt2
Dim C_GlLocAmt2
Dim C_DiffLocAmt2
Dim C_TempGlLocAmt2
Dim C_BatchLocAmt2
Dim C_GLNo2
Dim C_TempGlNo2
Dim C_BatchNo2
Dim C_GLInPutCd2   
Dim C_GLInPutNm2   
Dim C_BizAreaCd2
Dim C_BizAreaNm2
Dim C_BpCd2
Dim C_BpNm2
Dim C_TempGlDt2

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
Dim  lgNoteNoFr, lgNoteNoTo

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
			C_SttlLocAmt1	= 3
			C_GlLocAmt1		= 4
			C_DiffLocAmt1	= 5
			C_TempGlLocAmt1	= 6
			C_BatchLocAmt1	= 7			
			C_GLInPutCd1	= 8
			C_GLInPutNm1	= 9 
			C_BizAreaCd1	= 10
			C_BizAreaNm1	= 11
			C_BpCd1			= 12
			C_BpNm1			= 13			
			
		Case "B"
			C_AcctCd2		= 1 
			C_AcctNm2		= 2
			C_NoteNo2		= 3
			C_SttlDt2		= 4
			C_GlDt2			= 5
			C_SttlLocAmt2	= 6			
			C_GlLocAmt2		= 7
			C_DiffLocAmt2	= 8
			C_TempGlLocAmt2	= 9
			C_BatchLocAmt2	= 10			
			C_GLNo2			= 11			
			C_TempGlNo2		= 12
			C_BatchNO2		= 13
			C_GLInPutCd2	= 14
			C_GLInPutNm2	= 15
			C_BizAreaCd2	= 16
			C_BizAreaNm2	= 17
			C_BpCd2			= 18
			C_BpNm2			= 19			
			C_TempGlDt2		= 20
	
	End Select 			
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
    
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'gSelframeFlg = TAB1
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim StartDate, FirstDate, LastDate

	StartDate	= "<%=GetSvrDate%>"
	FirstDate	= UNIGetFirstDay(UNIDateAdd("m", -1, StartDate, parent.gServerDateFormat),Parent.gServerDateFormat)
	LastDate	= UNIGetLastDay(FirstDate , Parent.gServerDateFormat)
	frm1.txtSttlFrDt.Text  = UniConvDateAToB(FirstDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtSttlToDt.Text  = UniConvDateAToB(LastDate, Parent.gServerDateFormat, Parent.gDateFormat)
	
    frm1.txtShowBiz.value = "N"
    frm1.txtShowBp.value = "N"    
    
    frm1.txtSttlFrDt.focus 	
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

	
	<% Call LoadInfTB19029A("Q","*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","MA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
  	Dim arrData

		'어음구분 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboNoteFg ,lgF0  ,lgF1  ,Chr(11))
	
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
			   .ReDraw = false
			   .MaxCols   = C_BpNm1 + 1                                                  ' ☜:☜: Add 1 to Maxcols
			   .Col =.MaxCols
			   .ColHidden = true

			   Call ggoSpread.ClearSpreadData()				   			   
			   Call GetSpreadColumnPos("A")					    
			                         'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)			   
			   ggoSpread.SSSetEdit    C_AcctCd1           ,"계정코드"           ,18    ,,,20     ,2
			   ggoSpread.SSSetEdit    C_AcctNm1           ,"계정코드명"         ,18    ,3
			                         'ColumnPosition     Header            Width   Grp                    IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 
			   ggoSpread.SSSetFloat   C_SttlLocAmt1       ,"결제금액(자국)"     ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
			   ggoSpread.SSSetFloat   C_GlLocAmt1         ,"회계전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True
			   ggoSpread.SSSetFloat   C_DiffLocAmt1      ,"차이금액"    ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True 
			   ggoSpread.SSSetFloat   C_TempGlLocAmt1     ,"결의전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True 
			   ggoSpread.SSSetFloat   C_BatchLocAmt1      ,"Batch금액(자국)"    ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True

			   ggoSpread.SSSetEdit    C_GLInPutCd1        ,"전표입력경로"       ,10    ,,,10     ,2
			   ggoSpread.SSSetEdit    C_GLInPutNm1        ,"전표입력경로명"     ,15    ,3
			   ggoSpread.SSSetEdit    C_BizAreaCd1        ,"사업장"             ,10    ,,,10     ,2
			   ggoSpread.SSSetEdit    C_BizAreaNm1        ,"사업장명"           ,15    ,3
			   ggoSpread.SSSetEdit    C_BpCd1             ,"거래처"             ,18    ,,,10     ,2
			   ggoSpread.SSSetEdit    C_BpNm1			  ,"거래처명"           ,18    ,3			   
			   
			   call ggoSpread.MakePairsColumn(C_AcctCd1,C_AcctNm1)
			   call ggoSpread.MakePairsColumn(C_GLInPutCd1,C_GLInPutNm1)		   
			   call ggoSpread.MakePairsColumn(C_BizAreaCd1,C_BizAreaNm1)
			   call ggoSpread.MakePairsColumn(C_BpCd1,C_BpNm1)
			   Call ggoSpread.SSSetColHidden(C_BizAreaCd1,C_BizAreaNm1,True)
			   Call ggoSpread.SSSetColHidden(C_BpCd1,C_BpNm1,True)
			   Call ggoSpread.SSSetColHidden(C_BatchLocAmt1,C_BatchLocAmt1,True)
			   .ReDraw = true
	
			   Call SetSpreadLock("A")
    
			End With		
					
		
		Case "B"

			With frm1.vspdData2
			   ggoSpread.Source = frm1.vspdData2
			   ggoSpread.Spreadinit "V20021227",, Parent.gAllowDragDropSpread
			   .ReDraw = false
			   .MaxCols   = C_TempGlDt2 + 1                                                  ' ☜:☜: Add 1 to Maxcols
			   .Col =.MaxCols
			   .ColHidden = true

			   Call ggoSpread.ClearSpreadData()				   			   
			   Call GetSpreadColumnPos("B")			   			   
			                         'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
			   ggoSpread.SSSetEdit    C_AcctCd2       ,"계정코드"           ,18    ,         ,    ,20       ,2
			   ggoSpread.SSSetEdit    C_AcctNm2       ,"계정코드명"         ,18    ,3
			   ggoSpread.SSSetEdit    C_NoteNo2       ,"어음번호"           ,18    ,2                  ,     ,15     ,2
			   
			   ggoSpread.SSSetDate    C_SttlDt2        ,"결산일자"           ,15    ,2                  ,Parent.gDateFormat   ,-1 
			   ggoSpread.SSSetDate    C_GlDt2          ,"회계전표일자"       ,15    ,2                  ,Parent.gDateFormat   ,-1			   						   			 
			                         'ColumnPosition     Header            Width   Grp                    IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 			   
			   ggoSpread.SSSetFloat   C_SttlLocAmt2   ,"결제금액(자국)"     ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
			   ggoSpread.SSSetFloat   C_GlLocAmt2     ,"회계전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  			   
			   ggoSpread.SSSetFloat   C_DiffLocAmt2   ,"차이금액"			,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  			   
			   ggoSpread.SSSetFloat   C_TempGlLocAmt2 ,"결의전표금액(자국)" ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  
			   ggoSpread.SSSetFloat   C_BatchLocAmt2  ,"Batch금액(자국)"    ,15     ,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  			   
			   ggoSpread.SSSetEdit    C_GLNo2         ,"회계전표번호"       ,18    ,2                  ,     ,18     ,2			   
			   ggoSpread.SSSetEdit    C_TempGlNo2     ,"결의전표번호"       ,10    ,0                  ,     ,18     ,2
			   ggoSpread.SSSetEdit    C_BatchNo2      ,"Batch 번호"         ,18    ,2                  ,     ,18     ,2
			   ggoSpread.SSSetEdit    C_GLInPutCd2	  ,"전표입력경로"       ,18    ,         ,    ,10       ,2	
			   ggoSpread.SSSetEdit    C_GLInPutNm2    ,"전표입력경로명"     ,18    ,         ,    ,20       ,2
			   ggoSpread.SSSetEdit    C_BizAreaCd2    ,"사업장"             ,10    ,         ,    ,10       ,2
			   ggoSpread.SSSetEdit    C_BizAreaNm2    ,"사업장명"           ,15    ,3
			   ggoSpread.SSSetEdit    C_BpCd2     ,"거래처"             ,18    ,         ,    ,20       ,2
			   ggoSpread.SSSetEdit    C_BpNm2     ,"거래처명"           ,18    ,3				   
			   ggoSpread.SSSetDate    C_TempGlDt2     ,"결의전표일자"       ,15    ,2                  ,Parent.gDateFormat   ,-1 

			   call ggoSpread.MakePairsColumn(C_AcctCd2,C_AcctNm2)
			   call ggoSpread.MakePairsColumn(C_GLInPutCd2,C_GLInPutNm2)		   
			   call ggoSpread.MakePairsColumn(C_BizAreaCd2,C_BizAreaNm2)
			   call ggoSpread.MakePairsColumn(C_BpCd2,C_BpNm2)
			   Call ggoSpread.SSSetColHidden(C_BizAreaCd2,C_BizAreaNm2,True)
			   Call ggoSpread.SSSetColHidden(C_BpCd2,C_BpNm2,True)
			   Call ggoSpread.SSSetColHidden(C_BatchLocAmt2,C_BatchLocAmt2,True)
			   Call ggoSpread.SSSetColHidden(C_BatchNo2,C_BatchNo2,True)
			   .ReDraw = true				
			   Call SetSpreadLock("B")
    
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
	End Select
End Sub


'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
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
            ggoSpread.Source = frm1.vspdData1            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)               
			C_AcctCd1			= iCurColumnPos(1)    
			C_AcctNm1			= iCurColumnPos(2)			
			C_SttlLocAmt1		= iCurColumnPos(3)
			C_GlLocAmt1			= iCurColumnPos(4)
			C_DiffLocAmt1		= iCurColumnPos(5)
			C_TempGlLocAmt1     = iCurColumnPos(6)
			C_BatchLocAmt1		= iCurColumnPos(7)
			C_GLInPutCd1        = iCurColumnPos(8)			
			C_GLInPutNm1        = iCurColumnPos(9)
			C_BizAreaCd1        = iCurColumnPos(10)			
			C_BizAreaNm1        = iCurColumnPos(11)
			C_BpCd1				= iCurColumnPos(12)
			C_BpNm1				= iCurColumnPos(13)
			
       
		Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                  
			C_AcctCd2			= iCurColumnPos(1)
			C_AcctNm2			= iCurColumnPos(2)
			C_NoteNo2			= iCurColumnPos(3)	
			C_SttlDt2			= iCurColumnPos(4)
			C_GlDt2				= iCurColumnPos(5)			
			C_SttlLocAmt2		= iCurColumnPos(6)			
			C_GlLocAmt2			= iCurColumnPos(7)
			C_DiffLocAmt2		= iCurColumnPos(8)
			C_TempGlLocAmt2		= iCurColumnPos(9)
			C_BatchLocAmt2      = iCurColumnPos(10)
			C_GLNo2				= iCurColumnPos(11)			
			C_TempGlNo2			= iCurColumnPos(12)	
			C_BatchNo2			= iCurColumnPos(13)						
			C_GLInPutCd2		= iCurColumnPos(14)
			C_GLInPutNm2		= iCurColumnPos(15)
			C_BizAreaCd2		= iCurColumnPos(16)    
			C_BizAreaNm2        = iCurColumnPos(17)	
			C_BpCd2				= iCurColumnPos(18)
			C_BpNm2				= iCurColumnPos(19)			
			C_TempGlDt2			= iCurColumnPos(20)	
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
	
	With frm1.vspdData2
		.Row = .ActiveRow
		.Col = C_TempGlNo2
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
	
	With frm1.vspdData2
		.Row = .ActiveRow
		.Col = C_GLNo2
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
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5140ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData2
		.Row = .ActiveRow
		.Col = C_BatchNo2
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
			
		Case 1
			If frm1.txtAcctCd.className = parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "계정코드팝업"												' 팝업 명칭 
			arrParam(1) = "A_JNL_ACCT_ASSN A, A_ACCT B, A_ACCT_GP C"						' TABLE 명칭 
			arrParam(2) = Trim(strCode)														' Code Condition
			arrParam(3) = ""																' Name Condition
			arrParam(4) = " A.TRANS_TYPE = " & FilterVar("FN001", "''", "S") & "  "
			arrParam(4) = arrParam(4) & " AND ((A.JNL_CD = " & FilterVar("D1", "''", "S") & "  AND A.DR_CR_FG = " & FilterVar("CR", "''", "S") & " )  "
			arrParam(4) = arrParam(4) & "   OR (A.JNL_CD = " & FilterVar("D3", "''", "S") & "  AND A.DR_CR_FG = " & FilterVar("DR", "''", "S") & " )) "		
			arrParam(4) = arrParam(4) & " AND A.JNL_CD = " & FilterVar(frm1.cboNoteFg.Value, "''", "S")
			arrParam(4) = arrParam(4) & " AND A.ACCT_CD = B.ACCT_CD "
			arrParam(4) = arrParam(4) & " AND B.GP_CD = C.GP_CD AND B.DEL_FG <> " & FilterVar("Y", "''", "S") & "  "		' Where Condition			
			arrParam(5) = "계정코드"													' 조건필드의 라벨 명칭 
			
			arrField(0) = "A.ACCT_CD"														' Field명(0)
			arrField(1) = "B.ACCT_NM "														' Field명(1)
			arrField(2) = "B.GP_CD"															' Field명(2)
			arrField(3) = "C.GP_NM"															' Field명(3)
		 
			arrHeader(0) = "계정코드"													' Header명(0)
			arrHeader(1) = "계정코드명"													' Header명(1)
			arrHeader(2) = "그룹코드"													' Header명(2)
			arrHeader(3) = "그룹명"	
			
		Case 2	
			If frm1.txtBpCd.className = parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "거래처팝업"								' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER"									' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(4) = ""											' Where Condition
			arrParam(5) = "거래처"   
 
			arrField(0) = "BP_CD"									' Field명(0)
			arrField(1) = "BP_NM"									' Field명(1)    
			 
			arrHeader(0) = "거래처"									' Header명(0)
			arrHeader(1) = "거래처명"		
		
		Case 3
			If frm1.txtNoteNoFr.className = parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "어음번호팝업"						' 팝업 명칭 
			arrParam(1) = " f_note"									' TABLE 명칭 
			arrParam(2) = strCode						 				' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "note_fg = " & FilterVar(frm1.cboNoteFg.Value, "''", "S")		' Where Condition	
			arrParam(4) = arrParam(4) & "and note_sts = " & FilterVar("SM", "''", "S") & "  "		
			arrParam(5) = "어음번호"			
			arrField(0) = "note_no"									' Field명(0)
			arrField(1) = "issue_dt"								' Field명(1)
			arrField(2) = "due_dt"									' Field명(2)
			arrField(3) = "bp_cd"									' Field명(3)
    
			arrHeader(0) = "어음번호"							' Header명(0)
			arrHeader(1) = "발행일"								' Header명(1)			
			arrHeader(0) = "만기일"								' Header명(2)
			arrHeader(1) = "거래처"								' Header명(3)
		
		Case 4
			If frm1.txtNoteNoTo.className = parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "어음번호팝업"						' 팝업 명칭 
			arrParam(1) = " f_note"									' TABLE 명칭 
			arrParam(2) = strCode						 				' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "note_fg = " & FilterVar(frm1.cboNoteFg.Value, "''", "S")		' Where Condition	
			arrParam(4) = arrParam(4) & "and note_sts = " & FilterVar("SM", "''", "S") & "  "		
			arrParam(5) = "어음번호"			
			arrField(0) = "note_no"									' Field명(0)
			arrField(1) = "issue_dt"								' Field명(1)
			arrField(2) = "due_dt"									' Field명(2)
			arrField(3) = "bp_cd"									' Field명(3)
    
			arrHeader(0) = "어음번호"							' Header명(0)
			arrHeader(1) = "발행일"								' Header명(1)			
			arrHeader(0) = "만기일"								' Header명(2)
			arrHeader(1) = "거래처"								' Header명(3)
						
			
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

'------------------------------------------  EscPopup()  ------------------------------------------------
'	Name : EscPopup()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
		    Case "0"
				.txtBizAreaCd.focus
			Case "1"
				.txtAcctCd.focus
			Case "2"
				.txtBpCd.focus
			Case "3"
				.txtNoteNoFr.focus
			Case "4"
				.txtNoteNoTo.focus			

	    End Select
	End With
End Function

'------------------------------------------  SetPopup()  ------------------------------------------------
'	Name : SetPopup()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetPopup(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		     Case "0"
		     	.txtBizAreaCd.value = arrRet(0)
				.txtBizAreaNM.value = arrRet(1)
				.txtBizAreaCd.focus
			 Case "1"
				.txtAcctCd.value    = arrRet(0)
				.txtAcctNM.value    = arrRet(1)
				.txtAcctCd.focus
			Case "2"
				.txtBpCd.value = arrRet(0)
				.txtBpNM.value = arrRet(1)
				.txtBpCd.focus
			Case "3"
				.txtNoteNoFr.value  = arrRet(0)				
				lgNoteNoFr = arrRet(0)				
				.txtNoteNoFr.focus
			Case "4" 
				.txtNoteNoTo.value = arrRet(0)				
				lgNoteNoTo = arrRet(0)
				.txtNoteNoTo.focus 
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

'======================================================================================================
'	기능: 
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function MoveJmpClick()
Dim innerString

	Select Case gSelframeFlg
		Case TAB1						
			RefView.innerHTML = "<font color=""#777777"">회계전표&nbsp;|&nbsp;결의전표</font>"
			
			inputTypeView1.innerHTML =  " "
			inputTypeView2.innerHTML =  " "
			inputTypeView3.innerHTML =  " "	
			inputTypeView4.innerHTML =  " "								
											
			
		Case TAB2							
			RefView.innerHTML = "<A href='vbscript:OpenPopupGL()'>회계전표</A>&nbsp;|&nbsp;<A href='vbscript:OpenPopupTempGL()'>결의전표</A>"
			inputTypeView1.innerHTML =  "어음번호"
			
			innerString = " <INPUT TYPE=TEXT NAME=txtNoteNoFr ALT='검색시작어음번호' SIZE=20 MAXLENGTH=30 tag='11NXXU'><IMG SRC='../../../CShared/image/btnPopup.gif' STYLE='HEIGHT: 21px; WIDTH: 16px' NAME=btnNoteNoFr ALIGN=TOP TYPE=BUTTON onClick='vbscript:Call OpenPopup(frm1.txtNoteNoFr.value, 3)' > &nbsp;~&nbsp; "
			innerString = innerString & " <INPUT TYPE=TEXT NAME=txtNoteNoTo ALT='검색종료어음번호'  SIZE=20 MAXLENGTH=30 tag='11NXXU'><IMG SRC='../../../CShared/image/btnPopup.gif' STYLE='HEIGHT: 21px; WIDTH: 16px' NAME=btnNoteNoTo ALIGN=TOP TYPE=BUTTON onClick='vbscript:Call OpenPopup(frm1.txtNoteNoTo.value, 4)' ></TD>"			
			inputTypeView2.innerHTML = innerString
			inputTypeView3.innerHTML = "&nbsp;"
			inputTypeView4.innerHTML = "&nbsp;"			
			
			frm1.txtNoteNoFr.value = lgNoteNoFr
			frm1.txtNoteNoTo.value = lgNoteNoTo			
			
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

	Call cboNoteFg_OnChange()     
	
    Call InitSpreadSheet("A")																		'Setup the Spread sheet      
    Call InitSpreadSheet("B")    	
    
    Call InitComboBox
	
	Call ggoOper.SetReqAttr(frm1.txtBizAreaCd, "Q")
	Call ggoOper.SetReqAttr(frm1.txtBpCd, "Q")
	Call ggoOper.SetReqAttr(frm1.txtNoteNoFr, "Q")
	Call ggoOper.SetReqAttr(frm1.txtNoteNoTo, "Q")
	
	Call ClickTab1()																				'☜: Check Cookie	
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
    
    On Error Resume Next																			'☜: If process fails
    Err.Clear																						'☜: Clear error status

    FncQuery = False	
    Call InitVariables																				'☜: Processing is NG    
    
    
    If Not chkField(Document, "1") Then																'⊙: This function check indispensable field
		Exit Function
	End If
	
	If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If frm1.txtAcctCd.value = "" Then
		frm1.txtAcctNm.value = ""
	End If
	
	If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = ""
	End If

	
	
	Select Case gSelframeFlg
		Case TAB1
			ggoSpread.Source = Frm1.vspdData1
			If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
				IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")						'☜: "Will you destory previous data"		
				If IntRetCD = vbNo Then
					Exit Function
				End If
			End If    

			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.ClearSpreadData()
    
			If DbQuery("A") = False Then															'☜: Query db data
			   Exit Function
			End If
		Case TAB2
			ggoSpread.Source = Frm1.vspdData2
			If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
				IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")						'☜: "Will you destory previous data"		
				If IntRetCD = vbNo Then
					Exit Function
				End If
			End If    
	
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()
    
			If DbQuery("B") = False Then															'☜: Query db data
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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

	'------ Developer Coding part (Start )   -------------------------------------------------------------- 
	Call Parent.FncExport(Parent.C_MULTI)
	'------ Developer Coding part (End   )   -------------------------------------------------------------- 

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

	'------ Developer Coding part (Start )   -------------------------------------------------------------- 
	Call Parent.FncFind(Parent.C_MULTI, True)
	'------ Developer Coding part (End   )   -------------------------------------------------------------- 

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

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.SaveSpreadColumnInf()

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
 
    DbQuery = False                                                               '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
	
	
	
	Select Case pDirect
		Case "A"
			strVal = BIZ_PGM_ID		&"?txtMode="        & Parent.UID_M0001                      '☜: Query
			strVal = strVal			&"&txtMaxRows="		& Frm1.vspdData1.MaxRows				'☜: Max fetched data		
			
		Case "B"		
			strVal = BIZ_PGM_ID1	& "?txtMode="       & Parent.UID_M0001						'☜: Query
			strVal = strVal			& "&txtNoteNoFr="	& frm1.txtNoteNoFr.value			
			strVal = strVal			& "&txtNoteNoTo="	& frm1.txtNoteNoTo.value				'☜:
			strVal = strVal			& "&txtMaxRows="	& Frm1.vspdData2.MaxRows				'☜: Max fetched data
			strVal = strVal			& "&lgStrPrevKey="	& lgStrPrevKey              
	End Select 
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    strVal = strVal		& "&txtSttlFrDt="    & frm1.txtSttlFrDt.text							'☜:
    strVal = strVal		& "&txtSttlToDt="    & frm1.txtSttlToDt.text							'☜:
    strVal = strVal		& "&cboNoteFg="      & frm1.cboNoteFg.value								'☜:            
    strVal = strVal		& "&txtBizAreaCd="   & frm1.txtBizAreaCd.value							'☜:
    strVal = strVal		& "&txtBpCd="		 & frm1.txtBpCd.value								'☜:		
    strVal = strVal     & "&txtAcctCd="		 & frm1.txtAcctCd.value								'☜:    
    
    strVal = strVal		& "&txtShowBiz="     & frm1.txtShowBiz.value							'☜:
    strVal = strVal     & "&txtShowBp="		 & frm1.txtShowBp.value								'☜:
    strVal = strVal     & "&DispMeth="		 & Trim(frm1.RdoDiff.checked )                      '☜:

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                            '☜:  Run biz logic

    If Err.number = 0 Then
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	lgIntFlgMode      = Parent.OPMD_UMODE                                                '⊙: Indicates that current mode is Create mode

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    Select Case gSelframeFlg
		Case TAB1 
			Frm1.vspdData1.focus
		Case TAB2
			Frm1.vspdData2.focus
		
	End Select	
	
	Call DOSUM()	 
	Call SetToolbar("1100000000001111")                                           '☆: Developer must customize
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	'Call ggoOper.LockField(Document, "Q")

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
	Dim dbTotClsLocAmt
	Dim dbTotGlLocAmt
	DIm dbTotBatchLocAmt
	Dim dbTotTempGlLocAmt
	
	Select Case gSelframeFlg
		Case TAB1
			With frm1.vspdData1
				dbTotClsLocAmt	= FncSumSheet(frm1.vspdData1,C_SttlLocAmt1, 1, .MaxRows, False, -1, -1, "V")
				dbTotGlLocAmt		= FncSumSheet(frm1.vspdData1,C_GlLocAmt1, 1, .MaxRows, False, -1, -1, "V")				
				dbTotTempGlLocAmt	= FncSumSheet(frm1.vspdData1,C_TempGlLocAmt1, 1, .MaxRows, False, -1, -1, "V")
				dbTotBatchLocAmt	= FncSumSheet(frm1.vspdData1,C_DiffLocAmt1, 1, .MaxRows, False, -1, -1, "V")
				
				frm1.txtTotClsLocAmt1.text = UNIConvNumPCToCompanyByCurrency(dbTotClsLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				frm1.txtTotGlLocAmt1.text = UNIConvNumPCToCompanyByCurrency(dbTotGlLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    				
				frm1.txtTotTempGlLocAmt1.text = UNIConvNumPCToCompanyByCurrency(dbTotTempGlLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")        
				frm1.txtTotBatchLocAmt1.text = UNIConvNumPCToCompanyByCurrency(dbTotBatchLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    
			End With 
		Case TAB2
'			With frm1.vspdData2
'				dbTotClsLocAmt	= FncSumSheet(frm1.vspdData2,C_SttlLocAmt2, 1, .MaxRows, False, -1, -1, "V")
'				dbTotGlLocAmt		= FncSumSheet(frm1.vspdData2,C_GlLocAmt2, 1, .MaxRows, False, -1, -1, "V")				
'				dbTotTempGlLocAmt	= FncSumSheet(frm1.vspdData2,C_TempGlLocAmt2, 1, .MaxRows, False, -1, -1, "V")
'				dbTotBatchLocAmt	= FncSumSheet(frm1.vspdData2,C_DiffLocAmt2, 1, .MaxRows, False, -1, -1, "V")
'				
'				frm1.txtTotClsLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dbTotClsLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
'				frm1.txtTotGlLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dbTotGlLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    				
'				frm1.txtTotTempGlLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dbTotTempGlLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")        
'				frm1.txtTotBatchLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dbTotBatchLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    
'			End With 				
			 
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
	
	IF CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(frm1.txtBizAreaCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBizAreaNm.value= TRim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("124200","X","X","X")  	
		frm1.txtBizAreaCd.focus
	End IF

End Sub
'========================================================================================================
'   Event Name : txtAcctCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtAcctCd_onChange()

	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtAcctCd.value = "" Then Exit Sub
	
	IF CommonQueryRs("ACCT_NM", "A_ACCT ", " ACCT_CD=  " & FilterVar(frm1.txtAcctCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtAcctNm.value= TRim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("110100","X","X","X")  	
		frm1.txtAcctCd.focus
	End IF

End Sub


'========================================================================================================
'   Event Name : txtdealbpCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtBpCd_onChange()

	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtBpCd.value = "" Then Exit Sub
    	
	IF CommonQueryRs("BP_NM", "B_BIZ_PARTNER ", " BP_CD=  " & FilterVar(frm1.txtBpCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBpNm.value= TRim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("126100","X","X","X")  	
		frm1.txtBpCd.focus
	End IF

End Sub

'========================================================================================================
'   Event Name : cboNoteFg1_OnChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub cboNoteFg_OnChange()		
	frm1.txtAcctCd.value = ""
	frm1.txtAcctNm.value = ""
	
	If gSelframeFlg = TAB2 Then
		frm1.txtNoteNoFr.value = ""
		frm1.txtNoteNoTo.value = ""
	End If 
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			Case C_ZipCodePopUp
				.Col = Col - 1
				.Row = Row
				Call OpenZipCode(.Text,Row)
			End Select
		End If
    
	End With
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
    Select Case Col
         Case  C_StudyOnOffnM
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_StudyOnOffCd
                Frm1.vspdData.value = iDx
         Case Else
    End Select    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1111111111")    
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If

        Exit Sub
    End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
'7.1. SpreadSheet의 이벤트명[DblClick]에 로직 추가 
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
  
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery("B") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End if
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'=======================================================================================================
'   Event Name : txtLoanDtFr_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtSttlFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtSttlToDt.Focus
	   Call MainQuery
	End If   
End Sub
'=======================================================================================================
'   Event Name : txtLoanDtTo_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtSttlToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtSttlFrDt.Focus
	   Call MainQuery
	End If   
End Sub

'========================================================================================================
' Name : txtSttlFrDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtSttlFrDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtSttlFrDt.Action = 7                                    ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("M")	
       frm1.txtSttlFrDt.Focus
    End If
End Sub
'========================================================================================================
' Name : txtSttlToDt_DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtSttlToDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtSttlToDt.Action = 7                                    ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("M")	
       frm1.txtSttlToDt.Focus
    End If
End Sub

'========================================================================================================
' Name : txtGlFrDt_DblClick
' Desc : developer describe this line
'========================================================================================================
'Sub txtGlFrDt_DblClick(Button)
'    If Button = 1 Then
'       frm1.txtGlFrDt.Action = 7                                    ' 7 : Popup Calendar ocx
'       Call SetFocusToDocument("M")	
'       frm1.txtGlFrDt.Focus
'    End If
'End Sub
'========================================================================================================
' Name : txtGlToDtt_DblClick
' Desc : developer describe this line
'========================================================================================================
'Sub txtGlToDt_DblClick(Button)
'    If Button = 1 Then
'       frm1.txtGlToDt.Action = 7                                    ' 7 : Popup Calendar ocx
'       Call SetFocusToDocument("M")	
'       frm1.txtGlToDt.Focus
'    End If
'End Sub

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
	Else
		frm1.txtShowBiz.value = "N"	
		frm1.txtBizAreaCd.value = ""
		frm1.txtBizAreaNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCd, "Q")		
		
		ggoSpread.Source = frm1.vspdData1
		Call ggoSpread.SSSetColHidden(C_BizAreaCd1,C_BizAreaNm1,True)			

		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.SSSetColHidden(C_BizAreaCd2,C_BizAreaNm2,True)
	End If
End Sub

'========================================================================================================
' Name : chkShowBp_onchange
' Desc : 
'========================================================================================================
Sub chkShowBp_onchange()
	If frm1.chkShowBp.checked = True Then
		frm1.txtShowBp.value = "Y"
		Call ggoOper.SetReqAttr(frm1.txtBpCd, "D")	
		
		ggoSpread.Source = frm1.vspdData1
		Call ggoSpread.SSSetColHidden(C_BpCd1,C_BpNm1,FALSE)

		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.SSSetColHidden(C_BpCd2,C_BpNm2,FALSE)	
	Else
		frm1.txtShowBp.value = "N"	
		frm1.txtBpCd.value = ""
		frm1.txtBpNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtBpCd, "Q")	
		
		ggoSpread.Source = frm1.vspdData1
		Call ggoSpread.SSSetColHidden(C_BpCd1,C_BpNm1,True)

		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.SSSetColHidden(C_BpCd2,C_BpNm2,True	)	
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
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>만기어음처리정보</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
		    		<TD WIDTH=* align=right><span id="RefView"><A href="vbscript:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;
		    													<A href="vbscript:OpenPopupTempGL()">결의전표</A> </SPAN></TD>					
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
									<TD CLASS="TD5" NOWRAP>결제일자</TD>
									<TD CLASS="TD6" NOWRAP>
									    <script language =javascript src='./js/a5456ma1_fpDateTime1_txtSttlFrDt.js'></script>
									    &nbsp;~&nbsp;
									    <script language =javascript src='./js/a5456ma1_fpDateTime2_txtSttlToDt.js'></script>
									</TD>		
									<TD CLASS=TD5 NOWRAP>어음구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg" NAME="cboNoteFg" ALT="어음구분" STYLE="WIDTH: 132px" tag="12X"></SELECT></TD>
								</TR>
								<TR>		
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME=txtBizAreaCd ALT="사업장코드" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnBizAreaCd ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtBizAreaCd.value, 0)">
														<INPUT TYPE=TEXT NAME=txtBizAreaNm ALT="사업장명" SIZE="18" style="HEIGHT: 20px; " tag="14" >
														<INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkShowBiz ID=chkShowBiz tag="1" onclick=chkShowBiz_onchange()></TD>
									<TD CLASS=TD5 NOWRAP>계정코드</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME=txtAcctCd ALT="계정코드" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnAcctCd ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtACctCd.value, 1)">
														<INPUT TYPE=TEXT NAME=txtAcctNm ALT="계정코드명" SIZE="18" style="HEIGHT: 20px; " tag="14" ></TD>														
								</TR>						
								<TR>
									<TD CLASS=TD5 NOWRAP>거래처</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME=txtBpCd ALT="거래처코드" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=20 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnBpCd ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtBpCd.value, 2)">
														<INPUT TYPE=TEXT NAME=txtBpNm ALT="거래처명" SIZE="18" style="HEIGHT: 20px; " tag="14" >
														<INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkShowBp ID=chkShowBp tag="1" onclick=chkShowBp_onchange()></TD>
									<TD CLASS=TD5 NOWRAP>조회방식</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoType" ID="RdoDiff" VALUE="S" TAG="11" ><LABEL FOR="rdoReport1">차이금액</LABEL>&nbsp;&nbsp
														 <INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoType" ID="RdoTotal" VALUE="D" TAG="11" Checked><LABEL FOR="rdoReport2">전체금액</LABEL></TD>		
								</TR>								
								<TR>
									<TD CLASS=TD5 ID=inputTypeView1>어음번호</TD>								
									<TD CLASS=TD6 ID=inputTypeView2><INPUT TYPE=TEXT NAME=txtNoteNoFr ALT="검색시작어음번호" STYLE="HEIGHT: 20px; WIDTH: 80px" SIZE=30 MAXLENGTH=30 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnNoteNoFr ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtNoteNoFr.value, 3)" > &nbsp;~&nbsp;																	
																	<INPUT TYPE=TEXT NAME=txtNoteNoTo ALT="검색종료어음번호" STYLE="HEIGHT: 20px; WIDTH: 80px" SIZE=30 MAXLENGTH=30 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnNoteNoTo ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtNoteNoTo.value, 4)"  ismap></TD>
									<TD CLASS="TD5" ID=inputTypeView3>&nbsp;</TD>
									<TD CLASS="TD6" ID=inputTypeView4>&nbsp;</TD>																	
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
										<script language =javascript src='./js/a5456ma1_OBJECT1_vspdData1.js'></script>
									</TD>
								</TR>
								<TR>
									<TD HEIGHT=40 WIDTH=25%>
										<FIELDSET CLASS="CLSFLD">
											<TABLE  <%=LR_SPACE_TYPE_20%>>
												<TR>
													<TD CLASS="TDt" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;결제금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5456ma1_OBJECT1_txtTotClsLocAmt1.js'></script>
													</TD>
													<TD class=TDT STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TDt" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;차이금액합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5456ma1_OBJECT1_txtTotBatchLocAmt1.js'></script>
													</TD>
												</TR>
												<TR>
													<TD CLASS="TDt" NOWRAP>회계전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5456ma1_OBJECT1_txtTotGlLocAmt1.js'></script>
													</TD>
													<TD class=TDT STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TDt" NOWRAP>결의전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5456ma1_OBJECT1_txtTotTempGlLocAmt1.js'></script>
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
										<script language =javascript src='./js/a5456ma1_OBJECT1_vspdData2.js'></script>
									</TD>
								</TR>
								<TR>
									<TD HEIGHT=40 WIDTH=25%>
										<FIELDSET CLASS="CLSFLD">
											<TABLE  <%=LR_SPACE_TYPE_20%>>
												<TR>
													<TD CLASS="TDt" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;결제금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5456ma1_OBJECT1_txtTotClsLocAmt2.js'></script>
													</TD>
													<TD class=TDT STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TDt" NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;차이금액합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5456ma1_OBJECT1_txtTotBatchLocAmt2.js'></script>
													</TD>
												</TR>
												<TR>
													<TD CLASS="TDt" NOWRAP>회계전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5456ma1_OBJECT1_txtTotGlLocAmt2.js'></script>
													</TD>
													<TD class=TDT STYLE="WIDTH : 0px;"></TD>												
													<TD CLASS="TDt" NOWRAP>결의전표금액(자국)합계</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/a5456ma1_OBJECT1_txtTotTempGlLocAmt2.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
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
