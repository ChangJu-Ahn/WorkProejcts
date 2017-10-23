<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5962MA1
'*  4. Program Name         : 상여금월별현황 
'*  5. Program Desc         : 회계관리 / 월차계산 / 
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


<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================

Const BIZ_PGM_ID      = "a5962mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

Dim C_BUNAME  
Dim C_BTN     
Dim C_TYPE    
Dim C_BTN1    
Dim C_YEAR    
Dim C_ONE     
Dim C_TWO     
Dim C_THREE   
Dim C_FOUR    
Dim C_FIVE    
Dim C_SIX     
Dim C_SEVEN   
Dim C_EIGHT   
Dim C_NINE    
Dim C_TEN     
Dim C_ELEVEN  
Dim C_TWEL    
Dim C_BUCODE  
Dim C_ORG
Dim C_INTERNAL_CD	'2003.1.13 문희정추가 
Dim C_BIZ_AREA_CD   '2003.1.13 문희정추가  
Dim C_TYPECD  

'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #2
'--------------------------------------------------------------------------------------------------------

Dim C_BUNAME1  
Dim C_BTN2     
Dim C_TYPE1    
Dim C_BTN21    
Dim C_YEAR1    
Dim C_ONE1     
Dim C_TWO1     
Dim C_THREE1   
Dim C_FOUR1    
Dim C_FIVE1    
Dim C_SIX1     
Dim C_SEVEN1   
Dim C_EIGHT1   
Dim C_NINE1    
Dim C_TEN1     
Dim C_ELEVEN1  
Dim C_TWEL1    
Dim C_BUCODE1  
Dim C_ORG1     
Dim C_TYPECD1  
																		'Column constant for Spread Sheet  
Const C_SHEETMAXROWS_D1 = 2                                          '☜: Fetch count at a time

Const COOKIE_SPLIT      = 4877	                                      'Cookie Split String
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================

Dim lgIsOpenPop
Dim IsOpenPop  

<%
lsSvrDate   = GetSvrDate                                                                  'Get Server DB Date
%>



'========================================================================================================
Sub InitSpreadPosVariables()
	C_BUNAME   =  1                                                 'Column ant for Spread Sheet 
	C_BTN      =  2                                                 'Column ant for Spread Sheet 
	C_TYPE     =  3                                                 'Column ant for Spread Sheet
	C_BTN1     =  4                                                 'Column ant for Spread Sheet 
	C_YEAR     =  5															
	C_ONE      =  6
	C_TWO      =  7
	C_THREE    =  8 
	C_FOUR     =  9 
	C_FIVE     =  10
	C_SIX      =  11
	C_SEVEN    =  12 
	C_EIGHT    =  13
	C_NINE     =  14
	C_TEN      =  15
	C_ELEVEN   =  16
	C_TWEL     =  17
	C_BUCODE   =  18                                                'Column ant for Spread Sheet 
	C_ORG      =  19 
    C_INTERNAL_CD = 20
	C_BIZ_AREA_CD = 21	 
	C_TYPECD   =  22                                                'Column ant for Spread Sheet 


	C_BUNAME1   =  1                                                 'Column ant for Spread Sheet 
	C_BTN2      =  2                                                 'Column ant for Spread Sheet 
	C_TYPE1     =  3                                                 'Column ant for Spread Sheet
	C_BTN21     =  4                                                 'Column ant for Spread Sheet 
	C_YEAR1     =  5															
	C_ONE1      =  6
	C_TWO1      =  7 
	C_THREE1    =  8 
	C_FOUR1     =  9 
	C_FIVE1     =  10
	C_SIX1      =  11 
	C_SEVEN1    =  12 
	C_EIGHT1    =  13
	C_NINE1     =  14
	C_TEN1      =  15
	C_ELEVEN1   =  16
	C_TWEL1     =  17
	C_BUCODE1   =  18                                                'Column ant for Spread Sheet 
	C_ORG1      =  19      
	C_TYPECD1   =  20                                                'Column ant for Spread Sheet 
																		'Column constant for Spread S

End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode       = Parent.OPMD_CMODE
	lgBlnFlgChgValue   = False
    lgStrPrevKey       = ""
    lgStrPrevKeyIndex  = ""
    lgStrPrevKeyIndex1 = ""
    lgSortKey          = 1
		
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
		Dim strYear,strMonth,strDay
	
    frm1.fpdtWk_yymm.focus	
    Call ggoOper.FormatDate(frm1.fpdtWk_yymm, Parent.gDateFormat, 3)
    Call ExtractDateFrom("<%=lsSvrDate%>",Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
    frm1.fpdtWk_yymm.Year = strYear    

	
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream = frm1.fpdtWk_yymm.year & Parent.gColSep       'You Must append one character(Parent.gColSep)
    lgKeyStream = lgKeyStream & frm1.txtpayCD.value & Parent.gColSep
    lgKeyStream = lgKeyStream & frm1.txtFactoryCD.value & Parent.gColSep  
End Sub        


'========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

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
	
    	.MaxCols   = C_TYPECD + 1                                                 ' ☜:☜: Add 1 to Maxcols
		.Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:
		
		ggoSpread.Source= frm1.vspdData
		ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos("A")
        ggoSpread.SSSetEdit  C_BUNAME , "부서명", 20,   ,, 50,2
       ggoSpread.SSSetButton  C_BTN
       ggoSpread.SSSetEdit  C_TYPE   , "계정타입",10,   ,, 50,2       
       ggoSpread.SSSetButton  C_BTN1  
       ggoSpread.SSSetFloat C_YEAR   , "년계"  , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_ONE    , "1월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_TWO    , "2월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_THREE  , "3월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_FOUR   , "4월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_FIVE   , "5월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_SIX    , "6월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_SEVEN  , "7월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_EIGHT  , "8월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_NINE   , "9월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_TEN    , "10월"  , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_ELEVEN , "11월"  , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_TWEL   , "12월"  , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetEdit  C_BUCODE , "부서",10,   ,, 50,2
       ggoSpread.SSSetEdit  C_ORG    , "조직코드",10,   ,, 50,2
       ggoSpread.SSSetEdit  C_INTERNAL_CD	,	"내부부서코드"  ,10,,,10,2
       ggoSpread.SSSetEdit  C_BIZ_AREA_CD   ,   "사업장"   	 ,10,,,10,2       
       ggoSpread.SSSetEdit  C_TYPECD  , "계정명",10,   ,, 50,2	
       
	   call ggoSpread.MakePairsColumn(C_BUNAME,C_BTN)
	   call ggoSpread.MakePairsColumn(C_TYPE,C_BTN1)	       
       Call ggoSpread.SSSetColHidden(C_BUCODE,C_BUCODE,True)
       Call ggoSpread.SSSetColHidden(C_ORG,C_ORG,True)
       Call ggoSpread.SSSetColHidden(C_TYPECD,C_TYPECD,True)
       Call ggoSpread.SSSetColHidden(C_INTERNAL_CD,C_INTERNAL_CD,True)
       Call ggoSpread.SSSetColHidden(C_BIZ_AREA_CD,C_BIZ_AREA_CD,True)
       
		
		
	.ReDraw = true
    
    End With
 
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.Spreadinit "V20021103",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
        
	With frm1.vspdData1
	
	   .ReDraw = false
	   
       .MaxCols   = C_TYPECD1 + 1                                                 ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

       ggoSpread.Source= frm1.vspdData1
    ggoSpread.ClearSpreadData
    
        Call GetSpreadColumnPos("B")
        ggoSpread.SSSetEdit  C_BUNAME1 , "사업장명", 20,   ,, 50,2
       ggoSpread.SSSetButton  C_BTN2
       ggoSpread.SSSetEdit  C_TYPE1   , "",10,   ,, 50,2       
       ggoSpread.SSSetButton  C_BTN21  
       ggoSpread.SSSetFloat C_YEAR1   , "년계"  , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_ONE1    , "1월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_TWO1    , "2월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_THREE1  , "3월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_FOUR1   , "4월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_FIVE1   , "5월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_SIX1    , "6월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_SEVEN1  , "7월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_EIGHT1  , "8월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_NINE1   , "9월"   , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_TEN1    , "10월"  , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_ELEVEN1 , "11월"  , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetFloat C_TWEL1   , "12월"  , 10, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
       ggoSpread.SSSetEdit  C_BUCODE1 , "",10,   ,, 50,2
       ggoSpread.SSSetEdit  C_ORG1    , "",10,   ,, 50,2
       ggoSpread.SSSetEdit  C_TYPECD1 , "",10,   ,, 50,2	
       
       
	   call ggoSpread.MakePairsColumn(C_BUNAME1,C_BTN2)
	   call ggoSpread.MakePairsColumn(C_TYPE1,C_BTN21)
	          
       Call ggoSpread.SSSetColHidden(C_BUCODE1,C_BUCODE1,True)
       Call ggoSpread.SSSetColHidden(C_ORG1,C_ORG1,True)
       Call ggoSpread.SSSetColHidden(C_TYPECD1,C_TYPECD1,True)
		
	.ReDraw = true

    End With

    Call SetSpreadLock 

End Sub


'======================================================================================================
Sub SetSpreadLock()
                  With frm1
    				ggoSpread.Source = Frm1.vspdData
                      .vspdData.ReDraw = False
                        ggoSpread.Spreadlock    C_BUNAME, -1, C_BUNAME  
                        ggoSpread.Spreadlock    C_BTN    , -1, C_BTN
						ggoSpread.Spreadlock    C_TYPE   , -1, C_TYPE
						ggoSpread.Spreadlock    C_BTN1   , -1, C_BTN1
						ggoSpread.Spreadlock    C_YEAR   , -1, C_YEAR                            
						ggoSpread.SpreadUnlock    C_ONE    , -1, C_ONE      
						ggoSpread.SpreadUnlock    C_TWO    , -1, C_TWO      
						ggoSpread.SpreadUnlock    C_THREE  , -1, C_THREE      
						ggoSpread.SpreadUnlock    C_FOUR   , -1, C_FOUR      
						ggoSpread.SpreadUnlock    C_FIVE   , -1, C_FIVE            
						ggoSpread.SpreadUnlock    C_SIX    , -1, C_SIX
						ggoSpread.SpreadUnlock    C_SEVEN  , -1, C_SEVEN      
						ggoSpread.SpreadUnlock    C_EIGHT  , -1, C_EIGHT      
						ggoSpread.SpreadUnlock    C_NINE   , -1, C_NINE      
	  					ggoSpread.SpreadUnlock    C_TEN    , -1, C_TEN      
      					ggoSpread.SpreadUnlock    C_ELEVEN , -1, C_ELEVEN            
      					ggoSpread.SpreadUnlock    C_TWEL   , -1, C_TWEL    
      					ggoSpread.SpreadLock	.vspdData.MaxCols, -1,.vspdData.MaxCols                       
                      .vspdData.ReDraw = True
                      
                  End With
                  
                  With Frm1
                  ggoSpread.Source = Frm1.vspdData1
                      .vspdData1.ReDraw = False                       
						    ggoSpread.SSSetProtected    C_BUNAME1, -1, C_BUNAME1  
							ggoSpread.SSSetProtected    C_BTN2    , -1, C_BTN2
							ggoSpread.SSSetProtected    C_TYPE1   , -1, C_TYPE1
							ggoSpread.SSSetProtected    C_BTN21   , -1, C_BTN21					
                            ggoSpread.SSSetProtected    C_YEAR1   , -1, C_YEAR1                                                        
                            ggoSpread.SSSetProtected    C_ONE1    , -1, C_ONE1      
                            ggoSpread.SSSetProtected    C_TWO1    , -1, C_TWO1      
                            ggoSpread.SSSetProtected    C_THREE1  , -1, C_THREE1      
	                        ggoSpread.SSSetProtected    C_FOUR1   , -1, C_FOUR1      
                            ggoSpread.SSSetProtected    C_FIVE1   , -1, C_FIVE1            
                            ggoSpread.SSSetProtected    C_SIX1    , -1, C_SIX1
                            ggoSpread.SSSetProtected    C_SEVEN1  , -1, C_SEVEN1      
                            ggoSpread.SSSetProtected    C_EIGHT1  , -1, C_EIGHT1      
                            ggoSpread.SSSetProtected    C_NINE1   , -1, C_NINE1      
	                        ggoSpread.SSSetProtected    C_TEN1    , -1, C_TEN1      
                            ggoSpread.SSSetProtected    C_ELEVEN1 , -1, C_ELEVEN1            
                            ggoSpread.SSSetProtected    C_TWEL1   , -1, C_TWEL1 
                            ggoSpread.SpreadLock	.vspdData1.MaxCols, -1,.vspdData1.MaxCols           
                      .vspdData1.ReDraw = True
                  End With
   
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
  	
            With Frm1
                 ggoSpread.Source = Frm1.vspdData
                    .vspdData.ReDraw = False
				        ggoSpread.SSSetRequired    C_BUNAME , pvStartRow, pvEndRow
						ggoSpread.SSSetRequired    C_TYPE   , pvStartRow, pvEndRow
						ggoSpread.SSSetProtected   C_YEAR   , pvStartRow, pvEndRow
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
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData1.Col    = iDx
              Frm1.vspdData1.Row    = iRow
              Frm1.vspdData1.Action = 0 ' go to 
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
			C_BUNAME		= iCurColumnPos(1)
			C_BTN			= iCurColumnPos(2)
			C_TYPE			= iCurColumnPos(3)    
			C_BTN1			= iCurColumnPos(4)
			C_YEAR			= iCurColumnPos(5)
			C_ONE			= iCurColumnPos(6)
			C_TWO			= iCurColumnPos(7)
			C_THREE			= iCurColumnPos(8)
			C_FOUR			= iCurColumnPos(9)
			C_FIVE			= iCurColumnPos(10)
			C_SIX			= iCurColumnPos(11)
			C_SEVEN			= iCurColumnPos(12)
			C_EIGHT			= iCurColumnPos(13)
			C_NINE			= iCurColumnPos(14)
			C_TEN			= iCurColumnPos(15)
			C_ELEVEN		= iCurColumnPos(16)
			C_TWEL			= iCurColumnPos(17)
			C_BUCODE		= iCurColumnPos(18)
			C_ORG			= iCurColumnPos(19)
			C_INTERNAL_CD	= iCurColumnPos(20)
			C_BIZ_AREA_CD	= iCurColumnPos(21)
			C_TYPECD		= iCurColumnPos(22)
			
			
     Case "B"
		ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_BUNAME1		= iCurColumnPos(1)
			C_BTN2			= iCurColumnPos(2)
			C_TYPE1			= iCurColumnPos(3)    
			C_BTN21			= iCurColumnPos(4)
			C_YEAR1			= iCurColumnPos(5)
			C_ONE1			= iCurColumnPos(6)
			C_TWO1			= iCurColumnPos(7)
			C_THREE1		= iCurColumnPos(8)
			C_FOUR1			= iCurColumnPos(9)
			C_FIVE1			= iCurColumnPos(10)
			C_SIX1			= iCurColumnPos(11)
			C_SEVEN1		= iCurColumnPos(12)
			C_EIGHT1		= iCurColumnPos(13)
			C_NINE1			= iCurColumnPos(14)
			C_TEN1			= iCurColumnPos(15)
			C_ELEVEN1		= iCurColumnPos(16)
			C_TWEL1			= iCurColumnPos(17)
			C_BUCODE1		= iCurColumnPos(18)
			C_ORG1			= iCurColumnPos(19)
			C_TYPECD1		= iCurColumnPos(20)
     
    End Select    
End Sub
	   
'========================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
	Call InitVariables
    Call SetDefaultVal

	Call SetToolbar("1100111100101111")                                         '☆: Developer must customize	'------ Developer Coding part (End )   End Sub
	
End Sub
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
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
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			         '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents  Field

    If Not chkField(Document, "1") Then									         '⊙: This function check indispensable field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call InitVariables															 '⊙: Initializes local global variables
    lgCurrentSpd = "M"
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
End Function
	
'========================================================================================================
Function FncNew()
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
Function FncDelete()
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

       
    
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = Frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call MakeKeyStream("x")
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

'    If Trim(lgActiveSpd) = "" Then
'       lgActiveSpd = "M"
'    End If

'    Select Case UCase(Trim(lgActiveSpd))
'       Case  "M"
                 If Frm1.vspdData.MaxRows < 1 Then
                    Exit Function
                 End If
    
                 With Frm1
    
		              If .vspdData.ActiveRow > 0 Then
			              .vspdData.ReDraw = False
		
                          ggoSpread.Source = .vspdData	
                          ggoSpread.CopyRow
                          SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 
						frm1.vspdData.Row = frm1.vspdData.ActiveRow
						frm1.vspdData.col = C_BUNAME 
						frm1.vspdData.text = ""
						frm1.vspdData.Row = frm1.vspdData.ActiveRow
						frm1.vspdData.col = C_BUCODE 
						frm1.vspdData.text = ""
			
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
                        .vspdData.ReDraw = True
                        .vspdData.focus
		              End If
	              End With


    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                           '☜: Processing is NG
    Err.Clear                                                                   '☜: Clear err status
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  

    Set gActiveElement = document.ActiveElement   
    FncCancel = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'====================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
	
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
    
    If Not chkField(Document, "1") Then									         '⊙: This function check indispensable field
       Exit Function
    End If


	  With Frm1
			 .vspdData.ReDraw = False
			 .vspdData.Focus
			  ggoSpread.Source = .vspdData
			  ggoSpread.InsertRow .vspdData.ActiveRow, imRow
			  SetSpreadColor .vspdData.ActiveRow , .vspdData.ActiveRow + imrow -1
			 .vspdData.ReDraw = True
	  End With
    
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
   If Trim(lgActiveSpd) = "" Then
      lgActiveSpd = "M"
   End If
       

	Select Case UCase(Trim(lgActiveSpd))
	Case  "M"
	If Frm1.vspdData.MaxRows < 1 then
		Exit function
	End if	
	With Frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
	End With
	End Select

    If Frm1.vspdData1.MaxRows < 1 then
       Exit function
    End if	

    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                         '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False
    Err.Clear
	Call Parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================================
Function FncPrev() 
    FncPrev = False
    Err.Clear
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True
End Function

'========================================================================================================
Function FncNext() 
    FncNext = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncNext = True
End Function

'========================================================================================================
Function FncExcel() 
    FncExcel = False
    Err.Clear
	Call Parent.FncExport(Parent.C_MULTI)
    FncExcel = True
End Function

'========================================================================================================
Function FncFind() 
    FncFind = False
    Err.Clear
	Call Parent.FncFind(Parent.C_MULTI, True)
    FncFind = True
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
	if UCase(gActiveSpdSheet.name ) = "VSPDDATA" then
		ggoSpread.Source = gActiveSpdSheet
	    Call ggoSpread.RestoreSpreadInf()
	    Call InitSpreadSheet()      
	'    Call InitComboBox
		ggoSpread.Source = gActiveSpdSheet
		Call ggoSpread.ReOrderingSpreadData()
		lgCurrentSpd = "S"
		Call MakeKeyStream("X")
		Call DbQuery()
	else
		ggoSpread.Source = gActiveSpdSheet
	    Call ggoSpread.RestoreSpreadInf()
	    Call InitSpreadSheet()      
	    ggoSpread.Source = gActiveSpdSheet
		Call ggoSpread.ReOrderingSpreadData()
		lgCurrentSpd = "M1"
		Call MakeKeyStream("X")
		Call DbQuery()
	end if
'	Call InitData()
End Sub

'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	ggoSpread.Source = Frm1.vspdData1
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function DbQuery()
    Dim strVal

    Err.Clear                                                                        '☜: Clear err status
    DbQuery = False                                                                  '☜: Processing is NG
    
       if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
    
	If lgCurrentSpd = "M" or lgCurrentSpd = "M1" Then
       strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex              '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '☜: Max fetched data
    Else   

		
       strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex1             '☜: Next key tag
       strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D1)        '☜: Max fetched data at a time
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData1.MaxRows         '☜: Max fetched data

    End If   
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic

    DbQuery = True                                                                   '☜: Processing is NG

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
    Err.Clear                                                                      '☜: Clear err status
    DbSave = False                                                                 '☜: Processing is NG

    Call LayerShowHide(1)                                                          '☜: Show Processing Message
		
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
  	With Frm1
		.txtMode.value      = Parent.UID_M0002                                            '☜: Delete
		.txtKeyStream.value = lgKeyStream
	End With

    strVal  = ""
    strDel  = ""
    lGrpCnt = 1

	With Frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0
            
           Select Case .vspdData.Text 
               Case ggoSpread.InsertFlag                                      '☜: Update

                                                     strVal = strVal & "C" & Parent.gColSep                    '0
                                                     strVal = strVal & lRow & Parent.gColSep                    '1
                                                     strVal = strVal & frm1.fpdtWk_yymm.text & Parent.gColSep   '2
                                                     strVal = strVal & frm1.txtpayCD.value & Parent.gColSep      '3
 '                                                    strVal = strVal & frm1.txtFactoryCD.value & Parent.gColSep  '                                                     
                    
                    .vspdData.Col = C_BUCODE      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '4
                    .vspdData.Col = C_TYPECD      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '5
                    
                    .vspdData.Col = C_ORG         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '6
                    .vspdData.Col = C_INTERNAL_CD         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep  '7
                    .vspdData.Col = C_BIZ_AREA_CD         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep  '8
                    
                    .vspdData.Col = C_ONE         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '9
                    
                    .vspdData.Col = C_TWO         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '10
                    .vspdData.Col = C_THREE       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '11
                    
                    .vspdData.Col = C_FOUR        : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '12
                    .vspdData.Col = C_FIVE        : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '13
                    .vspdData.Col = C_SIX         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '14
                    .vspdData.Col = C_SEVEN       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '15
                    .vspdData.Col = C_EIGHT       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '16
                    .vspdData.Col = C_NINE        : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '17
                    .vspdData.Col = C_TEN         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '18
                    .vspdData.Col = C_ELEVEN      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '19
                    .vspdData.Col = C_TWEL        : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   '20                    
                    lGrpCnt = lGrpCnt + 1
	
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                     strVal = strVal & "U" & Parent.gColSep					'0
                                                     strVal = strVal & lRow & Parent.gColSep                    '1                                                                                      
                                                     strVal = strVal & frm1.fpdtWk_yymm.text & Parent.gColSep	'2
                                                     strVal = strVal & frm1.txtpayCD.value & Parent.gColSep		'3
                    .vspdData.Col = C_BUCODE  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			'4
                    .vspdData.Col = C_TYPECD  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			'5
                    .vspdData.Col = C_ORG	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep			'6
                    .vspdData.Col = C_INTERNAL_CD         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep    '7
                    .vspdData.Col = C_BIZ_AREA_CD         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep    '8
                    .vspdData.Col = C_ONE     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_TWO     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_THREE   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_FOUR    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_FIVE    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_SIX     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_SEVEN   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_EIGHT   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_NINE    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_TEN     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_ELEVEN  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_TWEL    : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep                    
                    lGrpCnt = lGrpCnt + 1
               
               
               
               Case ggoSpread.DeleteFlag                                      '☜: Delete
               
                                                     strDel = strDel & "D" & Parent.gColSep
                                                     strDel = strDel & lRow & Parent.gColSep
                                                     strDel = strDel & frm1.fpdtWk_yymm.text & Parent.gColSep   '2
                                                     strDel = strDel & frm1.txtpayCD.value & Parent.gColSep      '3
					.vspdData.Col = C_BUCODE      :  strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep   '4
                    .vspdData.Col = C_TYPECD      :  strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep   '5                                                                                           
                    lGrpCnt = lGrpCnt + 1
                    
           End Select
           
       Next
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

		
	End With
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                                '☜: Processing is OK
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Sub DbQueryOk()
	Dim lRow
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("1100111100111111") 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
 	If lgCurrentSpd = "M" Then
 	'Call SetSpreadColor()
	   lgCurrentSpd   = "S"
      ' Call InitData()
	   Call MakeKeyStream("X")
       Call DbQuery()      
    Else
      ' Call InitData()
	End If	
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
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
   Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData1.MaxRows = 0
    lgCurrentSpd = "M"

	Call SetToolbar("1100111100111111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DBQuery()
    Set gActiveElement = document.ActiveElement   
End Sub


	
'========================================================================================================
Sub DbDeleteOk()
End Sub

'======================================================================================================
'	Name : OpenCurrencypay()
'	Description : Major PopUp
'=======================================================================================================%>
Function OpenCurrencypay()
	Dim arrRet
	Dim arrParam(6), arrField(5), arrHeader(5)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	arrParam(0) = "상여종류 팝업"		        <%' 팝업 명칭 %>
	arrParam(1) = "a_bonus_base a,b_minor b"	 	<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtpayCD.Value	    		<%' Code Cindition%>
	arrParam(3) = ""								<%' Name Condition%>
	If Trim(frm1.fpdtWk_yymm.Text) <> "" Then
		arrParam(4) = "b.major_cd = " & FilterVar("H0040", "''", "S") & "  AND a.pay_type = b.minor_cd and a.yyyy = " & FilterVar(frm1.fpdtWk_yymm.Text, "''", "S")
	Else
		IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data.
		IsOpenPop = False
        Exit Function
    End If  	
	'arrParam(4) = "major_cd = 'H0040'AND (MINOR_CD>= '2' AND MINOR_CD <='9')"
    arrParam(5) = "상여종류"

    arrField(0) = "b.minor_CD"					    <%' Field명(0)%>
    arrField(1) = "b.minor_NM"	     			    <%' Field명(1)%>

    arrHeader(0) = "상여코드"				<%' Header명(0)%>
    arrHeader(1) = "상여코드명"	       	<%' Header명(1)%>


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtpayCD.focus
		Exit Function
	Else
		Call SetMinor(arrRet)

	End If

End Function


'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetMinor(Byval arrRet)
	With frm1
		.txtpayCD.focus
		.txtpayCD.value = arrRet(0)
		.txtpayNM.value = arrRet(1)
	End With

End Function



'======================================================================================================
'	Name : OpenCurrency()
'	Description : Major PopUp
'=======================================================================================================%>
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(6), arrField(5), arrHeader(5)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	arrParam(0) = "사업장-팝업"		    	    <%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_AREA"					 	<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtFactoryCD.Value	    	<%' Code Cindition%>
	arrParam(3) = ""								<%' Name Condition%>
	arrParam(4) = ""
    arrParam(5) = "사업장"

    arrField(0) = "BIZ_AREA_CD"					<%' Field명(0)%>
    arrField(1) = "BIZ_AREA_NM"	     			<%' Field명(1)%>

    arrHeader(0) = "사업장 코드"				<%' Header명(0)%>
    arrHeader(1) = "사업장명"				<%' Header명(1)%>


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtFactoryCD.focus
		Exit Function
	Else
		Call SetMajor(arrRet)

	End If

End Function

'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetMajor(Byval arrRet)
	With frm1
		.txtFactoryCD.focus
		.txtFactoryCD.value = arrRet(0)
		.txtFactoryNM.value = arrRet(1)
	End With

End Function



'===========================================================================
' Function Name : OpenDept
' Function Desc : OpenCode Reference Popup
'===========================================================================
Function OpenDept(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5)
	Dim strYear, strMonth, strDay, strDate
	'------ Developer Coding part (Start ) --------------------------------------------------------------
		
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode									'  Code Condition
	
	strDate = FilterVar(frm1.fpdtWk_yymm.year,"2999","SNM") & "-" & FilterVar(frm1.fpdtWk_yymm.Month,"12","SNM") & "-" & "01"
	strDate = DateAdd("D",-1, DateAdd("M",1,cdate(strDate)))
   	arrParam(1) = UNIDateClientFormat(strDate)
	arrParam(2) = ""							' 자료권한 Condition  
	arrParam(3) = "T"									' 결의일자 상태 Condition  
	arrParam(4) = iWhere
	arrParam(5) = Trim(frm1.txtFactoryCD.value)
	
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDt3.asp", Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet, iWhere, Row)
	End If	
			
End Function


'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetCost()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(arrRet, iWhere, Row)


	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1,2
		        .vspdData.Col = C_BUCODE
		    	.vspdData.text = arrRet(0)
		        .vspdData.Col = C_BUNAME
		    	.vspdData.text = arrRet(1)
				.vspdData.Col = C_BIZ_AREA_CD
				.vspdData.text = arrRet(2)
				.vspdData.Col = C_ORG
				.vspdData.text = arrRet(3)
				.vspdData.Col = C_INTERNAL_CD	
				.vspdData.text = arrRet(4)		    		    	
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
		End Select

		lgBlnFlgChgValue = True

	End With

End Function

'===========================================================================
Function OpenDept1(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
	        arrParam(0) = "계정타입-팝업"		   				    ' TextBox 명칭 
	    	arrParam(1) = " B_MINOR"					' TABLE 명칭 
	    	arrParam(2) = ""     	                        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "MAJOR_CD = " & FilterVar("H0071", "''", "S") & "  "			 		' Where Condition
	    	arrParam(5) = "계정타입"		   				    ' TextBox 명칭 

	    	arrField(0) = "MINOR_CD"		                ' Field명(0)
	    	arrField(1) = "MINOR_NM"    						' Field명(1)

	    	arrHeader(0) = "계정타입코드"		        		' Header명(0)
	    	arrHeader(1) = "계정타입명"	      				' Header명(1)

	End Select

    arrParam(3) = ""
	'arrParam(0) = arrParam(5)								    ' 팝업 명칭 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_BTN,frm1.vspdData.ActiveRow ,"M","X","X")
		Exit Function
	Else
		Call SetCost1(arrRet, iWhere, Row)
	End If

End Function

'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetCost()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCost1(arrRet, iWhere, Row)


	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_TYPECD
		    	.vspdData.text = arrRet(0)
		        .vspdData.Col = C_TYPE
		    	.vspdData.text = arrRet(1)
		    	
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
		End Select
		Call SetActiveCell(.vspdData,C_BTN,.vspdData.ActiveRow ,"M","X","X")

		lgBlnFlgChgValue = True

	End With

End Function

'=======================================================================================================
'   Event Name : fpdtWk_yymm_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================

Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
 		Call SetFocusToDocument("M")
		Frm1.fpdtWk_yymm.Focus
	End If
End Sub

'=======================================================================================================
'   Event Name : fpdtWk_yymm_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub fpdtWk_yymm_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub



'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Select Case Col				'추가부분을 위해..select로..
	    Case C_BTN
	        frm1.vspdData.Col = C_BUCODE
	        Call OpenDept(frm1.vspdData.Text, 2, Row)
  	    Case C_BTN1
	        frm1.vspdData.Col = C_TYPE
	        Call OpenDept1(frm1.vspdData.Text, 1, Row)
	End Select
	Call SetActiveCell(frm1.vspdData,Col - 1,frm1.vspdData.ActiveRow ,"M","X","X")
End Sub



'========================================================================================================
Sub vspdData_Click(Col, Row)

	Call SetPopupMenuItemInf("1101111111")   
    
    gMouseClickStatus = "SPC"  

    Set gActiveSpdSheet = frm1.vspdData


    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
    
    	frm1.vspdData.Row = Row  
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

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_LeaveCell(Col,Row,NewCol,NewRow,Cancel)
 
End Sub
'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
Sub vspdData_OnFocus()
    lgActiveSpd      = "M"
End Sub



'========================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_Change( Col ,  Row)

    Dim iDx
    Frm1.vspdData1.Row = Row
    Frm1.vspdData1.Col = Col

    
    Call CheckMinNumSpread(frm1.vspdData1,Col,Row)
   
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row

End Sub
'========================================================================================================
Sub vspdData1_Click(Col, Row)
    Call SetPopupMenuItemInf("0000111111") 
    gMouseClickStatus = "SP1C"
    Set gActiveSpdSheet = frm1.vspdData1   
End Sub
'========================================================================================================
'   Event Name : vspdData1_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_OnFocus()
    lgActiveSpd      = "S"
End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
End Sub    


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        frm1.vspdData.LeftCol=NewLeft
        frm1.vspdData1.LeftCol=NewLeft
    End If
    
End Sub

Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        frm1.vspdData1.LeftCol=NewLeft
        frm1.vspdData.LeftCol=NewLeft
    End If
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY SCROLL="no" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>상여금월별현황</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
                    <TD WIDTH=* ALIGN=RIGHT>&nbsp;</A></TD>
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
									<TD CLASS="TD5">년도</TD>
									<TD CLASS="TD6">										
										<script language =javascript src='./js/a5962ma1_fpdtWk_yymm_fpdtWk_yymm.js'></script>
									</TD>
									<TD CLASS="TD5">상여종류</TD>
									<TD CLASS="TD6">										
										<INPUT TYPE=TEXT   NAME="txtpayCD" SIZE=10 MAXLENGTH=1 tag="12XXXU" ALT="상여종류" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCurrencypay()">
										<INPUT TYPE=TEXT   NAME="txtpayNM" SIZE=22 MAXLENGTH=50 tag="14XXXU" >
									</TD>
								</TR>
								<TR>	
									<TD CLASS="TD5" nowrap>사업장</TD>
									<TD CLASS="TD6" nowrap>																				
										<INPUT TYPE=TEXT   NAME="txtFactoryCD" SIZE=10 MAXLENGTH=10 tag="12XXXU"  ALT="사업장" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCurrency()">
										<INPUT TYPE=TEXT   NAME="txtFactoryNM" SIZE=22 MAXLENGTH=50 tag="14XXXU" >
									</TD>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT="72%" WIDTH="100%" COLSPAN=4>
						<script language =javascript src='./js/a5962ma1_vaSpread1_vspdData.js'></script>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=* WIDTH="100%" COLSPAN=4>
						<script language =javascript src='./js/a5962ma1_vaSpread1_vspdData1.js'></script>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread"       TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN       NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtPrevNext"     TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtMaxRows"      TAG="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>

