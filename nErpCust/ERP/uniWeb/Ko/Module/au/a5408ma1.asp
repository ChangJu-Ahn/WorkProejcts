<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Account
*  2. Function Name        : 미결관리 
*  3. Program ID           : a5408ma1
*  4. Program Name         : 미결잔액명세서조회 
*  5. Program Desc         : 미결잔액명세서조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 
*  8. Modified date(Last)  : 
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">    </SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "a5408mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

' 미결관리1, 미결관리2, 적요, 기초금액, 발생금액,정리금액,잔액 
  
Dim C_MGNT_CD1
Dim C_UNSETTLED1	
Dim C_MGNT_NM1
DIM C_MGNT_CD2
Dim C_UNSETTLED2
Dim C_MGNT_NM2
Dim C_AMT1			
Dim C_AMT2			
Dim C_AMT3			
Dim C_AMT			


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim IsOpenPop
Dim gSelframeFlg																	'☜: Tab의 현재위치 

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 
                                                 


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
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	Dim strSvrDate, strDayCnt, strTemp
	Dim ServerDate
	
		ServerDate	= "<%=GetSvrDate%>"
	
		frm1.txtDate.text = UniConvDateAToB(ServerDate ,parent.gServerDateFormat,parent.gDateFormat) 
		frm1.txtDocCur.value	= Parent.gCurrency

		Call ggoOper.FormatDate(frm1.txtDate, Parent.gDateFormat, 2)
		
	
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

	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>                               '☆: 
	

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
   Select Case pOpt
	
       Case "Q"
			lgKeyStream = UNIGetFirstDay(frm1.txtDate.Year & Parent.gServerDateType & frm1.txtDate.Month & Parent.gServerDateType & "01" ,Parent.gServerDateFormat) & Parent.gColSep       'You Must append one character(Parent.gColSep)
			lgKeyStream = lgKeyStream & UNIGetLastDay(frm1.txtDate.Year & Parent.gServerDateType & frm1.txtDate.Month & Parent.gServerDateType & "01" ,Parent.gServerDateFormat) & Parent.gColSep       'You Must append one character(Parent.gColSep)
			lgKeyStream = lgKeyStream & Frm1.txtAcctCd.value & Parent.gColSep
			lgKeyStream = lgKeyStream & Frm1.txtDocCur.value & Parent.gColSep
       Case "R"
            lgKeyStream = UNIGetFirstDay(frm1.txtDate.Year & Parent.gServerDateType & frm1.txtDate.Month & Parent.gServerDateType & "01" ,Parent.gServerDateFormat) & Parent.gColSep       'You Must append one character(Parent.gColSep)
			lgKeyStream = lgKeyStream & UNIGetLastDay(frm1.txtDate.Year & Parent.gServerDateType & frm1.txtDate.Month & Parent.gServerDateType & "01" ,Parent.gServerDateFormat) & Parent.gColSep       'You Must append one character(Parent.gColSep)
            lgKeyStream = lgKeyStream & Frm1.txtAcctCd.value & Parent.gColSep
			lgKeyStream = lgKeyStream & Frm1.txtDocCur.value & Parent.gColSep
                  
   End Select 
                   
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
	


'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With Frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			'.Col = C_StudyOnOffCd : intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
			'.Col = C_StudyOnOffNm : .Value = intindex					
		Next	
	End With
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	 C_MGNT_CD1         = 1
	 C_UNSETTLED1		= 2
	 C_MGNT_NM1         = 3
	 C_MGNT_CD2         = 4
	 C_UNSETTLED2		= 5
	 C_MGNT_NM2         = 6
	 C_AMT1				= 7
	 C_AMT2				= 8
	 C_AMT3				= 9
	 C_AMT				= 10
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	With frm1.vspdData															' 1st Spread
	
      .MaxCols   = C_AMT +1
      .Col   = .MaxCols
      .ColHidden = True
      
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030204",,parent.gAllowDragDropSpread    
		ggoSpread.ClearSpreadData

	   .ReDraw = false
	   
	   Call GetSpreadColumnPos("A")
	
       Call AppendNumberPlace("6","4","2")

        
	    '미결관리1, 미결관리2, 적요, 기초금액, 발생금액,정리금액,잔액 
	    
	    ggoSpread.SSSetEdit	 C_MGNT_CD1	    ,""					,10, , , 10
        ggoSpread.SSSetEdit	 C_UNSETTLED1	,"미결관리1"	,10  ,0  ,	,40	,2 
        ggoSpread.SSSetEdit	 C_MGNT_NM1		,"관리항목명1"	,15   , , , 30
        ggoSpread.SSSetEdit	 C_MGNT_CD2	    ,""					,10, , , 10
		ggoSpread.SSSetEdit	 C_UNSETTLED2	,"미결관리2"	,10  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_MGNT_NM2		,"관리항목명2"	,15   , , , 30

		Call SetSpreadFloat (C_AMT1			,"기초금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT2			,"발생금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT3			,"정리금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT			,"잔액"			,17	,1	,Parent.ggAmtOfMoneyNo)
		
		Call ggoSpread.SSSetColHidden(C_MGNT_CD1,C_MGNT_CD1,True )
		Call ggoSpread.SSSetColHidden(C_MGNT_CD2,C_MGNT_CD2,True )
       .ReDraw = true
	
       Call SetSpreadLock("A") 
    
    End With
    
    With frm1.vspdData1																' 1st Spread
		.ReDraw = False

'	    .MaxRows = 0      '조회 상태에서 다시 조회 버튼 눌렀을 때,해당 필드들을 Clear하기 위해 필요한 문장.
	    .MaxRows = 1
		.MaxCols = C_AMT
		'.Col = MaxCols
		'.ColHidden = True
		
		.RowHeaderDisplay = 0
		.Row = 0
		.RowHidden = True 
		
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20030204",,parent.gAllowDragDropSpread    
		ggoSpread.ClearSpreadData
		
		Call GetSpreadColumnPos("B")
		
		Call AppendNumberPlace("6","4","2")
		
		'msgbox "Parent.ggAmtOfMoneyNo:" & Parent.ggAmtOfMoneyNo
											'ColumnPosition     Header  Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
        
	    ggoSpread.SSSetEdit	 C_MGNT_CD1	    ,""					,10, , , 10        
        ggoSpread.SSSetEdit	 C_UNSETTLED1	,"미결관리1"	,10  ,0  ,	,40	,2 
        ggoSpread.SSSetEdit	 C_MGNT_NM1		,"미결코드명1"	,15   , , , 30
        ggoSpread.SSSetEdit	 C_MGNT_CD2	    ,""					,10, , , 10
		ggoSpread.SSSetEdit	 C_UNSETTLED2	,"미결관리2"	,10  ,0  ,	,40	,2 
    	ggoSpread.SSSetEdit	 C_MGNT_NM2		,"미결코드명2"	,15   , , , 30
    	
		Call SetSpreadFloat (C_AMT1			,"기초금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
	    Call SetSpreadFloat (C_AMT2			,"발생금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT3			,"정리금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT			,"잔액"			,17	,1	,Parent.ggAmtOfMoneyNo)
		
		Call ggoSpread.SSSetColHidden(C_MGNT_CD1,C_MGNT_CD1,True )
		Call ggoSpread.SSSetColHidden(C_MGNT_CD2,C_MGNT_CD2,True )

		.ScrollBars = 0		'ScrollBarsNone
		
		.ReDraw = True
		Call SetSpreadLock("B") 
	End With
   

End SUb


Sub InitSpreadSheet2()

	Call initSpreadPosVariables()    
	With frm1.vspdData															' 1st Spread
	
      .MaxCols   = C_AMT +1
      .Col   = .MaxCols
      .ColHidden = True
      
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030204",,parent.gAllowDragDropSpread    
		ggoSpread.ClearSpreadData

	   .ReDraw = false
	   
	   Call GetSpreadColumnPos("A")
	
       Call AppendNumberPlace("6","4","2")

        
	    '미결관리1, 미결관리2, 적요, 기초금액, 발생금액,정리금액,잔액 
	    
	    ggoSpread.SSSetEdit	 C_MGNT_CD1	    ,""					,10, , , 10
        ggoSpread.SSSetEdit	 C_UNSETTLED1	,"미결관리1"	,10  ,0  ,	,40	,2 
        ggoSpread.SSSetEdit	 C_MGNT_NM1		,"미결관리명1"	,15   , , , 30
        ggoSpread.SSSetEdit	 C_MGNT_CD2	    ,""					,10, , , 10
		ggoSpread.SSSetEdit	 C_UNSETTLED2	,"미결관리2"	,10  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_MGNT_NM2		,"미결관리명2"	,15   , , , 30

		Call SetSpreadFloat (C_AMT1			,"기초금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT2			,"발생금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT3			,"정리금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT			,"잔액"			,17	,1	,Parent.ggAmtOfMoneyNo)
		
		Call ggoSpread.SSSetColHidden(C_MGNT_CD1,C_MGNT_CD1,True )
		Call ggoSpread.SSSetColHidden(C_MGNT_CD2,C_MGNT_CD2,True )
       .ReDraw = true
	
       Call SetSpreadLock("A") 
    
    End With
    
    With frm1.vspdData1																' 1st Spread
		.ReDraw = False

'	    .MaxRows = 0      '조회 상태에서 다시 조회 버튼 눌렀을 때,해당 필드들을 Clear하기 위해 필요한 문장.
	    .MaxRows = 1
		.MaxCols = C_AMT
		'.Col = MaxCols
		'.ColHidden = True
		
		.RowHeaderDisplay = 0
		.Row = 0
		.RowHidden = True 
		
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20030204",,parent.gAllowDragDropSpread    
		ggoSpread.ClearSpreadData
		
		Call GetSpreadColumnPos("B")
		
		Call AppendNumberPlace("6","4","2")
		
		'msgbox "Parent.ggAmtOfMoneyNo:" & Parent.ggAmtOfMoneyNo
											'ColumnPosition     Header  Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
        
	    ggoSpread.SSSetEdit	 C_MGNT_CD1	    ,""					,10, , , 10        
        ggoSpread.SSSetEdit	 C_UNSETTLED1	,"미결관리1"	,10  ,0  ,	,40	,2 
        ggoSpread.SSSetEdit	 C_MGNT_NM1		,"미결코드명1"	,15   , , , 30
        ggoSpread.SSSetEdit	 C_MGNT_CD2	    ,""					,10, , , 10
		ggoSpread.SSSetEdit	 C_UNSETTLED2	,"미결관리2"	,10  ,0  ,	,40	,2 
    	ggoSpread.SSSetEdit	 C_MGNT_NM2		,"미결코드명2"	,15   , , , 30
    	
		Call SetSpreadFloat (C_AMT1			,"기초금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
	    Call SetSpreadFloat (C_AMT2			,"발생금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT3			,"정리금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT			,"잔액"			,17	,1	,Parent.ggAmtOfMoneyNo)
		
		Call ggoSpread.SSSetColHidden(C_MGNT_CD1,C_MGNT_CD1,True )
		Call ggoSpread.SSSetColHidden(C_MGNT_CD2,C_MGNT_CD2,True )

		.ScrollBars = 0		'ScrollBarsNone
		
		.ReDraw = True
		Call SetSpreadLock("B") 
	End With
   

End SUb


'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock(ByVal pOpt)
	If pOpt = "A" Then
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
    ElseIF pOpt = "B" Then  
      ggoSpread.Source = frm1.vspdData1
      ggoSpread.SpreadLockWithOddEvenRowColor()
    End If 
    frm1.vspdData.operationmode = 3  
    
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                'Col          Row   Row2
      'ggoSpread.SSSetRequired    C_CDNo      , lRow, lRow
      ggoSpread.SSSetRequired    C_UNSETTLED1      ,pvStartRow	,pvEndRow
                            
                                'Col          Row   Row2
      ggoSpread.SSSetProtected   C_AMT,				,pvStartRow	,pvEndRow
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
			C_MGNT_CD1			= iCurColumnPos(1)
			C_UNSETTLED1    	= iCurColumnPos(2)
			C_MGNT_NM1          = iCurColumnPos(3)
			C_MGNT_CD2			= iCurColumnPos(4)
			C_UNSETTLED2  		= iCurColumnPos(5)
			C_MGNT_NM2			= iCurColumnPos(6)
			C_AMT1   			= iCurColumnPos(7)
			C_AMT2   			= iCurColumnPos(8)
			C_AMT3    			= iCurColumnPos(9)
			C_AMT    			= iCurColumnPos(10)
		
	   Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_MGNT_CD1			= iCurColumnPos(1)
			C_UNSETTLED1    	= iCurColumnPos(2)
			C_MGNT_NM1          = iCurColumnPos(3)
			C_MGNT_CD2			= iCurColumnPos(4)
			C_UNSETTLED2  		= iCurColumnPos(5)
			C_MGNT_NM2			= iCurColumnPos(6)
			C_AMT1   			= iCurColumnPos(7)
			C_AMT2   			= iCurColumnPos(8)
			C_AMT3    			= iCurColumnPos(9)
			C_AMT    			= iCurColumnPos(10)
		
    End Select    
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
'    Call InitComboBox
	Call initData
	Call ggoSpread.ReOrderingSpreadData()
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
	
    Call ggoOper.LockField(Document, "Q")                                   '⊙: Lock Suitable Field
									 ' N : 신규, Q :조회 
    Call InitVariables
    Call SetDefaultVal
    Call InitSpreadSheet                                                             'Setup the Spread sheet
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize
	frm1.txtDate.Focus
	Set gActiveElement = document.ActiveElement
	
	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

Sub txtDate_DblClick(Button)
	if Button = 1 then
		frm1.txtDate.Action = 7
	End if
End Sub

Sub txtDateTo_DblClick(Button)
	if Button = 1 then
		frm1.txtDateTo.Action = 7
	End if
End Sub


Sub txtDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

Sub txtDateTo_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

	
'       'Single                    'Multi
'    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
'		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"		
'		If IntRetCD = vbNo Then
'			Exit Function
'		End If
'   End If    
	
    'Call ggoOper.ClearField(Document, "2")										  '⊙: Clear Contents  Field
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
    Call InitVariables															  '⊙: Initializes local global variables

	If frm1.txtMgntCd1Fr.value <> "" And frm1.txtMgntCd1To.value <> "" Then
		If frm1.txtMgntCd1Fr.value > frm1.txtMgntCd1To.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtMgntCd1Fr.Alt, frm1.txtMgntCd1To.Alt)
			frm1.txtMgntCd1Fr.focus 
			Exit Function
		End If
	End If
	If frm1.txtMgntCd2Fr.value <> "" And frm1.txtMgntCd2To.value <> "" Then
		If frm1.txtMgntCd2Fr.value > frm1.txtMgntCd2To.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtMgntCd2Fr.Alt, frm1.txtMgntCd2To.Alt)
			frm1.txtMgntCd2Fr.focus 
			Exit Function
		End If
	End If	    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    If DbQuery("Q") = False Then                                                       '☜: Query db data
        Call LayerShowHide(0)                                                        '☜: Show Processing Message
        Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False	                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True	                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
      
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
       IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
       If IntRetCD = vbNo Then
          Exit Function
       End If
	End If
    	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables													         '⊙: Initializes local global variables

    If DbQuery("P") = False Then                                                       '☜: Query db data
       Exit Function
    End If
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
    ggoSpread.Source = frm1.vspdData

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
       IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
       If IntRetCD = vbNo Then
          Exit Function
       End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														     '⊙: Initializes local global variables

    If DbQuery("N") = False Then                                                       '☜: Query db data
       Exit Function
    End If
    
	'------ Developer Coding part (Start )   -------------------------------------------------------------- 
	'------ Developer Coding part (End   )   -------------------------------------------------------------- 

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

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			         '⊙: Data is changed.  Do you want to exit? 
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
Function DbQuery(pDirect)

	Dim strVal
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
            
    DbQuery = False                                                              '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    Call MakeKeyStream(pDirect)
	
	'msgbox lgStrPrevKeyIndex
    strVal = BIZ_PGM_ID	& "?txtMode="			& Parent.UID_M0001                     '☜: Query
    strVal = strVal		& "&txtKeyStream="		& lgKeyStream                   '☜: Query Key
    strVal = strVal		& "&txtPrevNext="		& pDirect                       '☜: Direction
    strVal = strVal		& "&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal		& "&txtMaxRows="		& Frm1.vspdData.MaxRows         '☜: Max fetched data
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
			
	If frm1.ProcessOpt1.checked = True Then
		strVal = strVal & "&ProcessOption=" & "1"	' All
	Else
		strVal = strVal & "&ProcessOption=" & "2"
	End If
	
	If frm1.QPointOpt1.checked = True Then
		strVal = strVal & "&QueryOption=" & "1"	' All
	Else
		strVal = strVal & "&QueryOption=" & "2"
	End If

	strVal = strVal		& "&txtMgntCd1Fr="		& Trim(frm1.txtMgntCd1Fr.value)
	strVal = strVal		& "&txtMgntCd1To="		& Trim(frm1.txtMgntCd1To.value)
	strVal = strVal		& "&txtMgntCd2Fr="		& Trim(frm1.txtMgntCd2Fr.value)
	strVal = strVal		& "&txtMgntCd2To="		& Trim(frm1.txtMgntCd2To.value)


	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 


	If pDirect <> "R" Then
		
		frm1.vspdData.MaxRows = 0
		frm1.vspdData1.MaxRows = 0
		
	End If
	
	If pDirect = "Q" Then
		'Call InitSpreadSheet
		
		frm1.vspdData.MaxRows = 0
		frm1.vspdData1.MaxRows = 0
	End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
    
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	'lgIntFlgMode      = Parent.OPMD_UMODE                                                   '⊙: Indicates that current mode is Create mode
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'If  gSelframeFlg = TAB1 Then
'		Frm1.vspdData.focus
	'Else
	'	Frm1.vspdData1.focus
	'End If
		
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	call txtDocCur_OnChange()
	 lgBlnFlgChgValue = False											
    Set gActiveElement = document.ActiveElement   
           
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================


 '**********************  회계전표 Popup  ****************************************
'	기능: 
'   설명: 
'************************************************************************************** 
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a5120ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function

	If frm1.vspdData.MaxRows > 0 Then
		With frm1.vspdData
			.Row = .ActiveRow
							
			arrParam(0) = Trim(.Text)	'결의전표번호 
			arrParam(1) = ""			'Reference번호 
		End With
	Else
		Call DisplayMsgBox("900002", "X","X","X")
		Exit Function
	
	End If
	
	If arrParam(0) = "전표번호" Then Exit Function		' 조회 이전 
	
	IsOpenPop = True   
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'************************************************************************************** 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd, strTempBankCd

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0
			
			arrParam(0) = "사업장팝업"						' 팝업 명칭 
			arrParam(1) = "B_Biz_AREA"							' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Condition
			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If
			arrParam(5) = "사업장코드"			
			
			arrField(0) = "BIZ_AREA_CD"								' Field명(0)
			arrField(1) = "BIZ_AREA_NM"								' Field명(1)
		    arrHeader(0) = "사업장코드"							' Header명(0)
			arrHeader(1) = "사업장명"							' Header명(1)

		Case 1
			arrParam(0) = "계정코드팝업"						' 팝업 명칭 
			arrParam(1) = "A_ACCT"								' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = " MGNT_FG=" & FilterVar("Y", "''", "S") & " "									' Where Condition
			arrParam(5) = "계정코드"			
		
		    arrField(0) = "ACCT_CD"								' Field명(0)
			arrField(1) = "ACCT_NM"								' Field명(1)
	    
		    arrHeader(0) = "계정코드"							' Header명(0)
			arrHeader(1) = "계정명"							' Header명(1)
		Case 2
			arrParam(0) = "통화코드팝업"				' 팝업 명칭 
			arrParam(1) = "B_Currency"	    			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "통화코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "Currency"	    			' Field명(0)
			arrField(1) = "Currency_desc"	    		' Field명(1)
    
			arrHeader(0) = "통화코드"					' Header명(0)
			arrHeader(1) = "통화코드명"
		Case 3					
			arrParam(0) = "전표경로팝업"						' 팝업 명칭 
			arrParam(1) = "b_minor B,a_daily_subledger A"		' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "a.gl_dt between  " & FilterVar(UNIConvDate(frm1.txtDate.Text), "''", "S") & " and  " & FilterVar(UNIConvDate(frm1.txtDateTo.Text), "''", "S") & " "
			arrParam(4) = arrParam(4) & " and (1=1 or 1=2 and a.biz_area_cd=" & FilterVar("2", "''", "S") & ") "
			arrParam(4) = arrParam(4) & " and (2=2 or 2=3 and a.biz_unit_cd=" & FilterVar("3", "''", "S") & ") "
			arrParam(4) = arrParam(4) & " and  b.minor_cd=a.gl_input_type and b.major_cd=" & FilterVar("A1001", "''", "S") & " "									' Where Condition
			arrParam(5) = "경로"			
		
		    arrField(0) = "A.gl_input_type"								' Field명(0)
			arrField(1) = "B.minor_nm"								' Field명(1)
	    
		    arrHeader(0) = "경로"					' Header명(0)
			arrHeader(1) = "경로명"							' Header명(1)
		    
		Case Else
			Exit Function
	End Select
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	 Select Case iWhere
	   Case 1
		frm1.txtAcctCd.focus
	   Case 2	
		frm1.txtDocCur.focus
	 End Select	
		Exit Function
	Else
	  Select Case iWhere
		Case 1	' Account
			frm1.txtAcctCd.focus
			frm1.txtAcctCd.value = arrRet(0)
			frm1.txtAcctNm.value = arrRet(1)
			
			Call txtAcctcd_Onchange()
		Case 2
			frm1.txtDocCur.focus
			frm1.txtDocCur.Value = arrret(0)
	 End Select
	End If	

End Function


Function txtAcctCd_Onchange()
	With frm1
		
    
		Call CommonQueryRs(" A_ACCt.ACCT_CD, ACCT_NM ","A_ACCT","A_ACCT.ACCT_CD = '" & .txtAcctCd.value & "' AND isnull(A_ACCT.mgnt_type,'')  <> ''" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
		If (lgF0 <> "X") And (Trim(lgF0) <> "") Then 
		
			
			Call CommonQueryRs("CTRL_NM", _
                       "A_ACCT A, A_CTRL_ITEM B", _
                       "A.mgnt_cd1 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(frm1.txtAcctCd.value, "''", "S"),_
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			 ggoSpread.Source = frm1.vspdData
			if lgF0 <> "" then
				CtrlCd.innerHTML = REPLACE(lgF0,Chr(11),"") 
				ggoSpread.SSSetEdit	 C_UNSETTLED1	,REPLACE(lgF0,Chr(11),"")	,10  ,0  ,	,40	,2 
			Else
				CtrlCd.innerHTML = "미결코드1" 
				
			End if
			
			Call CommonQueryRs("CTRL_NM", _
                       "A_ACCT A, A_CTRL_ITEM B", _
                       "A.mgnt_cd2 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(frm1.txtAcctCd.value, "''", "S"),_
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			if lgF0 <> "" then
				CtrlCd2.innerHTML = REPLACE(lgF0,Chr(11),"")
				ggoSpread.SSSetEdit	 C_UNSETTLED2	,REPLACE(lgF0,Chr(11),"")	,10  ,0  ,	,40	,2 
			Else
				CtrlCd2.innerHTML = "미결코드2"
			End if
			
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd1,		"D")
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"D")
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"D")
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd1Nm,		"Q")
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd2Nm,		"Q")
			.txtAcctCd.focus
			
		Else       
			.txtAcctCd.value = ""
			.txtAcctNm.value = ""
			.txtMgntCd1.value = ""
			.txtMgntCd1Nm.value = ""
			.txtMgntCd2.value = ""
			.txtMgntCd2Nm.value = ""
			CtrlCd.innerHTML = "미결코드1"
			CtrlCd2.innerHTML = "미결코드2"
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd1,		"Q")
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"Q")
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"Q")
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd1Nm,		"Q")
			'Call ggoOper.SetReqAttr(frm1.txtMgntCd2Nm,		"Q")      
			'.txtCtrlVal.value = ""
			'.txtCtrlValNm.value = ""       
			.txtAcctCd.focus       
		End If   
	End With
	

    txtAcctCd_OnChange = True
End Function




'************************************************************************************** 
Function OpenMgntPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd, strTempBankCd
	Dim IntRetCD, IntRetCD1
	Dim strFrom, strWhere, strFrom1, strWhere1
	Dim arrVal, arrVal1, arrVal2, arrVal3, arrVal4, arrVal5, arrVal6, arrVal7 ,arrMajor
	DIm stbl_id, scol_id, sdata_id, stbl_id2, scol_id2, sdata_id2 

	If Trim(frm1.txtAcctCd.value) = "" Then
        Call DisplayMsgBox("110131","x","x","x")
        Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0, 1

			IntRetCD = CommonQueryRs("TBL_ID, DATA_COLM_ID, DATA_COLM_NM ,ISNULL(LTRIM(RTRIM(MAJOR_CD)),'') ","A_OPEN_ACCT A, A_CTRL_ITEM B","A.mgnt_cd1 = B.CTRL_CD AND A.ACCT_CD = " & FilterVar(Trim(frm1.txtAcctCd.value), "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			If IntRetCD = true Then

				arrVal = Split(lgF0, Chr(11)) 
				stbl_id = arrVal(0)

				arrVal1 = Split(lgF1, Chr(11)) 
				scol_id = arrVal1(0)

				arrVal2 = Split(lgF2, Chr(11)) 
				arrVal3 = arrVal2(0)

				arrVal   =  Split(lgF3, Chr(11)) 
				arrMajor =  arrVal(0)				
			Else
				IntRetCD1 = DisplayMsgBox("900014","x","x","x") '☜ 바뀐부분 
				IsOpenPop = False
				Exit Function
			End If

			strFrom = " A_OPEN_ACCT A, " & stbl_id & " B "
			strWhere = " ACCT_CD =  " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ""
			strWhere = strWhere  & " AND A.MGNT_VAL1 = B."&scol_id

			If arrMajor  <> "" Then	
				strWhere = strWhere  & " and major_cd = '" & arrMajor & "'"
			End If

			arrParam(0) = "미결코드1팝업"			' 팝업 명칭 
			arrParam(1) = strFrom		    			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = strWhere						' Where Condition
			arrParam(5) = "미결코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "A.MGNT_VAL1"	    			' Field명(0)
			arrField(1) = "B."&arrVal3 	    			' Field명(1)

			arrHeader(0) = "미결관리1"				' Header명(0)
			arrHeader(1) = "미결코드"
		Case 2, 3
			IntRetCD1 =  CommonQueryRs("TBL_ID, DATA_COLM_ID, DATA_COLM_NM , ISNULL(LTRIM(RTRIM(MAJOR_CD)),'')","A_OPEN_ACCT A, A_CTRL_ITEM B","A.mgnt_cd2 = B.CTRL_CD AND A.ACCT_CD = " & FilterVar(Trim(frm1.txtAcctCd.value), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			If IntRetCD1 = True Then
				arrVal4 = Split(lgF0, Chr(11)) 
				stbl_id2 = arrVal4(0)

				arrVal5 = Split(lgF1, Chr(11)) 
				scol_id2 = arrVal5(0)

				arrVal6 = Split(lgF2, Chr(11)) 
				arrVal7 = arrVal6(0)
					
				arrVal   =  Split(lgF3, Chr(11)) 
				arrMajor =  arrVal(0)					
			Else
				IntRetCD1 = DisplayMsgBox("900014","x","x","x") '☜ 바뀐부분 
				IsOpenPop = False	
				Exit Function				
			End If

			strFrom1 = " A_OPEN_ACCT A, " & stbl_id2 & " B "
			strWhere1 = " ACCT_CD =  " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ""
			strWhere1 = strWhere1  & " AND A.MGNT_VAL2 = B."&scol_id2

			If arrMajor  <> "" Then
				strWhere1 = strWhere1  & " and major_cd = '" & arrMajor & "'"
			End If

			arrParam(0) = "미결코드2팝업"			' 팝업 명칭 
			arrParam(1) = strFrom1	    				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = strWhere1                      ' Where Condition
			arrParam(5) = "미결코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "MGNT_VAL2"	    			' Field명(0)
			arrField(1) = "B."&arrVal7 	    			' Field명(1)
   
			arrHeader(0) = "미결관리2"				' Header명(0)
			arrHeader(1) = "미결코드"
	End Select

	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0
				frm1.txtMgntCd1Fr.focus
			Case 1
				frm1.txtMgntCd1To.focus
			Case 2
				frm1.txtMgntCd2Fr.focus
			Case 3
				frm1.txtMgntCd2To.focus 
		End Select     
		Exit Function
	Else
		Select Case iWhere
			Case 0
				frm1.txtMgntCd1Fr.focus
				frm1.txtMgntCd1Fr.value = arrRet(0)
			Case 1	
				frm1.txtMgntCd1To.focus
				frm1.txtMgntCd1To.value = arrRet(0)
			Case 2
				frm1.txtMgntCd2Fr.focus
				frm1.txtMgntCd2Fr.Value = arrret(0)
			Case 3	 
				frm1.txtMgntCd2To.focus
				frm1.txtMgntCd2To.value = arrRet(0)
		End Select
	End If	
End Function

'========================================================================================================

Sub FillDateField()
Dim strDateFr, strDateTo

	If frm1.cboSEQ1.Value <> "" Then'
					                   'Select                 From                Where										Return value list
		Call CommonQueryRs(" Distinct(f_dt), t_dt "," A_Monthly_Subledger ",  " YEAR =  " & FilterVar(frm1.cboYear.Value , "''", "S") & " And SEQ=  " & FilterVar(frm1.cboSeq.value , "''", "S") & " "   ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		               'ComboObject Name      Name   Value   Separator
		strDateFr = Left(lgF0,4) & "-" & Mid(lgF0,5,2) & "-" & Mid(lgF0,7,2)
		strDateTo = Left(lgF1,4) & "-" & Mid(lgF1,5,2) & "-" & Mid(lgF1,7,2)
		frm1.txtDate.Value = UNIDateClientFormat(strDateFr)
		frm1.txtDateTo.Value = UNIDateClientFormat(strDateTo)
	End If
End Sub


'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			Case C_AMT1
				.Col = Col - 1
				.Row = Row
				Call OpenZipCode(.Text,Row)
			End Select
		End If
    
	End With
End Sub


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
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row <= 0 Then
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


'======================================================================================================
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
        frm1.vspddata1.Leftcol = NewLeft
        Exit Sub
    End If
	
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
      if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
           If DbQuery("R") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End if
End Sub


'========================================================================================================
Sub fpdtFoundDt_ButtonHit(Button, NewIndex)
	On Error Resume Next
    lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub fpdtCloseDt_ButtonHit(Button, NewIndex)
	On Error Resume Next
    lgBlnFlgChgValue = True
End Sub


'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY = " & FilterVar(frm1.txtDocCur.value , "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumSprSheet()
	END IF	    
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_AMT1,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_AMT2,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_AMT3,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_AMT, -1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub


'========================================================================================================
Function FncPrint() 

Dim StrEbrFile, StrUrl
Dim IntRetCd

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If
		
	If frm1.txtMgntCd1Fr.value <> "" And frm1.txtMgntCd1To.value <> "" Then
		If frm1.txtMgntCd1Fr.value > frm1.txtMgntCd1To.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtMgntCd1Fr.Alt, frm1.txtMgntCd1To.Alt)
			frm1.txtMgntCd1Fr.focus 
			Exit Function
		End If
	End If
	If frm1.txtMgntCd2Fr.value <> "" And frm1.txtMgntCd2To.value <> "" Then
		If frm1.txtMgntCd2Fr.value > frm1.txtMgntCd2To.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtMgntCd2Fr.Alt, frm1.txtMgntCd2To.Alt)
			frm1.txtMgntCd2Fr.focus 
			Exit Function
		End If
	End If	    


	Call SetPrintCond(StrEbrFile, StrUrl)

	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	Call FncEBRPrint(EBAction,ObjName,StrUrl)
	
End Function



'========================================================================================================
Function FncPreview()
 
Dim StrEbrFile, StrUrl
Dim IntRetCd

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If
		
	If frm1.txtMgntCd1Fr.value <> "" And frm1.txtMgntCd1To.value <> "" Then
		If frm1.txtMgntCd1Fr.value > frm1.txtMgntCd1To.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtMgntCd1Fr.Alt, frm1.txtMgntCd1To.Alt)
			frm1.txtMgntCd1Fr.focus 
			Exit Function
		End If
	End If
	If frm1.txtMgntCd2Fr.value <> "" And frm1.txtMgntCd2To.value <> "" Then
		If frm1.txtMgntCd2Fr.value > frm1.txtMgntCd2To.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtMgntCd2Fr.Alt, frm1.txtMgntCd2To.Alt)
			frm1.txtMgntCd2Fr.focus 
			Exit Function
		End If
	End If	    

	Call SetPrintCond(StrEbrFile, StrUrl)

	ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	
	Call FncEBRPreview(ObjName,StrUrl)
			
End Function


'=======================================================================================================
Sub SetPrintCond(StrEbrFile, StrUrl)
	
	Dim ValDateFr, ValDateTo, ValAcctCd, ValDocCur, ValProcessOpt, ValProcessOptNm
	Dim ValMgntCd1Fr, ValMgntCd1To, ValMgntCd2Fr, ValMgntCd2To
	Dim	strAuthCond
	
	If frm1.QPointOpt1.checked = True Then
		' 작업시점 
		StrEbrFile	= "a5408ma1"
	Else
		' 현재시점 
		StrEbrFile	= "a5408ma2"
	End If
	
	With frm1
		ValDateFr	= UNIGetFirstDay(frm1.txtDate.Year & Parent.gServerDateType & frm1.txtDate.Month & Parent.gServerDateType & "01" ,Parent.gServerDateFormat)
		ValDateTo	= UNIGetLastDay(frm1.txtDate.Year & Parent.gServerDateType & frm1.txtDate.Month & Parent.gServerDateType & "10" ,Parent.gServerDateFormat)
		ValAcctCd	= UCase(Trim(.txtAcctCd.value))
		ValDocCur	= UCase(Trim(.txtDocCur.value))
		
		If frm1.ProcessOpt1.checked = True Then
			' 전체 
			ValProcessOpt	= " 1=1 "
			ValProcessOptNm	= "전체"
		Else
			' 잔액 
			ValProcessOpt	= " sum(amt1+amt2-amt3) <> 0 "
			ValProcessOptNm	= "잔액"
		End If
		
		ValMgntCd1Fr = ""
		ValMgntCd1To = "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"
		ValMgntCd2Fr = ""
		ValMgntCd2To = "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"
		
		If Trim(.txtMgntCd1Fr.value) <> "" Then		ValMgntCd1Fr = Trim(.txtMgntCd1Fr.value)
		If Trim(.txtMgntCd1To.value) <> "" Then		ValMgntCd1To = Trim(.txtMgntCd1To.value)
		If Trim(.txtMgntCd2Fr.value) <> "" Then		ValMgntCd2Fr = Trim(.txtMgntCd2Fr.value)
		If Trim(.txtMgntCd2To.value) <> "" Then		ValMgntCd2To = Trim(.txtMgntCd2To.value)
	End With

	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND b.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND b.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND b.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND b.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


	StrUrl = StrUrl & "DateFr|"			& ValDateFr
	StrUrl = StrUrl & "|DateTo|"		& ValDateTo
	StrUrl = StrUrl & "|AcctCd|"		& ValAcctCd
	StrUrl = StrUrl & "|DocCur|"		& ValDocCur
	StrUrl = StrUrl & "|ValMgntCd1Fr|"	& ValMgntCd1Fr
	StrUrl = StrUrl & "|ValMgntCd1To|"	& ValMgntCd1To
	StrUrl = StrUrl & "|ValMgntCd2Fr|"	& ValMgntCd2Fr
	StrUrl = StrUrl & "|ValMgntCd2To|"	& ValMgntCd2To
	StrUrl = StrUrl & "|ProcessOpt|"	& ValProcessOpt
	StrUrl = StrUrl & "|ProcessOptNm|"	& ValProcessOptNm

	StrUrl = StrUrl & "|strAuthCond|"		& strAuthCond






End Sub
	

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
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
						<TD WIDTH=* ALIGN=RIGHT></TD>
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
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDate CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="작업년월"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>										  
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계정코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCd.Value,1)"> 
														 <INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="계정명">
									</TD>
									<TD CLASS="TD5" NOWRAP>거래통화</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="거래통화"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.Value,2)"> 
								</TR>


								<TR>
									<TD CLASS="TD5"  ID="CtrlCd" NOWRAP>미결관리1</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtMgntCd1Fr" SIZE=15 MAXLENGTH=30 tag="11XXXU" ALT="미결관리1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgntCd1Fr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMgntPopup(frm1.txtMgntCd1Fr.Value,0)">&nbsp;~&nbsp;
														   <INPUT TYPE="Text" NAME="txtMgntCd1To" SIZE=15 MAXLENGTH=30 tag="11XXXU" ALT="미결관리1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgntCd1To" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMgntPopup(frm1.txtMgntCd1To.Value,1)">
									</TD>
									<TD CLASS="TD5"  ID="CtrlCd2" NOWRAP>미결관리2</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtMgntCd2Fr" SIZE=15 MAXLENGTH=30 tag="11XXXU" ALT="미결관리2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgntCd2Fr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMgntPopup(frm1.txtMgntCd2Fr.Value,2)">&nbsp;~&nbsp;
														   <INPUT TYPE="Text" NAME="txtMgntCd2To" SIZE=15 MAXLENGTH=30 tag="11XXXU" ALT="미결관리2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgntCd2To" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMgntPopup(frm1.txtMgntCd2To.Value,3)">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>진행사항</TD>
									<TD CLASS="TD6" NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ProcessOpt" CHECKED ID="ProcessOpt1" VALUE="Y" tag="22"><LABEL FOR="ProcessOpt1">전체</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ProcessOpt" ID="ProcessOpt2" VALUE="N" tag="22"><LABEL FOR="ProcessOpt2">잔액</LABEL></SPAN>
									</TD>
									<TD CLASS=TD5 NOWRAP>조회시점</TD>
									<TD CLASS="TD6" NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="QPointOpt" CHECKED ID="QPointOpt1" VALUE="Y" tag="22"><LABEL FOR="QPointOpt1">작업시점</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="QPointOpt" ID="QPointOpt2" VALUE="N" tag="22"><LABEL FOR="QPointOpt2">현재시점</LABEL></SPAN>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="94%" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="33" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="6%" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData1 NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="33" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncPrint()" Flag=1>인쇄</BUTTON>&nbsp;
					</TD>
					<TD WIDTH=* ALIGN=RIGHT></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24">

<INPUT TYPE=HIDDEN NAME="htxtAcctCd"     TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtAcctCd1"    TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtDate"     TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtDateTo"     TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtBankAcctNo" TAG="X4">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
<INPUT TYPE="HIDDEN" NAME="uname">
<INPUT TYPE="HIDDEN" NAME="dbname">
<INPUT TYPE="HIDDEN" NAME="filename">
<INPUT TYPE="HIDDEN" NAME="condvar">
<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>

