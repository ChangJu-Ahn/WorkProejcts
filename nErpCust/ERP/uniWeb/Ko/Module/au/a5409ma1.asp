
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Account
*  2. Function Name        : 미결관리 
*  3. Program ID           : a5409ma1
*  4. Program Name         : 미결잔액거래내역조회 
*  5. Program Desc         : 미결잔액거래내역조회 
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

Const BIZ_PGM_ID      = "a5409mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

' 미결잔액명세서   : 미결관리1, 미결관리2                          , 적요, 기초금액, 발생금액, 정리금액, 잔액 
' 미결잔액거래내역 : 미결관리1, 미결관리2, 전표번호, 발생일, 정리일, 적요          , 발생금액, 정리금액, 잔액 

Dim C_UNSETTLED1
Dim C_UNSETTLED2
Dim C_GLNO
Dim C_GLDT
Dim C_CLSDT
Dim C_NOTE
Dim C_AMT0
Dim C_AMT1
Dim C_AMT2
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
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
'		
		Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

		EndDate = "<%=GetSvrDate%>"
		Call ExtractDateFrom(EndDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)
		
		StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
		EndDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
	
		
		frm1.txtDocCur.value	= Parent.gCurrency
		frm1.txtDateFr.Text		= StartDate 
		frm1.txtDateTo.Text		= EndDate 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub MakeKeyStream(pOpt)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
   Select Case pOpt
	
       Case "Q"        
			lgKeyStream = UNIConvDateToYYYYMMDD(Trim(frm1.txtDateFr.Text),parent.gDateFormat,"") & Parent.gColSep
			lgKeyStream = lgKeyStream & UNIConvDateToYYYYMMDD(Trim(frm1.txtDateTo.Text),parent.gDateFormat,"") & Parent.gColSep       'You Must append one character(Parent.gColSep)
			lgKeyStream = lgKeyStream & Frm1.txtAcctCd.value & Parent.gColSep
			lgKeyStream = lgKeyStream & Frm1.txtDocCur.value & Parent.gColSep
       Case "R"
            lgKeyStream = UNIConvDateToYYYYMMDD(Trim(frm1.txtDateFr.Text),parent.gDateFormat,"") & Parent.gColSep
			lgKeyStream = lgKeyStream & UNIConvDateToYYYYMMDD(Trim(frm1.txtDateTo.Text),parent.gDateFormat,"") & Parent.gColSep       'You Must append one character(Parent.gColSep)
            lgKeyStream = lgKeyStream & Frm1.txtAcctCd.value & Parent.gColSep
			lgKeyStream = lgKeyStream & Frm1.txtDocCur.value & Parent.gColSep
                  
   End Select 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
	

'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With Frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			Next	
	End With
End Sub


'========================================================================================================
Sub initSpreadPosVariables()  
	 C_UNSETTLED1		= 1
	 C_UNSETTLED2		= 2
	 C_GLNO				= 3
	 C_GLDT				= 4
	 C_CLSDT			= 5
	 C_NOTE				= 6
	 C_AMT0				= 7
	 C_AMT1				= 8
	 C_AMT2				= 9
	 C_AMT				= 10
End Sub


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

       
       'C_UNSETTLED1, C_UNSETTLED2, C_GLNO, C_GLDT, C_CLSDT, C_NOTE, C_AMT1, C_AMT2, C_AMT
	    ggoSpread.SSSetEdit	 C_UNSETTLED1	,"미결관리1"	,15  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_UNSETTLED2	,"미결관리2"	,15  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_GLNO			,"전표번호"		,15  ,0  ,	,40	,2 
	
		
		ggoSpread.SSSetEdit	 C_GLDT			,"발생일"		,10  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_CLSDT		,"정리일"		,10  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_NOTE			,"적요"			,15  ,0  ,	,40	,2 
		Call SetSpreadFloat (C_AMT0			,"이월금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT1			,"발생금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT2			,"정리금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT			,"잔액"			,17	,1	,Parent.ggAmtOfMoneyNo)
       
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
		ggoSpread.Spreadinit "V20030723",,parent.gAllowDragDropSpread    
		ggoSpread.ClearSpreadData
		
		Call GetSpreadColumnPos("B")
		
		Call AppendNumberPlace("6","4","2")
		
		'msgbox "Parent.ggAmtOfMoneyNo:" & Parent.ggAmtOfMoneyNo
											'ColumnPosition     Header  Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
        
	    ggoSpread.SSSetEdit	 C_UNSETTLED1	,"미결관리1"	,15  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_UNSETTLED2	,"미결관리2"	,15  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_GLNO			,"전표번호"		,15  ,0  ,	,40	,2 
	
		
		ggoSpread.SSSetEdit	 C_GLDT			,"발생일"		,10  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_CLSDT		,"정리일"		,10  ,0  ,	,40	,2 
		ggoSpread.SSSetEdit	 C_NOTE			,"적요"			,15  ,0  ,	,40	,2 
		Call SetSpreadFloat (C_AMT0			,"이월금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT1			,"발생금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT2			,"정리금액"		,17	,1	,Parent.ggAmtOfMoneyNo)
		Call SetSpreadFloat (C_AMT			,"잔액"			,17	,1	,Parent.ggAmtOfMoneyNo)
       
		.ScrollBars = 0		'ScrollBarsNone
		
		.ReDraw = True
		Call SetSpreadLock("B") 
	End With

    
    
End SUb


'======================================================================================================
Sub SetSpreadLock(ByVal pOpt)
	If pOpt = "A" Then
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
    ElseIF pOpt = "B" Then  
      ggoSpread.Source = frm1.vspdData1
      ggoSpread.SpreadLockWithOddEvenRowColor()
    End If   
    frm1.vspddata.operationmode = 3
End Sub

'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                'Col          Row   Row2
      'ggoSpread.SSSetRequired    C_CDNo      , lRow, lRow
      ggoSpread.SSSetRequired    C_UNSETTLED1      ,pvStartRow	,pvEndRow
                            
                                'Col          Row   Row2
      ggoSpread.SSSetProtected   C_AMT				,pvStartRow	,pvEndRow
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
              Frm1.vspdData.Action = 0 ' go to 
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

			C_UNSETTLED1    	= iCurColumnPos(1)
			C_UNSETTLED2  		= iCurColumnPos(2)
			C_GLNO    			= iCurColumnPos(3)    
			C_GLDT   			= iCurColumnPos(4)
			C_CLSDT   			= iCurColumnPos(5)
			C_NOTE    			= iCurColumnPos(6)
			C_AMT0    			= iCurColumnPos(7)
			C_AMT1    			= iCurColumnPos(8)
			C_AMT2    			= iCurColumnPos(9)
			C_AMT    			= iCurColumnPos(10)
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_UNSETTLED1    	= iCurColumnPos(1)
			C_UNSETTLED2  		= iCurColumnPos(2)
			C_GLNO    			= iCurColumnPos(3)    
			C_GLDT   			= iCurColumnPos(4)
			C_CLSDT   			= iCurColumnPos(5)
			C_NOTE    			= iCurColumnPos(6)
			C_AMT0    			= iCurColumnPos(7)
			C_AMT1    			= iCurColumnPos(8)
			C_AMT2    			= iCurColumnPos(9)
			C_AMT    			= iCurColumnPos(10)		
 
    End Select    
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
'    Call InitComboBox
	Call initData
	Call ggoSpread.ReOrderingSpreadData()
End Sub


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
	frm1.txtDateFr.Focus
	
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
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

Sub txtDateFr_DblClick(Button)
	if Button = 1 then
		frm1.txtDateFr.Action = 7
	End if
End Sub

Sub txtDateTo_DblClick(Button)
	if Button = 1 then
		frm1.txtDateTo.Action = 7
	End if
End Sub


Sub txtDateFr_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

Sub txtDateTo_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub


'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status
    
    
	ggoSpread.Source = Frm1.vspdData
	
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
    Call InitVariables															  '⊙: Initializes local global variables
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
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
 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    If DbQuery("Q") = False Then                                                       '☜: Query db data
        Call LayerShowHide(0)                                                        '☜: Show Processing Message
        Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                               '☜: Processing is OK
End Function


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
    

    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function


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
  
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

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
Function DbQuery(pDirect)

	Dim strVal
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
            
    DbQuery = False                                                              '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    Call MakeKeyStream(pDirect)
	
    strVal = BIZ_PGM_ID & "?txtMode="			& Parent.UID_M0001                     '☜: Query
    strVal = strVal		& "&txtKeyStream="		& lgKeyStream                   '☜: Query Key
    strVal = strVal		& "&txtPrevNext="		& pDirect                       '☜: Direction
    strVal = strVal		& "&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal		& "&txtMaxRows="		& Frm1.vspdData.MaxRows         '☜: Max fetched data
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

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
		'Call InitSpreadSheet2()
		frm1.vspdData.MaxRows = 0
		frm1.vspdData1.MaxRows = 0
		
	End If
	
	If pDirect = "Q" Then
		frm1.vspdData.MaxRows = 0
		frm1.vspdData1.MaxRows = 0
		
	End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
    
End Function


'========================================================================================================
Sub DbQueryOk()
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	Call txtDocCur_OnChange()											
    Set gActiveElement = document.ActiveElement   
    lgBlnFlgChgValue = False
End Sub


'************************************************************************************** 
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	If frm1.vspdData.MaxRows > 0 Then
		With frm1.vspdData
			.Row = .ActiveRow
			.Col =  C_GLNO
				
			arrParam(0) = Trim(.Text)	'결의전표번호 
			arrParam(1) = ""			'Reference번호 
		End With
	Else
		Call DisplayMsgBox("900002", "X","X","X")
		Exit Function
	
	End If
	
	If arrParam(0) = "전표번호" Then Exit Function		' 조회 이전 
	
	IsOpenPop = True   
   
	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function


'--------------------------------------------------------------------------------------------------------------
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim IntRetCD	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	If frm1.vspdData.MaxRows > 0 Then
		With frm1.vspdData
			.Row = .ActiveRow
			.Col =  C_GLNO
				
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
Function OpenMgntPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd, strTempBankCd
	Dim IntRetCD, IntRetCD1
	Dim strFrom, strWhere, strFrom1, strWhere1
	Dim arrVal, arrVal1, arrVal2, arrVal3, arrVal4, arrVal5, arrVal6, arrVal7
	
	Dim arrminor 
	
	DIm stbl_id, scol_id, sdata_id, stbl_id2, scol_id2, sdata_id2
    Dim strMwhere

	If Trim(frm1.txtAcctCd.value) = "" Then
        Call DisplayMsgBox("110131","x","x","x")
        Exit Function
	End If

	If IsOpenPop = True Then Exit Function

		Select Case iWhere
			Case 0, 1	
			
			strMwhere = "A.mgnt_cd1 = B.CTRL_CD AND A.ACCT_CD = '" & frm1.txtAcctCd.value  & "'" 
			
			
			IntRetCD = CommonQueryRs(" TOP 1  TBL_ID, DATA_COLM_ID, DATA_COLM_NM, ISNULL(LTRIM(RTRIM(MAJOR_CD)),'')  ","A_OPEN_ACCT A, A_CTRL_ITEM B",strMwhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			If IntRetCD = true Then

				arrVal = Split(lgF0, Chr(11)) 
				stbl_id = arrVal(0)
				
				arrVal1 = Split(lgF1, Chr(11)) 
				scol_id = arrVal1(0)
				
				arrVal2 = Split(lgF2, Chr(11)) 
				arrVal3 = arrVal2(0)
				
				arrVal   =  Split(lgF3, Chr(11)) 
				arrminor =  arrVal(0)
				
			Else
				IntRetCD = DisplayMsgBox("900014","x","x","x") '☜ 바뀐부분 
				IsOpenPop = False	
				Exit Function				
			End If
			 				
				strFrom = " A_OPEN_ACCT A, " & stbl_id & " B "
				strWhere = " ACCT_CD =  " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ""
				strWhere = strWhere  & " AND A.MGNT_VAL1 = B."&scol_id
						 
						 
				IF arrminor  <> "" then	
					strWhere = strWhere  & " and major_cd = '" & arrminor & "'"
				end if 
				
				
				
						 
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
			
		   	     strMwhere = "A.mgnt_cd2 = B.CTRL_CD AND A.ACCT_CD = '" & frm1.txtAcctCd.value  & "'" 
		   	     
				IntRetCD1 =  CommonQueryRs("TOP 1 TBL_ID, DATA_COLM_ID , DATA_COLM_NM,ISNULL(LTRIM(RTRIM(MAJOR_CD)),'')","A_OPEN_ACCT A, A_CTRL_ITEM B",strMwhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
				If IntRetCD1 = true Then

					arrVal4 = Split(lgF0, Chr(11)) 
					stbl_id2 = arrVal4(0)

					arrVal5 = Split(lgF1, Chr(11)) 
					scol_id2 = arrVal5(0)

					arrVal6 = Split(lgF2, Chr(11)) 
					arrVal7 = arrVal6(0)
					
					
					arrVal   =  Split(lgF3, Chr(11)) 
					
				    arrminor =  arrVal(0)
				Else
					IntRetCD1 = DisplayMsgBox("900014","x","x","x") '☜ 바뀐부분 
					IsOpenPop = False	
					Exit Function				
				End If
			
				strFrom1 = " A_OPEN_ACCT A, " & stbl_id2 & " B "
				strWhere1 = " ACCT_CD =  " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ""
				strWhere1 = strWhere1  & " AND A.MGNT_VAL2 = B."&scol_id2
		 
				IF arrminor  <> "" then	
					strWhere1 = strWhere1  & " and major_cd = '" & arrminor & "'"
				end if  
				
				
				arrParam(0) = "미결코드2팝업"			' 팝업 명칭 
				arrParam(1) = strFrom1	    				' TABLE 명칭 
				arrParam(2) = strCode						' Code Condition
				arrParam(3) = ""							' Name Cindition
				arrParam(4) = strWhere1                      ' Where Condition
				arrParam(5) = "미결코드"				' 조건필드의 라벨 명칭 

				arrField(0) = "MGNT_VAL2"	    			' Field명(0)
				arrField(1) = "B."&arrVal7 	    			' Field명(1)
   
				arrHeader(0) = "미결관리1"				' Header명(0)
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
	   case 3
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
			arrParam(4) = " MGNT_FG=" & FilterVar("Y", "''", "S") & "  "									' Where Condition
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
		Case 2
		frm1.txtDocCur.focus
		frm1.txtDocCur.Value = arrret(0)
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

'	With frm1.vspdData 
'		ggoSpread.Source = frm1.vspdData
'		If Row > 0 Then
'			Select Case Col
'			Case C_AMT1
'				.Col = Col - 1
'				.Row = Row
'				Call OpenZipCode(.Text,Row)
'			End Select
'		End If
'	End With
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
Sub vspdData_Click(Col, Row)

	Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
'    If Row <= 0 Then
'        ggoSpread.Source = frm1.vspdData
'        If lgSortKey = 1 Then
'            ggoSpread.SSSort
'            lgSortKey = 2
'        Else
'            ggoSpread.SSSort ,lgSortKey
'            lgSortKey = 1
'        End If
'    End If
    
End Sub


'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub  

'
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
' Name : fpdtFoundDt_ButtonHit
' Desc : developer describe this line
'========================================================================================================
Sub fpdtFoundDt_ButtonHit(Button, NewIndex)
	On Error Resume Next
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : fpdtCloseDt_ButtonHit
' Desc : developer describe this line
'========================================================================================================
Sub fpdtCloseDt_ButtonHit(Button, NewIndex)
	On Error Resume Next
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
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
		ggoSpread.SSSetFloatByCellOfCur C_AMT0,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_AMT1,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_AMT2,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_AMT, -1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub


'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line 
'========================================================================================================
Function FncPrint() 

Dim StrEbrFile, StrUrl
Dim IntRetCd
	
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

	Dim ValDateFr, ValDateTo, ValAcctCd, ValDocCur,ValMgntCd1Fr, ValMgntCd1To, ValMgntCd2Fr, ValMgntCd2To

	StrEbrFile = "a5409ma1"
	
	With frm1
		ValDateFr		= UniConvDateToYYYYMMDD(frm1.txtdateFr.Text,Parent.gDateFormat, "") '     frm1.txtdatefr.Text
		ValDateTo		= UniConvDateToYYYYMMDD(frm1.txtdateTo.Text,Parent.gDateFormat, "")
		ValAcctCd		= UCase(Trim(.txtAcctCd.value))
		ValDocCur		= UCase(Trim(.txtDocCur.value))
		ValMgntCd1Fr	= ""
		ValMgntCd1To	= "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"
		ValMgntCd2Fr	= ""
		ValMgntCd2To	= "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"

		If Trim(.txtMgntCd1Fr.value) <> "" Then		ValMgntCd1Fr = Trim(.txtMgntCd1Fr.value)
		If Trim(.txtMgntCd1To.value) <> "" Then		ValMgntCd1To = Trim(.txtMgntCd1To.value)
		If Trim(.txtMgntCd2Fr.value) <> "" Then		ValMgntCd2Fr = Trim(.txtMgntCd2Fr.value)
		If Trim(.txtMgntCd2To.value) <> "" Then		ValMgntCd2To = Trim(.txtMgntCd2To.value)
	End With

	' 권한관리 추가 
	If lgAuthBizAreaCd	= "" Then	lgAuthBizAreaCd		= "%"
	If lgInternalCd		= "" Then	lgInternalCd		= "%"
'	If lgSubInternalCd	= "" Then	lgSubInternalCd		= "%"
	If lgAuthUsrID		= "" Then	lgAuthUsrID			= "%"

	StrUrl = StrUrl & "DateFr|"			& ValDateFr
	StrUrl = StrUrl & "|DateTo|"		& ValDateTo
	StrUrl = StrUrl & "|AcctCd|"		& ValAcctCd
	StrUrl = StrUrl & "|ValMgntCd1Fr|"	& ValMgntCd1Fr
	StrUrl = StrUrl & "|ValMgntCd1To|"	& ValMgntCd1To
	StrUrl = StrUrl & "|ValMgntCd2Fr|"	& ValMgntCd2Fr
	StrUrl = StrUrl & "|ValMgntCd2To|"	& ValMgntCd2To	
	StrUrl = StrUrl & "|DocCur|"		& ValDocCur

	StrUrl = StrUrl & "|lgAuthBizAreaCd|"	& lgAuthBizAreaCd
	StrUrl = StrUrl & "|lgInternalCd|"		& lgInternalCd
	StrUrl = StrUrl & "|lgSubInternalCd|"	& lgSubInternalCd & "%"
	StrUrl = StrUrl & "|lgAuthUsrID|"		& lgAuthUsrID

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
					<TD WIDTH=* align=right><A href="vbscript:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;
											<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A></TD>
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
									<TD CLASS="TD5" NOWRAP>작업일자</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDateFr CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="시작일자"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTo name=txtDateTo CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="종료일자"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>거래통화</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="거래통화"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.Value,2)"> 										  
								</TR>

								<TR>
									<TD CLASS=TD5 NOWRAP>계정코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCd.Value,1)"> 
														 <INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="계정명">
									</TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>미결관리1</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtMgntCd1Fr" SIZE=15 MAXLENGTH=30 tag="11XXXU" ALT="미결관리1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgntCd1Fr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMgntPopup(frm1.txtMgntCd1Fr.Value,0)">&nbsp;~&nbsp;
														   <INPUT TYPE="Text" NAME="txtMgntCd1To" SIZE=15 MAXLENGTH=30 tag="11XXXU" ALT="미결관리1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgntCd1To" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMgntPopup(frm1.txtMgntCd1To.Value,1)">
									</TD>
									<TD CLASS="TD5" NOWRAP>미결관리2</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtMgntCd2Fr" SIZE=15 MAXLENGTH=30 tag="11XXXU" ALT="미결관리2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgntCd2Fr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMgntPopup(frm1.txtMgntCd2Fr.Value,2)">&nbsp;~&nbsp;
														   <INPUT TYPE="Text" NAME="txtMgntCd2To" SIZE=15 MAXLENGTH=30 tag="11XXXU" ALT="미결관리2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMgntCd2To" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMgntPopup(frm1.txtMgntCd2To.Value,3)">
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
</DIV>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
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

