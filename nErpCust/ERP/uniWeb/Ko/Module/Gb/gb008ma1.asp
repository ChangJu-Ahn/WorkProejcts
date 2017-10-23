
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 경영손익 
*  2. Function Name        : 
*  3. Program ID           : GB008MA1
*  4. Program Name         : 손익추정액 등록 
*  5. Program Desc         : Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/11/28
*  8. Modified date(Last)  : 2001/12/28
*  9. Modifier (First)     : PARK JAI HONG
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
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
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

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "GB008MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Dim C_ACCT_CD  
Dim C_ACCT_PB  
Dim C_ACCT_NM  
Dim C_COST_CD  
Dim C_COST_PB  
Dim C_COST_NM  
Dim C_AMOUNT  

'Const C_SHEETMAXROWS_D  = 30                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 

Const COOKIE_SPLIT      = 4877	                                      '☆: Cookie Split String
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
'Dim lgIsOpenPop
Dim IsOpenPop                   



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
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
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

	Dim StartDate
	StartDate	= "<%=GetSvrDate%>"                                               'Get Server DB Date
	
	frm1.fpdtWk_yymm.focus
	frm1.fpdtWk_yymm.text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.fpdtWk_yymm, Parent.gDateFormat, 2)
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

	<% Call loadInfTB19029A("I", "G", "NOCOOKIE", "MA") %>

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
' Desc : Make key stream of query or delete condition data
'========================================================================================================
Sub MakeKeyStream(pRow)
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.fpdtWk_yymm.text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth
    lgKeyStream = strYYYYMM & Parent.gColSep       'You Must append one character(Parent.gColSep)
    lgKeyStream = lgKeyStream & frm1.txtMinorCd.value & Parent.gColSep                      'Cost Center
    lgKeyStream = lgKeyStream & Frm1.txtMajorCd.value & Parent.gColSep                      '계정코드     

  '------ Developer Coding part (End   ) -------------------------------------------------------------- 

End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
		'	.Col = C_TYPECd
			intIndex = .value
		'	.col = C_TYPENm
			.value = intindex					
		Next	
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Select Case Col
	    Case C_ACCT_PB   ' 계정코드 
	        frm1.vspdData.Col = C_ACCT_CD
	        Call OpenCost(frm1.vspdData.Text, 1, Row)
	    Case C_COST_PB		'COST CENTER
	        frm1.vspdData.Col = C_COST_CD
	        Call OpenCost(frm1.vspdData.Text, 2, Row)
	End Select    
	Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")   
End Sub



'===========================================================================
' Function Name : OpenCost
' Function Desc : OpenCost Reference Popup
'===========================================================================
Function OpenCost(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function



	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(1) = "a_acct a, g_acct b"	' TABLE 명칭 
	    	arrParam(2) = Trim(strCode)	                        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
   	        arrParam(4) = "a.acct_cd = b.acct_cd and (a.temp_fg_3 in  (" & FilterVar("G2", "''", "S") & "," & FilterVar("G3", "''", "S") & "," & FilterVar("G4", "''", "S") & "," & FilterVar("G5", "''", "S") & "," & FilterVar("G6", "''", "S") & "," & FilterVar("G7", "''", "S") & ") or (a.temp_fg_3= " & FilterVar("G1", "''", "S") & " and b.acct_type = " & FilterVar("T", "''", "S") & " )) and a.DEL_FG <> " & FilterVar("Y", "''", "S") & "  "          <%' Where Condition%>	       
	    	arrParam(5) = "계정코드"		   				    ' TextBox 명칭 
	
	    	arrField(0) = "a.acct_cd"		                ' Field명(0)
	    	arrField(1) = "a.acct_nm"    						' Field명(1)%>
    
	    	arrHeader(0) = "계정코드"		        		' Header명(0)%>
	    	arrHeader(1) = "계정명"	        					' Header명(1)%>
			arrParam(0) = arrParam(5)								  ' 팝업 명칭 


	    Case 2
	       arrParam(1) = "B_COST_CENTER"	<%' TABLE 명칭 %>
	       arrParam(2) = Trim(strCode)	    <%' Code Condition%>
	       arrParam(3) = "" 		            		<%' Name Cindition%>
	       arrParam(4) = "COST_TYPE = " & FilterVar("O", "''", "S") & " "     <%' Where Condition%>
	       arrParam(5) = "Cost Center"			
	
           arrField(0) = "cost_cd"	     			<%' Field명(1)%>
           arrField(1) = "cost_nm"					<%' Field명(0)%>
   
    
           arrHeader(0) = "Cost Center"			    	<%' Header명(0)%>
           arrHeader(1) = "Cost Center명"				<%' Header명(1)%>
			arrParam(0) = arrParam(5)								  ' 팝업 명칭	    
	End Select


	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet, iWhere, Row)
	End If	
	
End Function


'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetCost()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCost(arrRet, iWhere, Row)


	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_ACCT_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_ACCT_NM
		    	.vspdData.text = arrRet(1)
		    Case 2
		        .vspdData.Col = C_COST_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_COST_NM
		    	.vspdData.text = arrRet(1)
		End Select

		lgBlnFlgChgValue = True

	End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
End Function

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ACCT_CD		= 1
	C_ACCT_PB		= 2
	C_ACCT_NM		= 3
	C_COST_CD		= 4
	C_COST_PB		= 5
	C_COST_NM		= 6
	C_AMOUNT		= 7
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	With frm1.vspdData
	
       .MaxCols = C_AMOUNT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
       
        ggoSpread.Source = Frm1.vspdData
		ggoSpread.Spreadinit "V20021130", ,parent.gAllowDragDropSpread

		Call ggoSpread.ClearSpreadData()
		
	   .ReDraw = false
	   
	   Call GetSpreadColumnPos("A") 

       ggoSpread.SSSetEdit  C_ACCT_CD , "계정코드" ,20,   ,, 20
       ggoSpread.SSSetButton C_ACCT_PB
       ggoSpread.SSSetEdit  C_ACCT_NM , "계정명"   ,30,   ,, 30
       ggoSpread.SSSetEdit  C_COST_CD , "Cost Center" ,20,   ,, 10
       ggoSpread.SSSetButton C_COST_PB
       ggoSpread.SSSetEdit  C_COST_NM , "Cost Center명"   ,20,   ,, 20
       ggoSpread.SSSetFloat  C_AMOUNT ,  "금액",20,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	
	   Call ggoSpread.MakePairsColumn(C_ACCT_CD,C_ACCT_PB)	
	   Call ggoSpread.MakePairsColumn(C_COST_CD,C_COST_PB)
	
	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SpreadLock      C_ACCT_CD , -1, C_ACCT_CD
      ggoSpread.SpreadLock      C_ACCT_PB , -1, C_ACCT_PB
      ggoSpread.SpreadLock      C_ACCT_NM , -1, C_ACCT_NM
      ggoSpread.SpreadLock      C_COST_CD , -1, C_COST_CD
      ggoSpread.SpreadLock      C_COST_PB , -1, C_COST_PB
      ggoSpread.SpreadLock      C_COST_NM , -1, C_COST_NM
      ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetRequired    C_ACCT_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_ACCT_NM , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_COST_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_COST_NM , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_AMOUNT, pvStartRow, pvEndRow

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

			C_ACCT_CD			= iCurColumnPos(1)
			C_ACCT_PB			= iCurColumnPos(2)
			C_ACCT_NM   		= iCurColumnPos(3)    
			C_COST_CD      		= iCurColumnPos(4)
			C_COST_PB      		= iCurColumnPos(5)
			C_COST_NM      		= iCurColumnPos(6)
			C_AMOUNT      		= iCurColumnPos(7)
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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
            
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	Call SetToolbar("1100110100101111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
'	Call CookiePage (0)                                                              '☜: Check Cookie
			
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
	
    Call InitVariables																
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

  
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
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
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
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
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
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

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                              '☜: Processing is OK
End Function

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
			SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	With frm1.vspdData
		.col=C_ACCT_CD
		.text=""
		.col=C_ACCT_NM
		.text=""
	END WITH
	'---------------------------------------------------------------------------------------------------- 

	

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

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
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
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
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
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

    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

'    Call LayerShowHide(1)
	If	LayerShowHide(1) = False Then
		Exit Function
	End If
                                                            '☜: Show Processing Message
    
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
'        strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)    '☜: Max fetched data at a time
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
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
    Dim iColSep 
    Dim iRowSep   
	
    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
'    Call LayerShowHide(1)                                                        '☜: Show Processing Message
	If	LayerShowHide(1) = False Then
		Exit Function
	End If

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	  


    Dim strYYYYMM
    Dim strYear,strMonth,strDay

    Call ExtractDateFrom(frm1.fpdtWk_yymm.text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM =   strYear & strMonth

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update추가 
                                                  strVal = strVal & "C" & iColSep
                                                  strVal = strVal & lRow & iColSep
                                                  strval = strval & strYYYYMM & iColSep
                    .vspdData.Col = C_COST_CD	: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_ACCT_CD	: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_AMOUNT	: strVal = strVal & Trim(.vspdData.Text) & iRowSep
                     lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & iColSep
                                                  strVal = strVal & lRow & iColSep
                                                  strval = strval & strYYYYMM& iColSep
                    .vspdData.Col = C_COST_CD	: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_ACCT_CD	: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_AMOUNT	: strVal = strVal & Trim(.vspdData.Text) & iRowSep
                     lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & iColSep
                                                  strDel = strDel & lRow & iColSep
                                                  strDel = strDel & strYYYYMM & iColSep
                    .vspdData.Col = C_COST_CD	: strDel = strDel & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_ACCT_CD	: strDel = strDel & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_AMOUNT	: strDel = strDel & Trim(.vspdData.Text) & iRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = Parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With


	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                               '☜: Processing is OK

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
    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	
    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("1100111100111111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitData()
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
	
    Call MakeKeyStream("X")
	Call SetToolbar("1100111100011111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DBQuery()
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : OpenCondAreaPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    
	arrParam(0) = "계정코드"		    	' 팝업 명칭 
	arrParam(1) = "a_acct a, g_acct b"		' TABLE 명칭 
	arrParam(2) = frm1.txtMajorCd.value	    ' Code Condition
	arrParam(3) = "" 		            		' Name Cindition
    arrParam(4) = "a.acct_cd = b.acct_cd and (a.temp_fg_3 in (" & FilterVar("G2", "''", "S") & "," & FilterVar("G3", "''", "S") & "," & FilterVar("G4", "''", "S") & "," & FilterVar("G5", "''", "S") & "," & FilterVar("G6", "''", "S") & "," & FilterVar("G7", "''", "S") & ") or (a.temp_fg_3 = " & FilterVar("G1", "''", "S") & " and b.acct_type = " & FilterVar("T", "''", "S") & " )) and a.DEL_FG <> " & FilterVar("Y", "''", "S") & "  "          <%' Where Condition%>	       
	arrParam(5) = "계정코드"			
	
    arrField(0) = "a.acct_cd"					' Field명(0)
    arrField(1) = "a.acct_nm"	     			' Field명(1)
    
    arrHeader(0) = "계정코드"				' Header명(0)
    arrHeader(1) = "계정코드명"				' Header명(1)


	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMajorCd.focus
		Exit Function
	Else
		Call SetMajor(arrRet)
	End If	

End Function

'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetMajor(Byval arrRet)
	With frm1
		.txtMajorCd.focus
		.txtMajorCd.value = arrRet(0)
		.txtMajorName.value = arrRet(1)		
	End With
End Function


'========================================================================================================
' Name : OpenCondAreaPopup2()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Cost Center"		    	<%' 팝업 명칭 %>
	arrParam(1) = "B_COST_CENTER"		<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtMinorCd.value	    <%' Code Condition%>
	arrParam(3) = "" 		            		<%' Name Cindition%>
	arrParam(4) = "COST_TYPE = " & FilterVar("O", "''", "S") & " "   <%' Where Condition%>
	arrParam(5) = "Cost Center"			
	
    arrField(0) = "COST_CD"					<%' Field명(0)%>
    arrField(1) = "COST_NM"	     			<%' Field명(1)%>
    
    arrHeader(0) = "Cost Center"				<%' Header명(0)%>
    arrHeader(1) = "Cost Center명"				<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMinorCd.focus
		Exit Function
	Else
		Call SetMinor(arrRet)
	End If	

End Function

'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetMinor(Byval arrRet)
	With frm1
		.txtMinorCd.focus
		.txtMinorCd.value = arrRet(0)
		.txtMinorName.value = arrRet(1)		
	End With
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
 		Call SetFocusToDocument("M")
	    frm1.fpdtWk_yymm.focus
	End If
End Sub

'======================================================================================================
' Name : fpdtWk_yymm_KeyPress
' Desc : Call Mainquery
'=======================================================================================================
Sub fpdtWk_yymm_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
		End If
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(Col , Row )
    Dim iDx
    Dim IntRetCD,EFlag


    Dim COST_CD
    Dim ACCT_CD

    EFlag = False
 
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '=============================계정코드 값 체크 시작 ==================================================
    Frm1.vspdData.Col = C_ACCT_CD
    ACCT_CD = Frm1.vspdData.Text

    If ACCT_CD <>"" Then
        IntRetCD = CommonQueryRs(" A.ACCT_NM "," A_ACCT A,G_ACCT B "," A.ACCT_CD=B.ACCT_CD AND (a.temp_fg_3 in (" & FilterVar("G2", "''", "S") & "," & FilterVar("G3", "''", "S") & "," & FilterVar("G4", "''", "S") & "," & FilterVar("G5", "''", "S") & "," & FilterVar("G6", "''", "S") & "," & FilterVar("G7", "''", "S") & ") or (a.temp_fg_3 = " & FilterVar("G1", "''", "S") & " and b.acct_type = " & FilterVar("T", "''", "S") & "  )) AND A.ACCT_CD=" & FilterVar(ACCT_CD, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD = False Then
		    Call DisplayMsgBox("110100","X","X","X")
		    Frm1.vspdData.Col = C_ACCT_CD
		    Frm1.vspdData.Text = ""
		    Frm1.vspdData.Col = C_ACCT_NM
		    Frm1.vspdData.Text = ""
		    Frm1.vspdData.Action = 0
		    Set gActiveElement = document.activeElement  
		    EFlag = True
	    Else
		    Frm1.vspdData.Col = C_ACCT_NM
		    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
	    End If
    End If
    '=============================계정코드 값 체크 끝 ==================================================


    '=============================cost center 값 체크 시작 ==================================================
    Frm1.vspdData.Col = C_COST_CD
    COST_CD = Frm1.vspdData.Text

    If COST_CD <>"" Then
        IntRetCD = CommonQueryRs("cost_nm","b_cost_center","cost_type = " & FilterVar("O", "''", "S") & "  and cost_cd = " & FilterVar(COST_CD, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD = False Then
		    Call DisplayMsgBox("124400","X","X","X")
		    Frm1.vspdData.Col = C_COST_CD
		    Frm1.vspdData.Text = ""
		    Frm1.vspdData.Col = C_COST_NM
		    Frm1.vspdData.Text = ""
		    Frm1.vspdData.Action = 0
		    Set gActiveElement = document.activeElement  
		    EFlag = True
	    Else
		    Frm1.vspdData.Col = C_COST_NM
		    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
	    End If
    End If
    '=============================cost center 값 체크 끝 ==================================================

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
		
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    '------ Developer Coding part (Start ) --------------------------------------------------------------
    '데이터 확인시 틀린데이터에 대해 undo 해준다.
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0
        
    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()				
	End If
	'------ Developer Coding part (End   ) --------------------------------------------------------------
    
    
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
	Call SetPopupMenuItemInf("1101111111")
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   DbQuery
    	End If
    End if
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>손익조정액 등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>대상년월</TD>
									<TD CLASS="TD6" NOWRAP ><script language =javascript src='./js/gb008ma1_fpdtWk_yymm_fpdtWk_yymm.js'></script></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>							
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>Cost Center</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtMinorCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup2()">
										<INPUT TYPE="Text" NAME="txtMinorName" SiZE=20 MAXLENGTH=20 tag="14XXXU" ALT="Cost Center명">
									</TD>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtMajorCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup()">
										<INPUT TYPE="Text" NAME="txtMajorName" SiZE=20 MAXLENGTH=20 tag="14XXXU" ALT="계정코드명"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/gb008ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
							    	<TD CLASS=TD6 NOWRAP></TD>
							    	<TD CLASS=TD6 NOWRAP></TD>
							    	<TD CLASS=TD5 NOWRAP>합계</TD>
							    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/gb008ma1_fpDoubleSingle2_txtSpouse_amt.js'></script></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>

				</TR>
				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%> ><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS=HIDDEN NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMode"    tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hPayCd"     tag="24" TABINDEX = "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

