
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 경영손익 
*  2. Function Name        :
*  3. Program ID           : GH003MA1
*  4. Program Name         : 품목그룹별 배부기준 DATA 등록 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/06
*  8. Modified date(Last)  : 2002/01/04
*  9. Modifier (First)     : Kim Kyoung Ho
* 10. Modifier (Last)      : Kim Kyoung Ho
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
Const BIZ_PGM_ID = "gh003mb1.asp"                                      'Biz Logic ASP
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Dim C_SalesGrp
Dim C_SalesGrp_bt
Dim C_SalesGrpNm
Dim C_center                                                     'Column Dimant for Spread Sheet
Dim C_center_bt 
Dim C_centernm  
Dim C_data      

Const COOKIE_SPLIT      = 4877	                                      '☆: Cookie Split String
'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop
Dim lsConcd


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
Sub MakeKeyStream(pOpt)

    Dim strYYYYMM
    Dim strYear,strMonth,strDay

   '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.fpdtWk_yymm.text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth
    lgKeyStream = strYYYYMM & Parent.gColSep       'You Must append one character(Parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtCurrencyCode.value & Parent.gColSep                      '계정그룹 


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
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	 C_SalesGrp	 = 1
	 C_SalesGrp_bt	 = 2
	 C_SalesGrpNm	 = 3
	 C_center    = 4                                                 'Column  Spread Sheet
	 C_center_bt = 5
	 C_centernm  = 6
	 C_data      = 7

End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	With frm1.vspdData

       .MaxCols = C_data + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

        ggoSpread.Source = Frm1.vspdData
		ggoSpread.Spreadinit "V20021217", ,parent.gAllowDragDropSpread

		ggoSpread.ClearSpreadData
		
	   .ReDraw = false

       Call GetSpreadColumnPos("A")

       Call AppendNumberPlace("6","3","0")
       
       ggoSpread.SSSetEdit  C_SalesGrp   , "영업그룹"      ,15,,, 35,2
       ggoSpread.SSSetButton C_SalesGrp_bt
       ggoSpread.SSSetEdit  C_SalesGrpNm , "영업그룹명"   ,25,,, 45,2
       ggoSpread.SSSetEdit  C_center   , "품목그룹"      ,15,,, 35,2
       ggoSpread.SSSetButton C_center_bt
       ggoSpread.SSSetEdit  C_centernm , "품목그룹명"   ,25,,, 45,2
       ggoSpread.SSSetFloat C_data     , "배부 Data"        ,20, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	  	
	   Call ggoSpread.MakePairsColumn(C_SalesGrp,C_SalesGrp_bt)	
	   Call ggoSpread.MakePairsColumn(C_center,C_center_bt)
	   	
	   .ReDraw = true


       Call SetSpreadLock(-1,-1)

    End With
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal lRow  , ByVal lRow2 )
    With frm1

    .vspdData.ReDraw = False

      ggoSpread.SpreadLock		C_SalesGrp,		lRow,  C_SalesGrp,		lRow2      
      ggoSpread.SpreadLock		C_SalesGrp_bt,	lRow,  C_SalesGrp_bt,		lRow2     
      ggoSpread.SpreadLock		C_SalesGrpNm,	lRow,  C_SalesGrpNm,		lRow2     
      ggoSpread.SpreadLock		C_center,		lRow,  C_center,		lRow2      
      ggoSpread.SpreadLock		C_center_bt,	lRow,  C_center_bt,		lRow2     
      ggoSpread.SpreadLock		C_centernm,		lRow,  C_centernm,		lRow2     
      ggoSpread.SpreadLock		C_data,			lRow,  C_data,			lRow2 
      ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub



'Sub SetSpreadYn_Lock(ByVal lRow  , ByVal lRow2 )
'    With frm1		
'    .vspdData.ReDraw = False
'      ggoSpread.SpreadLock		C_center,		lRow,  C_center,		lRow2      
'      ggoSpread.SpreadLock		C_center_bt,	lRow,  C_center_bt,		lRow2     
'      ggoSpread.SpreadLock		C_centernm,		lRow,  C_centernm,		lRow2     
'      ggoSpread.Spreadlock		C_data,			lRow,  C_data,			lRow2 
'      ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
'    .vspdData.ReDraw = True

'    End With
'End Sub


'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

			.vspdData.ReDraw = False
			ggoSpread.SSSetRequired    C_SalesGrp		,pvStartRow, pvEndRow
			ggoSpread.SSSetProtected    C_SalesGrpNm	,pvStartRow, pvEndRow
			
			ggoSpread.SSSetRequired    C_center		,pvStartRow, pvEndRow
			ggoSpread.SSSetProtected    C_centernm	,pvStartRow, pvEndRow
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
'======================================================================================================
'	Name : OpenCurrency()
'	Description : Major PopUp
'=======================================================================================================
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(6), arrField(5), arrHeader(5)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	arrParam(0) = "배부기준"		    	    <%' 팝업 명칭 %>
	arrParam(1) = "b_configuration a, b_minor b" 	<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtCurrencyCode.value		<%' Code Cindition%>
	arrParam(3) = ""								<%' Name Condition%>
	arrParam(4) = "b.major_cd = " & FilterVar("G1004", "''", "S") & " "
    arrParam(4) = arrParam(4) & "  and a.minor_cd = b.minor_cd "
	arrParam(4) = arrParam(4) & "  and a.major_cd = b.major_cd "
	arrParam(4) = arrParam(4) & "  AND  a.seq_no =5  and reference = " & FilterVar("Y", "''", "S") & "   "				<%' Where Condition%>
	arrParam(5) = "배부기준 코드"

    arrField(0) = "a.minor_cd"					<%' Field명(0)%>
    arrField(1) = "b.minor_nm"	     			<%' Field명(1)%>


    arrHeader(0) = "배부기준 코드"				<%' Header명(0)%>
    arrHeader(1) = "배부기준 코드명"				<%' Header명(1)%>


	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCurrencyCode.focus
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
		.txtCurrencyCode.focus
		.txtCurrencyCode.value = arrRet(0)
		.txtCurrency.value	   = arrRet(1)


			 Call CommonQueryRs("reference","b_configuration", " seq_no =1  and major_cd = " & FilterVar("G1004", "''", "S") & " and minor_cd =  " & FilterVar(arrRet(0), "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if Trim(Replace(lgF0,Chr(11),"")) = "X" then
					.txtYntag.value = ""
		       else
					.txtYntag.value = Trim(Replace(lgF0,Chr(11),""))
			   end if

		if .txtYntag.value = UCase("Y") then
		   .txtYn.value = "자동생성"
	   		call spreadflag()
		else
		   .txtYn.value = "수작업입력"
			call spreadflag()
		end if

	End With

End Function
'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Select Case Col				'추가부분을 위해..select로..
	    Case C_center_bt        'Cost center
	        frm1.vspdData.Col = C_center
	        Call OpenCost(frm1.vspdData.Text, 1, Row)
	    Case C_SalesGrp_bt        'Cost center
	        frm1.vspdData.Col = C_SalesGrp
	        Call OpenCost(frm1.vspdData.Text, 2, Row)

	End Select
	Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")
End Sub

'===========================================================================
' Function Name : OpenCode
' Function Desc : OpenCode Reference Popup
'===========================================================================
Function OpenCost(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
			arrParam(1) = "B_ITEM_GROUP"           <%' TABLE 명칭 %>
			arrParam(2) = ""                        <%' Code Condition%>
			arrParam(3) = "" 		            	<%' Name Cindition%>
			arrParam(4) = ""                        <%' Where Condition%>
			arrParam(5) = "품목그룹"

			arrField(0) = "ITEM_GROUP_CD"	     			<%' Field명(1)%>
			arrField(1) = "ITEM_GROUP_NM"					<%' Field명(0)%>

			arrHeader(0) = "품목그룹코드"			    	<%' Header명(0)%>
			arrHeader(1) = "품목그룹명"				<%' Header명(1)%>
	    Case 2
			arrParam(1) = "B_SALES_GRP"           <%' TABLE 명칭 %>
			arrParam(2) = ""                        <%' Code Condition%>
			arrParam(3) = "" 		            	<%' Name Cindition%>
			arrParam(4) = ""                        <%' Where Condition%>
			arrParam(5) = "영업그룹"

			arrField(0) = "SALES_GRP"	     			<%' Field명(1)%>
			arrField(1) = "SALES_GRP_NM"					<%' Field명(0)%>

			arrHeader(0) = "영업그룹코드"			    	<%' Header명(0)%>
			arrHeader(1) = "영업그룹명"				<%' Header명(1)%>
			
	End Select

    arrParam(3) = ""
	arrParam(0) = arrParam(5)														 ' 팝업 명칭 

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
		        .vspdData.Col = C_center
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_centernm
		    	.vspdData.text = arrRet(1)
		    Case 2
		        .vspdData.Col = C_SalesGrp
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_SalesGrpNm
		    	.vspdData.text = arrRet(1)
		    	
		End Select

		lgBlnFlgChgValue = True

	End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Function

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
			C_SalesGrp            	= iCurColumnPos(1)
			C_SalesGrp_bt       		= iCurColumnPos(2)
			C_SalesGrpNm           	= iCurColumnPos(3)    
			C_center            	= iCurColumnPos(4)
			C_center_bt       		= iCurColumnPos(5)
			C_centernm           	= iCurColumnPos(6)    
			C_data                	= iCurColumnPos(7)
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
'   Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
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

	Call SetToolbar("1100000000001111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   --------------------------------------------------------------
'    Call InitComboBox
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

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
	Call InitVariables
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------

'    Call InitVariables                                                               '⊙: Initializes local global variables
'    Call SetDefaultVal
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   --------------------------------------------------------------

    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncQuery = True

End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																 '☜: Processing is NG
    Err.Clear
                                                                '☜: Clear err status
    Call ggoOper.ClearField(Document, "1")                                        '☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                        '☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")
	'------ Developer Coding part (Start ) --------------------------------------------------------------
   	Call SetToolbar("1100111100001111")                                           '☆: Developer must customize
    Call SetDefaultVal
    Call InitVariables
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
    Err.Clear

   If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           'Check if there is retrived data
   Call DisplayMsgBox("900002","X","X","X")                                 '☜: Please do Display first.
   Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		                 '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

                                                                  '☜: Clear err status
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

    If Not chkField(Document, "2") Then
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData


If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If


    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call MakeKeyStream("X")
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
	' Clear key field
	'----------------------------------------------------------------------------------------------------

	With Frm1.VspdData
           .Col  = C_center
           .Row  = .ActiveRow
           .Text = ""
    End With

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
    
    FncInsertRow = False                                                               '☜: Clear err status

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If


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
    FncDeleteRow = True
                                                 '☜: Processing is OK
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

    if LayerShowHide(1) = false then
 exit function
end if
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
	Dim strVal
	Dim strDel
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    Dim iColSep 
    Dim iRowSep   

    Err.Clear                                                                    '☜: Clear err status

    DbSave = False                                                               '☜: Processing is NG
    if LayerShowHide(1) = false then
        exit function
    end if

    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    Call ExtractDateFrom(frm1.fpdtWk_yymm.text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth
  	With frm1
		.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    lGrpCnt = 1
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	  

  	With frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

               Case ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & iColSep
                                                  strVal = strVal & lRow & iColSep'
                                                  strval = strval & Trim(.txtCurrencyCode.value) & iColSep
                    .vspdData.Col = C_SalesGrp	: strVal = strVal & Trim(.vspdData.text) & iColSep                                                  
                    .vspdData.Col = C_center	: strVal = strVal & Trim(.vspdData.text) & iColSep
                    .vspdData.Col = C_data   	: strVal = strVal & Trim(.vspdData.text) & iRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.UpdateFlag                                      '☜: Update

                                                  strVal = strVal & "U" & iColSep
                                                  strVal = strVal & lRow & iColSep
                                                  strval = strval & strYYYYMM& iColSep
                                                  strval = strval & Trim(.txtCurrencyCode.value) & iColSep
                    .vspdData.Col = C_SalesGrp	: strVal = strVal & Trim(.vspdData.text) & iColSep                                                  
                   .vspdData.Col = C_center 	: strVal = strVal & Trim(.vspdData.Text) & iColSep
                   .vspdData.Col = C_data   	: strVal = strVal & Trim(.vspdData.Text) & iRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & iColSep
												  strDel = strDel & lRow & iColSep
                                                  strDel = strDel & strYYYYMM & iColSep
                                                  strDel = strDel & Trim(.txtCurrencyCode.value) & iColSep
                    .vspdData.Col = C_SalesGrp	: strVal = strVal & Trim(.vspdData.text) & iColSep                                                  
                    .vspdData.Col = C_center 	: strDel = strDel & Trim(.vspdData.Text) & iRowSep

                                        lGrpCnt = lGrpCnt + 1
           End Select
       Next


  	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      =  strDel & strVal

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
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
 if LayerShowHide(1) = false then
 exit function
end if


    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    DbDelete = True                                                              '☜: Processing is OK
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()

    lgIntFlgMode = Parent.OPMD_UMODE
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  --------------------------------------------------------------
    Call MakeKeyStream("X")
     Call ggoOper.ClearField(Document, "2")									     '⊙: Clear Contents  Field
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
	Call InitVariables()
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD,Input_alloc,  EFlag
	
	EFlag = False
	
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) --------------------------------------------------------------
				
	IF Col = C_center Then	
		Input_alloc = Frm1.vspdData.Text
	
		IntRetCD = CommonQueryRs(" item_nm ","B_COST_CENTER A  "," cost_cd = " & FilterVar(Input_alloc, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
		If IntRetCD = False Then		
			Call DisplayMsgBox("GB3101","X","X","X")
			Frm1.vspdData.Col = C_center
			Frm1.vspdData.Action = 0 		
			EFlag = True
		Else
			Frm1.vspdData.Col = C_centernm
			Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
		End If
	ElseIF Col = C_SalesGrp  Then
		Input_alloc = Frm1.vspdData.Text
	
		IntRetCD = CommonQueryRs(" item_nm ","B_SALES_GRP A  "," SalesGrp = " & FilterVar(Input_alloc, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
		If IntRetCD = False Then		
			Call DisplayMsgBox("GB3101","X","X","X")
			Frm1.vspdData.Col = C_SalesGrp
			Frm1.vspdData.Action = 0 		
			EFlag = True
		Else
			Frm1.vspdData.Col = C_SalesGrpNm
			Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
		End If
	ENd IF	
	
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    If EFlag Then
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.EditUndo Row
		Set gActiveElement = document.ActiveElement
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
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

Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
    	Call SetFocusToDocument("M")
		frm1.fpdtWk_yymm.focus
	End If
End Sub

'========================================================================================================
'   Event Name : spreadflag
'   Event Desc : This function is toolbar control
'========================================================================================================
Sub spreadflag()

	if frm1.txtYNtag.value = UCase("y") then

		Call SetToolbar("1100000000011111")
		Call SetSpreadLock(-1,-1)
	else
		Call SetToolbar("1100111100111111")
		Call SetSpreadLock(-1,-1)
	end if

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





</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목그룹별 기준Data</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/gh003ma1_fpDateTime3_fpdtWk_yymm.js'></script>
										<INPUT TYPE=HIDDEN NAME=ALLOC_KINDS VALUE = 1>
									</TD>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>
								<TR>
									<TD NOWRAP CLASS="TD5">배부기준</TD>
									<TD NOWRAP CLASS="TD6">

										<INPUT TYPE=TEXT   NAME="txtCurrencyCode" SIZE=5 MAXLENGTH=1 tag="12XXXU"  ALT="배부기준" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCurrency()">
										<INPUT TYPE=TEXT   NAME="txtCurrency" TAG="14XXU" SIZE=25 MAXLENGTH"30">

									</TD>
									<TD NOWRAP CLASS="TD5">자동생성여부</TD>
									<TD NOWRAP CLASS="TD6">
										<INPUT TYPE=HIDDEN NAME="txtYNtag" SIZE=11 MAXLENGTH=2 tag="14XX">
										<INPUT TYPE=TEXT NAME="txtYN" SIZE=20 MAXLENGTH=10 tag="14XX">
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
										<script language =javascript src='./js/gh003ma1_OBJECT1_vspdData.js'></script>
								</TD>
								</TR>
							</TABLE>
					</TD>
				</TR>
				<TR>
					<TD>
							<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_30%>>
										<TR HEIGHT=20>
											<TD CLASS=TD6 NOWRAP></TD>
							    			<TD CLASS=TD6 NOWRAP></TD>
							    			<TD CLASS=TD5 NOWRAP>합계</TD>
							    			<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/gh003ma1_fpDoubleSingle2_txtDataAmt.js'></script></TD>
										</TR>
							</TABLE>
							</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              
