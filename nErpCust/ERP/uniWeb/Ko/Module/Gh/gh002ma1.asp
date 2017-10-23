
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : GE
*  2. Function Name        :
*  3. Program ID           : gh002ma1
*  4. Program Name         : 경영손익 품목그룹별 배부기준 등록 
*  5. Program Desc         : Single-Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/08
*  8. Modified date(Last)  : 2001/12/31
*  9. Modifier (First)     : song sang min
* 10. Modifier (Last)      : song sang min
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

Const BIZ_PGM_ID      = "gh002mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_alloc_base_CODE  
Dim C_alloc_base_old   
Dim C_alloc_from       
Dim C_ACCT_PB          
Dim C_cost_nm		   
Dim C_acct_gp          
Dim c_code_gb          
Dim C_acct_gm          
Dim C_acct_cd 		 
Dim c_code_pb        
Dim C_acct_nm	     
Dim C_alloc_base	
Dim C_curr_PB       
Dim C_alloc_rate	

Const C_SHEETMAXROWS_D   = 100                                          '☜: Fetch count at a time
Const C_SHEETMAXROWS     = 100
Const COOKIE_SPLIT       = 4877	                                      'Cookie Split String

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop
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
	lgStrPrevKey = ""                                           'initializes Previous Key
    lgLngCurRows = 0                                            'initializes Deleted Rows Count
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

   <% Call loadInfTB19029A("I", "G", "NOCOOKIE", "MA") %>  'batch= B , print = P , input = I

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
    Dim strYYYYMM,strYYYYMM1
    Dim strYear,strMonth,strDay
    Dim strYear1,strMonth1,strDay1
    Dim temp_date	

    '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM =   strYear & strMonth
    
    temp_date = UNIDateAdd("M",-1,frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat)
    Call ExtractDateFrom(temp_date,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear1,strMonth1,strDay1)	
    strYYYYMM1 =   strYear1 & strMonth1


    lgKeyStream = strYYYYMM & Parent.gColSep       '날짜 
    lgKeyStream = lgKeyStream & Trim(Frm1.txtCode.Value)    & Parent.gColSep '계정그룹코드 
    lgkeyStream = lgkeyStream & Trim(frm1.fpdtWk_yymm.text) & Parent.gColSep '날짜 
    lgkeyStream = lgkeyStream & Trim(Frm1.txtCost.Value)    & Parent.gColSep 'cost center
    lgkeyStream = lgkeyStream & strYYYYMM1    & Parent.gColSep
    lgkeyStream = lgkeyStream & Trim(Frm1.txtCurrencyCode.Value) & Parent.gColSep
   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
    Dim iDx
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생하는 콤보 박스 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    Dim acct_gp

	Select Case Col
	    Case C_ACCT_PB
	        frm1.vspdData.Col = C_alloc_from
	        Call OpenCost(frm1.vspdData.Text, 1, Row)
	    Case c_code_pb
	        frm1.vspdData.Col = C_acct_cd
	        Call OpenCost(frm1.vspdData.Text, 2, Row)
	    Case C_curr_PB
	        frm1.vspdData.Col = C_alloc_base
	        Call OpenCost(frm1.vspdData.Text, 3, Row)
	    Case c_code_gb
	        frm1.vspdData.Col = C_acct_gp
	        Call OpenCost(frm1.vspdData.Text,4 , Row)

	        Frm1.vspdData.Row = Row
    		acct_gp       = Frm1.vspdData.Text

    		ggoSpread.Source = Frm1.vspdData

    	 	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	 		if acct_gp = "" then
      	    		ggoSpread.SpreadLock    C_acct_cd, Row, C_acct_cd,Row
    	    		ggoSpread.SpreadLock    C_code_pb, Row, C_code_pb,Row
    	   		else
    	   			ggoSpread.SpreadUnLock    C_acct_cd, Row, C_acct_cd,Row
    	    		ggoSpread.SpreadUnLock    C_code_pb, Row, C_code_pb,Row
           		end if
		End If
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
    Dim acct_gp

	If IsOpenPop = True Then Exit Function

    Frm1.vspdData.Col = C_acct_gp
	Frm1.vspdData.Row = Row
    acct_gp       = Frm1.vspdData.Text


	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(1) = "b_sales_grp"	' TABLE 명칭 
	    	arrParam(2) = Trim(strCode) 	' Code Condition
	    	arrParam(3) = "" 				' Name Cindition
	    	arrParam(4) = "" 		        ' Where Condition
	    	arrParam(5) = "영업그룹"	' TextBox 명칭 

	    	arrField(0) = "sales_grp"		    ' Field명(0)
	    	arrField(1) = "sales_grp_nm"    		' Field명(1)%>

	    	arrHeader(0) = "영업그룹"	' Header명(0)%>
	    	arrHeader(1) = "영업그룹명"	' Header명(1)%>

	    Case 2
	       arrParam(1) = "a_acct a, g_acct b"	            <%' TABLE 명칭 %>
	       arrParam(2) = Trim(strCode)	     	    <%' Code Condition%>
	       arrParam(3) = "" 		                <%' Name Cindition%>
	       arrParam(4) = "a.acct_cd = b.acct_cd and a.gp_cd = " & FilterVar(acct_gp, "''", "S") & " and (a.temp_fg_3 in (" & FilterVar("G2", "''", "S") & "," & FilterVar("G3", "''", "S") & "," & FilterVar("G4", "''", "S") & "," & FilterVar("G5", "''", "S") & "," & FilterVar("G6", "''", "S") & "," & FilterVar("G7", "''", "S") & ") or (a.temp_fg_3 = " & FilterVar("G1", "''", "S") & " and b.acct_type = " & FilterVar("T", "''", "S") & " )) "          <%' Where Condition%>
	       arrParam(5) = "계정코드"

           arrField(0) = "a.acct_cd"	     	  	<%' Field명(1)%>
           arrField(1) = "a.acct_nm"			    <%' Field명(0)%>

           arrHeader(0) = "계정코드"	  		    <%' Header명(0)%>
           arrHeader(1) = "계정명"		  	    <%' Header명(1)%>

	     Case 3
	    	arrParam(1) = " b_minor BM,B_CONFIGURATION BC"	' TABLE 명칭 
	    	arrParam(2) = ""	                	    ' Code Condition
	    	arrParam(3) = ""						' Name Cindition
	    	arrParam(4) = "BM.major_cd = " & FilterVar("G1004", "''", "S") & " AND   BM.MINOR_CD = BC.MINOR_CD AND   BC.SEQ_NO = 5 AND  REFERENCE  = " & FilterVar("Y", "''", "S") & " "' Where Condition
	    	arrParam(5) = "배부기준"  					 ' TextBox 명칭 

	    	arrField(0) = "BM.MINOR_CD"						 ' Field명(0)
	    	arrField(1) = "BM.MINOR_NM"    					 ' Field명(1)

	    	arrHeader(0) = "배부기준"		        	 ' Header명(0)
	    	arrHeader(1) = "배부기준명"	       			 ' Header명(1)
      Case 4
	        arrParam(1) = "a_acct_gp "	                    <%' TABLE 명칭 %>
	        arrParam(2) = Trim(strCode)	                    <%' Code Condition%>
	        arrParam(3) = "" 		                        <%' Name Cindition%>
	        arrParam(4) = "gp_cd in (select distinct gp_cd from a_acct where temp_fg_3 in (" & FilterVar("G1", "''", "S") & "," & FilterVar("G2", "''", "S") & "," & FilterVar("G3", "''", "S") & "," & FilterVar("G4", "''", "S") & "," & FilterVar("G5", "''", "S") & "," & FilterVar("G6", "''", "S") & "," & FilterVar("G7", "''", "S") & ")) "            <%' Where Condition%>
	        arrParam(5) = "계정그룹"

           arrField(0) = "gp_cd"	     	  <%' Field명(1)%>
           arrField(1) = "gp_nm"			  <%' Field명(0)%>

           arrHeader(0) = "계정그룹"	  <%' Header명(0)%>
           arrHeader(1) = "계정그룹명"		  <%' Header명(1)%>
	End Select

    arrParam(3) = ""
  	arrParam(0) = arrParam(5)								  ' 팝업 명칭 

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet, iWhere, Row)
	End If

End Function

'------------------------------------------  SetCode()  --------------------------------------------------
'	Name : SetCode()
'	Description : OpenCode Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCost(arrRet, iWhere, Row)


	With frm1
    .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_alloc_from
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_cost_nm
		    	.vspdData.text = arrRet(1)
		    Case 2
		        .vspdData.Col = C_acct_cd
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_acct_nm
		    	.vspdData.text = arrRet(1)
		    Case 3
		        .vspdData.Col = C_alloc_base
		    	.vspdData.text = arrRet(1)
		    	.vspdData.Col = C_alloc_base_CODE
		    	.vspdData.text = arrRet(0)
		   Case 4
		        .vspdData.Col = C_acct_gp
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_acct_gm
		    	.vspdData.text = arrRet(1)
		End Select

	lgBlnFlgChgValue = True

	End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Function

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
	C_alloc_base_CODE		= 1
	C_alloc_base_old		= 2
	C_alloc_from			= 3
	C_ACCT_PB				= 4
	C_cost_nm				= 5
	C_acct_gp				= 6
	c_code_gb				= 7
	C_acct_gm				= 8
	C_acct_cd 				= 9
	c_code_pb				= 10
	C_acct_nm				= 11
	C_alloc_base			= 12
	C_curr_PB				= 13
	C_alloc_rate			= 14	 
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	With frm1.vspdData

       .MaxCols   = C_alloc_rate + 1                                                 ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021202", ,parent.gAllowDragDropSpread

		ggoSpread.ClearSpreadData

	   .ReDraw = false

       Call AppendNumberPlace("6","3","0")

        Call GetSpreadColumnPos("A") 

        Select case CStr(frm1.txtCurrencyCode.value)
            case "1"
				ggoSpread.SSSetEdit     C_alloc_base_CODE ,     "배부코드"       ,2,,,5,2
				ggoSpread.SSSetEdit     C_alloc_base_old ,      "배부코드old"       ,2,,,5,2
				ggoSpread.SSSetEdit     C_alloc_from      ,     "영업그룹"    ,12,,,10,2
				ggoSpread.SSSetButton   C_ACCT_PB
				ggoSpread.SSSetEdit     C_cost_nm         ,     "영업그룹명"  ,22,,,20
				ggoSpread.SSSetEdit     C_acct_gp         ,     "계정그룹"       ,15,,,20,2
				ggoSpread.SSSetButton   c_code_gb
				ggoSpread.SSSetEdit     C_acct_gm         ,     "계정그룹명"     ,30,,,30
				ggoSpread.SSSetEdit     C_acct_cd         ,     "계정코드"       ,20,,,20,2
				ggoSpread.SSSetButton   c_code_pb
				ggoSpread.SSSetEdit     C_acct_nm         ,     "계정명"         ,30,,,30
				ggoSpread.SSSetEdit     C_alloc_base      ,     "배부기준"       ,15,,,50,2
				ggoSpread.SSSetButton   C_curr_PB
				ggoSpread.SSSetFloat    C_alloc_rate      ,     "배부기준율"     ,15,6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

                ggoSpread.SSSetRequired	C_alloc_rate, -1, -1
               
                Call ggoSpread.SSSetColHidden(C_alloc_from  ,C_alloc_from  ,True)
				Call ggoSpread.SSSetColHidden(C_ACCT_PB  ,C_ACCT_PB  ,True)
				Call ggoSpread.SSSetColHidden(C_cost_nm  ,C_cost_nm  ,True)

                frm1.txtCode.style.visibility   = "visible"    ' visible
                frm1.btnCode.style.visibility   = "visible"
                frm1.txtCost.style.visibility   = "hidden"
                frm1.btnCode1.style.visibility  = "hidden"
                frm1.txtCodeh.style.visibility  = "visible"
                frm1.txtCosth.style.visibility  = "hidden"
                
                TitleCC.innerHTML = ""	
			
			case "2"
				ggoSpread.SSSetEdit     C_alloc_base_CODE ,     "배부코드"       ,2,,,5,2
				ggoSpread.SSSetEdit     C_alloc_base_old ,      "배부코드old"       ,2,,,5,2
				ggoSpread.SSSetEdit     C_alloc_from      ,     "영업그룹"    ,30,,,10,2
				ggoSpread.SSSetButton   C_ACCT_PB
				ggoSpread.SSSetEdit     C_cost_nm         ,     "영업그룹명"  ,40,,,20,2
				ggoSpread.SSSetEdit     C_acct_gp         ,     "계정그룹"       ,12,,,20,2
				ggoSpread.SSSetButton   c_code_gb
				ggoSpread.SSSetEdit     C_acct_gm         ,     "계정그룹명"     ,22,,,30,2
				ggoSpread.SSSetEdit     C_acct_cd         ,     "계정코드"       ,12,,,20,2
				ggoSpread.SSSetButton   c_code_pb
				ggoSpread.SSSetEdit     C_acct_nm         ,     "계정명"         ,22,,,30,2
				ggoSpread.SSSetEdit     C_alloc_base      ,     "배부기준"       ,25,,,50,2
				ggoSpread.SSSetButton   C_curr_PB
				ggoSpread.SSSetFloat    C_alloc_rate      ,     "배부기준율"     ,25,6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
			
   		        ggoSpread.SSSetRequired	C_alloc_rate, -1, -1
		        
				Call ggoSpread.SSSetColHidden(C_acct_gp  ,C_acct_gp  ,True)
				Call ggoSpread.SSSetColHidden(c_code_pb  ,c_code_pb  ,True)
				Call ggoSpread.SSSetColHidden(C_acct_gm  ,C_acct_gm  ,True)
				Call ggoSpread.SSSetColHidden(C_acct_cd  ,C_acct_cd  ,True)
				Call ggoSpread.SSSetColHidden(c_code_pb  ,c_code_pb  ,True)
				Call ggoSpread.SSSetColHidden(C_acct_nm  ,C_acct_nm  ,True)

		        frm1.txtCode.style.visibility   = "hidden"    ' visible
                frm1.btnCode.style.visibility   = "hidden"
                frm1.txtCost.style.visibility   = "visible"
                frm1.btnCode1.style.visibility  = "visible"
                frm1.txtCodeh.style.visibility  = "hidden"
                frm1.txtCosth.style.visibility  = "visible"
                
                TitleACCT.innerHTML = ""

	        case "3"
		        ggoSpread.SSSetEdit     C_alloc_base_CODE ,     "배부코드"       ,2,,,5,2
				ggoSpread.SSSetEdit     C_alloc_base_old ,      "배부코드old"       ,2,,,5,2
				ggoSpread.SSSetEdit     C_alloc_from      ,     "영업그룹"    ,12,,,10,2
				ggoSpread.SSSetButton   C_ACCT_PB
				ggoSpread.SSSetEdit     C_cost_nm         ,     "영업그룹명"  ,20,,,20
				ggoSpread.SSSetEdit     C_acct_gp         ,     "계정그룹"       ,12,,,20,2
				ggoSpread.SSSetButton   c_code_gb
				ggoSpread.SSSetEdit     C_acct_gm         ,     "계정그룹명"     ,22,,,30
				ggoSpread.SSSetEdit     C_acct_cd         ,     "계정코드"       ,12,,,20,2
				ggoSpread.SSSetButton   c_code_pb
				ggoSpread.SSSetEdit     C_acct_nm         ,     "계정명"         ,22,,,30
				ggoSpread.SSSetEdit     C_alloc_base      ,     "배부기준"       ,12,,,50,2
				ggoSpread.SSSetButton   C_curr_PB
				ggoSpread.SSSetFloat    C_alloc_rate      ,     "배부기준율"     ,12,6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

               	ggoSpread.SSSetRequired	C_alloc_rate, -1, -1
               	
		        frm1.txtCode.style.visibility   = "visible"    ' visible
                frm1.btnCode.style.visibility   = "visible"
                frm1.txtCost.style.visibility   = "visible"
                frm1.btnCode1.style.visibility  = "visible"
                frm1.txtCodeh.style.visibility  = "visible"
                frm1.txtCosth.style.visibility  = "visible"
                frm1.txtCode.value = ""
                frm1.txtCodeh.value = ""
                frm1.txtCost.value = ""
                frm1.txtCosth.value = ""
		end select
       
		Call ggoSpread.MakePairsColumn(C_alloc_from,C_ACCT_PB)
		Call ggoSpread.MakePairsColumn(C_acct_gp,c_code_gb)
		Call ggoSpread.MakePairsColumn(C_acct_cd,c_code_pb)
		Call ggoSpread.MakePairsColumn(C_alloc_base,C_curr_PB)
		
		Call ggoSpread.SSSetColHidden(C_ALLOC_BASE_CODE,C_ALLOC_BASE_CODE,True)	
		Call ggoSpread.SSSetColHidden(C_alloc_base_old,C_alloc_base_old,True)

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
      ggoSpread.SpreadLock      C_alloc_from, -1, C_alloc_from
      ggoSpread.SpreadLock      C_ACCT_PB, -1, C_ACCT_PB
      ggoSpread.SpreadLock      C_cost_nm, -1, C_cost_nm
      ggoSpread.SpreadLock      C_acct_gp, -1, C_acct_gp
      ggoSpread.SpreadLock      c_code_gb, -1, c_code_gb
      ggoSpread.SpreadLock      C_acct_gm, -1, C_acct_gm
      ggoSpread.SpreadLock      C_acct_cd, -1, C_acct_cd
      ggoSpread.SpreadLock      c_code_pb, -1, c_code_pb
      ggoSpread.SpreadLock      C_acct_nm, -1, C_acct_nm
      ggoSpread.SSSetRequired	C_alloc_base, -1, -1
      ggoSpread.SSSetRequired	C_alloc_rate, -1, -1
      ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True
    End With
End Sub
'======================================================================================================
'	Name : initMinor()
'	Description : 폼 로드시에 배부유형을 셋팅한다.
'=======================================================================================================%>
Function initMinor()
	Dim intRetCD   	  
	intRetCD = CommonQueryRs(" bm.minor_Cd, bm.minor_nm "," g_option go,b_minor bm","go.minor_Cd = bm.minor_cd and  go.major_cd = " & FilterVar("g1012", "''", "S") & " and  bm.major_cd = go.major_cd" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if intRetCd = False then
		Call CommonQueryRs(" bm.minor_Cd, bm.minor_nm ","b_minor bm"," bm.major_cd = " & FilterVar("g1012", "''", "S") & " and  bm.minor_cd = " & FilterVar("1", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		frm1.txtCurrencyCode.value= Trim(Replace(lgF0,Chr(11),""))
		frm1.txtCurrency.value= Trim(Replace(lgF1,Chr(11),""))
	else
		frm1.txtCurrencyCode.value= Trim(Replace(lgF0,Chr(11),""))
		frm1.txtCurrency.value= Trim(Replace(lgF1,Chr(11),""))
	End IF
End Function

'======================================================================================================
'	Name : OpenCode()
'	Description : Major PopUp
'=======================================================================================================%>
Function OpenCode()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    arrParam(0) = "계정그룹"		    <%' 팝업 명칭 %>
	arrParam(1) = "a_acct_gp "	                <%' TABLE 명칭 %>
	arrParam(2) = frm1.txtCode.value        <%' Code Condition%>
	arrParam(3) = ""   		                <%' Name Cindition%>
	arrParam(4) = "gp_cd in (select distinct gp_cd from a_acct where temp_fg_3 in (" & FilterVar("G1", "''", "S") & "," & FilterVar("G2", "''", "S") & "," & FilterVar("G3", "''", "S") & "," & FilterVar("G4", "''", "S") & "," & FilterVar("G5", "''", "S") & "," & FilterVar("G6", "''", "S") & "," & FilterVar("G7", "''", "S") & ")) "         <%' Where Condition%>
	arrParam(5) = "계정그룹"

    arrField(0) = "gp_cd"	     			<%' Field명(1)%>
    arrField(1) = "gp_nm"					<%' Field명(0)%>


    arrHeader(0) = "계정그룹"			<%' Header명(0)%>
    arrHeader(1) = "계정그룹명"				<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCode.focus
		Exit Function
	Else
		Call SetCode(arrRet)
	End If

End Function
'======================================================================================================
'	Name : SetCode()
'	Description : 계정코드 Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetCode(Byval arrRet)
	With frm1
		.txtCode.focus
		.txtCode.value = arrRet(0)
		.txtCodeh.value = arrRet(1)
	End With
End Function
'======================================================================================================
'	Name : OpenCostCon()
'	Description : Major PopUp
'=======================================================================================================%>
Function OpenCodeCon()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	arrParam(0) = "영업그룹"		<%' 팝업 명칭 %>
	arrParam(1) = "b_sales_grp"	    <%' TABLE 명칭 %>
	arrParam(2) = frm1.txtCost.value    <%' Code Condition%>
	arrParam(3) = ""   		            <%' Name Cindition%>
	arrParam(4) = ""                    <%' Where Condition%>
	arrParam(5) = "영업그룹"

    arrField(0) = "sales_grp"	     	<%' Field명(1)%>
    arrField(1) = "sales_grp_nm"		<%' Field명(0)%>


    arrHeader(0) = "영업그룹"	 <%' Header명(0)%>
    arrHeader(1) = "영업그룹명"	 <%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCost.focus
		Exit Function
	Else
		Call SetCode1(arrRet)
	End If

End Function
'======================================================================================================
'	Name : SetCode1()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetCode1(Byval arrRet)
	With frm1
		.txtCost.focus
		.txtCost.value = arrRet(0)
		.txtCosth.value = arrRet(1)
	End With
End Function
'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    .vspdData.ReDraw = False
      select case CStr(frm1.txtCurrencyCode.value)
           case "1"
                  ggoSpread.SSSetRequired	C_acct_gp, pvStartRow, pvEndRow
                  ggoSpread.SSSetProtected	C_acct_gm, pvStartRow, pvEndRow
                  ggoSpread.SSSetProtected	C_acct_nm, pvStartRow, pvEndRow
           case "2"
		          ggoSpread.SSSetRequired	C_alloc_from, pvStartRow, pvEndRow
                  ggoSpread.SSSetProtected	C_cost_nm, pvStartRow, pvEndRow
           case "3"
                  ggoSpread.SSSetRequired	C_alloc_from, pvStartRow, pvEndRow
                  ggoSpread.SSSetProtected	C_cost_nm, pvStartRow, pvEndRow
                  ggoSpread.SSSetRequired	C_acct_gp, pvStartRow, pvEndRow
                  ggoSpread.SSSetProtected	C_acct_nm, pvStartRow, pvEndRow
                  ggoSpread.SSSetProtected	C_acct_gm, pvStartRow, pvEndRow
		end select
	        ggoSpread.SpreadLock     C_acct_cd, pvStartRow, pvEndRow
            ggoSpread.SpreadLock     C_code_pb, pvStartRow, pvEndRow
            ggoSpread.SSSetRequired	C_alloc_base, pvStartRow, pvEndRow
            ggoSpread.SSSetRequired	C_alloc_rate, pvStartRow, pvEndRow
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
              Frm1.vspdData.Action = 0
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

			C_alloc_base_CODE			= iCurColumnPos(1)
			C_alloc_base_old 			= iCurColumnPos(2)
			C_alloc_from     			= iCurColumnPos(3)    
			C_ACCT_PB        			= iCurColumnPos(4)
			C_cost_nm        			= iCurColumnPos(5)
			C_acct_gp        			= iCurColumnPos(6)
			c_code_gb       			= iCurColumnPos(7)
			C_acct_gm         			= iCurColumnPos(8)
			C_acct_cd 					= iCurColumnPos(9)
			c_code_pb					= iCurColumnPos(10)
			C_acct_nm    				= iCurColumnPos(11)    
			C_alloc_base       			= iCurColumnPos(12)
			C_curr_PB        			= iCurColumnPos(13)
			C_alloc_rate        		= iCurColumnPos(14)
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
    Call InitComboBox
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
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>
                                                                            <%'Format Numeric Contents Field%>
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call initMinor()
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>

    Call SetDefaultVal
    Call InitComboBox
    If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")									<%'버튼 툴바 제어 %>
    Else
       Call SetToolbar("1100111100111111")									<%'버튼 툴바 제어 %>
    End if
'    Call txtCurrencyCode_OnChange()
    Call CookiePage(0)

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

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										  '⊙: Clear Contents  Field
  '  Call SetDefaultVal
    Call InitVariables															  '⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									          '⊙: This function check indispensable field
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call MakeKeyStream("X")

   '------ Developer Coding part (End )   --------------------------------------------------------------
    If DbQuery = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncQuery = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "1")                                        '☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                        '☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")									<%'버튼 툴바 제어 %>
    Else
       Call SetToolbar("1100111100111111")									<%'버튼 툴바 제어 %>
    End if
    Call SetDefaultVal
    Call InitVariables
   '------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncNew = True															      '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                 '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		                 '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbDelete = False Then                                                     '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim currency_code

    FncSave = False                                                              '☜: Processing is NG
    Err.Clear

    currency_code = CStr(frm1.txtCurrencyCode.value)                                                           '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------


	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbSave = False Then                                                       '☜: Query db data
       Call LayerShowHide(0)
       Exit Function
    End If
    Set gActiveElement = document.ActiveElement
    FncSave = True                                                               '☜: Processing is OK
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

	With Frm1

	If .vspdData.ActiveRow > 0 Then
    	.vspdData.ReDraw = False

		ggoSpread.Source = frm1.vspdData
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	'------ Developer Coding part (Start ) --------------------------------------------------------------
			.vspdData.Col   =		C_alloc_from
			.vspdData.value = ""
			.vspdData.Col   =		C_cost_nm
			.vspdData.value = ""
			.vspdData.Col   =		C_acct_gp
			.vspdData.value = ""
			.vspdData.Col   =		C_acct_gm
			.vspdData.value = ""
			.vspdData.Col   =		C_acct_cd
			.vspdData.value = ""
			.vspdData.Col   =		C_acct_nm
			.vspdData.value = ""
	'------ Developer Coding part (End )   --------------------------------------------------------------

		.vspdData.ReDraw = True
		.vspdData.focus
	End If
	End With
    Set gActiveElement = document.ActiveElement
    FncCopy = True                                                                '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel()
    FncCancel = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCancel = True                                                             '☜: Processing is OK
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

    FncDeleteRow = False														 '☜: Processing is NG
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

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area

  '  Call SetDefaultVal
    Call InitVariables													         '⊙: Initializes local global variables

    If LayerShowHide(1) = false then
	    Exit Function
	End if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction


	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
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

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area

    Call SetDefaultVal
    Call InitVariables														     '⊙: Initializes local global variables

    if LayerShowHide(1) = false then
	    Exit Function
	end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction


	'------ Developer Coding part (Start )   --------------------------------------------------------------
	'------ Developer Coding part (End   )   --------------------------------------------------------------

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
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
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = false then
	    Exit Function
	End if                                                     '☜: Show Processing Message

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
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
    DIm IntRetCD
    Dim strYYYYMM
    Dim strYear,strMonth,strDay

    Err.Clear                                                                    '☜: Clear err status

    Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM = strYear & strMonth


    DbSave = False                                                               '☜: Processing is NG
    if LayerShowHide(1) = false then
	    Exit Function
	end if                                                     '☜: Show Processing Message

	'------ Developer Coding part (Start)  --------------------------------------------------------------
    With frm1
        .txtMode.value        = Parent.UID_M0002                                        '☜: Delete
        .txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
    End With

    '------ Developer Coding part (add)  --------------------------------------------------------------

    '------ Developer Coding part (add)  --------------------------------------------------------------
    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

               Case ggoSpread.InsertFlag                                      '☜: Update
                                                       strVal = strVal & "C" & Parent.gColSep
                                                       strVal = strVal & lRow & Parent.gColSep
                                                       strval = strval & strYYYYMM& Parent.gColSep
                                                       strval = strval & Trim("4") & Parent.gColSep
                     .vspdData.Col = C_alloc_from    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_acct_cd       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_acct_gp    	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '6
                     .vspdData.Col = C_alloc_base_code : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                     .vspdData.Col = C_alloc_rate  : strVal = strVal & UNICDbl(Trim(.vspdData.value)) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                        strVal = strVal & "U" & Parent.gColSep
                                                        strVal = strVal & lRow & Parent.gColSep
                                                        strval = strval & strYYYYMM& Parent.gColSep
                                                        strval = strval & Trim("4") & Parent.gColSep
                    .vspdData.Col = C_alloc_from      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_acct_cd  	      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_acct_gp       		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '6
                    .vspdData.Col = C_alloc_base_code : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_alloc_rate  	  : strVal = strVal & Trim(.vspdData.value) & Parent.gColSep
                    .vspdData.Col = C_alloc_base_old  : strVal = strVal & Trim(.vspdData.value) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                        strDel = strDel & "D" & Parent.gColSep
                                                        strDel = strDel & lRow & Parent.gColSep
                                                        strDel = strDel & strYYYYMM& Parent.gColSep
                                                        strDel = strDel & Trim("4") & Parent.gColSep
                    .vspdData.Col = C_alloc_from      : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_acct_cd  	      : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_acct_gp         : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '6
                    .vspdData.Col = C_alloc_base_code : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_alloc_rate  	  : strDel = strDel & Trim(.vspdData.value) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal

	End With

	'------ Developer Coding part (End )   --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True


    Set gActiveElement = document.ActiveElement
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()

    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
    'Call LayerShowHide(1)                                                        '☜: Show Processing Message

	                            '☜: Delete
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------

    DbDelete = True                                                              '☜: Processing is OK
	'Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()


	lgIntFlgMode      = Parent.OPMD_UMODE                                                   '⊙: Indicates that current mode is Create mode
	'------ Developer Coding part (Start)  --------------------------------------------------------------
    If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")									<%'버튼 툴바 제어 %>
    Else
       Call SetToolbar("1100111100111111")									<%'버튼 툴바 제어 %>
    End if

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
	Call InitVariables
	'------ Developer Coding part (Start)  --------------------------------------------------------------
    Frm1.fpdtWk_yymm.text =  Frm1.fpdtWk_yymm.text
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")									<%'버튼 툴바 제어 %>
    Else
       Call SetToolbar("1100111100111111")									<%'버튼 툴바 제어 %>
    End if
	'------ Developer Coding part (End )   --------------------------------------------------------------
    FncQuery()
    Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	Call InitVariables()

	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")									<%'버튼 툴바 제어 %>
    Else
       Call SetToolbar("1100111100111111")									<%'버튼 툴바 제어 %>
    End if
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call FncNew()
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
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD,EFlag
    Dim grp_cd
    Dim alloc_base
    Dim alloc_from
    Dim acct_cd
    Dim acct_gp  
    Dim currency_code

    EFlag = False

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	currency_code = CStr(frm1.txtCurrencyCode.value)
    
	'------ Developer Coding part (Start ) --------------------------------------------------------------
    '=============================cost center 값 체크 시작 ==================================================
   Select Case Col
     Case C_alloc_from
    
    alloc_from = Frm1.vspdData.Text
		If currency_code = "2" or currency_code = "3" then
			If alloc_from <>"" Then
				IntRetCD = CommonQueryRs("sales_grp_nm","b_sales_grp","sales_grp = " & FilterVar(alloc_from, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD = False Then
				    Call DisplayMsgBox("GB2301","X","X","X")
				    Frm1.vspdData.Col = C_alloc_from
				    frm1.vspdData.Text = ""
				    Frm1.vspdData.Col = C_cost_nm
				    frm1.vspdData.Text = ""
				    Frm1.vspdData.Col = Col
				    Frm1.vspdData.Action = 0
				    Set gActiveElement = document.activeElement  
				    EFlag = True
				    Else
					    Frm1.vspdData.Col = C_cost_nm
					    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
				 End If
			End If
		End If
    '=============================cost center 값 체크 끝 ==================================================

   '=============================계정그룹 값 체크 시작 ==================================================
   Case C_acct_gp

    ggoSpread.Source = frm1.vspdData
	grp_cd = Frm1.vspdData.Text
	If currency_code = "1" or currency_code = "3" then
			IntRetCD = CommonQueryRs(" gp_nm "," a_acct_gp ","gp_cd in (select distinct gp_cd from a_acct where temp_fg_3 LIKE " & FilterVar("G%", "''", "S") & ") and gp_cd =  " & FilterVar(grp_cd , "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If IntRetCD = False Then
				Call DisplayMsgBox("110200","X","X","X")
				Frm1.vspdData.Col = C_acct_gp
				Frm1.vspdData.Text = ""
				Frm1.vspdData.Col = C_acct_gm
				Frm1.vspdData.Text = ""
			    Frm1.vspdData.Col = C_acct_cd
				Frm1.vspdData.Text = ""
				Frm1.vspdData.Col = C_acct_nm
				Frm1.vspdData.Text = "" 	
				'=============================계정그룹 값 체크 끝 ==================================================
				'계정 그룹코드의 값이 선택되어야 계정코드의 값이 활성화 된다.
				If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
					frm1.vspdData.ReDraw = false
					ggoSpread.SpreadLock    C_acct_cd, Row, C_acct_cd,Row
					ggoSpread.SpreadLock    C_acct_pb, Row, C_acct_pb,Row
					frm1.vspdData.ReDraw = True
				End IF
					Frm1.vspdData.Col = Col
					Frm1.vspdData.Action = 0
				Set gActiveElement = document.activeElement  
				EFlag = True
			Else
				Frm1.vspdData.Col = C_acct_gm
				Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
				If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
					frm1.vspdData.ReDraw = false
					ggoSpread.SpreadUnLock    C_acct_cd, Row, C_acct_cd,Row
					ggoSpread.SpreadUnLock    C_acct_pb, Row, C_acct_pb,Row
					frm1.vspdData.ReDraw = True
				End If
			End If
	End If
   '<========== end
   '=============================배부기준 값 체크 시작 ==================================================
   Case C_alloc_base
    alloc_base = Frm1.vspdData.Text

    If alloc_base <>"" Then
	    IntRetCD = CommonQueryRs("BM.MINOR_NM,BM.MINOR_CD","b_minor BM,B_CONFIGURATION BC","BM.major_cd = " & FilterVar("G1004", "''", "S") & " AND   BM.MINOR_CD = BC.MINOR_CD AND   BC.SEQ_NO = 5 AND  REFERENCE  = " & FilterVar("Y", "''", "S") & "  and BM.MINOR_NM = " & FilterVar(alloc_base, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    If IntRetCD = False Then
		    Call DisplayMsgBox("GB0403","X","X","X")
		    Frm1.vspdData.Col = C_alloc_base
		    Frm1.vspdData.Action = 0
		    EFlag = True
	    Else
		    Frm1.vspdData.Col = C_alloc_base
		    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
		    Frm1.vspdData.Col = C_alloc_base_code
		    Frm1.vspdData.Text = Trim(Replace(lgF1,Chr(11),""))
		End If
    End If
    
    Case C_acct_cd         
'=============================계정코드 값 체크 시작 ==================================================			
			ggoSpread.Source = frm1.vspdData
			Frm1.vspdData.Col = C_acct_cd
			acct_cd = Frm1.vspdData.Text
			Frm1.vspdData.Col = C_acct_gp
			acct_gp = Frm1.vspdData.Text

			If currency_code = "1" or currency_code = "3" then
				IntRetCD = CommonQueryRs(" a.acct_nm, a.acct_cd"," a_acct a, g_acct b "," a.acct_cd = b.acct_cd and a.gp_cd = " & FilterVar(acct_gp, "''", "S") & " and a.acct_cd = " & FilterVar(acct_cd, "''", "S") & " and (a.temp_fg_3 in (" & FilterVar("G2", "''", "S") & "," & FilterVar("G3", "''", "S") & "," & FilterVar("G4", "''", "S") & "," & FilterVar("G5", "''", "S") & "," & FilterVar("G6", "''", "S") & "," & FilterVar("G7", "''", "S") & ") or (a.temp_fg_3 = " & FilterVar("G1", "''", "S") & " and b.acct_type = " & FilterVar("T", "''", "S") & " )) and a.DEL_FG <> " & FilterVar("Y", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
			    If IntRetCD = False and acct_cd <> "" Then
				    Call DisplayMsgBox("110100","X","X","X")
				    Frm1.vspdData.Col = C_acct_cd
				    Frm1.vspdData.Text = ""
				    Frm1.vspdData.Col = C_acct_nm
				    Frm1.vspdData.Text = ""
'				    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
'						frm1.vspdData.ReDraw = false
'						ggoSpread.SpreadLock    C_acct_cd, Row, C_acct_cd,Row
'						ggoSpread.SpreadLock    C_code_pb, Row, C_code_pb,Row
'						frm1.vspdData.ReDraw = True
'					End If
				    Frm1.vspdData.Col = Col
				    Frm1.vspdData.Action = 0
				    Set gActiveElement = document.activeElement
'				    EFlag = True
			    Else
				    Frm1.vspdData.Col = C_acct_nm
				    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
				End If
			End If
    
    End Select
    
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

    If Row = 0 Then
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
'   Event Name : vspdData_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
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
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
       If lgStrPrevKeyIndex <> "" Then
          lgCurrentSpd = "M"
          Call MakeKeyStream("X")
          Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
       End If
    End if

End Sub

'========================================================================================================
'   Event Name : cboYesNo_OnChange
'   Event Desc :
'========================================================================================================
Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
End Sub


'======================================================================================================
' Name : ExeReflect
' Desc : 일괄 생성 
'=======================================================================================================
Function ExeReflect()
	Dim IntRetCD
    Dim lGrpCnt
    Dim strVal
    Dim strDel
    Dim lRow
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    Dim currency_code

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	ExeReflect = False


    '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM =   strYear & strMonth
    currency_code = Trim(Frm1.txtCurrencyCode.Value)

    If currency_code = "3" Then
        Call CommonQueryRs("count(*)","g_alloc_Base","yyyymm = " & FilterVar(strYYYYMM, "''", "S") & " and alloc_kinds = " & FilterVar("4", "''", "S") & " and from_alloc <>" & FilterVar("*", "''", "S") & "  and acct_gp <>" & FilterVar("*", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
	            IntRetCD = DisplayMsgBox("GA0009",Parent.VB_YES_NO,"X","X")
            End if
    Elseif currency_code = "2" Then
        Call CommonQueryRs("count(*)","g_alloc_Base","yyyymm = " & FilterVar(strYYYYMM, "''", "S") & " and alloc_kinds = " & FilterVar("4", "''", "S") & " and from_alloc <>" & FilterVar("*", "''", "S") & "  and acct_gp = " & FilterVar("*", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
	            IntRetCD = DisplayMsgBox("GA0009",Parent.VB_YES_NO,"X","X")
            End if
    Else
        Call CommonQueryRs("count(*)","g_alloc_Base","yyyymm = " & FilterVar(strYYYYMM, "''", "S") & " and alloc_kinds = " & FilterVar("4", "''", "S") & " and acct_gp <>" & FilterVar("*", "''", "S") & "  and from_alloc = " & FilterVar("*", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
	            IntRetCD = DisplayMsgBox("GA0009",Parent.VB_YES_NO,"X","X")
            End if
    End If

	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call MakeKeyStream("X")


	If LayerShowHide(1) = false then
	    Exit Function
	End if

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0006                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
                             '☜: 비지니스 ASP 를 가동 
    Call LayerShowHide(0)
	ExeReflect = True                                                           '⊙: Processing is NG
End Function
'======================================================================================================
' Name : fpdtWk_yymm_DblClick
' Desc :
'=======================================================================================================
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
        Call FncQuery
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
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목그룹별 배부기준</font></td>
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
									<TD CLASS=TD5 NOWRAP WIDTH=14%>작업년월</TD>
									<TD CLASS=TD6 NOWRAP WIDTH=86%><script language =javascript src='./js/gh002ma1_fpDateTime3_fpdtWk_yymm.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>배부유형</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=TEXT NAME="txtCurrencyCode" SIZE=5 MAXLENGTH=5 tag="14XXXU"  ALT="배부유형코드" >
									<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=27 MAXLENGTH=30 tag="14XXXU"  ALT="배부유형">
									</TD>
									</TR>
							    	<TR>
                                    <TD CLASS=TD5 id = "TitleCC" NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP>
									    <INPUT TYPE=TEXT NAME="txtCost" SIZE=10 MAXLENGTH=30 tag=11  ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCode1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCodeCon()">
     								    <INPUT TYPE=TEXT NAME="txtCosth" SIZE=20 MAXLENGTH=10 tag=14XXXU  ALT="Cost Center명">
									</TD>
									<TD CLASS=TD5 id = "TitleACCT" NOWRAP>계정그룹</TD>
									<TD CLASS=TD6 NOWRAP>
									    <INPUT TYPE=TEXT NAME="txtCode" SIZE=10 MAXLENGTH=30 tag=11  ALT="계정그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCode" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode()">
     								    <INPUT TYPE=TEXT NAME="txtCodeh" SIZE=20 MAXLENGTH=10 tag=14XXXU  ALT="계정그룹명">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/gh002ma1_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>일괄생성</BUTTON>
 		</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>> <IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO  noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


