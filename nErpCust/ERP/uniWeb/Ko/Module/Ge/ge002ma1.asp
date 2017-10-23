
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        :
*  3. Program ID           : GE002MA1
*  4. Program Name         : 경영손익 품목별 배부기준 등록 
*  5. Program Desc         : 경영손익 품목별 배부기준 등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/08
*  8. Modified date(Last)  : 2001/12/31
*  9. Modifier (First)     : Song Sang Min
* 10. Modifier (Last)      : Song Sang Min
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "ge002mb1.asp"						           '☆: Biz Logic ASP Name

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


Const COOKIE_SPLIT       = 4877	                                      'Cookie Split String

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgIsOpenPop
Dim IsOpenPop


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
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) --------------------------------------------------------------

   <% Call loadInfTB19029A("I", "G", "NOCOOKIE", "MA") %>  'batch= B , print = P , input = I

	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) --------------------------------------------------------------

   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

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
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
    Dim iDx
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

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
Sub InitData()
End Sub

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

'=======================================================================================================%>
Function SetCode(Byval arrRet)
	With frm1
		.txtCode.focus
		.txtCode.value = arrRet(0)
		.txtCodeh.value = arrRet(1)
	End With
End Function

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

'=======================================================================================================%>
Function SetCode1(Byval arrRet)
	With frm1
		.txtCost.focus
		.txtCost.value = arrRet(0)
		.txtCosth.value = arrRet(1)
	End With
End Function

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
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub

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
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

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

    Call ggoOper.ClearField(Document, "A")										  '☜: Clear Condition Field and Contents Field
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
Function FncSave()
    Dim IntRetCD
    Dim currency_code

    FncSave = False                                                              '☜: Processing is NG
    Err.Clear

    currency_code = CStr(frm1.txtCurrencyCode.value)                                                           '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
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
'    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction


	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
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
'    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction


	'------ Developer Coding part (Start )   --------------------------------------------------------------
	'------ Developer Coding part (End   )   --------------------------------------------------------------

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz
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
'    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
End Function

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
    Dim iColSep 
    Dim iRowSep  

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
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	  

	With Frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

               Case ggoSpread.InsertFlag                                      '☜: Update
                                                       strVal = strVal & "C" & iColSep
                                                       strVal = strVal & lRow & iColSep
                                                       strval = strval & strYYYYMM& iColSep
                                                       strval = strval & Trim("4") & iColSep
                     .vspdData.Col = C_alloc_from    : strVal = strVal & Trim(.vspdData.Text) & iColSep
                     .vspdData.Col = C_acct_cd       : strVal = strVal & Trim(.vspdData.Text) & iColSep
                     .vspdData.Col = C_acct_gp    	 : strVal = strVal & Trim(.vspdData.Text) & iColSep '6
                     .vspdData.Col = C_alloc_base_code : strVal = strVal & Trim(.vspdData.Text) & iColSep
                     .vspdData.Col = C_alloc_rate  : strVal = strVal & UNICDbl(Trim(.vspdData.value)) & iRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                        strVal = strVal & "U" & iColSep
                                                        strVal = strVal & lRow & iColSep
                                                        strval = strval & strYYYYMM& iColSep
                                                        strval = strval & Trim("4") & iColSep
                    .vspdData.Col = C_alloc_from      : strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_acct_cd  	      : strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_acct_gp       		: strVal = strVal & Trim(.vspdData.Text) & iColSep '6
                    .vspdData.Col = C_alloc_base_code : strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_alloc_rate  	  : strVal = strVal & Trim(.vspdData.value) & iColSep
                    .vspdData.Col = C_alloc_base_old  : strVal = strVal & Trim(.vspdData.value) & iRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                        strDel = strDel & "D" & iColSep
                                                        strDel = strDel & lRow & iColSep
                                                        strDel = strDel & strYYYYMM& iColSep
                                                        strDel = strDel & Trim("4") & iColSep
                    .vspdData.Col = C_alloc_from      : strDel = strDel & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_acct_cd  	      : strDel = strDel & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_acct_gp         : strDel = strDel & Trim(.vspdData.Text) & iColSep '6
                    .vspdData.Col = C_alloc_base_code : strDel = strDel & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_alloc_rate  	  : strDel = strDel & Trim(.vspdData.value) & iRowSep
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
Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
End Sub

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
'    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
                             '☜: 비지니스 ASP 를 가동 
    Call LayerShowHide(0)
	ExeReflect = True                                                           '⊙: Processing is NG
End Function

'=======================================================================================================
Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.fpdtWk_yymm.focus
	End If
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별 배부기준</font></td>
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
									<TD CLASS=TD6 NOWRAP WIDTH=86%><script language =javascript src='./js/ge002ma1_fpDateTime3_fpdtWk_yymm.js'></script></TD>
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
									<script language =javascript src='./js/ge002ma1_vaSpread_vspdData.js'></script>
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


