<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Accounting
*  2. Function Name        : 외화계정관리항목별집계
*  3. Program ID           : a5154oa1_ko441
*  4. Program Name         : 
*  5. Program Desc         : 계정잔액증감명세서(외화)
*  6. Comproxy List        :
*  7. Modified date(First) : 2008/06/24
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : ajc
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

<!-- #Include file="../../inc/IncServer.asp" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/JpQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "a5235mb1_ko441.asp"                                      '비지니스 로직 ASP명

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const C_ACCTCD		= 1
Const C_ACCTNM		= 2

Const C_CRTLVAL1	= 3
Const C_CRTLNM1		= 4
Const C_CRTLVAL2	= 5

Const C_CUR		= 6
Const C_PRE		= 7
Const C_PRE_KRW		= 8
Const C_INC		= 9
Const C_INC_KRW		= 10
Const C_DEC		= 11
Const C_DEC_KRW		= 12
Const C_JAN		= 13
Const C_JAN_KRW		= 14


Const C_SHEETMAXROWS    = 100	                                      '한 화면에 보여지는 최대갯수*1.5%>
Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
<% EndDate= Year(Date) & "-" & Right("0" & Month(Date),2) & "-" & Right("0" & Day(Date),2) %>


Dim lsConcd
Dim IsOpenPop          
Dim lsCol, lsCon, lsTbl, lsMAJ, lsCal

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
	lgIntFlgMode      = OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
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

	Dim strYear, strMonth, strDay

	Call ExtractDateFrom("<%=EndDate%>" , Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)


	Call ggoOper.FormatDate(frm1.txtFromGLDt, Parent.gDateFormat,2)
	Call ggoOper.FormatDate(frm1.txtToGLDt, Parent.gDateFormat,2)
	frm1.txtFromGLDt.year = strYear 
	frm1.txtFromGLDt.month = strMonth -1
'	frm1.txtFromGLDt.day = "01" 
	frm1.txtToGLDt.text   = "<%=EndDate%>"	

   
End Sub
	

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
	
	Dim iCodeArr 
    Dim iNameArr
    Dim iDx
   
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
			.Col = C_ACCT
			If  Trim(.Value) = "총합계" THEN
			    
			    .Col = -1 
			    .Col2 = -1
			    .BackColor = RGB(255,230,255)
		   
			End If
			
			If  Trim(.Value) = "소계" THEN
			    
			    .Col = -1 
			    .Col2 = -1
			    .BackColor = RGB(230,255,255)
		   
			End If
    		next
    End With    
	

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "QA") %>
	
End Sub

'========================================================================================
' Function Name : WCookie
' Function Desc : Cookie Write, Clear
'========================================================================================
Function WCookie( ByVal Mode, Byval Row)

    If Mode = 1 then
       WriteCookie "FDT", frm1.txtFromGlDt.Text
       WriteCookie "TDT", frm1.txtToGlDt.text
     '  WriteCookie "BIZ", frm1.txtbizAreaCD.value		
       WriteCookie "ACCT", frm1.txtAcctFr.value
       WriteCookie "ACCTFR", frm1.txtAcctTo.value
       frm1.vspdData.Row = Row
       frm1.vspdData.Col = C_ACCTCD
       WriteCookie "CTRLVAL", frm1.vspdData.Text 

    Else

       WriteCookie "FDT", ""
       WriteCookie "TDT", ""
       WriteCookie "BIZ", ""
       WriteCookie "ACCTFR", ""
       WriteCookie "ACCTTO", ""
       WriteCookie "CtrlVal", ""

    end if

End Function

'========================================================================================
' Function Name : CookiePage
' Function Desc : Jump시 해당 조건값 Query
'========================================================================================
Function CookiePage(ByVal Kubun, ByVal Row)

	On Error Resume Next

	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	
	Dim strTemp, arrVal
		
	If Kubun = 1 Then  'jump할경우
            
	    Call WCookie(1, Row)
		
	ElseIf Kubun = 0 Then 'jump 한 경우
            strTemp = ReadCookie("FDT")
            if strTemp = "" then Exit Function
             
            frm1.txtFromGLDT.Text = strTemp
            frm1.txtToGLDT.Text = ReadCookie("TDT")
          '  frm1.txtBizAreaCD.value = ReadCookie("BIZ")
            frm1.txtAcctFr.value = ReadCookie("ACCTFR")
            frm1.txtAcctTo.value = ReadCookie("ACCTTO")

            If Err.number <> 0 Then
		   Err.Clear
		   Call WCookie(0, 0)
		   Exit Function 
            End If 

            Call FncQuery()
	    Call WCookie(0, 0)		
	
        End If

End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

	Dim IntRetCd
	Dim strBiz, strBiz1, strAcctFr, strAcctTo, VarFiscDt, strCurrencyCd, strCtrlCd, strCtrlVal
   
	strBiz = Trim(frm1.txtBizAreaCd.value )
	If strBiz = "" Then
	   strBiz = "0"
	End if
	
	strBiz1 = Trim(frm1.txtBizAreaCd1.value )
	If strBiz1 = "" Then
	   strBiz1 = "zzzzzzz"
	End if

	strCurrencyCd = Trim(frm1.txtCurrencyCd.value )
	if strCurrencyCd = "" Then
		strCurrencyCd = "%"
	End if	

	strCtrlCd = Trim(frm1.txtCtrlCd.value )
	if strCtrlCd = "" Then
		strCtrlCd = "%"
	End if	

	strCtrlVal = Trim(frm1.txtCtrlval.value )
	if strCtrlVal = "" Then
		strCtrlVal = "%"
	End if	


	lgKeyStream   = Replace(frm1.txtFromGLDt.text,"-","") & gColSep     
	lgKeyStream   = lgKeyStream & Replace(frm1.txtToGLDt.text,"-","")   & gColSep    
	lgKeyStream   = lgKeyStream & Trim(frm1.txtAcctFr.value )  & gColSep
	lgKeyStream   = lgKeyStream & Trim(frm1.txtAcctTo.value )  & gColSep    
	lgKeyStream   = lgKeyStream & strBiz & gColSep    	
	lgKeyStream   = lgKeyStream & strBiz1 & gColSep
        lgKeyStream   = lgKeyStream & strCurrencyCd & gColSep       

	lgKeyStream   = lgKeyStream & strCtrlCd   & gColSep
	lgKeyStream   = lgKeyStream & strCtrlVal  & gColSep

End Sub        

'========================================================================================================
'   Event Name : GetFiscDate()
'   Event Desc : 
'========================================================================================================
Function GetFiscDate( ByVal strFromDate)

	Dim strFiscYYYY, strFiscMM, strFiscDD
	Dim strFromYYYY, strFromMM, strFromDD

	GetFiscDate				="19000101"	

	Call ExtractDateFrom(Parent.gFiscStart,	Parent.gServerDateFormat,	Parent.gServerDateType,	strFiscYYYY,	strFiscMM,	strFiscDD)
	Call ExtractDateFrom(strFromDate,	Parent.gDateFormat,		Parent.gComDateType,		strFromYYYY,	strFromMM,	strFromDD)

	strFiscYYYY =  strFromYYYY

	If isnumeric(strFromYYYY) And isnumeric(strFromMM) And isnumeric(strFiscMM) Then
		If Cint(strFiscMM) > Cint(strFromMM)  then
		   GetFiscDate	= Cstr(Cint(strFromYYYY) - 1) & strFiscMM & "00" 'strFiscDD
		Else
		   GetFiscDate	= strFromYYYY & strFiscMM & "00" 'strFiscDD
		End If
	End If

End Function

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	With frm1.vspdData	
	
		.MaxCols = C_JAN_KRW + 1				   								    <%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		.ColHidden = True
		                
		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData

		.ReDraw = false
				    
		ggoSpread.Spreadinit

		Call AppendNumberPlace("6","18","2")

		ggospread.SSSetEdit   C_ACCTCD		, "계정코드"		, 10, 2,, , 2 
		ggoSpread.SSSetEdit   C_ACCTNM		, "계정명"		, 14

		ggoSpread.SSSetEdit   C_CRTLVAL1	, "관리항목값"		, 12
		ggoSpread.SSSetEdit   C_CRTLNM1		, "관리항목명"		, 20
		ggoSpread.SSSetEdit   C_CRTLVAL2	, "관리항목값2"		, 12

		ggoSpread.SSSetEdit   C_CUR		, "화 폐"		,  5, 2,, , 2 
		ggoSpread.SSSetFloat  C_PRE		, "이월금액(외화)"	, 12, "6",		ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_PRE_KRW     	, "이월금액(원화)"	, 12, ggAmtOfMoneyNo,	ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec		
		ggoSpread.SSSetFloat  C_INC		, "차변금액(외화)"	, 12, "6",		ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_INC_KRW		, "차변금액(원화)"	, 12, ggAmtOfMoneyNo,	ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_DEC		, "대변금액(외화)"	, 12, "6",		ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_DEC_KRW		, "대변금액(원화)"	, 12, ggAmtOfMoneyNo,	ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_JAN		, "외화잔액"		, 12, "6",		ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_JAN_KRW		, "원화잔액"		, 12, ggAmtOfMoneyNo,	ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec

		.ReDraw = true
			
		Call SetSpreadLock 
    
    End With

End Sub

Sub InitSpreadSheet2()
	Dim StrName1,StrName2,StrName3

	StrName1 = frm1.txtCtrlNm.Value


	With frm1.vspdData
	
		.MaxCols = C_JAN_KRW + 1
		.Col = .MaxCols
		.ColHidden = True
		                
		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData

		.ReDraw = false
				    
		ggoSpread.Spreadinit
		Call AppendNumberPlace("6","18","2")

		ggospread.SSSetEdit   C_ACCTCD		, "계정코드"		, 10, 2,, , 2 
		ggoSpread.SSSetEdit   C_ACCTNM		, "계정명"		, 14

		ggoSpread.SSSetEdit   C_CRTLVAL1	, StrName1		, 12
		ggoSpread.SSSetEdit   C_CRTLNM1		, "관리항목명"		, 20
		ggoSpread.SSSetEdit   C_CRTLVAL2	, "관리항목값2"		, 12

		ggoSpread.SSSetEdit   C_CUR		, "화 폐"		,  5, 2,, , 2 
		ggoSpread.SSSetFloat  C_PRE		, "이월금액(외화)"	, 12, "6",		ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_PRE_KRW     	, "이월금액(원화)"	, 12, ggAmtOfMoneyNo,	ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec		
		ggoSpread.SSSetFloat  C_INC		, "차변금액(외화)"	, 12, "6",		ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_INC_KRW		, "차변금액(원화)"	, 12, ggAmtOfMoneyNo,	ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_DEC		, "대변금액(외화)"	, 12, "6",		ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_DEC_KRW		, "대변금액(원화)"	, 12, ggAmtOfMoneyNo,	ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_JAN		, "외화잔액"		, 12, "6",		ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
		ggoSpread.SSSetFloat  C_JAN_KRW		, "원화잔액"		, 12, ggAmtOfMoneyNo,	ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec

		.ReDraw = true
			
		Call SetSpreadLock 
    
    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock    C_ACCTCD, -1, C_JAN_KRW
        
        .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(lRow)
    With frm1
    
       .vspdData.ReDraw = False
         
       .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart, gDateFormat, gComNum1000, gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart, gDateFormat, gComNum1000, gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
	            
	Call InitSpreadSheet                                                            'Setup the Spread sheet
	Call InitVariables   
	Call InitComboBox                                                           'Initializes local global variables    
	    
	Call FuncGetAuth(gStrRequestMenuID , gUsrId, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
	            
	Call SetToolbar("1100000000001111")										        '버튼 툴바 제어
	Call SetDefaultVal

'	Call CookiePage(0,0)
	frm1.txtAcctFr.focus
	
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If   

 '----------------------------------------------------------------------------------------------   
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                        '⊙: Initializes local global variables    
    Call MakeKeyStream("X")

    Call InitSpreadSheet2()

   '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim strReturn_value, strSQL
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If    
       
    
    If DbSave = False Then
       Exit Function
    End If				                                                    '☜: Save db data                      '☜: Processing is OK
    
    FncSave = True                                            
    
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
	
	     .ReDraw = True 
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
    
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
	
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow
       .vspdData.ReDraw = True       
    End With
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()

    Call parent.FncPrint()
    
End Function

'========================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function FncPrev() 
    On Error Resume Next                                                  '☜: Protect system from crashing
End Function

'========================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 

    Call parent.FncExport(C_MULTI)                                         '☜: 화면 유형
    
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 

    Call parent.FncFind(C_MULTI, False)                                    '☜:화면 유형, Tab 유무
    
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
	Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
	
	iColumnLimit = 5
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
        Frm1.vspdData.Col = iColumnLimit : Frm1.vspdData.Row = 0  :	iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 
    
    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False

	 Call LayerShowHide(1)

	 Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex="  & lgStrPrevKeyIndex                 '☜: Next key tag
        strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)            '☜: Max fetched data at a time
    End With

	 Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                          strVal = strVal & "C"  & gColSep
                                                          strVal = strVal & lRow & gColSep                                                          
                    
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U"  & gColSep
                                                          strVal = strVal & lRow & gColSep												          
                    lGrpCnt = lGrpCnt + 1
                    
                    
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D"  & gColSep
                                                  strDel = strDel & lRow & gColSep                                              
                    lGrpCnt = lGrpCnt + 1
           End Select
           
       Next
	
       .txtMode.value        = UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If    
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100000000011111")										        '버튼 툴바 제어

	
End Function


'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables
	
	Call FncQuery()
	
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================

Function OpenPopUp(iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim StrS


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    
    SELECT CASE iWhere
         CASE 0     
	         arrParam(0) = "사업장 팝업"				    ' 팝업 명칭
	         arrParam(1) = "B_BIZ_AREA"					    ' TABLE 명칭
	         arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	         arrParam(3) = ""							    ' Name Cindition
	         arrParam(4) = ""							    ' Where Condition
	         arrParam(5) = "사업장 코드"			
	
             arrField(0) = "BIZ_AREA_CD"					' Field명(0)
             arrField(1) = "BIZ_AREA_FULL_NM"				' Field명(1)
    
             arrHeader(0) = "사업장코드"				    ' Header명(0)
	         arrHeader(1) = "사업장명"				        ' Header명(1)
	         
         CASE 1     
	         arrParam(0) = "사업장 팝업"				    ' 팝업 명칭
	         arrParam(1) = "B_BIZ_AREA"					    ' TABLE 명칭
	         arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	         arrParam(3) = ""							    ' Name Cindition
	         arrParam(4) = ""							    ' Where Condition
	         arrParam(5) = "사업장 코드"			
	
             arrField(0) = "BIZ_AREA_CD"					' Field명(0)
             arrField(1) = "BIZ_AREA_FULL_NM"				' Field명(1)
    
             arrHeader(0) = "사업장코드"				    ' Header명(0)
	         arrHeader(1) = "사업장명"				        ' Header명(1)		         	
	              
	     CASE 2
	         arrParam(0) = "계정코드 팝업"				
	         arrParam(1) = "A_ACCT"					
	         arrParam(2) = Trim(frm1.txtAcctFr.Value)
	         arrParam(3) = ""							
'	         arrParam(4) = " /* acct_cd in (select acct_cd from a_acct where fx_eval_fg = 'y') */ "							
	         arrParam(4) = " "							
	         arrParam(5) = "계정 코드"			
	
             arrField(0) = "ACCT_CD"				
             arrField(1) = "ACCT_NM"				
    
             arrHeader(0) = "계정코드"				
	         arrHeader(1) = "계정코드명"					     
	     
	     CASE 3
         
	         arrParam(0) = "계정코드 팝업"				
	         arrParam(1) = "A_ACCT"					
	         arrParam(2) = Trim(frm1.txtAcctTo.Value)
	         arrParam(3) = ""							
'	         arrParam(4) = " /* acct_cd in (select acct_cd from a_acct where fx_eval_fg = 'y') */ "							
	         arrParam(4) = " "							
	         arrParam(5) = "계정 코드"			
	
             arrField(0) = "ACCT_CD"				
             arrField(1) = "ACCT_NM"				
    
             arrHeader(0) = "계정코드"				
	         arrHeader(1) = "계정코드명"	


	     CASE 4
         
	         arrParam(0) = "화폐코드 팝업"				
	         arrParam(1) = "B_CURRENCY"					
	         arrParam(2) = Trim(frm1.txtCurrencyCd.Value)
	         arrParam(3) = ""							
	         arrParam(4) = " "							
	         arrParam(5) = "계정 코드"			
	
             	 arrField(0) = "CURRENCY"				
             	 arrField(1) = "CURRENCY_DESC"				
    
             	 arrHeader(0) = "화폐코드"				
	         arrHeader(1) = "화폐코드명"	

	     CASE 5
		If frm1.txtAcctFr.value = "" Then 	         
			msgbox "계정코드를 먼저 선택하십시오."
			IsOpenPop = false
			Exit Function
	        End If
	        
	        If frm1.txtAcctTo.value <> "" Then 	         
			StrS = frm1.txtAcctTo.value
		else
			StrS = "zzz"
	        End If


	        arrParam(0) = "관리항목 팝업"				
	        arrParam(1) = "(select distinct A_ACCT_CTRL_ASSN.CTRL_CD,CTRL_NM from A_ACCT_CTRL_ASSN, A_CTRL_ITEM  where A_ACCT_CTRL_ASSN.Ctrl_CD = A_CTRL_ITEM.CTRL_CD AND ACCT_CD >= '" & frm1.txtAcctFr.value & "' AND ACCT_CD <= '" & frm1.txtAcctTo.value & "' ) a "
	        arrParam(2) = Trim(frm1.txtCtrlCd.Value)
	        arrParam(3) = ""							
	        arrParam(4) = ""							
	        arrParam(5) = "계정 코드"			

            	arrField(0) = "a.CTRL_CD"				
            	arrField(1) = "a.CTRL_NM"				

            	arrHeader(0) = "관리항목"
	        arrHeader(1) = "관리항목명"

	    Case 6
		arrParam(0) = Trim(frm1.txtCtrlNm.value)							' 팝업 명칭 
		arrParam(1) = Trim(frm1.hTblId.value) 
		arrParam(2) = ""												' Code Condition
		arrParam(3) = ""												' Name Cindition
			
		arrParam(4) = ""

		arrParam(5) = Trim(frm1.txtCtrlNm.value)									' 조건필드의 라벨 명칭 

		arrField(0) = Trim(frm1.hDataColmID.value)			' Field명(0)
		arrField(1) = Trim(frm1.hDataColmNm.value)						' Field명(1)

		arrHeader(0) = Trim(frm1.hDataColmID.value)					' Header명(0)
		arrHeader(1) = Trim(frm1.hDataColmNm.value)						' Header명(1)
	     
	END SELECT         
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If	

End Function

'------------------------------------------  SetReturnVal()  ---------------------------------------------
'	Name : SetReturnVal()
'	Description : Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetReturnVal(Byval arrRet, Byval iWhere)
		
	With frm1
		
		SELECT CASE iWhere 
		
		    CASE 0
				.txtBizAreaCd.value = arrRet(0)
				.txtBizAreaNm.value = arrRet(1)
				if Trim(.txtBizAreaCd1.value) = "" then
				    .txtBizAreaCd1.value = arrRet(0)
				    .txtBizAreaNm1.value = arrRet(1)
				
				End if

		    CASE 1
				.txtBizAreaCd1.value = arrRet(0)
				.txtBizAreaNm1.value = arrRet(1)				  	
			  
			CASE 2
				.txtAcctFr.value = arrRet(0)
				.txtAcctFrNm.value = arrRet(1)	
				if Trim(.txtAcctTo.value) = "" then
				    .txtAcctTo.value = arrRet(0)
				    .txtAcctToNm.value = arrRet(1)
				
				End if
	  			
			CASE 3
				.txtAcctTo.value = arrRet(0)
				.txtAcctToNm.value = arrRet(1)	

			CASE 4
				.txtCurrencyCd.value = arrRet(0)
				.txtCurrencyNm.value = arrRet(1)

			Case 5
			  	.txtCtrlCd.focus
			 	.txtCtrlCd.value = trim(arrRet(0))
			 	.txtCtrlNm.value = trim(arrRet(1))

			 	 CtrlVal.innerHTML = frm1.txtCtrlNm.value 
			 	.txtCtrlVal.value	= ""
			 	.txtCtrlValNm.value	= ""
				Call ElementVisible(frm1.txtCtrlVal, 1)
				Call ElementVisible(frm1.txtCtrlValNm, 1)
				Call ElementVisible(frm1.btnCtrlVal, 1)
				
				call QueryCtrlVal3()	
				
			Case 6
			  .txtCtrlVal.focus
			  .txtCtrlVal.value = arrRet(0)
			  .txtCtrlValNm.value = arrRet(1)

		End SELECT
		
	End With
	
End Function


Function QueryCtrlVal()

    Dim ArrRet

    IF frm1.txtCtrlCd.value = "" Then
		Call DisplayMsgBox("205152", "X", "보조부항목","X")
		frm1.txtCtrlCd.focus
	END IF

    Call CommonQueryRs( "TBL_ID,DATA_COLM_ID,DATA_COLM_NM,COLM_DATA_TYPE" , _ 
				"A_CTRL_ITEM" , _
				 "CTRL_CD = " & FilterVar(frm1.txtCtrlCd.value, "''", "S"), _ 
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)


	ArrRet 	= Split(lgF0,Chr(11))
	
	IF Trim(ArrRet(0)) <> "" then

		frm1.hTblId.value  = ArrRet(0)
		
		ArrRet 	= Split(lgF1,Chr(11))
		frm1.hDataColmID.value  = ArrRet(0)
		ArrRet 	= Split(lgF2,Chr(11))
		frm1.hDataColmNm.value = ArrRet(0)

		Call OpenPopUp(6)
	ELSE

		if replace(lgF3,Chr(11),"") = "D" then
			 frm1.txtCtrlValNm.value = "YYYY-MM-DD"
		Elseif replace(lgF3,Chr(11),"") = "N" then
			 frm1.txtCtrlValNm.value = "숫자는 구분자없이"
		End if	 
				
		
	END IF

End Function

Function QueryCtrlVal3()

    Dim ArrRet

    IF frm1.txtCtrlCd.value = "" Then
		Call DisplayMsgBox("205152", "X", "보조부항목","X")
		frm1.txtCtrlCd.focus
	END IF

    Call CommonQueryRs( "TBL_ID,DATA_COLM_ID,DATA_COLM_NM,COLM_DATA_TYPE" , _ 
				"A_CTRL_ITEM" , _
				 "CTRL_CD = " & FilterVar(frm1.txtCtrlCd.value, "''", "S"), _ 
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)


	ArrRet 	= Split(lgF0,Chr(11))
	
	IF Trim(ArrRet(0)) <> "" then

		frm1.hTblId.value  = ArrRet(0)
		
		ArrRet 	= Split(lgF1,Chr(11))
		frm1.hDataColmID.value  = ArrRet(0)
		ArrRet 	= Split(lgF2,Chr(11))
		frm1.hDataColmNm.value = ArrRet(0)

	ELSE

		if replace(lgF3,Chr(11),"") = "D" then
			 frm1.txtCtrlValNm.value = "YYYY-MM-DD"
		Elseif replace(lgF3,Chr(11),"") = "N" then
			 frm1.txtCtrlValNm.value = "숫자는 구분자없이"
		End if	 
				
		
	END IF

End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	ggoSpread.Source = frm1.vspdData
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
    
    If Row > 0 Then
		Select Case Col
	'		Case C_EmpPopup
	'			Call OpenEmp(1)	
		End Select    
	End If
            
End Sub
'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    Dim iDx
    Dim IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    '------- Developer Coding part (End   ) -------------------------------------------------------------- 
             
   	If Frm1.vspdData.CellType = SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
'   Event Name : txtBizAreaCd_Onchange
'   Event Desc :
'========================================================================================================
Function txtBizAreaCd_Onchange()

    if frm1.txtBizAreaCd.value = "" then
	   frm1.txtBizAreaNm.value = ""
       Exit Function
    end if
    
    Call CommonQueryRs("distinct BIZ_AREA_FULL_NM ", " B_BIZ_AREA ", " BIZ_AREA_CD = '" & frm1.txtBizAreaCd.value & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    
    if (lgF0 <> "X") AND (Trim(lgF0) <> "") then 
       frm1.txtBizAreaNm.value = Left(lgF0, Len(lgF0)-1)    
    else       
       msgbox "사업장 정보가 없습니다. 다시 선택하십시오"
       frm1.txtBizAreaNm.value = ""
       frm1.txtBizAreaCd.focus       
    end if   
    
    txtBizAreaCd_OnChange = True    
 

            
End Function

'========================================================================================================
'   Event Name : txtBizAreaCd_Onchange
'   Event Desc :
'========================================================================================================
Function txtBizAreaCd1_Onchange()

    If frm1.txtBizAreaCd1.value = "" then
	   frm1.txtBizAreaNm1.value = ""
       Exit Function
    End if
    
    Call CommonQueryRs("distinct BIZ_AREA_FULL_NM ", " B_BIZ_AREA ", " BIZ_AREA_CD = '" & frm1.txtBizAreaCd1.value & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    
    If (lgF0 <> "X") AND (Trim(lgF0) <> "") Then 
       frm1.txtBizAreaNm1.value = Left(lgF0, Len(lgF0)-1)    
    Else       
       msgbox "사업장 정보가 없습니다. 다시 선택하십시오"
       frm1.txtBizAreaNm1.value = ""
       frm1.txtBizAreaCd1.focus       
    End if   
    
    txtBizAreaCd1_OnChange = True    
   
End Function

'========================================================================================================
'   Event Name : txtAcctFr_Onchange
'   Event Desc :
'========================================================================================================
Function txtAcctFr_Onchange()
    
    if frm1.txtAcctFr.value = "" Then
       frm1.txtAcctFrNm.value = ""
       Exit Function
    end if
    
    Call CommonQueryRs("ACCT_CD, ACCT_NM ", " A_ACCT", " ACCT_CD = '" & frm1.txtAcctFr.value & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    
    if (lgF0 <> "X") AND (Trim(lgF0) <> "") then 
       frm1.txtAcctFrNm.value = Left(lgF1, Len(lgF1)-1)    
       frm1.txtAcctFr.focus 
    else       
       frm1.txtAcctFrNm.value = ""
       frm1.txtAcctFr.focus       
    end if   
    
    txtAcctFr_OnChange = True
            
End Function


'========================================================================================================
'   Event Name : txtCtrlCd_Onchange
'   Event Desc :
'========================================================================================================
Function txtCtrlCd_Onchange()
	With frm1
		If .txtAcctFr.value = "" Then
		   msgbox "계정코드를 먼저 선택하십시오."
		   .txtCtrlCd.value = ""
		   .txtAcctFr.focus
		   Exit Function
		End if

		Call CommonQueryRs("A_ACCT_CTRL_ASSN.CTRL_CD, CTRL_NM ", " A_ACCT_CTRL_ASSN, A_CTRL_ITEM", "a_acct_ctrl_assn.ctrl_cd = a_ctrl_item.ctrl_cd AND ACCT_CD = '" & frm1.txtAcctFr.value & "' and a_acct_ctrl_assn.ctrl_cd = '" & frm1.txtCtrlCd.value & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    

		If (lgF0 <> "X") And (trim(lgF0) <> "") Then
			.txtCtrlNm.value = Left(lgF1, Len(lgF1)-1)
			 CtrlVal.innerHTML = frm1.txtCtrlNm.value 
			  .txtCtrlVal.value	= ""
			  .txtCtrlValNm.value	= ""
				Call ElementVisible(frm1.txtCtrlVal, 1)
				Call ElementVisible(frm1.txtCtrlValNm, 1)
				Call ElementVisible(frm1.btnCtrlVal, 1)
		.txtCtrlCd.focus	
		Else
		Call DisplayMsgBox("800054", "X", "X", "X")
		.txtCtrlCd.value = ""
		.txtCtrlNm.value = ""
				CtrlVal.innerHTML = "" 
				.txtCtrlVal.value	= ""
				.txtCtrlValNm.value	= ""
				Call ElementVisible(frm1.txtCtrlVal, 0)
				Call ElementVisible(frm1.txtCtrlValNm, 0)
				Call ElementVisible(frm1.btnCtrlVal, 0)
		.txtCtrlCd.focus
		End If
	End With

    txtCtrlCd_OnChange = True
End Function

FUNCTION txtCtrlVAL_OnChange()

    Dim ArrRet

    IF frm1.txtCtrlCd.value = "" Then
		Call DisplayMsgBox("205152", "X", "보조부항목","X")
		frm1.txtCtrlCd.focus
	END IF

    Call CommonQueryRs( "TBL_ID,DATA_COLM_ID,DATA_COLM_NM,COLM_DATA_TYPE" , _ 
				"A_CTRL_ITEM" , _
				 "CTRL_CD = " & FilterVar(frm1.txtCtrlCd.value, "''", "S"), _ 
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)


	ArrRet 	= Split(lgF0,Chr(11))
	
	IF Trim(ArrRet(0)) <> "" then

		frm1.hTblId.value  = ArrRet(0)
		
		ArrRet 	= Split(lgF1,Chr(11))
		frm1.hDataColmID.value  = ArrRet(0)
		ArrRet 	= Split(lgF2,Chr(11))
		frm1.hDataColmNm.value = ArrRet(0)
		
		Call CommonQueryRs( " " & frm1.hDataColmID.value & " , " & frm1.hDataColmNm.value & " " , _ 
				"  " & frm1.hTblId.value & "  " , _
				 " " & frm1.hDataColmID.value & "  = " & FilterVar(frm1.txtCtrlVAL.value, "''", "S"), _ 
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
	    IF REPLACE(lgF0,Chr(11),"") <> "" THEN
	       frm1.txtCtrlVALNM.value =   REPLACE(lgF1,Chr(11),"")
	    ELSE
	       frm1.txtCtrlVAL.value =   ""
	       frm1.txtCtrlVALNM.value =   ""
		END IF
		
	ELSE

		if replace(lgF3,Chr(11),"") = "D" then
			 frm1.txtCtrlValNm.value = "YYYY-MM-DD"
		Elseif replace(lgF3,Chr(11),"") = "N" then
			 frm1.txtCtrlValNm.value = "숫자는 구분자없이"
		Else
		     frm1.txtCtrlValNm.value = ""
		End if	 
				
		
	END IF

End Function


'========================================================================================================
'   Event Name : txtAcctFr_Onchange
'   Event Desc :
'========================================================================================================
Function txtAcctTo_Onchange()
  
    if frm1.txtAcctTo.value = "" Then
       frm1.txtAcctToNm.value = ""
       Exit Function
    end if
    Call CommonQueryRs("ACCT_CD, ACCT_NM ", " A_ACCT", " ACCT_CD = '" & frm1.txtAcctTo.value & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    
    if (lgF0 <> "X") AND (trim(lgF0) <> "") then 
       frm1.txtAcctToNm.value = Left(lgF1, Len(lgF1)-1)           
    else
       frm1.txtAcctTo.focus       
    end if   
    
    txtAcctTo_OnChange = True
            
End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
     gMouseClickStatus = "SPC" 

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 컬럼을 더블클릭할 경우 발생
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
    
'	Dim strUrl    
	    
'	if Row < 0 then Exit Function
'	Call JumpPgm()

  '  Call CookiePage(1, Row)
  '  Call PgmJump(BIZ_PGM_JUMP_ID)
     
End Function


Function JumpPgm()
	
	Dim pvSelmvid, pvFB_fg,pvKeyVal,StrNVar,StrNPgm,pvSingle
	Dim strBp
	
	if lgIntFlgMode     <> Parent.OPMD_UMODE then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if
	
	strBp = ""

	ggoSpread.Source = frm1.vspdData
	frm1.vspddata.row = frm1.vspddata.ActiveRow
    frm1.vspddata.col = C_ACCTCD


		pvKeyVal  = frm1.vspdData.value
		pvSingle  =	frm1.vspdData.value & chr(11) & _
					frm1.txtbizAreaCD.value & chr(11) & _
					strBp & chr(11) & _
					frm1.txtFromGlDt.text & chr(11) & _ 
					frm1.txtToGlDt.text & chr(11) 

		pvFB_fg   = "F"
		pvSelmvid = "ACCT_CD"

		Call Jump_Pgm (	pvSelmvid, _
						pvFB_fg, _
						pvSingle,  _
						pvKeyVal)				
	
End Function
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
	
'    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS_D Then	           
 '   	If lgStrPrevKeyIndex <> "" Then    
  '  	   'Call MakeKeyStream("X")  
   '   	  ' DbQuery
    '	End If
   ' End if
End Sub


'=======================================================================================================
'   Event Name : 
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromGlDt.Action = 7
    End If
End Sub

Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidDt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행
'=======================================================================================================
Sub txtFromGLDt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtToGLDt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtAcctFr_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtAcctTo_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtCtrlCd_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub


Function QueryCtrlVal()

    Dim ArrRet

    IF frm1.txtCtrlCd.value = "" Then
		Call DisplayMsgBox("205152", "X", "보조부항목","X")
		frm1.txtCtrlCd.focus
	END IF

    Call CommonQueryRs( "TBL_ID,DATA_COLM_ID,DATA_COLM_NM,COLM_DATA_TYPE" , _ 
				"A_CTRL_ITEM" , _
				 "CTRL_CD = " & FilterVar(frm1.txtCtrlCd.value, "''", "S"), _ 
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)


	ArrRet 	= Split(lgF0,Chr(11))
	
	IF Trim(ArrRet(0)) <> "" then

		frm1.hTblId.value  = ArrRet(0)
		
		ArrRet 	= Split(lgF1,Chr(11))
		frm1.hDataColmID.value  = ArrRet(0)
		ArrRet 	= Split(lgF2,Chr(11))
		frm1.hDataColmNm.value = ArrRet(0)

		Call OpenPopUp(6)
	ELSE

		if replace(lgF3,Chr(11),"") = "D" then
			 frm1.txtCtrlValNm.value = "YYYY-MM-DD"
		Elseif replace(lgF3,Chr(11),"") = "N" then
			 frm1.txtCtrlValNm.value = "숫자는 구분자없이"
		End if	 
				
		
	END IF

End Function



'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrUrl, StrEbrFile)
	Dim strBiz, strBiz1, strAcctFr, strAcctTo, VarFiscDt
	Dim STRCtrlCd, STRCtrlVal
   
	strBiz = Trim(frm1.txtBizAreaCd.value )
	If strBiz = "" Then
	   strBiz = "0"
	End if
	
	strBiz1 = Trim(frm1.txtBizAreaCd1.value )
	If strBiz1 = "" Then
	   strBiz1 = "zzzzzzz"
	End if

	STRCtrlCd     = frm1.txtCtrlCd.VALUE

	
	strAcctFr = Trim(frm1.txtAcctFr.value )
	if strAcctFr = "" Then
	   strAcctFr = "4"
	End if
	
	strAcctTo = Trim(frm1.txtAcctTo.value )
	If strAcctTo = "" Then
	   strAcctTo = "9"
	End if

	StrEbrFile = "a5233oa1_ko441"

	StrUrl = StrUrl & "DateFr|"     & Replace(frm1.txtFromGLDt.text,"-","")
	StrUrl = StrUrl & "|DateTo|"    & Replace(frm1.txtToGLDt.text,"-","") 
	StrUrl = StrUrl & "|AcctCdFr|"  & strAcctFr
	StrUrl = StrUrl & "|AcctCdTo|"  & strAcctTo
	StrUrl = StrUrl & "|BizAreaFr|" & strBiz
	StrUrl = StrUrl & "|BizAreaTo|" & strBiz1

End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function BtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile	
    Dim ObjName
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
	Call SetPrintCond(StrUrl, StrEbrFile)

    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)
		
End Function



'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview() 
	'On Error Resume Next                                                    '☜: Protect system from crashing
    
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile    
    Dim ObjName
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    

	Call SetPrintCond(StrUrl, StrEbrFile)

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)	
			
End Function


Function PgmJump1()
   Dim Row

    Row = frm1.vspdData.ActiveRow 
    if Row < 0 then Exit Function
    
    Call BtnDisabled(1)
    Call CookiePage(1, Row)
    Call PgmJump(BIZ_PGM_JUMP_ID)

	Call BtnDisabled(0)
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
	<!-- space Area-->

	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../image/table/seltab_up_bg.gif" NOWRAP><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정잔액증감명세서(외화)</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
                                <TD CLASS="TD5" NOWRAP>회계일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5108ma1_fpDateTime1_txtFromGlDt.js'></script>&nbsp;~&nbsp;
											           <script language =javascript src='./js/a5108ma1_fpDateTime2_txtToGlDt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBizAreaCd" ALT="시작사업장코드" Size="10" MAXLENGTH="10" STYLE="TEXT-ALIGN: LEFT" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenPopUp(0)">
													   <INPUT NAME="txtBizAreaNm" ALT="사업장명" Size="25" MAXLENGTH="40" STYLE="TEXT-ALIGN: LEFT" tag="14NXXU">&nbsp;~&nbsp;</TD>
						    </TR>	
						    <TR>
								<TD CLASS="TD5" NOWRAP>계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctFr" SIZE=10 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="시작계정코드"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcctFr" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(2)">
													   <INPUT TYPE=TEXT NAME="txtAcctFrNm" SIZE=25 tag="14">&nbsp;~</TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBizAreaCd1" ALT="종료사업장코드" Size="10" MAXLENGTH="10" STYLE="TEXT-ALIGN: LEFT" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenPopUp(1)">
													   <INPUT NAME="txtBizAreaNm1" ALT="사업장명" Size="25" MAXLENGTH="40" STYLE="TEXT-ALIGN: LEFT" tag="14NXXU"></TD>
						    </TR>	
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctTo" SIZE=10 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="종료계정코드"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcctTo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(3)">
													   <INPUT TYPE=TEXT NAME="txtAcctToNm" SIZE=25 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>화폐</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCurrencyCd" ALT="화폐코드" Size="10" MAXLENGTH="3" STYLE="TEXT-ALIGN: LEFT" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenPopUp(4)">
									 	       <INPUT NAME="txtCurrencyNm" ALT="화폐명" Size="25" MAXLENGTH="40" STYLE="TEXT-ALIGN: LEFT" tag="14NXXU">
								</TD>						
						    <TR>
								<TD CLASS="TD5" ID="CtrlCd" NOWRAP>관리항목코드1</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCtrlCd" SIZE=10 MAXLENGTH=20 tag="12XXXU" ALT="관리항목코드1" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCtrlCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(5)"> <INPUT TYPE="Text" NAME="txtCtrlNm" SIZE=25 tag="14X" ALT="보조부항목명"></TD>
								<TD CLASS="TD5" ID="CtrlVal" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCtrlVal" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCtrlVal" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call QueryCtrlVal()"> <INPUT TYPE="Text" NAME="txtCtrlValNm" SIZE=25 tag="14X" ALT=""></TD>
							</TR>
						    </TR>					
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE CLASS="BasicTB" CELLSPACING=0 >
							<TR>
								<TD HEIGHT=100% WIDTH=100% >
								<script language =javascript src='./js/a5124ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			<!--	<TR>
				   <TD height=20>
				   <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>                    	
	                	<TR>
	                	<TD width="16%" align="center" bgcolor="#d1e8f9" NOWRAP>&nbsp;합    계&nbsp;</TD>
	                	<TD width="21%" align="center" bgcolor="#eeeeec" NOWRAP>	                        					
							<script language =javascript src='./js/a5124ma1_OBJECT22_txtSSumAmt.js'></script></TD>								
	                	<TD width="21%" align="center" bgcolor="#eeeeec" NOWRAP>
							<script language =javascript src='./js/a5124ma1_OBJECT22_txtTDrAmt.js'></script></TD>								
	                	<TD width="21%" align="center" bgcolor="#eeeeec" NOWRAP>
							<script language =javascript src='./js/a5124ma1_OBJECT22_txtTCrAmt.js'></script></TD>								
                        <TD width="21%" align="center" bgcolor="#eeeeec" NOWRAP>
							<script language =javascript src='./js/a5124ma1_OBJECT22_txtTSumAmt.js'></script></TD>								
	                	</TR>	                	
	               </TABLE>  	
				   </TD>
				</TR>-->
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
	<!--				<TD><BUTTON NAME="bttnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()"  Flag=1>미리보기</BUTTON>&nbsp;	-->
	<!--					<BUTTON NAME="bttnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()"    Flag=1>인쇄</BUTTON></TD>	-->
	<!--				<TD WIDTH="*" ALIGN=RIGHT><a href  onClick="VBSCRIPT:JumpPgm()">계정별원장조회&nbsp;&nbsp;&nbsp;</a></TD>	-->
				</TR>
			</TABLE>	
		</TD>
	</TR>			
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no
		 noresize framespacing=0></IFRAME></TD>
	<!-- <TD WIDTH=100% HEIGHT=150><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=150 FRAMEBORDER=0 SCROLLING=no
		 noresize framespacing=0></IFRAME></TD> -->
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtGlno"       TAG="22">

<INPUT TYPE=hidden NAME="hTblId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDataColmID" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDataColmNm" tag="24" TABINDEX="-1">


<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">	<%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
