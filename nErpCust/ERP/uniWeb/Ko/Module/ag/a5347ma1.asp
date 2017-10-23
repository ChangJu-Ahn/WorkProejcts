<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Account
*  2. Function Name        : 
*  3. Program ID           : a5347ma1
*  4. Program Name         : 배치번호별 확인 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2002/12/06
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit																	'☜: indicates that All variables must be declared in advance

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_ID1 = "a5347mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "a5347mb2.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID3 = "a5347mb3.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID4 = "a5347mb4.asp"												'☆: 비지니스 로직 ASP명 

Const C_MaxKey          = 3					                          '☆: SpreadSheet의 키의 갯수 
Const C_MaxKey_A          = 3					                          '☆: SpreadSheet의 키의 갯수 
Const C_MaxKey_B          = 3					                          '☆: SpreadSheet의 키의 갯수 

Dim lgIsOpenPop                                          

Dim IsOpenPop          

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim lgPageNo_B                                              '☜: Next Key tag                          
Dim lgSortKey_B          

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim lgPageNo_C
Dim lgSortKey_C          

<% 
Dim dtToday
dtToday = GetSvrDate                                                 
%>


Dim C_BizAreaCd1
Dim C_TransDT 
Dim C_TransType 
Dim C_TransTypeNm 
Dim C_EventCD 
Dim C_EventNm
Dim C_TransAmt 
Dim C_RetFg

Dim C_BizAreaCd2
Dim C_GlDT 
Dim C_AcctCd 
Dim C_AcctNm
Dim C_TempDrAmt
Dim C_TempCrAmt
Dim C_GlDrAmt
Dim C_GlCrAmt

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.1 Common Group -1
' Description : This part declares 1st common function group
'=======================================================================================================
'*******************************************************************************************************

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE							'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgStrPrevKey     = ""	
    lgPageNo_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1    										'initializes Previous Key
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay,  EndDate, StartDate
	
	Call ExtractDateFrom("<%=dtToday%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateAdd("M", -1, EndDate, parent.gDateFormat)
	
	frm1.txtFromReqDt.text	=  StartDate
	frm1.txtToReqDt.text	=  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtFromReqDt.focus	
    frm1.hOrgChangeId.value = parent.gChangeOrgId   	
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE" , "QA") %>  
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>                              
End Sub


'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("A5347MA101","S","A", "V20030201", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")	
	Call SetZAdoSpreadSheet("A5347MA102","S","B", "V20030201", parent.C_SORT_DBAGENT,frm1.vspdData1,C_MaxKey, "X", "X")	
	Call SetZAdoSpreadSheet("A5347MA103","S","C", "V20030201", parent.C_SORT_DBAGENT,frm1.vspdData2,C_MaxKey, "X", "X")	
	Call SetSpreadLock()
End Sub


'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	With frm1
		ggoSpread.Source = .vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()

		ggoSpread.Source = .vspdData1
		ggoSpread.SpreadLockWithOddEvenRowColor()

		ggoSpread.Source = .vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End With
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
	
	If lgIsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = GetKeyPos("A",1)
		arrParam(0) = ""
	    arrParam(1) = Trim(.Text)
	End With

	lgIsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
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

	If lgIsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = GetKeyPos("A",1)
		arrParam(0) = ""					        '회계전표번호 
	    arrParam(1) = Trim(.Text)			
	End With

	lgIsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
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
		lgIsOpenPop = False
		Exit Function
	End If

	If lgIsOpenPop = True Then Exit Function

	With frm1.vspdData
		.Row = .ActiveRow
		.Col = GetKeyPos("A",2)
		arrParam(0) = Trim(.Text)							        '배치번호 
	    arrParam(1) = ""											'Reference번호	
	End With

	lgIsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _	
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
End Function

 '========================================== 2.4.2 Open???()  =============================================
'	Name : OpenPopUp()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "거래유형 팝업"    ' 팝업 명칭 
			arrParam(1) = "A_ACCT_TRANS_TYPE"    ' TABLE 명칭 
			arrParam(2) = strCode      ' Code Condition
			arrParam(3) = ""       ' Name Cindition
			arrParam(4) = ""       ' Where Condition
			arrParam(5) = "거래유형"     ' 조건필드의 라벨 명칭 

			arrField(0) = "TRANS_TYPE"     ' Field명(0)
			arrField(1) = "TRANS_NM"     ' Field명(1)

			arrHeader(0) = "거래유형코드"   ' Header명(0)
			arrHeader(1) = "거래유형명"    ' Header명(1)
		Case 1
			arrParam(0) = "사업장 팝업"  				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"	 				' TABLE 명칭 
			arrParam(2) = strCode							' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "사업장"	    				' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"							' Field명(0)
			arrField(1) = "BIZ_AREA_NM"							' Field명(1)
    
			arrHeader(0) = "사업장코드"	     				' Header명(0)
			arrHeader(1) = "사업장명"					' Header명(1)
		Case 2
			arrParam(0) = "전표입력경로팝업"								' 팝업 명칭 
			arrParam(1) = "B_MINOR" 										' TABLE 명칭 
			arrParam(2) = strCode										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "MAJOR_CD = " & FilterVar("A1001", "''", "S") & " "												' Where Condition
			arrParam(5) = "전표입력경로코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "MINOR_CD"										' Field명(0)
			arrField(1) = "MINOR_NM"										' Field명(1)
	
			arrHeader(0) = "전표입력경로코드"									' Header명(0)
			arrHeader(1) = "전표입력경로명"									' Header명(1)			
			
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function
'------------------------------------------  EscPopUp()  --------------------------------------------------
'	Name : EscPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtTransType.focus
			Case 1
				.txtBizCd.focus
			Case 2
				.txtGlInputType.focus
		End Select
	End With
	
End Function
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtTransType.value  = arrRet(0)
				.txtTransTypeNm.value  = arrRet(1)			    
				.txtTransType.focus
			Case 1
				.txtBizCd.value  = arrRet(0)
				.txtBizNm.value  = arrRet(1)			    
				.txtBizCd.focus
			Case 2'입력경로 
				.txtGlInputType.value = UCase(Trim(arrRet(0)))
				.txtGlInputTypeNm.value = arrRet(1)							
				.txtGlInputType.focus

		End Select
	End With
End Function

Sub txtBizCd_onBlur()	
	if frm1.txtBizCd.value = "" then
		frm1.txtBizNm.value = ""
	end if
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
    Call LoadInfTB19029()														
	
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables()	
	Call InitSpreadSheet()
	Call SetDefaultVal()	
    Call SetToolBar("1100000000000111")
End Sub

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Sub txtFromReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtToReqDt.focus
		Call FncQuery
	end if
End Sub

Sub txtToReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt.focus
		Call FncQuery
	End if
End Sub

Sub GIDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt.focus
		Call FncQuery
	end if
End Sub
'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If                                                  
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables() 											
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
		Exit Function
    End If

	If CompareDateByFormat(frm1.txtFromReqDt.text,frm1.txtToReqDt.text,frm1.txtFromReqDt.Alt,frm1.txtToReqDt.Alt, _
        	               "970025",frm1.txtFromReqDt.UserDefinedFormat,parent.gComDateType, True) = False Then
		frm1.txtFromReqDt.focus
		Exit Function
	End If    
	

    Call DbQuery()
    FncQuery = True													
'	Set gActiveElement = document.ActiveElement 
End Function

'========================================================================================================
Function FncNew()

End Function

'========================================================================================================
Function FncDelete()

End Function

'========================================================================================================
Function FncSave()

End Function
'========================================================================================================
Function FncCopy()

End Function

'========================================================================================================
Function FncCancel() 

End Function


'========================================================================================================
Function FncInsertRow()

End Function

'========================================================================================================
Function FncDeleteRow()

End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
	Set gActiveElement = document.activeElement  
End Function


'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
Function FncSplitColumn()
    
     If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Function

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

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
    Dim strVal1
    DbQuery = False
    Call LayerShowHide(1)
    
    Err.Clear																				'☜: Protect system from crashing

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData    
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData    
    
    With frm1
		strVal = BIZ_PGM_ID1
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
		strVal = strVal & "?txtMode="      & parent.UID_M0001							'☜:조회표시 
		strVal = strVal & "&txtBizCd="     & Trim(.txtBizCd.value)	 			    '☆: 조회 조건 데이타 
		strVal = strVal & "&txtTransType="    & Trim(.txtTransType.value)
		strVal = strVal	& "&rdoDiff="		& Trim(frm1.RdoDiff.checked)
'		strVal = strVal & "&RdoDispType="    & Trim(.RdoDiff.value)
		strVal = strVal & "&txtFromReqDt=" & UNIConvDate(Trim(.txtFromReqDt.Text))
		strVal = strVal & "&txtToReqDt="   & UNIConvDate(Trim(.txtToReqDt.Text))
		strVal = strVal & "&txtMaxRows="   & frm1.vspdData.MaxRows
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtGlInputType="   & Trim(.txtGlInputType.value)
		strVal = strVal & "&txtfrBatchNo="   & Trim(.txtfrbatchNo.value)
		strVal = strVal & "&txtToBatchNo="   & Trim(.txtTobatchNo.value)
		strVal = strVal & "&txtfrRefNo="   & Trim(.txtfrRefNo.value)
		strVal = strVal & "&txtToRefNo="   & Trim(.txtToRefNo.value)

    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        Call RunMyBizASP(MyBizASP, strVal)	                                         '☜: 비지니스 ASP 를 가동 
    End With

    DbQuery = True
End Function

'========================================================================================================
Function DbQuery2() 
	Dim strVal
    Dim strVal1

    DbQuery2 = False
    Call LayerShowHide(1)
    
    Err.Clear
        
    With frm1
		strVal = BIZ_PGM_ID2
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
		strVal = strVal & "?txtMode="      & parent.UID_M0001							'☜:조회표시 
		strVal = strVal & "&txtBatchNo="     & Trim(GetKeyPosVal("A", 2)) 
		strVal = strVal & "&txtMaxRows="   & frm1.vspdData1.MaxRows
'		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey_B
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo_B="       & lgPageNo_B                          '☜: Next key tag
        strVal = strVal & "&lgSelectListDT_B=" & GetSQLSelectListDataType("B")         
        strVal = strVal & "&lgTailList_B="     & MakeSQLGroupOrderByList("B")
		strVal = strVal & "&lgSelectList_B="   & EnCoding(GetSQLSelectList("B"))

        Call RunMyBizASP(MyBizASP, strVal)	                                         '☜: 비지니스 ASP 를 가동 
    End With


    DbQuery2 = True
End Function


Function DbQuery3() 
	Dim strVal
    Dim strVal1

    DbQuery3 = False
    Call LayerShowHide(1)

    Err.Clear
        
    With frm1
		strVal = BIZ_PGM_ID3
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
		strVal = strVal & "?txtGlNo=" & Trim(GetKeyPosVal("A", 3))
		strVal = strVal & "&txtMaxRows="   & frm1.vspdData1.MaxRows
'		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey_C
    '--------- Developer Coding Part (End) ------------------------------------------------------------

        strVal = strVal & "&lgPageNo_C="       & lgPageNo_C                          '☜: Next key tag
        strVal = strVal & "&lgSelectListDT_C=" & GetSQLSelectListDataType("C")         
        strVal = strVal & "&lgTailList_C="     & MakeSQLGroupOrderByList("C")
		strVal = strVal & "&lgSelectList_C="   & EnCoding(GetSQLSelectList("C"))

        Call RunMyBizASP(MyBizASP, strVal)	                                         '☜: 비지니스 ASP 를 가동 
    End With


    DbQuery3 = True
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()		
	
	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	Call vspdData_Click(1,1)
	frm1.vspdData.focus

	Call Dbquery2()
'	Call Dbquery3()
	
'    lgSaveRow        = 1
'	CALL vspdData_Click(1, 1)
	
End Function
Function DbQueryOk2()		
	
	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    Call DbQuery3()

	
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================


'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'========================================================================================================
Function OpenOrderPopup(ByVal pSpdNo)

	Dim arrRet

	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp", Array(ggoSpread.GetXMLData(pSpdNo),gMethodText), "dialogWidth=" & parent.SORTW_WIDTH & " px; dialogHeight=" & parent.SORTW_HEIGHT & " px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo, arrRet(0), arrRet(1))
       Call InitVariables()
	   If frm1.RdoDt.checked = True then
	       Call InitSpreadSheet()
	   Else
	       Call InitSpreadSheet()	   
	   End IF    
	End If

End Function

'*******************************************************************************************************
' Function Name : PopZAdoConfigGrid
' Function Desc : 
'========================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then	
		Exit Sub
	End If		
	
	If UCase(Trim(gActiveSpdSheet.Name)) = "VSPDDATA" Then
		Call OpenOrderPopup("A")	
	Elseif UCase(Trim(gActiveSpdSheet.Name)) = "VSPDDATA1" Then
		Call OpenOrderPopup("B")
	End If
	
End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"									'Split 상태코드 
    
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col							'Ascending Sort
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey				'Descending Sort
			lgSortKey = 1
		End If										
		Exit Sub
	End If		
    If Col < 1 Then Exit Sub
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    lgPageNo_B       = ""                                  'initializes Previous Key
    lgSortKey_B      = 1

End Sub	
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SP1C"									'Split 상태코드 
    
	Set gActiveSpdSheet = frm1.vspdData1
	
	If frm1.vspdData1.Maxrows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData1
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col							'Ascending Sort
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey				'Descending Sort
			lgSortKey = 1
		End If										
		Exit Sub
	End If		

	Call SetSpreadColumnValue("A",frm1.vspdData1,Col,Row)

End Sub	


'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub    
'========================================================================================================
' Function Name : vspdData1_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
	If Button = 2 And gMouseClickStatus = "SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End Sub 

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    gMouseClickStatus = "SPC"	'Split 상태코드    

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
		    ggoSpread.Source = frm1.vspdData
			If lgSortKey_A = 1 Then
				ggoSpread.SSSort, lgSortKey_A
	            lgSortKey_A = 2
		    Else
			    ggoSpread.SSSort, lgSortKey_A
				lgSortKey_A = 1
	        End If    
		    Exit Sub
	    End If
	    
		Call SetSpreadColumnValue("A",frm1.vspdData,Col,NewRow)	        
    
     
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
		
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

	    Call DbQuery2()
	
		lgPageNo_B       = ""                                  'initializes Previous Key
		lgSortKey_B      = 1
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : This event is spread sheet data Button Clicked
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then

        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq

            .hItemSeq.value = .vspdData.Text
            .vspdData2.MaxRows = 0
        End With

        frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End if

		lgCurrRow = NewRow
'		Call DbQuery2(lgCurrRow)
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub


'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
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
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If    
	
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then	
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub
'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If    
	
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	    
    	If lgPageNo_B <> "" Then	
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery2 = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub

'=======================================================================================================
Sub txtBizCd_onChange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtBizCd.value = "" Then frm1.txtBizNm.value = "":	Exit Sub

	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(frm1.txtBizCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBizNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtBizCd.alt,"X")  	
		frm1.txtBizCd.value = ""
		frm1.txtBizNm.value = ""
		frm1.txtBizCd.focus
	End If
End Sub	

'=======================================================================================================
Sub txtGlInputType_onchange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtGlInputType.value = "" Then frm1.txtGlInputTypeNm.value = "":	Exit Sub

	If CommonQueryRs("MINOR_NM", "B_MINOR ", " MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  AND MINOR_CD=  " & FilterVar(frm1.txtGlInputType.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtGlInputTypeNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtGlInputType.alt,"X")  	
		frm1.txtGlInputType.value = ""
		frm1.txtGlInputTypeNm.value = ""
		frm1.txtGlInputType.focus
	End If
	if frm1.txtGlInputType.value = "" then
		frm1.txtGlInputTypeNm.value = ""
	end if
End Sub	

'=======================================================================================================
Sub txtTransType_onchange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtTransType.value = "" Then frm1.txtTransTypeNm.value = "":	Exit Sub

	If CommonQueryRs("TRANS_NM", "A_ACCT_TRANS_TYPE ", " TRANS_TYPE=  " & FilterVar(frm1.txtTransType.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtTransTypeNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtTransType.alt,"X")  	
		frm1.txtTransType.value = ""
		frm1.txtTransTypeNm.value = ""
		frm1.txtTransType.focus
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromReqDt.Focus 
    End If
End Sub

Sub txtToReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToReqDt.Focus 
    End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtFromEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtToEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


'=======================================================================================================
'   Event Name : RdoDt_OnClick()
'   Event Desc :  
'=======================================================================================================
Function RdoDiff_OnClick() 
	If frm1.RdoDiff.checked = True then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
	End if
End Function

'=======================================================================================================
'   Event Name : RdoSum_OnClick()
'   Event Desc :  
'=======================================================================================================
Function RdoTotal_OnClick() 
	If frm1.RdoTotal.checked = True then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
	End if
End Function


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>배치번호별전표처리확인</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<A href="vbscript:OpenPopupTempGL()">결의전표</A> </TD>					
					<TD WIDTH=10></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% colspan=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100%  colspan=2>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5"NOWRAP>거래기간</TD>
									<TD CLASS="TD6"NOWRAP><script language =javascript src='./js/a5347ma1_fpDateTime1_txtFromReqDt.js'></script>&nbsp;~&nbsp; 
														<script language =javascript src='./js/a5347ma1_fpDateTime2_txtToReqDt.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>사업장</TD>										
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizCd"   ALT="사업장"   Size="10" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="11N"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizCd.Value, 1)">
														 <INPUT NAME="txtBizNm" ALT="사업장명" Size="20" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>거래유형</TD>										
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransType"   ALT="거래유형"   Size="10" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="11N"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtTransType.value, 0)">
														 <INPUT NAME="txtTransTypeNm" ALT="거래유형명" Size="20" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
									<TD CLASS=TD5 NOWRAP>배치번호</TD>				
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtfrBatchNo" SIZE=18 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="배치번호">&nbsp;~&nbsp;
														 <INPUT TYPE="Text" NAME="txttoBatchNo" SIZE=18 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="배치번호"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>전표입력경로</TD>
									<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtGlInputType" SIZE=10  MAXLENGTH=10 tag="11XXXU" ALT="전표입력경로"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtGlInputType.Value, 2)">
										 <INPUT TYPE=TEXT ID="txtGlInputTypeNm" NAME="txtGlInputTypeNm" SIZE=20 tag="14X" ALT="전표입력경로명">
									</TD>
									<TD CLASS=TD5 NOWRAP>참조번호</TD>				
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtfrRefNo" SIZE=18 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="참조번호">&nbsp;~&nbsp;
														 <INPUT TYPE="Text" NAME="txttoRefNo" SIZE=18 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="참조번호"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS=TD5 NOWRAP>표시구분</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoType" ID="RdoDiff" VALUE="D" TAG="11" ><LABEL FOR="RdoDiff" Id="RdoDiff">차이분</LABEL>&nbsp;&nbsp
										<INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoType" ID="RdoTotal" VALUE="T" TAG="11" Checked><LABEL FOR="RdoTotal" Id="RdoTotal">전체</LABEL></TD>
								</TR>

							</TABLE>
						</FIELDSET>	
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% colspan=2></TD>
				</TR>
				
				<TR HEIGHT="100%">
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD HEIGHT="60%" COLSPAN = "6">
								<script language =javascript src='./js/a5347ma1_vaSpread1_vspdData.js'></script></TD>

							<TR HEIGHT="40%">
								<TD><script language =javascript src='./js/a5347ma1_OBJECT1_vspdData1.js'></script></TD>		
								<TD><script language =javascript src='./js/a5347ma1_OBJECT2_vspdData2.js'></script></TD>
							</TR>						
								
								
							</TR>
<!--
							<TR>
								<TD HEIGHT=40 WIDTH=25% COLSPAN = "6">
									<FIELDSET CLASS="CLSFLD">
										<TABLE  CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS="TDt" STYLE="WIDTH : 0px"></TD>
												<TD CLASS="TDt" NOWRAP>배치금액(자국)</TD>
												<TD CLASS="TDt" STYLE="WIDTH : 0px"></TD>
												<TD CLASS="TDt" NOWRAP>차변(자국)</TD>
												<TD CLASS="TDt" STYLE="WIDTH : 0px"></TD>
												<TD CLASS="TDt" NOWRAP>대변(자국)</TD>
											</TR>
											<TR>
												<TD CLASS=TDt NOWRAP COLSPAN=2><script language =javascript src='./js/a5347ma1_OBJECT1_txtTotBatchLocAmt.js'></script></TD>
												<TD CLASS=TDt NOWRAP COLSPAN=2><script language =javascript src='./js/a5347ma1_OBJECT1_txtTotDrLocAmt.js'></script></TD>
												<TD CLASS=TDt NOWRAP COLSPAN=2><script language =javascript src='./js/a5347ma1_OBJECT1_txtTotCrLocAmt.js'></script></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
-->							
						</TABLE>
					</TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT= <%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hRadio" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtTransType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtFocus" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hFromReqDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hToReqDt" tag="24"TABINDEX="-1">
<INPUT	TYPE=hidden	 NAME="hOrgChangeId"	tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

