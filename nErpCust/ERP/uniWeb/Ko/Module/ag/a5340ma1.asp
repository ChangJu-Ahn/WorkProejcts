<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Account
*  2. Function Name        : 
*  3. Program ID           : a5340ma1
*  4. Program Name         : 전표처리확인작업 
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit																	'☜: indicates that All variables must be declared in advance

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_ID1 = "a5340mb1.asp" ' 일자별 쿼리 
Const BIZ_PGM_ID2 = "a5340mb2.asp" ' 일자별 detail												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID3 = "a5340mb3.asp" ' 전체 쿼리 
Const BIZ_PGM_ID4 = "a5340mb4.asp"	'전체 detail
Const BIZ_PGM_ID5 = "a5340mb5.asp"
Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Const C_MaxKey          = 10					                          '☆: SpreadSheet의 키의 갯수 

Dim lgIsOpenPop                                          

Dim IsOpenPop          
Dim  gSelframeFlg

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim lgPageNo_B                                              '☜: Next Key tag                          
Dim lgSortKey_B          

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
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay,  EndDate, StartDate
	
	Call ExtractDateFrom("<%=dtToday%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateAdd("M", -1, EndDate, parent.gDateFormat)
	
	frm1.txtFromReqDt.text	=  StartDate
	frm1.txtToReqDt.text	=  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtFromReqDt2.text	=  StartDate
	frm1.txtToReqDt2.text	=  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	
	frm1.txtFromReqDt.focus	
    frm1.hOrgChangeId.value = parent.gChangeOrgId   	
    gSelframeFlg = TAB1
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE" , "QA") %>  
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>                              
End Sub
'=======================================================================================================
Sub InitSpreadSheet(ByVal pvDtFg)

	Select Case UCase(pvDtFg)
		Case "D"
			Call SetZAdoSpreadSheet("A5340MA101","S","A", "V20030201", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
			Call SetZAdoSpreadSheet("A5340MA101_DTL","G","B", "V20030201", parent.C_GROUP_DBAGENT,frm1.vspdData1,C_MaxKey, "X", "X")
		Case "S"
			Call SetZAdoSpreadSheet("A5340MA102","S","A", "V20030201", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
			Call SetZAdoSpreadSheet("A5340MA102_DTL","G","B", "V20030201", parent.C_GROUP_DBAGENT,frm1.vspdData1,C_MaxKey, "X", "X")
		Case "T"
			Call SetZAdoSpreadSheet("A5340MA103","S","C", "V20030201", parent.C_SORT_DBAGENT,frm1.vspdData2,C_MaxKey, "X", "X")
	End Select 
	
	Call SetSpreadLock("A")
	Call SetSpreadLock("B")
	Call SetSpreadLock("C")
End Sub

'========================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	With frm1
		Select Case UCase(pvSpdNo)
			Case "A"
				ggoSpread.Source = .vspdData
				ggoSpread.SpreadLockWithOddEvenRowColor()
			Case "B"
				ggoSpread.Source = .vspdData1
				ggoSpread.SpreadLockWithOddEvenRowColor()
			Case "C"
				ggoSpread.Source = .vspdData2
				ggoSpread.SpreadLockWithOddEvenRowColor()
		End Select
	End With
End Sub


'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "거래유형 팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT_TRANS_TYPE"								' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Condition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "거래유형"									' 조건필드의 라벨 명칭 

			arrField(0) = "TRANS_TYPE"										' Field명(0)
			arrField(1) = "TRANS_NM"										' Field명(1)

			arrHeader(0) = "거래유형코드"								' Header명(0)
			arrHeader(1) = "거래유형명"									' Header명(1)
		Case 1,3
			arrParam(0) = "사업장 팝업"  								' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Condition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "사업장"	    								' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"										' Field명(0)
			arrField(1) = "BIZ_AREA_NM"										' Field명(1)
    
			arrHeader(0) = "사업장코드"	     							' Header명(0)
			arrHeader(1) = "사업장명"									' Header명(1)
		Case 2
			arrParam(0) = "전표입력경로팝업"							' 팝업 명칭 
			arrParam(1) = "B_MINOR" 										' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Condition
			arrParam(4) = "MAJOR_CD = " & FilterVar("A1001", "''", "S") & " "								' Where Condition
			arrParam(5) = "전표입력경로코드"							' 조건필드의 라벨 명칭 

			arrField(0) = "MINOR_CD"										' Field명(0)
			arrField(1) = "MINOR_NM"										' Field명(1)
	
			arrHeader(0) = "전표입력경로코드"							' Header명(0)
			arrHeader(1) = "전표입력경로명"								' Header명(1)			
		Case 4
			arrParam(0) = "모듈팝업"									' 팝업 명칭 
			arrParam(1) = "B_MINOR" 										' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Condition
			arrParam(4) = "MAJOR_CD = " & FilterVar("B0001", "''", "S") & " "								' Where Condition
			arrParam(5) = "모듈코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "MINOR_CD"										' Field명(0)
			arrField(1) = "MINOR_NM"										' Field명(1)
	
			arrHeader(0) = "모듈코드"									' Header명(0)
			arrHeader(1) = "모듈명"										' Header명(1)			
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
			Case 3
				.txtBizCd2.focus
			Case 4
				.txtModuleCd.focus
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
			Case 3
				.txtBizCd2.value  = arrRet(0)
				.txtBizNm2.value  = arrRet(1)			    
				.txtBizCd2.focus
			Case 4
				.txtModuleCd.value  = arrRet(0)
				.txtModuleNm.value  = arrRet(1)			    
				.txtModuleCd.focus
		End Select
	End With
End Function


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
	Call InitSpreadSheet("D")
	Call InitSpreadSheet("T")
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
Sub txtFromReqDt2_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtToReqDt2.focus
		Call FncQuery
	end if
End Sub

Sub txtToReqDt2_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt2.focus
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
'    Call ggoOper.ClearField(Document, "2")	

    Call InitVariables() 											
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
		Exit Function
    End If

	If gSelframeFlg = TAB1 Then
		If CompareDateByFormat(frm1.txtFromReqDt.text,frm1.txtToReqDt.text,frm1.txtFromReqDt.Alt,frm1.txtToReqDt.Alt, _
		    	               "970025",frm1.txtFromReqDt.UserDefinedFormat,parent.gComDateType, True) = False Then
			frm1.txtFromReqDt.focus
			Exit Function
		End If    
	Else
		If CompareDateByFormat(frm1.txtFromReqDt2.text,frm1.txtToReqDt2.text,frm1.txtFromReqDt2.Alt,frm1.txtToReqDt2.Alt, _
		    	               "970025",frm1.txtFromReqDt2.UserDefinedFormat,parent.gComDateType, True) = False Then
			frm1.txtFromReqDt2.focus
			Exit Function
		End If    
	End If

	Select Case gSelframeFlg
		Case TAB1
			ggoSpread.Source = Frm1.vspdData
			Call ggoSpread.ClearSpreadData()
			ggoSpread.Source = Frm1.vspdData1
			Call ggoSpread.ClearSpreadData()
		Case TAB2
			ggoSpread.Source = Frm1.vspdData2
			Call ggoSpread.ClearSpreadData()
	End Select

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
    
    With frm1
		If gSelframeFlg = TAB1 Then
			If frm1.RdoDt.checked = True Then	strVal = BIZ_PGM_ID1
			If frm1.RdoDt.checked = False Then 	strVal = BIZ_PGM_ID3

    '--------- Developer Coding Part (Start) ----------------------------------------------------------

			strVal = strVal & "?txtMode="      & parent.UID_M0001							'☜:조회표시 
			If lgIntFlgMode <> parent.OPMD_UMODE Then
				strVal = strVal & "&txtBizCd="     & Trim(.txtBizCd.value)	 			    '☆: 조회 조건 데이타 
				strVal = strVal & "&txtTransType="    & Trim(.txtTransType.value)
				strVal = strVal & "&Radio="    & Trim(.RdoDt.value)
				strVal = strVal & "&txtFromReqDt=" & UNIConvDate(Trim(.txtFromReqDt.Text))
				strVal = strVal & "&txtToReqDt="   & UNIConvDate(Trim(.txtToReqDt.Text))
				strVal = strVal & "&txtGlInputType="   & Trim(.txtGlInputType.value)		
			Else
				strVal = strVal & "&txtBizCd="		& Trim(.hBizCd.value)	 			    '☆: 조회 조건 데이타 
				strVal = strVal & "&txtTransType=" & Trim(.hTransType.value)
				strVal = strVal & "&Radio="			& Trim(.hRdoDt.value)
				strVal = strVal & "&txtFromReqDt="	& UNIConvDate(Trim(.hFromReqDt.value))
				strVal = strVal & "&txtToReqDt="	& UNIConvDate(Trim(.hToReqDt.value))
				strVal = strVal & "&txtGlInputType="   & Trim(.hGlInputType.value)		
			End If
			strVal = strVal & "&txtMaxRows="   & frm1.vspdData.MaxRows
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		Else
			strVal = BIZ_PGM_ID5
			strVal = strVal & "?txtMode="		& parent.UID_M0001							'☜:조회표시 
			strVal = strVal & "&txtBizCd="		& Trim(.txtBizCd2.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtModuleCd="	& Trim(.txtModuleCd.value)
			strVal = strVal & "&txtFromReqDt=" & UNIConvDate(Trim(.txtFromReqDt2.Text))
			strVal = strVal & "&txtToReqDt="	& UNIConvDate(Trim(.txtToReqDt2.Text))
			strVal = strVal & "&txtMaxRows="	& frm1.vspdData2.MaxRows
'			strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("C")         
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("C")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("C"))
		End If
    '--------- Developer Coding Part (End) ------------------------------------------------------------
			
		

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
    
    Err.Clear																				'☜: Protect system from crashing
 
    With frm1
		If frm1.RdoDt.checked = True then    
			strVal = BIZ_PGM_ID2
		Else
			strVal = BIZ_PGM_ID4
		End If	

    '--------- Developer Coding Part (Start) ----------------------------------------------------------

		strVal = strVal & "?txtMode="      & parent.UID_M0001							'☜:조회표시 
		strVal = strVal & "&txtBizCd="     & Trim(GetKeyPosVal("A", 1)) 
		strVal = strVal & "&txtTransType="    & Trim(GetKeyPosVal("A", 5))
		strVal = strVal & "&Radio="    & Trim(.RdoDt.value)
		strVal = strVal & "&txtGlInputType="    & Trim(GetKeyPosVal("A", 6))
		If frm1.RdoDt.checked = True then    
			strVal = strVal & "&txtTransDt=" & UNIConvDate(Trim(GetKeyPosVal("A", 2)))
		Else
			strVal = strVal & "&txtFromReqDt=" & UNIConvDate(Trim(.hFromReqDt.value))
			strVal = strVal & "&txtToReqDt="   & UNIConvDate(Trim(.hToReqDt.value))		
		end if

		strVal = strVal & "&JnlCd="   & Trim(GetKeyPosVal("A", 3))
		strVal = strVal & "&EventCd="   & Trim(GetKeyPosVal("A", 4))
		strVal = strVal & "&txtMaxRows="   & frm1.vspdData1.MaxRows
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo_B                          '☜: Next key tag
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")         
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
        Call RunMyBizASP(MyBizASP, strVal)	                                         '☜: 비지니스 ASP 를 가동 
    End With


    DbQuery2 = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()		
	
	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    If gSelframeFlg = TAB1 Then
		Call vspdData_Click(1,1)
		frm1.vspdData.focus
		Call Dbquery2()
	Else
		frm1.vspdData2.focus
	End If		
	
End Function
Function DbQueryOk2()		
	
	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode

	
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
	       Call InitSpreadSheet("D")
	   Else
	       Call InitSpreadSheet("S")	   
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
    
	    Call DbQuery2()
     
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
	
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

Sub txtBizCd_onChange()	
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtBizCd.value = "" Then frm1.txtBizNm.value = "":	Exit Sub

'	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD= '" & TRim(frm1.txtBizCd.value) & "'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
'		arrVal = Split(lgF0, Chr(11)) 
'		frm1.txtBizNm.value= Trim(arrVal(0)) 
'	Else
'		IntRetCD = DisplayMsgBox("970000","X",frm1.txtBizCd.alt,"X")  	
'		frm1.txtBizCd.value = ""
'		frm1.txtBizNm.value = ""
'		frm1.txtBizCd.focus
'	End If
End Sub	

Sub txtBizCd2_onchange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtBizCd2.value = "" Then frm1.txtBizNm2.value = "":	Exit Sub

	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(frm1.txtBizCd2.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBizNm2.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtBizCd2.alt,"X")  	
		frm1.txtBizCd2.value = ""
		frm1.txtBizNm2.value = ""
		frm1.txtBizCd2.focus
	End If
End Sub	

Sub txtGlInputType_onchange()	
	Dim IntRetCD
	Dim arrVal

	If frm1.txtGlInputType.value = "" Then frm1.txtGlInputTypeNm.value = "":	Exit Sub

'	If CommonQueryRs("MINOR_NM", "B_MINOR ", " MAJOR_CD ='A1001' AND MINOR_CD= '" & TRim(frm1.txtGlInputType.value) & "'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
'		arrVal = Split(lgF0, Chr(11)) 
'		frm1.txtGlInputTypeNm.value= Trim(arrVal(0)) 
'	Else
'		IntRetCD = DisplayMsgBox("970000","X",frm1.txtGlInputType.alt,"X")  	
'		frm1.txtGlInputType.value = ""
'		frm1.txtGlInputTypeNm.value = ""
'		frm1.txtGlInputType.focus
'	End If
End Sub	

Sub txtModuleCd_onchange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtModuleCd.value = "" Then frm1.txtModuleNm.value = "":	Exit Sub

	If CommonQueryRs("MINOR_NM", "B_MINOR ", " MAJOR_CD = " & FilterVar("B0001", "''", "S") & "  AND MINOR_CD=  " & FilterVar(frm1.txtModuleCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtModuleNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtModuleCd.alt,"X")  	
		frm1.txtModuleCd.value = ""
		frm1.txtModuleNm.value = ""
		frm1.txtModuleCd.focus
	End If
	if frm1.txtModuleCd.value = "" then
		frm1.txtModuleNm.value = ""
	end if
End Sub	

Sub txtTransType_OnChange()	
	Dim IntRetCD
	Dim arrVal
	
	If frm1.txtTransType.value = "" Then 
		frm1.txtTransTypeNm.value = ""
	   	Exit Sub
    End If
    
'	If CommonQueryRs("TRANS_NM", "A_ACCT_TRANS_TYPE ", " TRANS_TYPE= '" & TRim(frm1.txtTransType.value) & "'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
'		arrVal = Split(lgF0, Chr(11)) 
'		frm1.txtTransTypeNm.value= Trim(arrVal(0)) 
'	Else
'		IntRetCD = DisplayMsgBox("970000","X",frm1.txtTransType.alt,"X")  	
'		frm1.txtTransType.value = ""
'		frm1.txtTransTypeNm.value = ""
'		frm1.txtTransType.focus
'	End If
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

Sub txtFromReqDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt2.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromReqDt2.Focus 
    End If
End Sub

Sub txtToReqDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqDt2.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToReqDt2.Focus 
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

Sub fpdtToEnterDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

Sub fpdtFromEnterDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

Sub fpdtToEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub
'=======================================================================================================
'   Event Name : RdoDt_OnClick()
'   Event Desc :  
'=======================================================================================================
Function RdoDt_OnClick() 
	If frm1.RdoDt.checked = True then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData

       Call InitVariables()
       Call InitSpreadSheet("D")   
	End if
End Function

'=======================================================================================================
'   Event Name : RdoSum_OnClick()
'   Event Desc :  
'=======================================================================================================
Function RdoSum_OnClick() 
	If frm1.RdoSum.checked = True then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
       Call InitVariables()
	   Call InitSpreadSheet("S")   
	End if
End Function

Function ClickTab1()
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1

	Call SetToolbar("1100000000001111") 				 
End Function

Function ClickTab2()
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2

	Call SetToolbar("1100000000001111") 
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>전표처리확인작업</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><IMG height=23 src="../../image/table/tab_up_left.gif" width=9></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>모듈별확인</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<DIV ID="TabDiv" SCROLL=no>
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
										<TD CLASS="TD6"NOWRAP><script language =javascript src='./js/a5340ma1_fpDateTime1_txtFromReqDt.js'></script>&nbsp;~&nbsp; 
															<script language =javascript src='./js/a5340ma1_fpDateTime2_txtToReqDt.js'></script></TD>
										<TD CLASS=TD5 NOWRAP>사업장</TD>										
										<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizCd"   ALT="사업장코드"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizCd.Value, 1)">
															 <INPUT NAME="txtBizNm" ALT="사업장명" Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>거래유형</TD>										
										<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransType"   ALT="거래유형"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtTransType.value, 0)">
															 <INPUT NAME="txtTransTypeNm" ALT="거래유형명" Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag="14N"></TD>
										<TD CLASS=TD5 NOWRAP>합계유형</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoSumType" ID="RdoDt" VALUE="S" TAG="11" Checked><LABEL FOR="rdoReport1">일자별</LABEL>&nbsp;&nbsp
														 <INPUT TYPE=RADIO CLASS="RADIO" NAME="RdoSumType" ID="RdoSum" VALUE="D" TAG="11"><LABEL FOR="rdoReport2">합계</LABEL></TD>
									</TR>
									<TR>
										<TD CLASS="TD5"NOWRAP>전표입력경로</TD>
										<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtGlInputType" SIZE=12  MAXLENGTH=12 tag="11XXXU" ALT="전표입력경로"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtGlInputType.Value, 2)">
											 <INPUT TYPE=TEXT ID="txtGlInputTypeNm" NAME="txtGlInputTypeNm" SIZE=24 MAXLENGTH="24" tag="14X" ALT="전표입력경로명">										 
										</TD>
										<TD CLASS=TD5 NOWRAP></TD>				
										<TD CLASS=TD6 NOWRAP></TD>
									</TR>
								</TABLE>
							</FIELDSET>	
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100% colspan=2></TD>
					</TR>
					
	<!--												
					<TR HEIGHT="100%">
						<TD><script language =javascript src='./js/a5340ma1_OBJECT1_vspdData.js'></script></TD>		
						<TD><script language =javascript src='./js/a5340ma1_OBJECT2_vspdData1.js'></script></TD>
					</TR>						
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100% colspan=2></TD>
					</TR>
	-->
					<TR HEIGHT="60%">
						<TD WIDTH="100%" colspan="6">
						<script language =javascript src='./js/a5340ma1_OBJECT1_vspdData.js'></script></TD>
					</TR>
					<TR HEIGHT="40%">
						<TD WIDTH="100%" colspan="6">
						<script language =javascript src='./js/a5340ma1_OBJECT2_vspdData1.js'></script></TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100% colspan=2></TD>
					</TR>

					
				</TABLE>
			</DIV>
			<DIV ID="TabDiv" SCROLL=no>
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
										<TD CLASS="TD6"NOWRAP><script language =javascript src='./js/a5340ma1_fpDateTime1_txtFromReqDt2.js'></script>&nbsp;~&nbsp; 
															<script language =javascript src='./js/a5340ma1_fpDateTime2_txtToReqDt2.js'></script></TD>
										<TD CLASS=TD5 NOWRAP>사업장</TD>										
										<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizCd2"   ALT="사업장코드"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizCd2.Value, 3)">
															 <INPUT NAME="txtBizNm2" ALT="사업장명" Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5"NOWRAP>모듈</TD>
										<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtModuleCd" SIZE=12  MAXLENGTH=12 tag="11XXXU" ALT="모듈코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnModuleCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtModuleCd.Value, 4)">
											 <INPUT TYPE=TEXT ID="txtModuleNm" NAME="txtModuleNm" SIZE=24 MAXLENGTH="24" tag="14X" ALT="모듈명">										 
										</TD>
										<TD CLASS=TD5 NOWRAP></TD>				
										<TD CLASS=TD6 NOWRAP></TD>
									</TR>
								</TABLE>
							</FIELDSET>	
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100% colspan=2></TD>
					</TR>
					<TR>
						<TD WIDTH=100% HEIGHT=100%>
							<script language =javascript src='./js/a5340ma1_OBJECT1_vspdData2.js'></script>
						</TD>
					</TR>
				</TABLE>
			</DIV>		

		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT= <%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBizCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hRdoDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hTransType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hFromReqDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hToReqDt" tag="24"TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hGlInputType" tag="24"TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"	tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

