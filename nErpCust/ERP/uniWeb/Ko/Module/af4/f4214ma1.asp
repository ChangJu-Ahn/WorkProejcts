
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4214ma1
'*  4. Program Name         : 차입처별 차입금 현황조회 
'*  5. Program Desc         : Query of Loan Balance
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002.04.25
'*  8. Modified date(Last)  : 2003.05.19
'*  9. Modifier (First)     : An, do hyun
'* 10. Modifier (Last)      : 
'* 11. Comment              : 
'======================================================================================================= -->
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

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================


Const BIZ_PGM_ID = "F4214MB1.asp"												'☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 8                                           '☆: key count of SpreadSheet

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop

Dim lgSelectList
Dim lgSelectListDT

Dim lgSortFieldNm
Dim lgSortFieldCD

Dim lgMaxFieldCount

Dim lgPopUpR
Dim lgKeyPos
Dim lgKeyPosVal
Dim lgCookValue

Dim lgSaveRow

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
    lgStrPrevKey		= ""
    lgPageNo			= ""
    lgIntFlgMode		= parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue	= False                    'Indicates that no value changed
	lgSortKey			= 1
	lgSaveRow			= 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Function SetDefaultVal()
	Dim StartDate, FristDate, LastDate
	StartDate	= "<%=GetSvrDate%>"
	FristDate	= UNIGetFirstDay("<%=GetSvrDate%>",parent.gServerDateFormat)
	LastDate	= UNIGetLastDay (StartDate,parent.gServerDateFormat)

	frm1.txtBaseFromDt.Text	= UniConvDateAToB(FristDate, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtBaseToDt.Text	= UniConvDateAToB(LastDate, parent.gServerDateFormat, parent.gDateFormat)
End Function

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
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","MA") %>

End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1012", "''", "S") & "  AND MINOR_CD IN(" & FilterVar("U", "''", "S") & " ," & FilterVar("C", "''", "S") & " ) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboApSts ,lgF0  ,lgF1  ,Chr(11))

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("F4214MA101","S","A","V20030303",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    
End Sub


'========================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
        With frm1
			.vspdData.ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			.vspdData.ReDraw = True
        End With
    End if
End Sub

'==================================++++==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal lRow)

End Sub


'===========================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function



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
'	Call initMinor()
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

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	'---------Developer Coding part (Start)----------------------------------------------------------------
	'Call parent.GetAdoFieldInf("F4214ma101","S","A")												' G for Group , A for SpreadSheet No('A','B',....   
	    
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec,,,parent.ggStrMinPart,parent.ggStrMaxPart)
	Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec,,,parent.ggStrMinPart,parent.ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")												'⊙: Lock  Suitable  Field

	'	lgMaxFieldCount =  UBound(parent.gFieldNM)

	Redim lgPopUpR(parent.C_MaxSelList -1,1)
	    
	'Call parent.MakePopData(parent.gDefaultT,parent.gFieldNM,parent.gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,parent.C_MaxSelList)		'You must not this line
	    
	Call InitVariables																	'⊙: Initializes local global variables
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call InitComboBox
	Call txtLoanPlcfg_onchange()
	Call SetToolbar("110000000001111")													'⊙: 버튼 툴바 제어	
	frm1.txtBaseFromDt.focus

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

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoOper.ClearField(Document, "3")										'⊙: Clear Contents  Field
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtBaseFromDt.Text, frm1.txtBaseToDt.Text, frm1.txtBaseFromDt.Alt, frm1.txtBaseToDt.Alt, _
						"970025", frm1.txtBaseFromDt.UserDefinedFormat, parent.gComDateType, true) = False Then
		frm1.txtBaseFromDt.focus											'⊙: GL Date Compare Common Function
		Exit Function
	End if
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
    if frm1.txtBizAreaCd.value <> "" then
	  If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  	Call DisplayMsgBox("970000","X",frm1.txtBizAreaCd.alt,"X")            '☜ : No data is found. 
	  	frm1.txtBizAreaNm.value = ""
	  	frm1.txtBizAreaCd.focus
 	  	Exit Function
	  End If
	End If
	  
	if frm1.txtBizAreaCd1.value <> "" then
	  If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd1.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  	Call DisplayMsgBox("970000","X",frm1.txtBizAreaCd1.alt,"X")            '☜ : No data is found.
	  	frm1.txtBizAreaNm1.value = ""
	  	frm1.txtBizAreaCd1.focus
 	  	Exit Function
	  End If
	End If
	
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    On Error Resume Next
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow() 
    On Error Resume Next
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    On Error Resume Next
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True  
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_MULTI, True)

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

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncExit = True
    
End Function


'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim strLoanPlcFg

	Err.Clear                                                                   '☜: Protect system from crashing
	DbQuery = False

	Call LayerShowHide(1)

	If frm1.txtLoanPlcfg1.checked Then
		strLoanPlcFg = frm1.txtLoanPlcfg1.value
	ElseIf frm1.txtLoanPlcfg2.checked Then
		strLoanPlcFg = frm1.txtLoanPlcfg2.value
	End if

	With frm1
'		strVal = BIZ_PGM_ID
'	'---------Developer Coding part (Start)----------------------------------------------------------------
'		If lgIntFlgMode <> parent.OPMD_UMODE Then										'This means that it is first search
'			strVal = strVal & "?cboConfFg="			& Trim(.cboConfFg.value)			'☆: 조회 조건 데이타 
'			strVal = strVal & "&cboApSts="			& Trim(.cboApSts.value)
'			strVal = strVal & "&cboLoanFg="			& Trim(.cboLoanFg.value)
'			strVal = strVal & "&txtDocCur="			& Trim(.txtDocCur.value)
'			strVal = strVal & "&txtLoanPlcFg="		& Trim(strLoanPlcFg)
'			strVal = strVal & "&txtLoanPlcCd="		& Trim(.txtLoanPlcCd.value)
'			StrVal = strVal & "&txtLoanType="		& Trim(.txtLoanType.value)
'			StrVal = strVal & "&txtBaseFromDt="		& Trim(.txtBaseFromDt.Text)
'			StrVal = strVal & "&txtBaseToDt="		& Trim(.txtBaseToDt.Text)
'			strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
'			strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
'		Else
'			strVal = strVal & "?cboConfFg="			& Trim(.hConfFg.value)				'☆: 조회 조건 데이타 
'			strVal = strVal & "&cboApSts="			& Trim(.hApSts.value)
'			strVal = strVal & "&cboLoanFg="			& Trim(.hLoanFg.value)
'			strVal = strVal & "&txtDocCur="			& Trim(.hDocCur.value)
'			strVal = strVal & "&txtLoanPlcFg="		& Trim(.hLoanPlcFg.value)
'			strVal = strVal & "&txtLoanPlcCd="		& Trim(.hLoanPlcCd.value)
'			StrVal = strVal & "&txtLoanType="		& Trim(.hLoanType.value)
'			StrVal = strVal & "&txtBaseFromDt="		& Trim(.hBaseFromDt.value)
'			StrVal = strVal & "&txtBaseToDt="		& Trim(.hBaseToDt.value)
'			strVal = strVal & "&txtBizAreaCd="		& Trim(.htxtBizAreaCd.value)
'			strVal = strVal & "&txtBizAreaCd1="		& Trim(.htxtBizAreaCd1.value)
'		End if
'	'---------Developer Coding part (End)----------------------------------------------------------------
'		strVal = strVal & "&lgPageNo="			& lgPageNo								'Next key tag
'		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
'		strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
'		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))
'
'		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

		' 1024 byte 초과로 post 방식으로 전환 

		If lgIntFlgMode <> parent.OPMD_UMODE Then										'This means that it is first search
			.hConfFg.value			= Trim(.cboConfFg.value)			'☆: 조회 조건 데이타 
			.hApSts.value			= Trim(.cboApSts.value)
			.hLoanFg.value			= Trim(.cboLoanFg.value)
			.hDocCur.value			= Trim(.txtDocCur.value)
			.hLoanPlcFg.value		= Trim(strLoanPlcFg)
			.hLoanPlcCd.value		= Trim(.txtLoanPlcCd.value)
			.hLoanType.value		= Trim(.txtLoanType.value)
			.hBaseFromDt.value		= Trim(.txtBaseFromDt.Text)
			.hBaseToDt.value		= Trim(.txtBaseToDt.Text)
			.htxtBizAreaCd.value	= Trim(.txtBizAreaCd.value)
			.htxtBizAreaCd1.value	= Trim(.txtBizAreaCd1.value)
		End if

		.hlgPageNo.value		= lgPageNo								'Next key tag
		.hlgSelectListDT.value	= GetSQLSelectListDataType("A")
		.hlgTailList.value		= MakeSQLGroupOrderByList("A")
		.hlgSelectList.value	= EnCoding(GetSQLSelectList("A"))

		' 권한관리 추가 
		.hlgAuthBizAreaCd.value			= lgAuthBizAreaCd
		.hlgInternalCd.value			= lgInternalCd
		.hlgSubInternalCd.value			= lgSubInternalCd
		.hlgAuthUsrID.value				= lgAuthUsrID

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	
	Call FncSetToolBar("Query")
	Call txtLoanPlcfg_onchange()
	'Call SetToolbar("1100000000011111")
	Call CurFormatNumericOCX()
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtBaseFromDt.focus
	End If

	Set gActiveElement = document.activeElement 
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	Dim intRetCD
	Dim strByCurrency
	With frm1
		If Trim(.txtDocCur.value) = "" Then
            intRetCD = CommonQueryRs("top 1 currency"," b_numeric_format "," decimals  = (select max(decimals) from b_numeric_format where data_type = 2 ) and data_type=2 and form_type = " & FilterVar("Q", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            
			If intRetCD = True Then	
				strByCurrency = Trim(Replace(lgF0,Chr(11),""))
			Else
				strByCurrency = parent.gCurrency
			End If
			ggoOper.FormatFieldByObjectOfCur .txtPreTotAmt,	strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtPlanTotAmt,strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtLoanTotAmt,strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtNextTotAmt,strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
			ggoOper.FormatFieldByObjectOfCur .txtPayTotAmt,	strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBalTotAmt,	strByCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		Else

			ggoOper.FormatFieldByObjectOfCur .txtPreTotAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtPlanTotAmt,.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtLoanTotAmt,.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtNextTotAmt,.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
			ggoOper.FormatFieldByObjectOfCur .txtPayTotAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBalTotAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		End If
	End With

End Sub

'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
End Function


'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================
 '------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If lgIsOpenPop = True Then Exit Function
	if UCase(frm1.txtLoanPlcCd.className) = "PROTECTED" Then Exit Function

	
	lgIsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanPlcCd.focus
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If

End Function
 '------------------------------------------ OpenLoanNo() -------------------------------------------------
'	Name : OpenPopUp()
'	Description : Loan Number PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	Select Case iWhere
		case 0
			If frm1.txtLoanPlcCd.className = parent.UCN_PROTECTED Then Exit Function	
			If frm1.txtLoanPlcfg1.Checked = true Then
				arrParam(0) = "은행팝업"
				arrParam(1) = "B_BANK A"
				arrParam(2) = strCode
				arrParam(3) = ""
				arrParam(4) = ""
				arrParam(5) = "은행코드"

				arrField(0) = "A.BANK_CD"
				arrField(1) = "A.BANK_NM"
						    
				arrHeader(0) = "은행코드"
				arrHeader(1) = "은행명"
			Else
				Call OpenBp(strCode, iWhere)
				exit function
			End If
        
        Case 1	
			arrParam(0) = "차입용도팝업"			' 팝업 명칭 
			arrParam(1) = "b_minor" 				    ' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "major_cd=" & FilterVar("f1000", "''", "S") & " "	        ' Where Condition
			arrParam(5) = "차입용도"				' 조건필드의 라벨 명칭 

			arrField(0) = "minor_cd"						' Field명(0)
			arrField(1) = "minor_nm"						' Field명(1)
    
			arrHeader(0) = frm1.txtLoanType.Alt				' Header명(0)
			arrHeader(1) = frm1.txtLoanTypeNm.Alt				    ' Header명(1)
		Case 2
			arrParam(0) = "거래통화팝업"								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' 조건필드의 라벨 명칭 

		    arrField(0) = "CURRENCY"										' Field명(0)
		    arrField(1) = "CURRENCY_DESC"									' Field명(1)

		    arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화코드명"									' Header명(1)
		case 3,4
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition

			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "사업장 코드"			

			arrField(0) = "BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "BIZ_AREA_NM"					' Field명(1)

			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)

		Case Else
			Exit Function
	End Select
    
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' 거래처 
				frm1.txtLoanPlcCd.focus
			Case 1		'차입용도 
				frm1.txtLoanType.focus
			Case 2
				frm1.txtDocCur.focus
			Case 3
				frm1.txtBizAreaCd.focus
			Case 4
				frm1.txtBizAreaCd1.focus
		End Select
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  SetReturnPopUp()  --------------------------------------------------
'	Name : SetReturnPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnPopUp(Byval arrRet, Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' 거래처 
				frm1.txtLoanPlcCd.value = arrRet(0)
				frm1.txtLoanPlcNm.value = arrRet(1)
				frm1.txtLoanPlcCd.focus
			Case 1		'차입용도 
				frm1.txtLoanType.value = arrRet(0)
				frm1.txtLoanTypeNm.value = arrRet(1)
				frm1.txtLoanType.focus
			Case 2
				frm1.txtDocCur.value = arrRet(0)
				frm1.txtDocCur.focus
			Case 3
				frm1.txtBizAreaCd.Value		= arrRet(0)
				frm1.txtBizAreaNm.Value		= arrRet(1)
				frm1.txtBizAreaCd.focus
			Case 4
				frm1.txtBizAreaCd1.Value	= arrRet(0)
				frm1.txtBizAreaNm1.Value	= arrRet(1)
				frm1.txtBizAreaCd1.focus
		End Select

	End With
	
End Function

'========================================================================================================
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'========================================================================================================
Function OpenGroupPopup()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(parent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = parent.gMethodText
  
	For ii = 0 to parent.C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
      
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(lgSortFieldCD,lgSortFieldNm,arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to parent.C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function



'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
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
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    If frm1.vspdData.MaxRows = 0 then
        Exit Sub
    End If
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
        Exit Sub
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
'Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
'End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : 
'==========================================================================================
'Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
'End Sub


'======================================================================================================
'   Event Name : txtLoanPlcfg_onchange
'   Event Desc : 
'=======================================================================================================
Function txtLoanPlcfg_onchange()
	If frm1.txtLoanPlcfg0.checked = true then
		Call ggoOper.SetReqAttr(frm1.txtLoanPlcCd, "Q")
		frm1.txtLoanPlcCd.value = ""
		frm1.txtLoanPlcNm.value = ""
	Else
		Call ggoOper.SetReqAttr(frm1.txtLoanPlcCd, "D")
	End If
End Function

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
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub


'========================================================================================================
'   Event Name : fpLoanfrDt_DblClick()
'   Event Desc : 
'========================================================================================================
Sub txtBaseFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBaseFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBaseFromDt.Focus
	End if
End Sub

Sub txtBaseToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBaseToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBaseToDt.Focus
	End if
End Sub

Sub txtLoanPlcCd_DblClick(Button)
	if Button = 1 then
		frm1.txtLoanPlcCd.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoanPlcCd.Focus
	End if
End Sub

Sub txtLoanType_DblClick(Button)
	if Button = 1 then
		frm1.txtLoanType.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoanType.Focus
	End if
End Sub

'========================================================================================================
'   Event Name : fpLoanFrDt_KeyPress()
'   Event Desc : 
'========================================================================================================
Sub txtBaseFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtBaseToDt.Focus
	   Call MainQuery
	End If   
End Sub

Sub txtBaseToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtBaseFromDt.Focus
	   Call MainQuery
	End If   
End Sub

Sub txtLoanPlcCd_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub

Sub txtLoanType_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
  	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<!--<TD WIDTH="*" align=right><button name="btnAutoSel" class="clsmbtn" ONCLICK="OpenGroupPopup()">정렬순서</button></td>-->
					<TD WIDTH=10>&nbsp;</TD>

				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>기준일자</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpBaseFrDt name=txtBaseFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="시작기준일자"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpBaseToDt name=txtBaseToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="종료기준일자"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>거래통화</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.Value, 2)">
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCd.Value,3)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
									<TD CLASS="TD5" NOWRAP>장단기구분</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLoanFg" ALT="장단기구분" STYLE="WIDTH: 135px" tag="11"><OPTION VALUE=""></OPTION></SELECT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCd1.Value,4)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>차입용도</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtLoanType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="차입용도코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtLoanType.Value,1)">
														   <INPUT TYPE="Text" NAME="txtLoanTypeNm" SIZE=20 tag="14X" ALT="차입용도명">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입처구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg0 VALUE="BK" Checked tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg0>은행+거래처</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg1 VALUE="BK" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg1>은행</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg2 VALUE="BP" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg2>거래처</LABEL></TD>
									<TD CLASS="TD5" NOWRAP>차입처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanPlcCd" ALT="차입처" SIZE="10" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanPlcCd.Value, 0)">
															<INPUT NAME="txtLoanPlcNm" ALT="차입처명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">승인상태</TD>
									<TD CLASS="TD6"><SELECT ID="cboConfFg" NAME="cboConfFg" ALT="승인상태" STYLE="WIDTH: 135px" tag="1XN"><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS="TD5">진행상황</TD>
									<TD CLASS="TD6"><SELECT ID="cboApSts" NAME="cboApSts" ALT="진행상황" STYLE="WIDTH: 135px" tag="1XN"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=40 WIDTH=25%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>전월잔액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtPreTotAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="전월잔액" tag="34X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtPreTotLocAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="전월잔액(자국)" tag="34X2Z"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>만기도래액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtPlanTotAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="만기도래액" tag="34X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtPlanTotLocAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="만기도래액(자국)" tag="34X2Z"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입금액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtLoanTotAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="차입금액" tag="34X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtLoanTotLocAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="차입금액(자국)" tag="34X2Z"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>익월도래액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtNextTotAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="익월도래액" tag="34X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtNextTotLocAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="익월도래액(자국)" tag="34X2Z"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>원금상환액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtPayTotAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="원금상환액" tag="34X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtPayTotLocAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="원금상환액(자국)" tag="34X2Z"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>차입잔액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtBalTotAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="차입잔액" tag="34X2Z"></OBJECT>');</SCRIPT>&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtBalTotLocAmt" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="차입잔액(자국)" tag="34X2Z"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="hConfFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hApSts" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hDocCur" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanPlcFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanPlcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanType" tag="24">
<INPUT TYPE=HIDDEN NAME="hBaseFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hBaseToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1" tag="24">

<INPUT TYPE=HIDDEN NAME="hlgPageNo"			tag="24">
<INPUT TYPE=HIDDEN NAME="hlgSelectListDT"	tag="24">
<INPUT TYPE=HIDDEN NAME="hlgTailList"		tag="24">
<INPUT TYPE=HIDDEN NAME="hlgSelectList"		tag="24">

<INPUT TYPE=HIDDEN NAME="hlgAuthBizAreaCd"	tag="24">
<INPUT TYPE=HIDDEN NAME="hlgInternalCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="hlgSubInternalCd"	tag="24">
<INPUT TYPE=HIDDEN NAME="hlgAuthUsrID"		tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

