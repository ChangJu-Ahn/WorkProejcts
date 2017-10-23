
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : RECEIPT
'*  3. Program ID		    : f5105ma
'*  4. Program Name         : 기준일별어음조회 
'*  5. Program Desc         : 기준일별어음조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/10/16
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Hee Jung, Kim
'* 11. Comment              :
'======================================================================================================= 
-->
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

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "f5105mb1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 3					                          '☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          
Dim IsOpenPop                                          

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
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

    Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtFromIssueDt.text = frDt
	frm1.txtToIssueDt.Text = toDt
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

	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>                                '☆: 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
		
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteFg ,lgF0  ,lgF1  ,Chr(11))    

End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
       Call SetZAdoSpreadSheet("f5105ma1","S","A","V20021215",Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey,"X","X")
 	   Call SetSpreadLock("A") 
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(Byval iOpt)
	If iOpt = "A" Then
		With frm1
			.vspdData.ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
		End With
    End If
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

    Call LoadInfTB19029														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")											
    Call InitComboBox
    frm1.txtFromIssueDt.focus

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


	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub


'=======================================================================================================
'   Event Name : txtBasicDt1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromIssueDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromIssueDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromIssueDt.Focus
	End if
End Sub

Sub txtToIssueDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToIssueDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToIssueDt.Focus
	End if
End Sub


Sub txtFromIssueDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToIssueDt.focus
	   Call MainQuery
	End If   
End Sub

Sub txtToIssueDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtFromIssueDt.focus
	   Call MainQuery
	End If   
End Sub

'=======================================================================================================
'   Event Name : txtFromIssueDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================

Sub txtFromIssueDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtToIssueDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                            
    
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData						
    Call InitVariables 											
    
    Call InitSpreadSheet()
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtFromIssueDt.Text, frm1.txtToIssueDt.Text, frm1.txtFromIssueDt.Alt, frm1.txtToIssueDt.Alt, _
						"970025", frm1.txtFromIssueDt.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtFromIssueDt.focus											'⊙: GL Date Compare Common Function
			Exit Function
	End if

	If Trim(frm1.txtBizAreaCd.value) <> "" And Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If UCase(Trim(frm1.txtBizAreaCd.value)) > UCase(Trim(frm1.txtBizAreaCd1.value)) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
	If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If frm1.txtBizAreaCd1.value = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
    '-----------------------
    'Query function call area
    '-----------------------

    If DbQuery = False Then Exit Function

    FncQuery = True													
	Set gActiveElement = document.activeElement    
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
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

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
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

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
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = Frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
       Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

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

	Err.Clear                                                       
	DbQuery = False
	    
	Call LayerShowHide(1)
	With frm1
	    
		strVal = BIZ_PGM_ID

		'--------- Developer Coding Part (Start) ----------------------------------------------------------
		If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtMode	="			& Parent.UID_M0001
			strVal = strVal & "&txtFromIssueDt="	& Trim(frm1.txtFromIssueDt.text)
			strVal = strVal & "&txtToIssueDt="		& Trim(frm1.txtToIssueDt.text)
			strVal = strVal & "&cboNoteFg="			& Trim(frm1.cboNoteFg.value)
			strVal = strVal & "&txtBankCd="			& Trim(frm1.txtBankCd.value)
			strVal = strVal & "&txtBpCd="			& Trim(frm1.txtBpCd.value)
			strVal = strVal & "&txtStsCd="			& Trim(frm1.txtStsCd.value)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
			strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(frm1.txtBizAreaCd.alt)
			strVal = strVal & "&txtBizAreaCd1_Alt="	& Trim(frm1.txtBizAreaCd1.alt)
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		Else
			strVal = strVal & "?txtMode="			& Parent.UID_M0001				'☜: 비지니스 처리 ASP의 상태 
			strVal = strVal & "&txtFromIssueDt="	& Trim(.htxtFromIssueDt.value)	'☆: 조회 조건 데이타 
			strVal = strVal & "&txtToIssueDt="		& Trim(.htxtToIssueDt.value)	'☆: 조회 조건 데이타(to발행일)
			strVal = strVal & "&cboNoteFg="			& Trim(.hcboNoteFg.value)		'☆: 조회 조건 데이타(어음구분)
			strVal = strVal & "&txtBankCd="			& Trim(.htxtBankCd.value)		'☆: 조회 조건 데이타(은행코드)
			strVal = strVal & "&txtBpCd="			& Trim(.htxtBpCd.value)			'☆: 조회 조건 데이타(거래처코드)
			strVal = strVal & "&txtStsCd="			& Trim(.htxtStsCd.value)		'☆: 조회 조건 데이타(상태)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.htxtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.htxtBizAreaCd1.value)
			strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(frm1.txtBizAreaCd.alt)
			strVal = strVal & "&txtBizAreaCd1_Alt="	& Trim(frm1.txtBizAreaCd1.alt)
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		End If   
		           
		'--------- Developer Coding Part (End) ------------------------------------------------------------
		strVal = strVal & "&lgPageNo="			& lgPageNo         
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 


		Call RunMyBizASP(MyBizASP, strVal)							
	        
	End With
	    
	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1
    
    CALL vspdData_Click(1, 1)
    Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field

	Call SetToolbar("1100000000011111")	
End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'	Name : OpenGroupPopup()
'
'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
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

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	

	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0		' 은행 
				arrParam(0) = "은행 팝업"	' 팝업 명칭 
				arrParam(1) = "B_BANK"			 			' TABLE 명칭 
				arrParam(2) = strCode						' Code Condition
				arrParam(3) = ""							' Name Cindition
				arrParam(4) = ""							' Where Condition
				arrParam(5) = "은행"					' 조건필드의 라벨 명칭 

				arrField(0) = "BANK_CD"						' Field명(0)
				arrField(1) = "BANK_NM"						' Field명(1)
    
				arrHeader(0) = "은행코드"					' Header명(0)
				arrHeader(1) = "은행명"						' Header명(1)
		
		Case 2		' 어음상태 
			If frm1.txtStsCd.className = Parent.UCN_PROTECTED Then Exit Function

				arrParam(0) = "어음상태팝업"
				arrParam(1) = "B_MINOR A, B_CONFIGURATION B"								'popup할 sql문 
				arrParam(2) = strCode
				arrParam(3) = ""
			If frm1.cboNoteFg.Value = "D3" Then
				arrParam(4) = "A.MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  " _
							& " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD " _
							& " AND B.SEQ_NO = 2 AND B.REFERENCE = " & FilterVar("D3", "''", "S") & "  "					'
			ElseIf frm1.cboNoteFg.Value = "D1" Then
				arrParam(4) = "A.MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  " _
							& " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD " _
							& " AND B.SEQ_NO = 3 AND B.REFERENCE = " & FilterVar("D1", "''", "S") & "  "					'규칙 반영 
			Else
				arrParam(4) = "A.MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  " _
							& " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD " _
							& " AND B.SEQ_NO =  AND B.REFERENCE = " & FilterVar("D6", "''", "S") & "  "						'규칙 반영 
			End If												
				arrParam(5) = "어음상태"
	
				arrField(0) = "A.MINOR_CD"					' form1에 선택한 minor_cd,nm표시 
				arrField(1) = "A.MINOR_NM"
				    
				arrHeader(0) = "어음상태코드"		
				arrHeader(1) = "어음상태명"
			
		Case Else
			Exit Function
	End Select
    
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function 
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function EscPopUp(Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' 은행 
				.txtBankCd.focus
			Case 1		'거래처 
				.txtBpCd.focus
			Case 2		' 어음상태 
				.txtStsCd.focus
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' 은행 
				.txtBankCd.value = arrRet(0)
				.txtBankNM.value = arrRet(1)
				.txtBankCd.focus
			Case 1		'거래처 
				.txtBpCd.value	= arrRet(0)
				.txtBpNM.value	= arrRet(1)
				.txtBpCd.focus
			Case 2		' 어음상태 
				.txtStsCd.value = arrRet(0)
				.txtStsNm.value = arrRet(1)
				.txtStsCd.focus
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

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

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If
End Function


'=======================================================================================================
'	Name : SetReturnVal()
'	Description : 
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		case 0
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 1
			frm1.txtBizAreaCd1.Value = arrRet(0)
			frm1.txtBizAreaNm1.Value = arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
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
Sub vspdData_MouseDown(Button , Shift , x , y)

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
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	

	If Row = frm1.vspdData.ActiveRow Then
		Exit Sub
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

'	Next
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
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

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    If Row <> NewRow And NewRow > 0 Then
        
		ggoSpread.Source = frm1.vspdData

    End If

End Sub

'========================================================================================================
'   Event Name : fpdtFromEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtFromEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtFromEnterDt.Focus
	End If
End Sub
'========================================================================================================
'   Event Name : fpdtToEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtToEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtToEnterDt.Focus
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
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
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
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
					<TD WIDTH=100% HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>발행일</TD>
									<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtFromIssueDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작발행일자" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToIssueDt CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="종료발행일자" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>어음구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg" NAME="cboNoteFg" ALT="어음구분" STYLE="WIDTH: 132px" tag="12X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd.value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~</TD>
									<TD CLASS=TD5 NOWRAP>은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 0)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="14XXXU" ALT="발행은행"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd1.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>거래처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=10 MAXLENGTH=10   tag="11XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value, 1)"> &nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM" NAME="txtBpNM" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="14X" ALT="거래처"> </TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>상태</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT CLASS="clstxt" TYPE=TEXT ID="txtStsCd" NAME="txtStsCd" SIZE=10 MAXLENGTH=10   tag="11XXXU" ALT="발행구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtStsCd.Value, 2)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtStsNM" NAME="txtStsNM" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="14X" ALT="발행구분"> </TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<TD HEIGHT=100% WIDTH=100%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="21" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=10% WIDTH=100%>
									<FIELDSET CLASS="CLSFLD">
										<TABLE <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS=TD5 NOWRAP>어음금액합계</TD>
												<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtNoteSum style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
												<TD CLASS=TD5 NOWRAP>결제금액합계</TD>
												<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtSttlSum style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>
									</FIELDSET>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
		
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="2"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtFromIssueDt" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtToIssueDt" tag="2">
<INPUT TYPE=HIDDEN NAME="hcboNoteFg" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtBankCd" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtBpCd" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtStsCd" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1" tag="24">

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
