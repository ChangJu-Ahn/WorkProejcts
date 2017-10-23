<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--'**********************************************************************************************
*  1. Module Name          : 회계
*  2. Function Name        : 결산
*  3. Program ID           : A5312MA1_KO441
*  4. Program Name         : 
*  5. Program Desc         : 환평가전표조회
*  6. Comproxy List        :
*  7. Modified date(First) : 2003/12/01
*  8. Modified date(Last)  : 2005/05/04
*  9. Modifier (First)     : 
* 10. Modifier (Last)      : J
* 11. Comment              : 
'********************************************************************************************** -->
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
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"      SRC="../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit									'☜: indicates that All variables must be declared in advance
	
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "a5312mb1_ko441.asp"			'☆: 비지니스 로직 ASP명 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 500                                          '☆: Fetch max count at once
Const C_MaxKey          = 5                                           '☆: key count of SpreadSheet

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim lgIsOpenPop
Dim lgSaveRow

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인
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
Sub SetDefaultVal()
	Dim StartDate
	Dim EndDate
	Dim strYear, strMonth, strDay

	StartDate	= "<%=GetSvrDate%>"                           'Get Server DB Date

	Call ExtractDateFrom(StartDate,Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	frm1.txtYyyymm.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)
	
	
	Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat,2)

	' frm1.hOrgChangeid.value = parent.gChangeOrgId
	frm1.txtYyyymm.focus





End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A","COOKIE","QA")%>
	<% Call LoadBNumericFormatA("Q", "A", "COOKIE", "QA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub

'====================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================

Sub InitSpreadSheet()
 Call AppendNumberPlace("6","15","2")
    Call SetZAdoSpreadSheet("A5312MA1_KO441","S","A","V20090712",Parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock("A")

End Sub

Sub CurFormatNumericOCX()

	With frm1

		ggoOper.FormatFieldByObjectOfCur .txtAmtSum1,	  "USD", parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	
	End With

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock(Byval pOpt)
	if pOpt = "A" then
		ggoSpread.Source = frm1.vspdData
	    With frm1
			.vspdData.ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
	    End With
	end if
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================

Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
End Sub

'==================================================================================
' Name : PopZAdoConfigGrid
' Desc :
'==================================================================================
Sub PopZAdoConfigGrid()

	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")

End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	'Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
 

   Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)	
call CurFormatNumericOCX()
    Call ggoOper.LockField(Document, "N")                                   
								' G for Group , A for SpreadSheet No('A','B',....      
    Call InitVariables																	'⊙: Initializes local global variables
    Call SetDefaultVal
    Call InitSpreadSheet()
    Call SetToolbar("110000000001111")													'⊙: 버튼 툴바 제어	
   '---------Developer Coding part (Start)----------------------------------------------------------------

    frm1.txtYyyymm.focus



End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
Dim IntRetCd

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If


	
    If DbQuery = False Then Exit Function

    FncQuery = True													

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
	Dim imRow
	FncInsertRow = False
'	imRow = AskSpdSheetAddRowCount()
'	If imRow = "" then
'		Exit Function
'	End If

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If     
 With frm1
	.vspdData.focus
	ggoSpread.Source = .vspdData
	'.vspdData.EditMode = True
	.vspdData.ReDraw = False
	ggoSpread.InsertRow ,imRow
	SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	.vspdData.ReDraw = True
 End With
 Call SetToolbar("11001111001111")
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement  
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

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
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

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery()
	Dim strVal


	Err.Clear                                                               '☜: Protect system from crashing
	DbQuery = False

	Call LayerShowHide(1)

    With frm1
        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------



			strVal = strVal & "?txtYyyymm="			& Trim(frm1.txtYyyymm.year & Right("0" & frm1.txtYyyymm.month,2) )
			strVal = strVal & "&txtModuleCd="		& Trim(.txtModuleCd.value)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)



    '--------- Developer Coding Part (End) ------------------------------------------------------------

        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")		'lgSelectListDT
         
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")		'lgMaxFieldCount,lgPopUpR,parent.gFieldCD,parent.gNextSeq,parent.gTypeCD(0),parent.C_MaxSelList)
	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))



        Call RunMyBizASP(MyBizASP, strVal)							

    End With
  

		DBQuery = True


End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												


	lgBlnFlgChgValue = False
    lgIntFlgMode     = Parent.OPMD_UMODE
    lgSaveRow        = 1
  ' CALL InitData()
    Call SetToolBar("1100000000011111")	
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement

End Function


'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================

Sub OpenOrderByPopup(ByVal pSpdNo)

	Dim arrRet
	On Error Resume Next
	If lgIsOpenPop = True Then Exit Sub

	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp", _
									Array(ggoSpread.GetXMLData("A"),gMethodText), _
									"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & _
									parent.SORTW_HEIGHT & "px; ; center: Yes; help: No; resizable: No; status: No;")
									
	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
		Exit Sub
	Else
		Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
		Call InitVariables
		Call InitSpreadSheet()
	End If

End Sub




'========================================================================================================
' Name : OpenCondAreaPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	With frm1
		If IsOpenPop = True Then Exit Function 

		Select Case iWhere
			Case 1
				arrParam(0) = "사업장 팝업"		    	<%' 팝업 명칭 %>
				arrParam(1) = "B_BIZ_AREA"					<%' TABLE 명칭 %>
				arrParam(2) = frm1.txtBizAreaCd.value			<%' Code Condition%>
				arrParam(3) = "" 		            		<%' Name Cindition%>
				arrParam(4) = ""							<%' Where Condition%>
				arrParam(5) = "사업장"			
	
				arrField(0) = "BIZ_AREA_CD"					<%' Field명(0)%>
				arrField(1) = "BIZ_AREA_NM"	     			<%' Field명(1)%>
    
				arrHeader(0) = "사업장코드"				<%' Header명(0)%>
				arrHeader(1) = "사업장명"				<%' Header명(1)%>

			Case 2
				arrParam(0) = "모듈구분팝업"										    ' 팝업 명칭 
				arrParam(1) = " b_minor "													' TABLE 명칭 
				arrParam(2) = frm1.txtModuleCd.value											' Code Condition
				arrParam(3) = ""															' Name Cindition
				arrParam(4) = " MAJOR_CD = " & FilterVar("A1045","''","S")
				arrParam(5) = "모듈구분"												' 조건필드의 라벨 명칭 

				arrField(0) = "MINOR_CD"													' Field명(0)
				arrField(1) = "MINOR_NM"													' Field명(1)
			 
				arrHeader(0) = "모듈구분"												' Header명(0)
				arrHeader(1) = "모듈구분명"												' Header명(1)

		End Select    
	End With
	
	IsOpenPop = True
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.txtBizAreaCd.focus
			Case 2
				frm1.txtModuleCd.focus
		End Select    
		Exit Function
	Else
		Call SetMajor(arrRet, iWhere)
	End If	

End Function


'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetMajor(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtBizAreaCd.focus
				.txtBizAreaCd.value = arrRet(0)
				.txtBizAreaNm.value = arrRet(1)
				.htxtBizAreaCd.value = arrRet(0)
			Case 2
				.txtModuleCd.focus
				.txtModuleCd.value = arrRet(0)
				.txtModuleName.value = arrRet(1)		
		End Select    
	End With
End Function



Sub txtYyyymm_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYyyymm.Action = 7
 		Call SetFocusToDocument("M")
		Frm1.txtYyyymm.Focus
   End If
End Sub



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
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
 	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub
	

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001")
	Dim ii

	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
    
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
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

Sub vspdData_Change(ByVal Col, ByVal Row)
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)  
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
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
'   Event Name : txtDeptCd_Keypress()
'   Event Desc : 
'========================================================================================================
Sub txtDeptCd_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()		
	End If   
End Sub

'========================================================================================================
'   Event Name : txtAcctCd_Keypress()
'   Event Desc : 
'========================================================================================================
Sub txtAcctCd_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()		
	End If   
End Sub

'========================================================================================================
'   Event Name : txtCondAsstNo_Keypress()
'   Event Desc : 
'========================================================================================================
Sub txtCondAsstNo_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()		
	End If   
End Sub

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 



'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub





'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
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
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						  <TABLE <%=LR_SPACE_TYPE_40%>>
			            	                 <TR>
			            		                <TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript>ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtYyyymm" CLASS=FPDTYYYYMMDD tag="12X1" ALT="작업년월" Title="FPDATETIME"></OBJECT>')</script></TD>
								
			            		                <TD CLASS="TD5" NOWRAP>모듈구분</TD>
			            		                <TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtModuleCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="모듈구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnModuleCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup(2)">
									<INPUT TYPE="Text" NAME="txtModuleName" SiZE=22 MAXLENGTH=50 tag="14XXXU" ALT="모듈구분명">
			            		                 </TD>
                                                        </TR>
			                            	<TR>
			            	                 	<TD CLASS="TD5" NOWRAP>사업장</TD>
			            		                <TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup(1)">
									<INPUT TYPE="Text" NAME="txtBizAreaNm" SiZE=22 MAXLENGTH=50 tag="14XXXU" ALT="사업장명">
			            		                </TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR HEIGHT =100%>
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							 </TR>
						<TR>
								<TD HEIGHT=20 WIDTH=100%>
									<FIELDSET>
										<LEGEND>합계</LEGEND>
										<TABLE <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>잔액</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=FPDOUBLESINGLE name=txtAmtSum1 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="전기말상각누계액" tag="24X2"></OBJECT>');</SCRIPT></TD>
												<TD CLASS="TD5" NOWRAP>잔액(자국)</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtAmtSum2 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="전기말미상각잔액" tag="24X2"></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>평가손</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtAmtSum3 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="전기말상각누계액" tag="24X2"></OBJECT>');</SCRIPT></TD>
												<TD CLASS="TD5" NOWRAP>평가익</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtAmtSum4 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="전기말미상각잔액" tag="24X2"></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>평가금액(자국)</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtAmtSum5 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="전기말상각누계액" tag="24X2"></OBJECT>');</SCRIPT></TD>
												<TD CLASS="TD5" /TD>
												<TD CLASS="TD6" /TD>
											</TR>

										</TABLE>
									</FIELDSET>
								</TD>
							<TR>

						</TABLE>
					</TD>
			
				</TR>
				
			</TABLE>
		</TD>
	</TR>
		
	<TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="htxtFr_dt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtTo_dt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDeptCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtAcctCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtCondAsstNo"	tag="24">
<INPUT TYPE=HIDDEN NAME="hDurYrsFg"	        tag="24">

<INPUT TYPE=HIDDEN NAME="htxtBizUnitCd"	    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1"	tag="24">

<INPUT TYPE=TEXT NAME="hOrgChangeId"		tag="14" TABINDEX="-1">
<INPUT TYPE=TEXT NAME="hINternalCD"		tag="14" TABINDEX="-1">
				
				
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TabIndex="-1">	
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

