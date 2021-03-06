
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 표준원가관리 
'*  3. Program ID           : c2410ma1
'*  4. Program Name         : 표준원가 간접비 산출근거조회 
'*  5. Program Desc         : 공장별 표준계산시 간접비에 대한 산출근거를 한다.
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/09/02
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Ig Sung, Cho
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

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs">          </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "C2410MB1.asp"												'☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'Const C_SHEETMAXROWS_D  = 100                                          '☆: Fetch max count at once
Const C_MaxKey          = 5                                           '☆: key count of SpreadSheet

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim lgIsOpenPop

Dim lgMaxFieldCount

Dim lgCookValue

Dim lgSaveRow

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================						
Sub InitVariables()
    lgStrPrevKey		= ""
    lgPageNo			= ""
    lgIntFlgMode		= Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue	= False                    'Indicates that no value changed
	lgSortKey			= 1
	lgSaveRow			= 0

End Sub

'========================================================================================================
Sub SetDefaultVal()

'	frm1.fpdtFromEnterDt.Text	= "<%= FromDateOfDB %>"
'	frm1.fpdtToEnterDt.Text	= "<%= ToDateOfDB %>"

End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "C", "COOKIE", "QA")%>                                '☆: 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim strTemp, arrVal

	Const CookieSplit = 4877						

	If Kubun = 0 Then                                              ' Called Area
       strTemp = ReadCookie(CookieSplit)

       If strTemp = "" then Exit Function

       arrVal = Split(strTemp, Parent.gRowSep)

       Frm1.txtSchoolCd.Value = ReadCookie ("SchoolCd")
       Frm1.txtGrade.Value   = arrVal(0)
				
       Call MainQuery()

       WriteCookie CookieSplit , ""
	
	ElseIf Kubun = 1 then                                         ' If you want to call
		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue		
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End IF
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Function


'========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'Call SetCombo(frm1.cboPrcFlg, "T", "진단가")
	'Call SetCombo(frm1.cboPrcFlg, "F", "가단가")
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub InitSpreadSheet()

		Call SetZAdoSpreadSheet("c2410ma1","S","A","V20021211",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
		Call SetSpreadLock 
End Sub


'========================================================================================================
Sub SetSpreadLock()
'    With frm1
    
'    .vspdData.ReDraw = False
'    ggoSpread.SpreadLock 1, -1
'    .vspdData.ReDraw = True

'    End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

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
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    '---------Developer Coding part (Start)----------------------------------------------------------------
'    Call Parent.GetAdoFieldInf("c2410ma1","S","A")												' G for Group , A for SpreadSheet No('A','B',....   
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")												'⊙: Lock  Suitable  Field

'	lgMaxFieldCount =  UBound(Parent.gFieldNM)

 '   Redim lgPopUpR(Parent.C_MaxSelList -1,1)
    
'    Call Parent.MakePopData(Parent.gDefaultT,Parent.gFieldNM,Parent.gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,Parent.C_MaxSelList)		'You must not this line
    
    Call InitVariables																	'⊙: Initializes local global variables
    Call SetDefaultVal
    Call InitSpreadSheet()
    Call SetToolbar("110000000001111")													'⊙: 버튼 툴바 제어	
    frm1.txtPlantCd.focus
   	Set gActiveElement = document.activeElement		
End Sub


'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'========================================================================================

Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    					
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function


'========================================================================================
Function FncNew() 
End Function

'========================================================================================
Function FncSave() 
End Function


'========================================================================================
Function FncCancel() 
    On Error Resume Next
End Function

'=======================================================================================================
Function FncInsertRow() 
    On Error Resume Next
End Function

'=======================================================================================================
Function FncDeleteRow() 
    On Error Resume Next
End Function


'=======================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function


'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'=======================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'=======================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True  
End Function


'=======================================================================================================
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


'=======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncExit = True
    
End Function
 


'========================================================================================

Function DbQuery() 
	Dim strVal

	Err.Clear                                                                   '☜: Protect system from crashing
	DbQuery = False

	Call LayerShowHide(1)

	With frm1
		strVal = BIZ_PGM_ID
    '---------Developer Coding part (Start)----------------------------------------------------------------
		If lgIntFlgMode <> Parent.OPMD_UMODE Then										'This means that it is first search
			strVal = strVal & "?txtPlantCd="	& Trim(.txtPlantCd.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd="		& Trim(.txtItemCd.value)	 			    '☆: 조회 조건 데이타 
		Else
			strVal = strVal & "?txtPlantCd="	& Trim(.htxtPlantCd.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd="		& Trim(.htxtItemCd.value)	 			    '☆: 조회 조건 데이타 
		End if
	'---------Developer Coding part (End)----------------------------------------------------------------
		strVal = strVal & "&lgPageNo="			& lgPageNo								'Next key tag
'		strVal = strVal & "&lgMaxCount="		& C_SHEETMAXROWS_D					'한번에 가져올수 있는 데이타 건수 
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")					'field type
		strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
																				'order by 구문 만들어진다 
		strVal = strVal & "&lgSelectList=" & EnCoding(GetSQLSelectList("A"))
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True

End Function


'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue	=False
	lgIntFlgMode		= Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgSaveRow			= 1
'	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
'	Call SetToolbar("110000000001111")										'⊙: 버튼 툴바 제어 
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   	

End Function



'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'=======================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장팝업"				'팝업 명칭 
	arrParam(1) = "B_PLANT"						'TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	'Code Condition
	arrParam(3) = ""							'Name Cindition
	arrParam(4) = ""							'Where Condition
	arrParam(5) = "공장"					'TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					'Field명(0)
    arrField(1) = "PLANT_NM"					'Field명(1)
    
    arrHeader(0) = "공장코드"					'Header명(0)
    arrHeader(1) = "공장명"					'Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If
		
End Function

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("125000","x","x","x") '공장을 먼저 입력하세요 
		frm1.txtPlantCd.focus
		Exit Function
	End If

	lgIsOpenPop = True

'	arrParam(0) = "품목팝업"				'팝업 명칭 
	'arrParam(1) = "B_ITEM_BY_PLANT a, B_ITEM b, C_STD_COST_BY_WORK_STEP c "						'TABLE 명칭 
	'arrParam(2) = Trim(frm1.txtItemCd.Value)	'Code Condition
	'arrParam(3) = ""							'Name Cindition
	'arrParam(4) = "c.ITEM_CD = a.ITEM_CD AND c.PLANT_CD = a.PLANT_CD AND c.WORK_STEP = '02' AND a.ITEM_CD = b.ITEM_CD AND a.PLANT_CD = '" & Trim(frm1.txtPlantCd.Value) & "'"							'Where Condition
	'arrParam(5) = "품목"					'TextBox 명칭 
	
  '  	arrField(0) = "b.ITEM_CD"					'Field명(0)
  '  	arrField(1) = "b.ITEM_NM"					'Field명(1)
 
'	arrHeader(0) = "품목코드"	  				' Header명(0)
'	arrHeader(1) = "품목명"						' Header명(1)
		
'	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
'		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "15"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	

	arrField(0) = 1 								' Field명(0) :"ITEM_CD"
	arrField(1) = 2									' Field명(1) :"ITEM_NM"

	arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent,arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet)
	End If	

End Function

 '==========================================  2.4.3 SetPopup()  =============================================
'	Name : SetPopup()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet)
	With frm1
		 frm1.txtItemCd.focus
		.TxtItemCd.Value = arrRet(0)
		.TxtItemNm.Value = arrRet(1)

		lgBlnFlgChgValue = True
		
	End With
	
End Function
'=======================================================================================================
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.focus
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'	Name : OpenSchoolCd()
'	Description : Plant PopUp
'========================================================================================================
Function OpenSchoolCd()

End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
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
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
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
    End If
	
  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgCookValue = ""
'	
'	For ii = 1 to UBound(lgKeyPos)
'		frm1.vspdData.Col = lgKeyPos(ii)
'		frm1.vspdData.Row = Row
'		lgKeyPosVal(ii) = frm1.vspdData.text
'		lgCookValue = lgCookValue & Trim(lgKeyPosVal(ii)) & Parent.gRowSep 
'	Next
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
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

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
    	If lgPageNo <> "" Then                  '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
    	'If lgStrPrevKeyCostElmtCd <> "" Then                  '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
	      	DbQuery
    	End If

    End if
    
End Sub

'========================================================================================================
'   Event Name : txtPlantCd_Keypress()
'   Event Desc : 
'========================================================================================================
Sub txtPlantCd_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()		
	End If   
End Sub

'========================================================================================================
'   Event Name : txtPlantCd_Keypress()
'   Event Desc : 
'========================================================================================================
Sub txtPlantCd_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub



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

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
	
'======================================================================================================= -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>표준원가간접비산출근거조회</font></td>
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
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
										 <INPUT TYPE=TEXT ID="txtPlantNm" NAME="txtPlantNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenItem()">
										 <INPUT TYPE=TEXT ID="txtItemNm" NAME="txtItemNm" SIZE=30 tag="14X">
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c2410ma1_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="htxtPlantCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="htxtItemCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

