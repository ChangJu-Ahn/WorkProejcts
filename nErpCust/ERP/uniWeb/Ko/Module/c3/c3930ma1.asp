
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3605ma1
'*  4. Program Name         : 재료비재공내역 조회 
'*  5. Program Desc         : 재료비재공내역 조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/13
'*  8. Modified date(Last)  : 2002/03/25
'*  9. Modifier (First)     : Cho Ig sung
'* 10. Modifier (Last)      : jang yoon ki
'* 11. Comment              :
'====================================================================================================== -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "C3930MB1.asp"                              '☆: Biz Logic ASP Name

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

Dim lgMaxFieldCount
Dim lgCookValue 


Dim lgSaveRow 
														'☜: indicates that All variables must be declared in advance

'======================================================================================================
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'=======================================================================================================

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
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
	Dim StartDate
	
	StartDate	= "<%=GetSvrDate%>"


	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)

End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "QA") %>                                '☆: 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
  
	Call SetZAdoSpreadSheet("C3930MA101","G","A","V20021213",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock 

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    .vspdData.ReDraw = True
    
    End With
End Sub


'======================================================================================================
'	Name : OpenCostCd()
'	Description : Cost Center PopUp
'=======================================================================================================
Function OpenPlantCd(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장팝업"			'팝업 명칭 
	arrParam(1) = "B_PLANT"						'TABLE 명칭 
	arrParam(2) = strCode						'Code Condition
	arrParam(3) = ""							'Name Condition
	arrParam(4) = ""							'Where Condition
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"					    'Field명(0)
    arrField(1) = "PLANT_NM"					    'Field명(1)
    
    arrHeader(0) = "코스트센타코드"					'Header명(0)
    arrHeader(1) = "코스트센타명"					'Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlantCd(arrRet)
	End If	

End Function

'======================================================================================================
'	Name : SetCostCd()
'	Description : Cost Center Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetPlantCd(Byval arrRet)
	
	With frm1
		.txtPlantCd.focus
	   	.txtPlantCd.value = arrRet(0)
    	.txtPlantNm.value = arrRet(1)

	End With
	
End Function

'======================================================================================================
'	Name : OpenCostCd()
'	Description : Cost Center PopUp
'=======================================================================================================
Function OpenItemAcct(Byval strCode,ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	select case iWhere
		case 1
			arrParam(0) = "모품목계정팝업"			'팝업 명칭 
			arrParam(1) = "B_MINOR"						'TABLE 명칭 
			arrParam(2) = strCode						'Code Condition
			arrParam(3) = ""							'Name Condition
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1001", "''", "S") & " "							'Where Condition
			arrParam(5) = "모품목계정"			
	
			arrField(0) = "MINOR_CD"					    'Field명(0)
			arrField(1) = "MINOR_NM"					    'Field명(1)
    
			arrHeader(0) = "모품목계정"					'Header명(0)
			arrHeader(1) = "모품목계정명"					'Header명(1)
		case 2
			arrParam(0) = "자품목계정팝업"			'팝업 명칭 
			arrParam(1) = "B_MINOR"						'TABLE 명칭 
			arrParam(2) = strCode						'Code Condition
			arrParam(3) = ""							'Name Condition
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1001", "''", "S") & " "							'Where Condition
			arrParam(5) = "자품목계정"			
	
			arrField(0) = "MINOR_CD"					    'Field명(0)
			arrField(1) = "MINOR_NM"					    'Field명(1)
    
			arrHeader(0) = "자품목계정"					'Header명(0)
			arrHeader(1) = "자품목계정명"					'Header명(1)
	end select
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		select case iWhere
			case 1
				frm1.txtParentItemAcctCd.focus
			case 2
				frm1.txtChildItemAcctCd.focus
		end select
		
		Exit Function
	Else
		Call SetItemAcct(arrRet,iWhere)
	End If	

End Function

'======================================================================================================
'	Name : SetCostCd()
'	Description : Cost Center Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetItemAcct(Byval arrRet,ByVal iwhere)
	
	With frm1
		select case iWhere
			case 1
				.txtParentItemAcctCd.focus
	   			.txtParentItemAcctCd.value = arrRet(0)
    			.txtParentItemAcctNm.value = arrRet(1)
			case 2
				.txtChildItemAcctCd.focus
	   			.txtChildItemAcctCd.value = arrRet(0)
    			.txtChildItemAcctNm.value = arrRet(1)
		end select
	End With
	
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


'======================================================================================================
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'=======================================================================================================

'======================================================================================================
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()
	
	Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)  
    Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
	
	
'    lgMaxFieldCount =  UBound(Parent.gFieldNM)                      

'    ReDim lgPopUpR(Parent.C_MaxSelList - 1,1)

 '   Call Parent.MakePopData(Parent.gDefaultT,Parent.gFieldNM,Parent.gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,Parent.C_MaxSelList)

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")
    frm1.txtYyyymm.focus
    frm1.txtBasSum.allownull = False
    frm1.txtIssueSum.allownull = False
    frm1.txtRcptSum.allownull = False
    frm1.txtBalSum.allownull = False
    										
    'Call InitComboBox()
    'Call CookiePage(0)
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYyyymm_DblClick(Button)
    If Button = 1 Then
        frm1.txtYyyymm.Action = 7
   		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYyyymm_Change()
    lgBlnFlgChgValue = True
End Sub



'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_ItemPopUp Then
        .Col = Col
        .Row = Row
        
        ' Status
        .Col = C_ReqStatus 'ggoSpread.SSGetColsIndex(8)
        If .Text = "A" Then Exit Sub
        
        Call OpenItemInfo(.Text, 1)
        
    End If
    
    End With
    Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")   	
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
'	Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
'		Dim intIndex
'
'		 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
'		With frm1.vspdData
'		
'			.Row = Row
'	    
'			Select Case Col
'				Case  1
'					.Col = Col
'					intIndex = .Value
'					.Col = C_BillFG
'					.Value = intIndex
'			End Select
'		End With
'	End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub




Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'======================================================================================================
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'=======================================================================================================

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    On Error Resume Next
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
    Call InitVariables                                                      'Initializes local global variables
    Call InitTxtAmtSumClear
    															
  '-----------------------
    'Check condition area
    '----------------------- 
 
    
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Query function call area
    '-----------------------
   
	IF DbQuery = False Then
		Exit Function
	END IF
	       
    FncQuery = True															
    
End Function


'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 '☜: 화면 유형 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      '☜:화면 유형, Tab 유무 
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
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal
	Dim txtDate
	'Dim hDate
	
	txtDate = frm1.txtYyyyMm.Year & Right("0" & frm1.txtYyyyMm.Month,2)	
	'hDate =  frm1.hYyyymm.Year & Right("0" & frm1.hYyyymm.Month,2)	
	
	Err.Clear                                                                   '☜: Protect system from crashing
	DbQuery = False
	
	Call LayerShowHide(1)
	With frm1	
		strVal = BIZ_PGM_ID
		
    '---------Developer Coding part (Start)----------------------------------------------------------------
		If lgIntFlgMode <> Parent.OPMD_UMODE Then										'This means that it is first search
			strVal = strVal & "?txtYyyyMm="	& Trim(txtDate)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtPlantCd="	& Trim(.txtPlantCd.value)
			strVal = strVal & "&txtParentItemAcctCd="	& Trim(.txtParentItemAcctCd.value)	 	 
			strVal = strVal & "&txtChildItemAcctCd="	& Trim(.txtChildItemAcctCd.value)	 	 
		Else								'This means that it is first search
			strVal = strVal & "?txtYyyyMm="	& Trim(.hYyyymm.value)	 			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtPlantCd="	& Trim(.hPlantCd.value)	
			strVal = strVal & "&txtParentItemAcctCd="	& Trim(.hParentItemAcctCd.value)	 
			strVal = strVal & "&txtChildItemAcctCd="	& Trim(.hChildItemAcctCd.value)	 
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



'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================================
Function DbQueryOk()													'조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field

    Call SetToolbar("11000000000111")
	
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
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "3")	   
	   Exit Function        
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

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
'	Dim ii
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
	
'	For ii = 1 to Ubound(lgKeyPos)
'        frm1.vspdData.Col = lgKeyPos(ii)
'        frm1.vspdData.Row = Row
'        lgKeyPosVal(ii)   = frm1.vspdData.text
'		lgCookValue       = lgCookValue & Trim(lgKeyPosVal(ii)) & Parent.gRowSep 
'	Next
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
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
'   Event Name : fpdtFromEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
       frm1.txtYyyymm.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtYyyymm.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================


Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
' Function Name : ()
' Function Desc : 
'========================================================================================================
Sub InitTxtAmtSumClear()

	frm1.txtBasSum.text = "0"
	frm1.txtIssueSum.text = "0"
	frm1.txtRcptSum.text = "0"
	frm1.txtBalSum.text = "0"

end Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재료비재공내역조회</font></td>
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
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/c3930ma1_fpDateTime1_txtYyyymm.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlantCd frm1.txtPlantCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14"></TD>								
										</OBJECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>모품목계정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtParentItemAcctCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="모품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenItemAcct(frm1.txtParentItemAcctCd.value,1) ">&nbsp;<INPUT TYPE=TEXT NAME="txtParentItemAcctNm" SIZE=30 tag="14"></TD>								
									<TD CLASS="TD5" NOWRAP>자품목계정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtChildItemAcctCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="자품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenItemAcct(frm1.txtChildItemAcctCd.value,2) ">&nbsp;<INPUT TYPE=TEXT NAME="txtChildItemAcctNm" SIZE=30 tag="14"></TD>								
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
								<script language =javascript src='./js/c3930ma1_I958719408_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>기초금액합계</TD>
									<TD CLASS=TD6 NOWRAP>											
										<script language =javascript src='./js/c3930ma1_fpDoubleSingle2_txtBasSum.js'></script>&nbsp;
	                                </TD>
									<TD CLASS="TD5" NOWRAP>출고금액합계</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3930ma1_fpDoubleSingle2_txtIssueSum.js'></script>&nbsp;
									</TD>
								</TR>
							</TABLE>
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS="TD5" NOWRAP>입고금액합계</TD>
									<TD CLASS=TD6 NOWRAP>										
										<script language =javascript src='./js/c3930ma1_fpDoubleSingle2_txtRcptSum.js'></script>&nbsp;
									</TD>
									<TD CLASS="TD5" NOWRAP>기말금액합계</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3930ma1_fpDoubleSingle2_txtBalSum.js'></script>&nbsp;
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No  noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>

</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1" ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="hYyyymm" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hChildItemAcctCd" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hParentItemAcctCd" tag="24" TABINDEX = "-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

