<%@ LANGUAGE="VBSCRIPT"%>

<!--
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7132ma1
'*  4. Program Name         : 감가상각방법등록 
'*  5. Program Desc         : 상각방법별로 상각Option을 등록한다.
'*  6. Modified date(First) : 2003/09/19
'*  7. Modified date(Last)  : 2003/09/19
'*  8. Modifier (First)     : Park, Joon Won
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
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
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit         '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "a7132mb1.asp"   '☆: 비지니스 로직 ASP명 

'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================
Dim C_DeprMthd
Dim C_DeprMthdNm
Dim C_DeprFg
Dim C_DeprFgNm
Dim C_DeprUnit
Dim C_DeprUnitNm
Dim C_DeprMeth
Dim C_DeprMethNm
Dim C_DeprTerm
Dim C_DeprTermNm
Dim C_DeprInc
Dim C_DeprIncNm
Dim C_DeprSold
Dim C_DeprSoldNm
Dim C_ResRate
Dim C_DeprCloseT


On Error Resume Next
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
Dim lgRetFlag
Dim IsOpenPop        


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize column position value in spreadsheet
'========================================================================================================
Sub initSpreadPosVariables()
	C_DeprMthd		= 1
	C_DeprMthdNm	= 2
	C_DeprFg		= 3
	C_DeprFgNm		= 4
	C_DeprUnit		= 5
	C_DeprUnitNm	= 6	
	C_DeprMeth		= 7
	C_DeprMethNm	= 8	
	C_DeprTerm		= 9
	C_DeprTermNm	= 10
	C_DeprInc		= 11
	C_DeprIncNm		= 12
	C_DeprSold		= 13
	C_DeprSoldNm	= 14	
	C_ResRate		= 15
	C_DeprCloseT    = 16

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgPageNo     = "0"
    
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
 		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
    	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData

		.MaxCols = C_DeprCloseT +1       '☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols						
		.ColHidden = True
      
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030918",,parent.gAllowDragDropSpread    

		Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

		.ReDraw = false

		Call AppendNumberPlace("6","3","0")
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit  C_DeprMthd,    "상각방법",10,,,2,2'1
		ggoSpread.SSSetEdit  C_DeprMthdNm,  "상각방법명",30 '3
		ggoSpread.SSSetCombo C_DeprFg,		"상각여부", 5 '4
		ggoSpread.SSSetCombo C_DeprFgNm,	"상각여부", 14 '5
		ggoSpread.SSSetCombo C_DeprUnit,    "상각단위", 5
		ggoSpread.SSSetCombo C_DeprUnitNm,  "상각단위", 10   '15
		ggoSpread.SSSetCombo C_DeprMeth,    "상각방식", 5
		ggoSpread.SSSetCombo C_DeprMethNm,  "상각방식", 10  '25
		ggoSpread.SSSetCombo C_DeprTerm,    "월할여부", 5
		ggoSpread.SSSetCombo C_DeprTermNm,  "월할여부", 10  '25
		ggoSpread.SSSetCombo C_DeprInc,     "지출처리방법", 5
		ggoSpread.SSSetCombo C_DeprIncNm,   "지출처리방법", 15  '25
		ggoSpread.SSSetCombo C_DeprSold,    "매각월상각여부", 5
		ggoSpread.SSSetCombo C_DeprSoldNm,  "매각월상각여부", 15  '25
	    ggoSpread.SSSetFloat C_ResRate,	"잔존율(%)", 33,Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z",,999999999
	    ggoSpread.SSSetEdit  C_DeprCloseT,  "",2 '3
				

		Call ggoSpread.SSSetColHidden(C_DeprFg,C_DeprFg,True)
		Call ggoSpread.SSSetColHidden(C_DeprUnit,C_DeprUnit,True)
		Call ggoSpread.SSSetColHidden(C_DeprMeth,C_DeprMeth,True)
		Call ggoSpread.SSSetColHidden(C_DeprTerm,C_DeprTerm,True)
		Call ggoSpread.SSSetColHidden(C_DeprInc,C_DeprInc,True)
		Call ggoSpread.SSSetColHidden(C_DeprSold,C_DeprSold,True)
		Call ggoSpread.SSSetColHidden(C_DeprCloseT,C_DeprCloseT,True)
		
		Call ggoSpread.MakePairsColumn(C_DeprFg,C_DeprFgNm)
		Call ggoSpread.MakePairsColumn(C_DeprUnit,C_DeprUnitNm)
		Call ggoSpread.MakePairsColumn(C_DeprMeth,C_DeprMethNm)
		Call ggoSpread.MakePairsColumn(C_DeprTerm,C_DeprTermNm)		
		Call ggoSpread.MakePairsColumn(C_DeprInc,C_DeprIncNm)
		Call ggoSpread.MakePairsColumn(C_DeprSold,C_DeprSoldNm)
		
		Call InitComboBox
		
		.ReDraw = true
    
    End With
	Call SetSpreadLock     
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
	With frm1

		.vspdData.ReDraw = False
	
		ggoSpread.SpreadLock C_DeprMthd, -1, C_DeprMthd
		ggoSpread.SpreadLock C_DeprMthdNm, -1, C_DeprMthdNm
		ggoSpread.SpreadLock  C_DeprFgNm,     -1, C_DeprFgNm
		ggoSpread.SpreadLock  C_DeprUnitNm,   -1, C_DeprUnitNm
		ggoSpread.SpreadLock  C_DeprMethNm,   -1, C_DeprMethNm
		ggoSpread.SpreadLock  C_DeprTermNm,   -1, C_DeprTermNm
		ggoSpread.SpreadLock  C_DeprIncNm,    -1, C_DeprIncNm
		ggoSpread.SpreadLock  C_DeprSoldNm,   -1, C_DeprSoldNm
		ggoSpread.SpreadLock  C_ResRate,	  -1, C_ResRate
		ggoSpread.SpreadLock	.vspdData.MaxCols,-1,-1

	.vspdData.ReDraw = True
	
	End With
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False

    ggoSpread.SSSetRequired C_DeprMthd,	pvStarRow, pvEndRow
	ggoSpread.SSSetRequired C_DeprMthdNm,	pvStarRow, pvEndRow
    ggoSpread.SSSetRequired  C_DeprFgNm,	pvStarRow, pvEndRow
    ggoSpread.SSSetRequired  C_DeprUnitNm,	pvStarRow, pvEndRow
	ggoSpread.SSSetRequired  C_DeprMethNm,	pvStarRow, pvEndRow
	ggoSpread.SSSetRequired  C_DeprTermNm,	pvStarRow, pvEndRow
	ggoSpread.SSSetRequired  C_DeprIncNm,	pvStarRow, pvEndRow
	ggoSpread.SSSetRequired  C_DeprTermNm,	pvStarRow, pvEndRow
	ggoSpread.SSSetRequired  C_DeprSoldNm,	pvStarRow, pvEndRow
	ggoSpread.SSSetRequired  C_ResRate,		pvStarRow, pvEndRow
  
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_DeprMthd		= iCurColumnPos(1)
			C_DeprMthdNm	= iCurColumnPos(2)
			C_DeprFg		= iCurColumnPos(3)
			C_DeprFgNm		= iCurColumnPos(4)
			C_DeprUnit		= iCurColumnPos(5)
			C_DeprUnitNm	= iCurColumnPos(6)	
			C_DeprMeth		= iCurColumnPos(7)
			C_DeprMethNm	= iCurColumnPos(8)	
			C_DeprTerm		= iCurColumnPos(9)
			C_DeprTermNm	= iCurColumnPos(10)
			C_DeprInc		= iCurColumnPos(11)
			C_DeprIncNm		= iCurColumnPos(12)
			C_DeprSold		= iCurColumnPos(13)
			C_DeprSoldNm	= iCurColumnPos(14)	
			C_ResRate		= iCurColumnPos(15)
			C_DeprCloseT    = iCurColumnPos(16)
	End Select
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
' Name : InitComboBox()
' Description : Combo Display
'========================================================================================================= 

Sub InitComboBox()

' ------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim IntRetCD1
	Dim IntRetCD2
	Dim IntRetCD3
	  
	On Error Resume Next

	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = " & FilterVar("A2012", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_DeprFg
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_DeprFgNm
	End If

	IntRetCD2 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = " & FilterVar("A2013", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If IntRetCD2 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_DeprUnit
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_DeprUnitNm
	End If

	IntRetCD3 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = " & FilterVar("A2014", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If intRetCD <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_DeprMeth
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_DeprMethNm
	End If
	
	IntRetCD4 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = " & FilterVar("A2015", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If intRetCD <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_DeprTerm
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_DeprTermNm
	End If
	
	IntRetCD5 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = " & FilterVar("A2016", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If intRetCD <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_DeprInc
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_DeprIncNm
	End If
	
	IntRetCD6 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = " & FilterVar("A2017", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If intRetCD <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_DeprSold
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_DeprSoldNm
	End If
	
' ------ Developer Coding part (End )   --------------------------------------------------------------


end sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
' 기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'       하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
' Name : Open???()
' Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'      ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=========================================================================================================

Function OpenDepr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "상각방법팝업"	
	arrParam(1) = " a_asset_depr_method "
	arrParam(2) = Trim(frm1.txtDeprCd.Value)
	arrParam(3) = ""			
	arrParam(4) = ""			
	arrParam(5) = "상각방법"		
	
    arrField(0) = "depr_mthd"	
    arrField(1) = "depr_mthd_nm"		
    
    arrHeader(0) = "상각방법코드"		
    arrHeader(1) = "상각방법명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeprCd.focus
		Exit Function
	Else
		Call SetDepr(arrRet)
	End If
		
End Function


Function SetDepr(byval arrRet)
	frm1.txtDeprCd.focus
	frm1.txtDeprCd.Value    = arrRet(0)		
	frm1.txtDeprNm.Value    = arrRet(1)		
End Function



Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
'  IntRetCD = DisplayMsgBox(frm1.vspdData.Maxcols , parent.VB_YES_NO, "X", "X")
	

	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'            3. Event부 
' 기능: Event 함수에 관한 처리 
' 설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
' Window에 발생 하는 모든 Even 처리 
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call InitVariables                                                      '⊙: Initializes local global variables

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitComboBox
    'Call MASetToolbar("11001101001011")          '⊙: 버튼 툴바 제어 
    Call SetToolbar("11100100000011")   
    frm1.txtDeprCd.focus 
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
' Document의 TAG에서 발생 하는 Event 처리 
' Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
' Event간 충돌을 고려하여 작성한다.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag 처리  *********************************************
' Window에 발생 하는 모든 Even 처리 
'*********************************************************************************************************
Sub txtDeprCd_OnChange()
	If Trim(frm1.txtDeprCd.value) = "" Then
		frm1.txtAcctNm.value = ""
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================

Sub vspdData_Change(ByVal Col, ByVal Row)
    
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)  

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
 Dim strTemp
 Dim intPos1

 With frm1.vspdData 

 If Row > 0 And Col = C_AcctCdPopUp Then
     .Col = C_AcctCd
     .Row = Row
         
     Call OpenAcct(1)
 End If
     
 End With
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
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


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    Dim iColumnName
 	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
   
'    If Row <= 0 Then
'       frm1.vspdData.Row=Row
'       frm1.vspdData.Col=Col
'       iColumnName = frm1.vspdData.Text

'       iColumnName = AskSpdSheetColumnName(iColumnName)
        
'       If iColumnName <> "" Then
'          ggoSpread.Source = frm1.vspdData
'          Call ggoSpread.SSSetReNameHeader(Col,iColumnName)

          'Call SetSortFieldNM("A", frm1.vspdData,Col)
'       End If
        
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
'    End If
End Sub



'==========================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
 Dim index
 
 With frm1.vspdData
  .Row = Row

  .Col = Col
  index = .Value
   
  .Col = Col - 1
  .Value = index
 End With
 
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
'    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then '☜: 재쿼리 체크 
'  If lgStrPrevKey <> "" Then       '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
'   DbQuery
'  End If
 '   End if

End Sub


'#########################################################################################################
'            4. Common Function부 
' 기능: Common Function
' 설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 


'#########################################################################################################
'            5. Interface부 
' 기능: Interface
' 설명: 각각의 Toolbar에 대한 처리를 행한다. 
'       Toolbar의 위치순서대로 기술하는 것으로 한다. 
' << 공통변수 정의 부분 >>
'  공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'    통일하도록 한다.
'  1. 공통컨트롤을 Call하는 변수 
'        ADF (ADS, ADC, ADF는 그대로 사용)
'        - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
'  2. 공통컨트롤에서 Return된 값을 받는 변수 
'      strRetMsg
'#########################################################################################################

'********************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
' 설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                    '⊙: Processing is NG
    
    Err.Clear                                                           '☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")       '데이타가 변경되었습니다. 조회하시겠습니까?
     If IntRetCD = vbNo Then
       Exit Function
     End If
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then       '⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")        '⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables
                     '⊙: Initializes local global variables
 
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery                '☜: Query db data
       
    FncQuery = True                '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then
   Exit Function
  End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    
    Call SetToolbar("11000100000011")          '⊙: 버튼 툴바 제어 
    
    FncNew = True                                                           '⊙: Processing is OK

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
  Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")  '☜ 바뀐부분 
    If IntRetCD = vbNo Then
        Exit Function
    End If

    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
  '-----------------------
  'Precheck area
  '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                          'No data changed!!
        Exit Function
    End If
    
  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call   DbSave                                                     '☜: Save db data
    
' frm1.vspdData.ReDraw = False
' ggoSpread.SSDeleteFlag 1 , frm1.vspdData.MaxRows
'   Call SetSpreadLock
' frm1.vspdData.ReDraw = True

 FncSave = True                                                          '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData
  .ReDraw = False
 
  ggoSpread.Source = frm1.vspdData 
     ggoSpread.CopyRow
  SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
  'Key field clear
  .Col = C_AcctCd
  .Text = ""
  
  .Col = C_AcctNm
  .Text = ""

  .ReDraw = True
    End With
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel()


 Call SetToolbar("11001111001111")          '⊙: 버튼 툴바 제어 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData 
    ggoSpread.EditUndo             
    
    Call InitData 
                                         '☜: Protect system from crashing
 lgBlnFlgChgValue = False
     
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(Byval pvRowCnt) 
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


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
 If frm1.vspdData.MaxRows < 1 Then Exit Function
    
    With frm1.vspdData 
     .focus
  ggoSpread.Source = frm1.vspdData 
    
  lDelRows = ggoSpread.DeleteRow

    End With
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)            '☜: 화면 유형 
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
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
 Dim IntRetCD
 FncExit = False
    ggoSpread.Source = frm1.vspdData 
    If ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
  If IntRetCD = vbNo Then
   Exit Function
  End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
' 설명 : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

 Call ggoOper.ClearField(Document, "2")
 ggoSpread.Source = frm1.vspdData
 ggospread.ClearSpreadData		'Buffer Clear
 
 Call InitVariables
 Call LayerShowHide(1)

 Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtDeprCd=" & .hDeprCd.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtDeprCd=" & .txtDeprCd.value
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If

 Call RunMyBizASP(MyBizASP, strVal)          '☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()              '☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE            '⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")         '⊙: This function lock the suitable field
 call InitData
    Call SetToolbar("11001111001111")          '⊙: 버튼 툴바 제어 
 

End Function

Sub InitData()
 Dim intRow
 Dim intIndex 

 With frm1.vspdData
  For intRow = 1 To .MaxRows
   
   .Row = intRow
   
   .Col = C_DeprFg
   intIndex = .value
   .col = C_DeprFgNm
   .value = intindex
    
   .Col = C_DeprUnit
   intIndex = .value
   .col = C_DeprUnitNm
   .value = intindex
       
   .Col = C_DeprMeth
   intIndex = .value
   .col = C_DeprMethNm
   .value = intindex
   
   .Col = C_DeprTerm
   intIndex = .value
   .col = C_DeprTermNm
   .value = intindex
   
   .Col = C_DeprInc
   intIndex = .value
   .col = C_DeprIncNm
   .value = intindex
   
   .Col = C_DeprSold
   intIndex = .value
   .col = C_DeprSoldNm
   .value = intindex

  Next 
 End With
End Sub



Function DbSave() 
    Dim aAs0011     'As New AS0011ManageSvr
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	Dim iColSep
    Dim iRowSep   
 'Dim ColSep, RowSep
 
    DbSave = False                                                          '⊙: Processing is NG
    
    On Error Resume Next                                                   '☜: Protect system from crashing

 Call LayerShowHide(1)
 
 With frm1
  .txtMode.value = parent.UID_M0002
  
  '-----------------------
  'Data manipulate area
  '-----------------------
  lGrpCnt = 1
  strVal = ""
  strDel = ""
  
  iColSep = Parent.gColSep
  iRowSep = Parent.gRowSep	
    
  '-----------------------
  'Data manipulate area
  '-----------------------
  For lRow = 1 To .vspdData.MaxRows
    
      .vspdData.Row = lRow
      .vspdData.Col = 0
      
      Select Case .vspdData.Text

          Case ggoSpread.InsertFlag       '☜: 신규 
     
     strVal = strVal & "C" & iColSep & lRow & iColSep '☜: C=Create, Row위치 정보 

              .vspdData.Col = C_DeprMthd
              strVal = strVal & Trim(.vspdData.Text) & iColSep
              
              .vspdData.Col = C_DeprMthdNm
              strVal = strVal & Trim(.vspdData.Text) & iColSep
              
              .vspdData.Col = C_DeprFg
              strVal = strVal & Trim(.vspdData.Text) & iColSep
              
              .vspdData.Col = C_DeprUnit
              strVal = strVal & Trim(.vspdData.Text) & iColSep

              .vspdData.Col = C_DeprMeth
              strVal = strVal & Trim(.vspdData.Text) & iColSep

              .vspdData.Col = C_DeprTerm
              strVal = strVal & Trim(.vspdData.Text) & iColSep
              
              .vspdData.Col = C_DeprInc
              strVal = strVal & Trim(.vspdData.Text) & iColSep
              
              .vspdData.Col = C_DeprSold
              strVal = strVal & Trim(.vspdData.Text) & iColSep
              
              .vspdData.Col = C_ResRate
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
              
              .vspdData.Col = C_DeprCloseT
              strVal = strVal & "1" & iRowSep
              
                            
              lGrpCnt = lGrpCnt + 1
              
          Case ggoSpread.UpdateFlag       '☜: 수정 

'     strVal = strVal & "U" & iColSep & lRow & iColSep '☜: U=Update

'              .vspdData.Col = C_AcctCd
'              strVal = strVal & Trim(.vspdData.Text) & iColSep

'              .vspdData.Col = C_DeprMthd
'              strVal = strVal & Trim(.vspdData.Text) & iColSep
              
'              .vspdData.Col = C_DurYrs
'              strVal = strVal & Trim(.vspdData.Text) & iColSep

'              .vspdData.Col = C_AcctFg
'              strVal = strVal & Trim(.vspdData.Text) & iColSep

'              .vspdData.Col = C_DeprFg
'              strVal = strVal & Trim(.vspdData.Text) & iRowSep
              
'              lGrpCnt = lGrpCnt + 1
              
          Case ggoSpread.DeleteFlag       '☜: 삭제 

     strDel = strDel & "D" & iColSep & lRow & iColSep'☜: D=Delete

              .vspdData.Col = C_DeprMthd
              strDel = strDel & Trim(.vspdData.Text) & iRowSep
              
              lGrpCnt = lGrpCnt + 1
      End Select

  Next
  
  .txtMaxRows.value = lGrpCnt-1
  .txtSpread.value = strDel & strVal
  'msgbox GetUserPath 
  'msgbox BIZ_PGM_ID
  Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 

 End With
 
    DbSave = True                                                           '⊙: Processing is NG

End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()             '☆: 저장 성공후 실행 로직 
 
 Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
   
 Call InitVariables
 'Call InitSpreadSheet  '스프레드시트 초기화 막음 
    Call InitComboBox
 'lgBlnFlgChgValue = False
 
 Call DBQuery()
 'Call MainQuery()
 
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function

Sub txtDeprCd_onChange()
	Dim IntRetCD
	Dim arrVal

	If frm1.txtDeprCd.value = "" Then Exit Sub

	If CommonQueryRs("DEPR_MTHD_NM", "A_ASSET_DEPR_METHOD ", " DEPR_MTHD=  " & FilterVar(frm1.txtDeprCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtDeprNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("117420","X","X","X")  	
		frm1.txtDeprCd.focus
	End If
End Sub



'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!--
'#########################################################################################################
'            6. Tag부 
'######################################################################################################### 
 -->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
 <TR>
  <TD  <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>감가상각방법등록</font></td>
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
				<TD CLASS="TD5" NOWRAP>상각방법</TD>
				<TD CLASS="TD656" COLSPAN=3><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtDeprCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="상각방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDepr()">
				 <INPUT TYPE=TEXT ID="txtDeprNm" NAME="txtDeprNm" SIZE=25 tag="14X">
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
     <TD WIDTH=100% valign=top>
      <TABLE <%=LR_SPACE_TYPE_20%>>
       <TR>
        <TD WIDTH="100%" NOWRAP>
         <script language =javascript src='./js/a7132ma1_I547683215_vspdData.js'></script>
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
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
  </TD>
 </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hDeprCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



