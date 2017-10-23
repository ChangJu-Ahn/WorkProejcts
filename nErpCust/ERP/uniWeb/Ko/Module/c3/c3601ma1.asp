
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3601ma1
'*  4. Program Name         : CC별 배부내역 조회 
'*  5. Program Desc         : CC별 배부내역 조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/27
'*  8. Modified date(Last)  : 2002/03/05
'*  9. Modifier (First)     : Cho Ig Sung
'* 10. Modifier (Last)      : Jang Yoon Ki
'* 11. Comment              :
'======================================================================================================= -->
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
                                                              '☜: indicates that All variables must be declared in advance 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

Const BIZ_PGM_ID = "c3601mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "c3601mb2.asp"											 '☆: 비지니스 로직 ASP명 


'Const C_SHEETMAXROWS_D_A  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 

'Const C_SHEETMAXROWS_D_B  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey            = 3                                    '☆☆☆☆: Max key value

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================	
Dim lgIsOpenPop                                             '☜: Popup status                           
'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   

Dim lgPageNo_A                                              '☜: Next Key tag                          
Dim lgSortKey_A                                             '☜: Sort상태 저장변수                      

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   

Dim lgPageNo_B                                              '☜: Next Key tag                          
Dim lgSortKey_B                                             '☜: Sort상태 저장변수                      

'☜:--------Spreadsheet temp---------------------------------------------------------------------------   
                                                             '☜:--------Buffer for Spreadsheet -----   
'Dim lgKeyPos                                                '☜: Key위치                               
'Dim lgKeyPosVal                                             '☜: Key위치 Value                         
Dim lsIntFlgMode



 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgIntFlgMode    = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lsIntFlgMode    = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	
    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1
    
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
	Dim StartDate

	StartDate     = "<%=GetSvrDate%>"                                                                  'Get DB Server Date

	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call loadInfTB19029A("Q", "C", "NOCOOKIE", "QA")%>
End Sub

 '******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("C3601MA101","G","A","V20021213",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock("A") 
	Call SetZAdoSpreadSheet("C3601MA101","G","B","V20021213",parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X","X")
    Call SetSpreadLock("B")
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(iOpt)
	If iOpt = "A" Then
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
    Else
      ggoSpread.Source = frm1.vspdData2
      ggoSpread.SpreadLockWithOddEvenRowColor()
    End If   
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock1
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock1()

	ggoSpread.Source = frm1.vspdData2
    With frm1.vspdData2
		.ReDraw = False
		ggoSpread.SpreadLock C_RecvCostCd, -1, C_RecvCostCd		
		ggoSpread.SpreadLock C_RecvCostNm, -1, C_RecvCostNm
		ggoSpread.SpreadLock C_DstbFct, -1, C_DstbFct			
		ggoSpread.SpreadLock C_DstbAmt, -1, C_DstbAmt		
		ggoSpread.SpreadLock C_RecvAmt,-1 , C_RecvAmt		
		ggoSpread.SpreadLock C_DstbRate,-1 , C_DstbRate		
		.ReDraw = True
    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

		.vspdData.ReDraw = False

		' 필수 입력 항목으로 설정 
		' SSSetRequired(ByVal Col, ByVal Row, Optional ByVal Row2 = -10)
    
		.vspdData.ReDraw = True


		.vspdData2.ReDraw = False

		' 필수 입력 항목으로 설정 
		' SSSetRequired(ByVal Col, ByVal Row, Optional ByVal Row2 = -10)
    
		.vspdData2.ReDraw = True
    End With
End Sub


'======================================================================================================
'	Name : OpenCostCd()
'	Description : Cost Center PopUp
'=======================================================================================================
Function OpenCostCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop= True

	arrParam(0) = "코스트센타팝업"			'팝업 명칭 
	arrParam(1) = "B_COST_CENTER"						'TABLE 명칭 
	arrParam(2) = strCode						'Code Condition
	arrParam(3) = ""							'Name Condition
	arrParam(4) = ""							'Where Condition
	arrParam(5) = "코스트센타"			
	
    arrField(0) = "COST_CD"					    'Field명(0)
    arrField(1) = "COST_NM"					    'Field명(1)
    
    arrHeader(0) = "코스트센타코드"					'Header명(0)
    arrHeader(1) = "코스트센타명"					'Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop= False
	
	If arrRet(0) = "" Then
		frm1.txtCostCd.focus
		Exit Function
	Else
		Call SetCostCd(arrRet, iWhere)
	End If	

End Function

'======================================================================================================
'	Name : SetCostCd()
'	Description : Cost Center Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCostCd(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 0 Then
    		.txtCostCd.focus
    		.txtCostCd.value = arrRet(0)
    		.txtCostNm.value = arrRet(1)
    	Else
    		.vspdData.Col = C_CostCd
    		.vspdData.Text = arrRet(0)
    		.vspdData.Col = C_CostNm
    		.vspdData.Text = arrRet(1)
            
    		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		        '변경이 읽어났다고 알려줌 
    	End If
	
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


 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
 '******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
       
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    
    Call ggoOper.ClearField(Document, "1")										'⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field

	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()

    Call SetToolbar("11000000000011")								        '⊙: 버튼 툴바 제어 
    frm1.txtYyyyMm.focus 
    frm1.txtOriginTotAmt.allownull = False
    frm1.txtAllocAmtSum.allownull = False
    frm1.txtAmtSum.allownull = False
    '--------- Developer Coding Part (End  ) ----------------------------------------------------------
    Set gActiveElement = document.activeElement 
    
End Sub

'========================================================================================================
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'========================================================================================================
Function OpenGroupPopup()

	Dim arrRet
'	Dim arrParam
'	Dim TInf(5)
'	Dim ii
	
	On Error Resume Next
	
'	ReDim arrParam(Parent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
'    TInf(0) = Parent.gMethodText
  
'	For ii = 0 to Parent.C_MaxSelList * 2 - 1 Step 2
'      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
'      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
'    Next  
      
  
	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	

	lgIsOpenPop = False
	
	
	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData("A", arrRet(0), arrRet(1))
	
'	   For ii = 0 to Parent.C_MaxSelList * 2 - 1 Step 2
'           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
'           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
'       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
   
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
'   Event Name : txtYyyymm_KeyDown
'   Event Desc :
'==========================================================================================

Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

 '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

 '#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************

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
'======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	gMouseClickStatus = "SPC"	'Split 상태코드    

    If Row <> NewRow And NewRow > 0 Then
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
		Call SetSpreadColumnValue("A", frm1.vspdData, NewCol, NewRow)
	    Call DbQuery("M1Q")
	         
		ggoSpread.Source = frm1.vspdData2 
		ggoSpread.ClearSpreadData

	
		lgPageNo_B       = ""                                  'initializes Previous Key
		lgSortKey_B      = 1

	End If    
	    

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click( ByVal Col, ByVal Row)
'    Dim ii

	Call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row = 0 Then
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
    
'	 For ii = 1 to UBound(lgKeyPos)
'        frm1.vspdData.Col = lgKeyPos(ii)
'        frm1.vspdData.Row = Row
'        lgKeyPosVal(ii)   = frm1.vspdData.text        
'	 Next
	 
     ggoSpread.Source = frm1.vspdData2 
     ggoSpread.ClearSpreadData

     lgPageNo_B       = ""                                  'initializes Previous Key
     lgSortKey_B      = 1

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
    Call DbQuery("M1Q")
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click( ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001") 
	gMouseClickStatus = "SP2C"
	 Set gActiveSpdSheet = frm1.vspdData2
	 
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    Call SetSpreadColumnValue("B", frm1.vspdData2, Col, Row)
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
		End If
   End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery("M1N") = False Then
              Call RestoreToolBar()
              Exit Sub
          End if
		End If
   End if
    
End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
       
    FncQuery = False                                                        '⊙: Processing is NG
    Err.Clear     

    '-----------------------
    'Erase contents area
    '-----------------------
    
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
    
    Call InitVariables 														'⊙: Initializes local global variables
    Call InitTxtAmtSumClear
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								        '⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery("MQ") = False Then   
       Exit Function           
    End If     							

    FncQuery = True		
    
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)										
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
	Call parent.FncFind(Parent.C_MULTI,False)
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
    FncExit = True
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery(pDirect) 
	Dim strVal
	Dim strYear, strMonth, strDay
	
	Err.Clear  '☜: Protect system from crashing
        
    DbQuery = False
                 
	Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)  
	
	'--------- Developer Coding Part (Start) ----------------------------------------------------------
	With Frm1
	
	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
    Select Case pDirect
        Case "MQ","MN"
                            
			
                strVal = BIZ_PGM_ID & "?txtMode="			& Parent.UID_M0001						'Hidden의 검색조건으로 Query
				strVal = strVal		& "&txtYyyymm="			& strYear & strMonth
				strVal = strVal		& "&txtCostCd="			& .txtCostCd.value				
				strVal = strVal     & "&txtMaxRows="        & .vspdData.MaxRows 

'--------- Developer Coding Part (End) ----------------------------------------------------------
                strVal = strVal      & "&lgPageNo="          & lgPageNo_A                          '☜: Next key tag
                strVal = strVal      & "&lgSelectListDT="    & GetSQLSelectListDataType("A")
                strVal = strVal      & "&lgTailList="        & MakeSQLGroupOrderByList("A")
                strVal = strVal      & "&lgSelectList="      & EnCoding(GetSQLSelectList("A"))
'               strVal = strVal      & "&lgMaxCount="        & CStr(C_SHEETMAXROWS_D_A)            '☜: 한번에 가져올수 있는 데이타 건수 
            
			
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        Case "M1Q","M1N"
			
			Call InitTxtAmtSumClear
						
				strVal = BIZ_PGM_ID2 & "?txtMode="			& Parent.UID_M0001						'Hidden의 검색조건으로 Query				
    		    strVal = strVal		 & "&txtYyyyMm="		& strYear & strMonth
				strVal = strVal		 & "&txtCostCd="		& GetKeyPosVal("A",1)
				strVal = strVal		 & "&txtDiflag="		& GetKeyPosVal("A",2)
				strVal = strVal		 & "&txtAcctCd="		& GetKeyPosVal("A",3)				
				strVal = strVal      & "&txtMaxRows="       & .vspdData.MaxRows  

'--------- Developer Coding Part (End) ----------------------------------------------------------
                strVal = strVal      & "&lgPageNo="          & lgPageNo_B                          '☜: Next key tag
                strVal = strVal      & "&lgSelectListDT="    & GetSQLSelectListDataType("B")
                strVal = strVal      & "&lgTailList="        & MakeSQLGroupOrderByList("B")
                strVal = strVal      & "&lgSelectList="      & EnCoding(GetSQLSelectList("B"))
'               strVal = strVal      & "&lgMaxCount="        & CStr(C_SHEETMAXROWS_D_B)  '☜: 한번에 가져올수 있는 데이타 건수 
            
    End Select		
    
    'msgbox strval
	End with
    
    
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic    
	
    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(iOpt)							'☆: 조회 성공후 실행로직 
    
	If iOpt = 1 Then  	
		lgIntFlgMode     = Parent.OPMD_UMODE									'⊙: Indicates that current mode is Update mode           
       frm1.vspdData.focus
       Call vspdData_Click(1,1) 
    Else 
		lsIntFlgMode     = Parent.OPMD_UMODE									'⊙: Indicates that current mode is Update mode           
	End If							                                     '⊙: This function lock the suitable field

	Call ggoOper.LockField(Document, "Q")								 '⊙: This function lock the suitable field 
    
End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	Dim gPos
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       Case "VSPDDATA2"                  
	            gPos = "B"
    End Select     
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(gPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	

	If arrRet(0) = "X" Then
	   Exit Function
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "3")	   
	   Exit Function        
	Else
	   Call ggoSpread.SaveXMLData(gPos,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


'========================================================================================================
' Function Name : InitTxtAmtSumClear()
' Function Desc : 
'========================================================================================================

Sub InitTxtAmtSumClear()

	frm1.txtAmtSum.Value = "0"

end Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>C/C별배부내역조회</font></td>
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
									<TD CLASS="TD6"><script language =javascript src='./js/c3601ma1_fpDateTime1_txtYyyymm.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>코스트센타</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="코스트센타"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCostCd frm1.txtCostCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=30 tag="14"></TD>								
									</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=45% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/c3601ma1_vspdData_vspdData.js'></script>
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
									<TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD5 NOWRAP>합계</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3601ma1_fpDoubleSingle1_txtOriginTotAmt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
        </TR>  
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/c3601ma1_vspdData2_vspdData2.js'></script>
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
								<TD CLASS=TD5>총배부금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/c3601ma1_fpDoubleSingle2_txtAllocAmtSum.js'></script>&nbsp;
    							<TD CLASS=TD5>합계</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/c3601ma1_fpDoubleSingle2_txtAmtSum.js'></script>&nbsp;
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="2" TABINDEX = "-1" ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="2" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="2" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="2" TABINDEX = "-1" >

<INPUT TYPE=HIDDEN NAME="hYyyymm" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hHeadRows" tag="24" TABINDEX = "-1" >

<INPUT TYPE=HIDDEN NAME="txtGiveCostCd" tag="2" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtAcctCd" tag="2" TABINDEX = "-1" >

</FORM>
</BODY>
</HTML>

