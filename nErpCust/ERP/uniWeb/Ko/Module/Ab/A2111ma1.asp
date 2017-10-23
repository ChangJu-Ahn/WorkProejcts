<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : PrePayment management
'*  3. Program ID           : a2111ma1.asp
'*  4. Program Name         : 전표관리항목조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003.08.28
'*  8. Modified date(Last)  : 2003.08.28
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2001.01.13
'**********************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'############################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<Script Language="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID        = "a2111mb1.asp"							'☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "a2111mb2.asp"							'☆: Biz logic spread sheet for #2
Const BIZ_PGM_SAVE_ID   = "a2111mb3.asp"							'☆: Biz logic For Update Row Data

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey            = 0										'☆☆☆☆: Max key value
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  lgIsOpenPop													'☜: Popup status                           
Dim  lgKeyPosVal
Dim  IsOpenPop														'☜: Popup status   
Dim  lgPageNo_A														'☜: Next Key tag                          
Dim  lgSortKey_A													'☜: Sort상태 저장변수                     
Dim  lgPageNo_B														'☜: Next Key tag                          
Dim  lgSortKey_B													'☜: Sort상태 저장변수 
Dim  lgFncQuery

Dim  C_GL_CTRL_FLD 
Dim  C_GL_CTRL_NM  

'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------
 '#########################################################################################################
'												2. Function부 
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub  InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgIntFlgMode     = parent.OPMD_CMODE                   'Indicates that current mode is Create mode

    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B		 = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1
End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'========================================================================================================= 
Sub  SetDefaultVal()

End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub  LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("I", "A", "NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "A", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	C_GL_CTRL_FLD = 1
	C_GL_CTRL_NM  = 2
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()

    With frm1.vspdData
		.MaxCols	= C_GL_CTRL_NM + 1
		.Col		= .MaxCols
		.ColHidden	= True

		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread

		Call GetSpreadColumnPos()

		ggoSpread.SSSetEdit	C_GL_CTRL_FLD ,"전표관리항목"  , 30,,,30,2
		ggoSpread.SSSetEdit	C_GL_CTRL_NM  ,"전표관리항목명", 50

		Call ggoSpread.MakePairsColumn(C_GL_CTRL_FLD,C_GL_CTRL_NM,"1")
		
		.ReDraw = True
		Call SetSpreadLock_A()
    End With

    Call SetZAdoSpreadSheet("A2111MA1","S","A","V20021211",Parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X","X")
	Call SetSpreadLock_B()																		
End Sub

'=========================================================================================================
' Function Name : SetSpreadLock_A
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadLock_A()
	ggoSpread.SpreadLock C_GL_CTRL_FLD	, -1, C_GL_CTRL_FLD
	ggoSpread.SpreadLock C_GL_CTRL_NM	, -1, C_GL_CTRL_NM
End Sub

'=========================================================================================================
' Function Name : SetSpreadLock_B
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadLock_B()
	With frm1.vspdData2
		.ReDraw = False       
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()	
		.ReDraw = True
	End With 
End Sub

'=========================================================================================================
' Function Name : SetSpreadColor_A
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadColor_A()
	Dim ii 
	
	With frm1.vspddata
		ggoSpread.Source = frm1.vspddata
		For ii = 1 To .MaxRows
			.row = ii
			.col = C_GL_CTRL_FLD
			If UCase(Left(Trim(.value),7)) = "USER_DF" Then
				ggoSpread.SpreadUnLock C_GL_CTRL_NM	, ii, C_GL_CTRL_NM ,ii
			Else	
				ggoSpread.SpreadLock   C_GL_CTRL_NM	, ii, C_GL_CTRL_NM ,ii
			End If
		Next
	End With					
End Sub

'=========================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'=========================================================================================================
Sub GetSpreadColumnPos()
    Dim iCurColumnPos

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_GL_CTRL_FLD = iCurColumnPos(1)
	C_GL_CTRL_NM  = iCurColumnPos(2)
End Sub


'**********************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** 

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenGlCtrlPopUp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenGlCtrlPopUp(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function	
	
	lgIsOpenPop = True
	
	arrParam(0) = "전표관리항목팝업"								' 팝업 명칭 
	arrParam(1) = "A_SUBLEDGER_CTRL " 									' TABLE 명칭 
	arrParam(2) = Trim(strCode)											' Code Condition
	arrParam(3) = ""													' Name Condition
	arrParam(4) = ""													' Where Condition
	arrParam(5) = "전표관리항목"									' 조건필드의 라벨 명칭 

	arrField(0) = "GL_CTRL_FLD"											' Field명(0)
	arrField(1) = "ISNULL(GL_CTRL_NM,'')"								' Field명(1)
	arrField(2) = ""													' Field명(2)
	arrField(3) = ""													' Field명(3)
			
	arrHeader(0) = "전표관리항목"									' Header명(0)
	arrHeader(1) = "전표관리항목명"									' Header명(1)
	arrHeader(2) = ""													' Header명(2)
	arrHeader(3) = ""													' Header명(3)

	lgIsOpenPop = True
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then	    
		frm1.txtGlCtrlFld.Focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,1)
	End If
End Function
			
'=======================================================================================================
'	Name : SetBankAcct()
'	Description : Bank Account No Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		Case 1
			frm1.txtGlCtrlFld.value = arrRet(0)
			frm1.txtGlCtrlNm.value = arrRet(1)
			frm1.txtGlCtrlFld.focus				
	End Select
	
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : 
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	Dim iGridPos
	
	iGridPos = "B"
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(iGridPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData(iGridPos,arrRet(0),arrRet(1))
		Call InitVariables()
		Call InitSpreadSheet()       
   End If
End Function

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
Sub  Form_Load()
	Call LoadInfTB19029()			
	Call InitVariables()																	'⊙: Initializes local global variables
	Call SetDefaultVal()
	Call InitSpreadSheet()
    Call SetToolbar("1100100000011111")														'⊙: 버튼 툴바 제어 
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")

    gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData        
    
    If Row <= 0 Then
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
	
	Call DbQuery("2",Row)
    
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
    lgPageNo_B       = ""                                  'initializes Previous Key
    lgSortKey_B      = 1
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    gMouseClickStatus = "SPC"	'Split 상태코드    
	Set gActiveSpdSheet = frm1.vspdData        

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
		    Exit Sub
	    End If
	    
		Call DbQuery("2",NewRow)
     
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	
		lgPageNo_B       = ""                                  'initializes Previous Key
		lgSortKey_B      = 1
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")
	
    gMouseClickStatus = "SP2C"	'Split 상태코드 
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
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then													'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DbQuery("1","")
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then													'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DbQuery("2","")
		End If
   End if
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

' #########################################################################################################
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
Function  FncQuery()
    FncQuery = False															'⊙: Processing is NG
    Err.Clear     

	lgFncQuery = True
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then											'⊙: This function check indispensable field
		Exit Function
    End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData    
	    
    Call InitVariables() 														'⊙: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery("1","")														'☜: Query db data

    FncQuery = True		
	
	Set gActiveElement = document.activeElement
	lgFncQuery = False
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function  FncPrint() 
    Call parent.FncPrint()
    	
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
    	
	Set gActiveElement = document.activeElement    
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
Function  FncExit()
    FncExit = True
End Function

'=======================================================================================================
' Function Name : `
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1,var2
 
    FncSave = False                                                         
    
    On Error Resume Next
    Err.Clear                                                               
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    If var1 = False  Then											'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")				'⊙: Display Message(There is no changed data.)
		Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()													'☜: Save db data
    FncSave = True  

    Set gActiveElement = document.activeElement
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    Dim lngRows 
    Dim lGrpcnt
    DIM strVal 

    DbSave = False                                                          
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   
	Err.Clear 
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 
    lGrpCnt = 1
    strVal = ""
    
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			If .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & parent.gColSep & lngRows & parent.gColSep
			    .Col = C_GL_CTRL_FLD '1
			    strVal = strVal & Trim(.Text) & parent.gColSep
			    .Col = C_GL_CTRL_NM  '2
			    strVal = strVal & Trim(.Text) & parent.gRowSep
			          
			    lGrpCnt = lGrpCnt + 1          
			End If
		Next
	End With
 
	frm1.txtSpread.value =  strVal								'Spread Sheet 내용을 저장 
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'저장 비지니스 ASP 를 가동 
    
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk()											'☆: 저장 성공후 실행 로직 
	ggoSpread.Source = frm1.vspdData        				
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2        				
								
	Call DBquery(1,"")
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 
'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function  DbQuery(ByVal iOpt,ByVal Row) 
	Dim strVal
	Dim strCode
	Dim iRow
	
    Err.Clear																						'☜: Protect system from crashing
	On Error Resume Next
	
	If Row = "" Then 
		iRow = frm1.vspddata.ActiveRow
	Else
		iRow = Row		
	End If	
	
    DbQuery = False
    Call DisableToolBar(parent.TBC_QUERY)															'☜: Disable Query Button Of ToolBar
	Call LayerShowHide(1)
    
    With frm1
		Select Case iOpt 
			Case "1" 
				strVal = BIZ_PGM_ID & "?txtGlCtrlFld=" & Trim(.txtGlCtrlFld.value)
				strVal = strVal & "&txtGlCtrlFld_ALT=" & .txtGlCtrlFld.alt
			Case "2"
				.vspddata.row = iRow
				.vspddata.col = C_GL_CTRL_FLD
				strCode = .vspddata.value

				strVal = BIZ_PGM_ID1 & "?txtGlCtrlFld=" & strCode
				strVal = strVal & "&lgPageNo="        & lgPageNo									'☜: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList="      & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="    & EnCoding(GetSQLSelectList("A"))
		End Select 
      
		Call RunMyBizASP(MyBizASP, strVal)															'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'==================================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'==================================================================================================================
Function DbQueryOk(ByVal iOpt)																		'☆: 조회 성공후 실행로직 
    lgIntFlgMode = parent.OPMD_UMODE																'⊙: Indicates that current mode is Update mode
    
	If iOpt = 1 Then

       Call vspdData_Click(1,1)
       frm1.vspdData.focus
	End If																							'⊙: This function lock the suitable field

	Call ggoOper.LockField(Document, "Q")															'⊙: This function lock the suitable field 
	Call SetSpreadColor_A()
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>전표관리항목수정</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*" align=right>&nbsp;</td>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>전표관리항목</TD>
									<TD CLASS="TD6" COLSPAN=3 NOWRAP><INPUT TYPE=TEXT NAME="txtGlCtrlFld" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: Left" Tag="11XXXU" ALT="전표관리항목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGlCtrlFld" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenGlCtrlPopUp(frm1.txtGlCtrlFld.Value)">&nbsp;<INPUT TYPE=TEXT NAME="txtGlCtrlNm" SIZE=30 tag="14"></TD>
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
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=6>
								<script language =javascript src='./js/a2111ma1_I548417818_vspdData.js'></script></TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" colspan=6>
								<script language =javascript src='./js/a2111ma1_I369390183_vspdData2.js'></script></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA	CLASS=HIDDEN NAME=txtSpread	tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>

