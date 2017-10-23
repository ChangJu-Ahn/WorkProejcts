
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1504ma1.asp
'*  4. Program Name         : Shift Exception
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : RYU SUNG WON
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID  = "p1504mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "p1504mb2.asp"											'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const C_SHEETMAXROWS = 30

Dim C_ExceptionCd
Dim C_ExceptionNm
Dim C_ExceptionWorkFlgNm
Dim C_ExceptionStartDt
Dim C_ExceptionStartTime
Dim C_ExceptionEndDt
Dim C_ExceptionEndTime
Dim C_ExceptionType
Dim C_ExceptionWorkFlg


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
Dim iDBSYSDate
Dim StartDate, EndDate
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_ExceptionCd			= 1
	C_ExceptionNm			= 2
	C_ExceptionWorkFlgNm	= 3
	C_ExceptionStartDt		= 4
	C_ExceptionStartTime	= 5
	C_ExceptionEndDt		= 6
	C_ExceptionEndTime		= 7
	C_ExceptionType			= 8
	C_ExceptionWorkFlg		= 9
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
		
End Sub

Sub SetReadCookie
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value	= ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value	= ReadCookie("txtPlantNm")
		frm1.txtResourceCd.Value= ReadCookie("txtResourceCd")
		frm1.txtResourceNm.value= ReadCookie("txtResourceNm")
		frm1.txtShiftCd.Value	= ReadCookie("txtShiftCd")
		frm1.txtShiftNm.value	= ReadCookie("txtShiftNm")		
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtResourceCd", ""
	WriteCookie "txtResourceNm", ""
	WriteCookie "txtShiftCd", ""
	WriteCookie "txtShiftNm", ""	
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
	With frm1.vspdData
	ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    
    .MaxCols = C_ExceptionWorkFlg+1    
    .MaxRows = 0    

	.ReDraw = false
	
	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit 	C_ExceptionCd,			"예외코드",			13,,,10,2
	ggoSpread.SSSetEdit 	C_ExceptionNm,			"예외내용",			34,,,40
	ggoSpread.SSSetCombo	C_ExceptionWorkFlgNm,	"잔업/Break 구분",	20
	ggoSpread.SSSetDate 	C_ExceptionStartDt,		"시작일",			12, 2, parent.gDateFormat
	ggoSpread.SSSetTime 	C_ExceptionStartTime,	"시작시각",			12, 2, 1 ,1	    
	ggoSpread.SSSetDate 	C_ExceptionEndDt,		"종료일",			12, 2, parent.gDateFormat
	ggoSpread.SSSetTime 	C_ExceptionEndTime,		"종료시각",			12, 2, 1 ,1	   		
	ggoSpread.SSSetCombo	C_ExceptionType,		"예외타입",			12
	ggoSpread.SSSetCombo	C_ExceptionWorkFlg,		"잔업/Break 구분",	15

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(.MaxCols - 1, .MaxCols - 1, True)
	Call ggoSpread.SSSetColHidden(.MaxCols - 2, .MaxCols - 2, True)

	ggoSpread.SSSetSplit2(1)										'frozen 기능추가 
	.ReDraw = true

	Call SetSpreadLock 

    End With

End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock		C_ExceptionCd, -1, C_ExceptionCd
		ggoSpread.SSSetRequired 	C_ExceptionWorkFlgNm, -1	
		ggoSpread.SSSetRequired		C_ExceptionStartDt, -1
		ggoSpread.SSSetRequired		C_ExceptionEndDt, -1
		ggoSpread.SSSetRequired		C_ExceptionStartTime, -1
		ggoSpread.SSSetRequired		C_ExceptionEndTime, -1
		ggoSpread.SSSetProtected	.vspdData.MaxCols, -1
		.vspdData.ReDraw = True
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired 	C_ExceptionCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_ExceptionWorkFlgNm,	pvStartRow, pvEndRow		
		ggoSpread.SSSetRequired 	C_ExceptionStartDt,		pvStartRow, pvEndRow	
		ggoSpread.SSSetRequired 	C_ExceptionEndDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_ExceptionStartTime,	pvStartRow, pvEndRow	
		ggoSpread.SSSetRequired 	C_ExceptionEndTime,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	.vspdData.MaxCols,		pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ExceptionCd			= iCurColumnPos(1)
			C_ExceptionNm			= iCurColumnPos(2)
			C_ExceptionWorkFlgNm	= iCurColumnPos(3)
			C_ExceptionStartDt		= iCurColumnPos(4)
			C_ExceptionStartTime	= iCurColumnPos(5)
			C_ExceptionEndDt		= iCurColumnPos(6)
			C_ExceptionEndTime		= iCurColumnPos(7)
			C_ExceptionType			= iCurColumnPos(8)
			C_ExceptionWorkFlg		= iCurColumnPos(9)
    End Select    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim strCboCd
	Dim strCboNm
	
	strCboCd = ""
	strCboCd = "Y" & vbTab & "N"
	strCboNm = ""
	strCboNm = "OverTime" & vbTab & "Downtime"

	ggoSpread.SetCombo	strCboCd,	C_ExceptionWorkFlg
	ggoSpread.SetCombo	strCboNm,	C_ExceptionWorkFlgNm
	
	strCboCd = "R" & vbTab 
	    
    ggoSpread.SetCombo	strCboCd,	C_ExceptionType
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function
'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)


	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "자원팝업"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(frm1.txtResourceCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "			
	arrParam(5) = "자원"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "자원"		
    arrHeader(1) = "자원명"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtResourceCd.focus
		
End Function

'------------------------------------------  OpenShift()  -------------------------------------------------
'	Name : OpenShift()
'	Description : ShiftPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenShift()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	elseif frm1.txtResourceCd.value = "" Then
		Call DisplayMsgBox("971012","X", "자원","X")
		frm1.txtResourceCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True	
	
	arrParam(0) = "Shift팝업"	
	arrParam(1) = "P_SHIFT_HEADER,P_RESOURCE_ON_SHIFT"				
	arrParam(2) = Trim(frm1.txtShiftCd.Value)
	arrParam(3) = ""
	arrParam(4) = "P_RESOURCE_ON_SHIFT.RSC_PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " & _
				  " AND P_RESOURCE_ON_SHIFT.SHIFT_PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " & _
				  " AND P_RESOURCE_ON_SHIFT.RESOURCE_CD =  " & FilterVar(frm1.txtResourceCd.value, "''", "S") & " " & _
				  " AND P_SHIFT_HEADER.PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " & _
				  " AND P_RESOURCE_ON_SHIFT.SHIFT_CD = P_SHIFT_HEADER.SHIFT_CD " 			  
	
	arrParam(5) = "Shift명"			
	
    arrField(0) = "P_RESOURCE_ON_SHIFT.SHIFT_CD"	
    arrField(1) = "P_SHIFT_HEADER.DESCRIPTION"	
       
    
    arrHeader(0) = "Shift"		
    arrHeader(1) = "Shift명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetShift(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtShiftCd.focus
	
End Function
'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	frm1.txtResourceCd.Value    = arrRet(0)		
	frm1.txtResourceNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetShift()  --------------------------------------------------
'	Name : SetShift()
'	Description : Shift Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetShift(Byval arrRet)
	frm1.txtShiftCd.Value    = arrRet(0)		
	frm1.txtShiftNm.Value    = arrRet(1)		
End Function
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 


Sub InitData(ByVal lngStartRow, ByVal iPos)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		.ReDraw = False
		
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.Col = C_ExceptionWorkFlg
			intIndex = .value
			.col = C_ExceptionWorkFlgNm
			.value = intindex
		Next	
		.ReDraw = True
	End With
End Sub

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
	Err.Clear
	
	iDBSYSDate = "<%=GetSvrDate%>"											'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	StartDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                            '⊙: Lock  Suitable  Field    
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    
    Call InitComboBox
    Call SetToolbar("11001101000011")								'⊙: 버튼 툴바 제어 
    Call SetReadCookie   	
    	
    If Trim(frm1.txtPlantCd.value) = "" Then	
		If parent.gPlant <> "" then	
			frm1.txtPlantCd.value = parent.gPlant	
			frm1.txtPlantNm.value = parent.gPlantNm	
			frm1.txtResourceCd.focus 	
			Set gActiveElement = document.activeElement 
		Else	
			frm1.txtPlantCd.focus 	
			Set gActiveElement = document.activeElement 
		End If	
	End If 	
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub



'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspddata_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("1101111111")    
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
       
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
    End If
    
	If Row <= 0 Or Col < 0 Then
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Name : vspddata_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspddata_MouseDown(Button,Shift,x,y)
	If Button = "2" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

End Sub

'==========================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================
Sub vspdData_DblClick(ByVal Col , ByVal Row )       
    If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspddata_KeyPress(index , KeyAscii)
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
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
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row
		Select Case Col
			Case  C_ExceptionWorkFlgNm
				.Col = Col
				intIndex = .Value
				.Col = C_ExceptionWorkFlg
				.Value = intIndex
		End Select
    End With
End Sub

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

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    FncQuery = False															'⊙: Processing is NG
    Err.Clear																    '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtResourceCd.value = "" Then
		frm1.txtResourceNm.value = ""
	End If
	
	If frm1.txtShiftCd.value = "" Then
		frm1.txtShiftNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables
  
    If Not chkField(Document, "1") Then	Exit Function					'⊙: This function check indispensable field
    If DbQuery = False Then				Exit Function

    FncQuery = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False																'⊙: Processing is NG
    
    On Error Resume Next														'☜: Protect system from crashing
    Err.Clear																	'☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'⊙: No data changed!!
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then	Exit Function
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		If Not chkField(Document, "1") Then	Exit Function
    End IF

    If DbSave = False Then	Exit Function
    
    FncSave = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear
    
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
	FncCopy = False                                                               '☜: Processing is NG
	ggoSpread.Source = frm1.vspdData	
	
	With frm1.vspdData
		.ReDraw = False
		If .ActiveRow > 0 Then
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
			
			.EditMode = True
			.ReDraw = True
			.Focus
		End If
	End With
	
	If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo															'☜: Protect system from crashing
    Call InitData(1,1)
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow
	Dim newRow
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False
    
    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End If
    
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
		'--------------------------------------------
		' Default Setting
		' 추가할 Row갯수만큼 컬럼을 초기화 시킨다.
		'--------------------------------------------
		For newRow = 0 To Cint(imRow) - 1
			.vspdData.Col = C_ExceptionType
			.vspdData.Row = .vspdData.ActiveRow + newRow
			.vspdData.Text = "R"
    
			.vspdData.Col = C_ExceptionWorkFlg
			.vspdData.Row = .vspdData.ActiveRow + newRow
			.vspdData.Text = "Y"
    
			.vspdData.Col = C_ExceptionStartDt
			.vspdData.Row = .vspdData.ActiveRow + newRow 
			.vspdData.Text = startdate
    
			.vspdData.Col = C_ExceptionEndDt
			.vspdData.Row = .vspdData.ActiveRow  + newRow
			.vspdData.Text = startdate
		Next

		.vspdData.EditMode = True
		.vspdData.ReDraw = True
    End With    
	
	Set gActiveElement = document.activeElement 
	    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If 
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
   
    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
    
    With frm1.vspdData 
		.focus
		Set gActiveElement = document.activeElement 
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
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()	
	Call InitData(1,1)
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
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
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
	Dim strVal
	    
    DbQuery = False
    LayerShowHide(1) 
    Err.Clear                                                               '☜: Protect system from crashing
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtResourceCd=" & Trim(.hResourceCd.value)
			strVal = strVal & "&txtShiftCd=" & Trim(.hShiftCd.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtResourceCd=" & Trim(.txtResourceCd.value)
			strVal = strVal & "&txtShiftCd=" & Trim(.txtShiftCd.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call InitData(LngMaxRow,1)
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call SetToolbar("11001111001111")
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    
    Dim lRow        
    Dim lGrpCnt     
   	Dim strVal
	Dim strDel
	
    DbSave = False                                                          '⊙: Processing is NG
    
    LayerShowHide(1) 
		
    On Error Resume Next                                                   '☜: Protect system from crashing
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
    
		lGrpCnt = 1
    
		strVal = ""
		strDel = ""
    
		For lRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag												'☜: 신규 
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep					'☜: C=Create
					
		            .vspdData.Col = C_ExceptionCd	'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionNm	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_ExceptionWorkFlg	'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionStartDt	'4
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionStartTime	'5
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionEndDt	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_ExceptionEndTime	'7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_ExceptionType	'8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		                            
		            lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep					'☜: U=Update
					
		            .vspdData.Col = C_ExceptionCd	'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionNm	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionWorkFlg	'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionStartDt	'4
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionStartTime	'5
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionEndDt	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionEndTime	'7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_ExceptionType	'8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		                                                                           
		            lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag												'☜: 삭제 
					strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep
					
		            .vspdData.Col = C_ExceptionCd	'1
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		                            
		            lGrpCnt = lGrpCnt + 1
		    End Select
		            
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
	
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'☜: 비지니스 ASP 를 가동 
	End With
	
    DbSave = True																	'⊙: Processing is NG
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()																	'☆: 저장 성공후 실행 로직 
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
	
    Call MainQuery()
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 

End Function
Function DbDeleteOk()

End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Shift예외등록</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14" ALT="공장명"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>자원</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=15 MAXLENGTH=10 tag="12XXXU" ALT="자원"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=30 tag="14" ALT="자원명"></TD>								
									<TD CLASS=TD5 NOWRAP>Shift</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShiftCd" SIZE=8 MAXLENGTH=2 tag="12XXXU" ALT="Shift"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShiftCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenShift()">&nbsp;<INPUT TYPE=TEXT NAME="txtShiftNm" SIZE=25 tag="14" ALT="Shift명"></TD>
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
							<TD HEIGHT=100%>
								<script language =javascript src='./js/p1504ma1_I212316697_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hResourceCd" tag="24"><INPUT TYPE=HIDDEN NAME="hShiftCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

