<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--*******************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Change
'*  3. Program ID           : a7109ma1
'*  4. Program Name         : 고정자산부서이동등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormatF
'                             AS0031
'                             AS0039
'*  7. Modified date(First) : 2000/03/18
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--########################################################################################################
'												1. 선 언 부 
'   ##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* 

'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- 해당 위치에 따라 달라짐, 상대 경로  -->

<!--========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->							<!--처리ASP에서 서버작업이 필요한 경우  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit    												'☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>
											'비지니스 로직 ASP명 
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"			'환율정보 비지니스 로직 ASP명 


'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID    = "a7109mb1.asp"  
Const BIZ_PGM_ID2   = "a7109mb2.asp"  

Dim C_DeptCd
Dim C_DeptNm
Dim C_InvRate
Dim C_CostCenterNm
Dim C_BizAreaNm

Const C_SHEETMAXROWS = 10

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
'Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
'Dim lgIntGrpCount				    '☜: Group View Size를 조사할 변수 
'Dim lgIntFlgMode					'☜: Variable is for Operation Status
'Dim lgStrPrevKey

 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
Sub initSpreadPosVariables()
	 C_DeptCd		 = 1
	 C_DeptNm		 = 2
	 C_InvRate		 = 3
	 C_CostCenterNm = 4
	 C_BizAreaNm	 = 5
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'☆: 사용자 변수 초기화 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
    
	frm1.txthToOrgChangeId.value =parent.gChangeOrgId
	frm1.txthFrOrgChangeId.value =parent.gChangeOrgId
    
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
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

<%
Dim svrDate
svrDate = GetSvrDate
%>

	frm1.txtChgDt.text    = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	
	
End Sub

 '******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
Function OpenMasterRef()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If lgIntFlgMode = parent.OPMD_UMODE Then 
			Call DisplayMsgBox("200005", "X", "X", "X")
			Exit function
	End If	
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName & "?PID=" & gstrRequestMenuID, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPoRef(arrRet)
	End If	
	

End Function

 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(strRet)
	
    Dim strVal
    
    'lgMasterQueryFg = False
    
	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

	Call ggoOper.ClearField(Document, "1")								
	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call SetDefaultVal

    Call InitVariables
    
	frm1.txtAsstNo.value     = strRet(0)
	frm1.txtAsstNm.value	 = strRet(1)

	Call Dbquery_master("R")
		'lgMasterQueryFg = True
	
	lgBlnFlgChgValue = False
    lgIntFlgMode = parent.OPMD_CMODE
End Sub

'======================================================================================================
'   Function Name : OpenChgNoInfo()
'   Function Desc : 
'=======================================================================================================
Function OpenChgNoInfo()


	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a7107ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7107ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName & "?PID=" & gstrRequestMenuID , Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetChgNoInfo(arrRet)
	End If	


	frm1.txtChgNo.focus 
	
End Function

'======================================================================================================
'   Function Name : SetChgNoInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetChgNoInfo(Byval arrRet)
	frm1.txtChgNo.value  = arrRet(0)				
End Function

'======================================================================================================
'   Function Name : OpenDeptTO( )
'   Function Desc : 
'=======================================================================================================
Function OpenDeptTO()
	Dim arrRet
	Dim arrParam(8)
	Dim IntRetCd
	Dim  field_fg
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtToDeptCd.value		            '  Code Condition
   	arrParam(1) = frm1.txtChgDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "txtToDeptCd"
		Call SetPopUpReturn(arrRet,field_fg)
	End If	


END FUNCTION
'======================================================================================================
'   Function Name :OpenDeptFR( )
'   Function Desc : 
'=======================================================================================================
Function OpenDeptFR()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg, strAsstNo, strAcqSeq
	Dim IntRetCd


	arrParam(0) = "관리부서 팝업"		
	
	strAsstNo  = Trim(frm1.txtAsstNo.value)
		
	if strAsstNo = "" then
		IntRetCD = DisplayMsgBox("117326","X","X","X")    '자산번호를 입력하십시오.
		Exit Function
	end if
		
	arrParam(1) = "B_ACCT_DEPT A,A_ASSET_INFORM_OF_DEPT B, B_COST_CENTER C, B_BIZ_AREA D"
	arrParam(2) = Trim(frm1.txtFrDeptCd.Value)       	'Code Condition
		
	arrParam(3) = "" 
		
	arrParam(4) = "B.ASST_NO =  " & FilterVar(strAsstNo , "''", "S") & " AND B.DEPT_CD = A.DEPT_CD AND A.ORG_CHANGE_ID = B.ORG_CHANGE_ID " _
					& "AND A.COST_CD = C.COST_CD AND C.BIZ_AREA_CD = D.BIZ_AREA_CD " 'AND B.INV_QTY <> 0"
		
	arrParam(5) = "관리부서코드"			
	
	arrField(0) = "B.DEPT_CD"	
	arrField(1) = "A.DEPT_Nm"		
	arrField(2) = "D.BIZ_AREA_CD"	
	arrField(3) = "D.BIZ_AREA_Nm"		
	arrField(4) = "B.ORG_CHANGE_ID"		
		
	arrHeader(0) = "관리부서코드"		
	arrHeader(1) = "관리부서명"		
	arrHeader(2) = "사업장코드"		
	arrHeader(3) = "사업장명"		    
	arrHeader(4) = "조직변경ID"	
			
		

	IsOpenPop = True
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=650px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "txtFrDeptCd"
		Call SetPopUpReturn(arrRet,field_fg)
	End If	
End Function


'=======================================================================================================
'	Name : SetPopUpReturn()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetPopUpReturn(byval arrRet,byval field_fg)
	
	Select case field_fg
		case "txtFrDeptCd"
			frm1.txtFrDeptCd.value    = arrRet(0)		
			frm1.txtFrDeptCdNm.value  = arrRet(1)			
			frm1.txtFrBizAreaCd.value = arrRet(2)		
			frm1.txtFrBizAreaNm.value = arrRet(3)
			frm1.txthFrOrgChangeId.value = arrRet(4)
		case "txtToDeptCd"
		
			frm1.txtToDeptCd.value    = arrRet(0)				
			frm1.txtToDeptCdNm.value  = arrRet(1)
			frm1.txtChgDt.text  = arrRet(3)
			Call txtToDeptCd_onblur() 
	
	End Select
	
	lgBlnFlgChgValue = True	
	
End Function

Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function
'=======================================================================================================
'Description : 회계전표 생성내역 팝업 
'=======================================================================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	
End Function

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()
	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  
		.ReDraw = false	
		
		.MaxCols = C_BizAreaNm + 1                               '☜: 최대 Columns의 항상 1개 증가시킴 
		'.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear
		
		'Hidden Column 설정 
    	.Col = .MaxCols											'공통콘트롤 사용 Hidden Column
    	.ColHidden = True
    		
		Call GetSpreadColumnPos("A")
		'Call AppendNumberPlace("6","3","0")
		ggoSpread.SSSetEdit	  C_DeptCd,		   "관리부서코드", 20, 0, -1, 30
		ggoSpread.SSSetEdit	  C_DeptNm,		   "관리부서명",   27, 0, -1, 30
		ggoSpread.SSSetFloat  C_InvRate,	   "배분비율(%)",     19, parent.ggExchRateNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"		
		ggoSpread.SSSetEdit	  C_CostCenterNm,  "코스트센타명", 25, 0, -1, 30
		ggoSpread.SSSetEdit	  C_BizAreaNm,     "사업장명",     25, 0, -1, 30
		
		'Call ggoSpread.MakePairsColumn(C_DeptCd,C_DeptNm)
		.ReDraw = true
		
		Call SetSpreadLock 
		
	End With
    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SpreadLock C_DeptCd,      -1, C_DeptCd
		ggoSpread.SpreadLock C_DeptNm,      -1, C_DeptNm
		ggoSpread.SpreadLock C_InvRate,      -1, C_InvRate
		ggoSpread.SpreadLock C_CostCenterNm,-1, C_CostCenterNm
	    ggoSpread.SpreadLock C_BizAreaNm,   -1, C_BizAreaNm +1
		
		.vspdData.ReDraw = True
	End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_DeptCd		 = iCurColumnPos(1)
			C_DeptNm		 = iCurColumnPos(2)
			C_InvRate		 = iCurColumnPos(3)
			C_CostCenterNm   = iCurColumnPos(4)
			C_BizAreaNm	     = iCurColumnPos(5)
	End Select
End Sub
 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
'	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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
'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row   

    'lgBlnFlgChgValue = True

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"	'Split 상태코드 
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
'       End If
'    End If
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
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
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
    End if
        
End Sub

Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub
'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtChgDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtChgDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgDt.Action = 7
        lgBlnFlgChgValue = True	
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7
    End If
End Sub

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
        
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    

    
    Call AppendNumberPlace("6","11","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec) 
        
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call InitSpreadSheet()
    Call InitVariables																'⊙: Initializes local global variables
    Call SetDefaultVal
'    Call SetToolbar("1110100000000111")
    Call SetToolbar("1110100000001111")										' 처음 로드시 표준 에 따라 

	lgBlnFlgChgValue = False


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
		
    frm1.txtChgNo.focus 
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
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
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables															'⊙: Initializes local global variables
'    Call InitSpreadSheet																			'⊙: Initializes local global variables
	'frm1.vspdData.MaxRows = 0 ' InitSpreadSheet 대신 
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Query function call area
    '----------------------- 
    '''''Call DbQuery_master
    Call DbQuery																'☜: Query db data
           
    FncQuery = True																'⊙: Processing is OK        
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 

    FncNew = False                                                          '⊙: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True  Then   'Or ggoSpread.SSCheckChange = True 
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")           '⊙: "Will you destory previous data"
	    'intRetCD = MsgBox("데이타가 변경되었습니다. 신규입력을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    
    Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    Call InitVariables															'⊙: Initializes local global variables
'    Call InitSpreadSheet																			'⊙: Initializes local global variables
	'frm1.vspdData.MaxRows = 0 ' InitSpreadSheet 대신 
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
'	Call SetToolbar("11101000000011")
    Call SetToolbar("1110100000000111")										' 처음 로드시 표준 에 따라 	
    Call SetDefaultVal
    
    FncNew = True																'⊙: Processing is OK
	lgBlnFlgChgValue = False

    'SetGridFocus

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
	
	dim intRetCD
	    
    FncDelete = False														'⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
		IntRetCD = DisplayMsgBox("900002","X","X","X")                                
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    dim intRetVal
    Dim varDeptCd
    
    if intRetVal = false then
    '   exit function
    end if
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                          '⊙: No data changed!!
        Exit Function
    End If
    
  '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             '⊙: Check contents area
       Exit Function
    End If
    
	varDeptCd = frm1.txtToDeptCd.value 
	
	if varDeptCd = frm1.txtFrDeptCd.value then
		IntRetCD = DisplayMsgBox("117324","X","X","X")
		Exit Function
	end if	       	

  '-----------------------
    'Save function call area
    '-----------------------
    CAll DbSave				                                                '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call Parent.fncPrint()    
End Function




'======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                               
End Function
'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)										
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
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")
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
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 

    Err.Clear                                                               '☜: Protect system from crashing
    
    DbDelete = False														'⊙: Processing is NG
    Call LayerShowHide(1)    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtChgNo=" & Trim(frm1.txtChgNo.value)		'☜: 삭제 조건 데이타 
    strVal = strVal & "&CboChgFg=" & "05" 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
	
	lgBlnFlgChgValue = False        	
	Call FncNew()
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 


    
    DbQuery = False                                                         '⊙: Processing is NG
    Call LayerShowHide(1)
    Err.Clear                                                               '☜: Protect system from crashing
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode="   & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtChgNo="		& Trim(frm1.txtChgNo.value)  		'☆: 조회 조건 데이타 
    '''strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
'    Call SetToolbar("1111100000010111")
    Call SetToolbar("1111100000011111")
	
	'SetGridFocus
	
	Call DbQuery_master("Q")														'''자산 부서별 정보를 조회한다.
End Function


'========================================================================================
' Function Name : DbQuery_master
' Function Desc : This function is data query and displayf
'========================================================================================

Function DbQuery_master(ByVal OptMeth) 
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbQuery_master = False                                                         '⊙: Processing is NG
'    Call InitSpreadSheet																			'⊙: Initializes local global variables
	'frm1.vspdData.MaxRows = 0 ' InitSpreadSheet 대신 
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Dim strVal
    
    strVal = BIZ_PGM_ID2 & "?txtMode="   & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtAsstNo="		& Trim(frm1.txtAsstNo.value)  		'☆: 조회 조건 데이타 
    strVal = strVal & "&txtQueryFg="    & "Master"
	strval = strval & "&txtOptMeth="    & OptMeth
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery_master = True                                                          '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk_master()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    'lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	
'    Call SetToolbar("11111000000111")
	Call txtFrDeptCd_onblur()
	Call txtToDeptCd_onblur()
    If lgIntFlgMode = parent.OPMD_UMODE Then
		frm1.vspdData.focus
	Else
		frm1.txtChgDt.focus
	End If
	lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 
	Call LayerShowHide(1)
    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG        
     
	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With

    DbSave = True                                                           '⊙: Processing is NG


End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()															'☆: 저장 성공후 실행 로직	  
     
    Call InitVariables
    
    Call DBQuery()
	
	lgIntFlgMode = parent.OPMD_UMODE
	
End Function

	'==========================================================================================
'   Event Name : txtFrDeptCd_onblur
'   Event Desc : 
'==========================================================================================
Sub txtFrDeptCd_onblur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtChgDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtFrDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " B.BIZ_AREA_CD,  c.biz_area_nm, A.org_change_id, A.internal_cd"    		
		strFrom		=			 " B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C "
		strWhere	=		     " A.COST_CD = B.COST_CD and b.BIZ_AREA_CD=C.BIZ_AREA_CD "
		strWhere	= strWhere & " AND A.dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtFrDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and A.org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		'strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtChgDt.Text, gDateFormat,""), "''", "S") & "))"			
		strWhere    = strWhere & " from b_acct_dept where dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtFrDeptCd.value)), "''", "S") & " and org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtChgDt.Text, gDateFormat,""), "''", "S") & "))" 
		'부서코드 기준으로의 max(org_change_id)로 수정 
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtFrDeptCd.value = ""
			frm1.txtFrDeptCdNm.value = ""
			frm1.txtfrBizAreaCd.value = ""
			frm1.txtfrBizAreaNm.value = ""
			frm1.txthfrOrgChangeId.value = ""
			frm1.txtfrDeptCd.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.txtfrBizAreaCd.value = Trim(arrVal2(1))
				frm1.txtfrBizAreaNm.value = Trim(arrVal2(2))
				frm1.txthfrOrgChangeId.value = Trim(arrVal2(3))
				
				
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------

End Sub
	

'==========================================================================================
'   Event Name : txtToDeptCd_onblur
'   Event Desc : 
'==========================================================================================
Sub txtToDeptCd_onblur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtChgDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtToDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " B.BIZ_AREA_CD,  c.biz_area_nm, A.org_change_id, A.internal_cd"    		
		strFrom		=			 " B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C "
		strWhere	=		     " A.COST_CD = B.COST_CD and b.BIZ_AREA_CD=C.BIZ_AREA_CD "
		strWhere	= strWhere & " AND A.dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtToDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and A.org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtChgDt.Text, gDateFormat,""), "''", "S") & "))"			
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtToDeptCd.value = ""
			frm1.txtToDeptCdNm.value = ""
			frm1.txtToBizAreaCd.value = ""
			frm1.txtToBizAreaNm.value = ""
			frm1.txthToOrgChangeId.value = ""
			frm1.txtToDeptCd.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.txtToBizAreaCd.value = Trim(arrVal2(1))
				frm1.txtToBizAreaNm.value = Trim(arrVal2(2))
				frm1.txthToOrgChangeId.value = Trim(arrVal2(3))
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------

End Sub
	
	


'***************************************************************************************************************


Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
End Function

' For Developement
Function click_query()
	Call FncQuery()
End Function

Function click_save()
	Call FncSave()
End Function

Function click_delete()
	Call FncDelete()
End Function

Function click_new()
	Call FncNew()
End Function

Sub txtChgDt_onBlur()
    
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2

	lgBlnFlgChgValue = True
	With frm1
	
		If LTrim(RTrim(.txtToDeptCd.value)) <> "" and Trim(.txtChgDt.Text <> "") Then
			'----------------------------------------------------------------------------------------
				strSelect	=			 " Distinct org_change_id "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtToDeptCd.value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtChgDt.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.txthToOrgChangeId.value) Then
					'IntRetCD = DisplayMsgBox("124600","X","X","X") 
					.txtToDeptCd.value = ""
					.txtToDeptCdNm.value = ""
					.txtToBizAreaCd.value = ""
					.txtToBizAreaNm.value=""
					.txthToOrgChangeId.value = ""
					.txtToDeptCd.focus
			End if
		End If
	End With
'----------------------------------------------------------------------------------------

End Sub

Sub txtChgQty_Change()
    lgBlnFlgChgValue = true
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
<!-- #Include file="../../inc/uni2kcm.inc"  -->

</SCRIPT>
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenMasterRef()">자산마스터참조</A></TD>
			        <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- 본문내용  -->
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
		<!-- 첫번째 탭 내용  -->
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>자산변동번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtChgNo" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="자산변동번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenChgNoInfo"></TD>
								</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
							<TABLE <%=LR_SPACE_TYPE_60%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>자산번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstNo" SIZE=20 MAXLENGTH=18 TAG="24XXXU" ALT="자산번호"> <INPUT TYPE="Text" NAME="txtAsstNm" SIZE=30 MAXLENGTH=40 tag="24X" ALT="자산명"></TD>
												<TD CLASS="TDT" NOWRAP>&nbsp</TD>
												<TD CLASS="TD6" NOWRAP>&nbsp</TD>
											</TR>
											<TR>
												<TD WIDTH=* HEIGHT=100% COLSPAN=4>
												    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="24x2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
												</TD>
											</TR>

						</TABLE>
					</TD>
				</TR>		
				<TR>
					<TD WIDTH=100% HEIGHT=40 VALIGN=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>부서이동일자</TD>
												<TD CLASS="TD6" NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtChgDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="매각/폐기일자"> </OBJECT>');</SCRIPT>				
												<TD CLASS=TD5 NOWRAP>이동수량</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtChgQty style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 120px" title="FPDOUBLESINGLE" ALT="이동수량" tag="24X3Z"> </OBJECT>');</SCRIPT>&nbsp;
												</TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>주는부서</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFrDeptCd" SIZE=15 MAXLENGTH=10 tag="22XXXU" ALT="주는부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDeptFR">&nbsp;<INPUT TYPE=TEXT NAME="txtFrDeptCdNm" SIZE=20 tag="24" ></TD>
												<TD CLASS="TD5" NOWRAP>주는사업장</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFrBizAreaCd" SIZE=15 MAXLENGTH=10 tag="24XXXU" ALT="주는사업장">&nbsp;<INPUT TYPE=TEXT NAME="txtFrBizAreaNm" SIZE=20 tag="24" ></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>받는부서</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToDeptCd" SIZE=15 MAXLENGTH=10 tag="22XXXU" ALT="받는부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDeptTO">&nbsp;<INPUT TYPE=TEXT NAME="txtToDeptCdNm" SIZE=20 tag="24" ></TD>
												<TD CLASS="TD5" NOWRAP>받는사업장</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToBizAreaCd" SIZE=15 MAXLENGTH=10 tag="24XXXU" ALT="받는사업장">&nbsp;<INPUT TYPE=TEXT NAME="txtToBizAreaNm" SIZE=20 tag="24" ></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>적요</TD>
												<TD CLASS="TD6" NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtDesc" SIZE=90 MAXLENGTH=30 tag="2X" ALT="적요"></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>&nbsp</TD>
												<TD CLASS="TD6" NOWRAP>&nbsp</TD>
												<TD CLASS="TD5" NOWRAP>&nbsp</TD>
												<TD CLASS="TD6" NOWRAP>&nbsp</TD>
											</TR>

											<TR>
												<TD CLASS="TD5" NOWRAP>주는부서</TD>
												<TD CLASS="TD6" NOWRAP>&nbsp</TD>
												<TD CLASS="TD5" NOWRAP>받는부서</TD>
												<TD CLASS="TD6" NOWRAP>&nbsp</TD>
											</TR>

											<TR>
												<TD CLASS="TD5" NOWRAP>결의전표번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFrTempGLNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="주는부서 결의전표번호"></TD>
												<TD CLASS="TD5" NOWRAP>결의전표번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToTempGLNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="받는부서 결의전표번호"></TD>
											</TR>	
											<TR>
												<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFrGLNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="주는부서 전표번호"></TD>
												<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToGLNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="받는부서 전표번호"></TD>
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
	<TR height=10>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA><% '업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode"		tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"	tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"	tag="24" TABINDEX = "-1">
<INPUT TYPE=hidden NAME="txthChgRefNo"	tag="24" TABINDEX = "-1">
<INPUT TYPE=hidden NAME="txthAsstNo"	tag="24" TABINDEX = "-1">
<INPUT TYPE=hidden NAME="txthFrOrgChangeId"	tag="24" TABINDEX = "-1">
<INPUT TYPE=hidden NAME="txthToOrgChangeId"	tag="24" TABINDEX = "-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

