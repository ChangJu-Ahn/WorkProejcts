<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		: 공정별원가 
'*  3. Program ID			: C4005MA1.asp
'*  4. Program Name			:배부요소DATA등록 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4005MB1.asp
'*						
'*  7. Modified date(First)	: 2005/09/05
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: 
'* 11. Comment				: 
'* 12. History              : 
'*                          : 
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" --> 
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID	= "C4005MB1.asp"			'☆:  비지니스 로직 ASP명 
Const BIZ_PGM_ID2	= "C4005MB2.asp"			'☆:  비지니스 로직 ASP명 


Dim LocSvrDate
Dim StartDate
Dim EndDate

LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)													

'spread wp
Dim C_ContractFlg
Dim C_WPCd			
Dim C_WPPop		
Dim C_WPNm				
Dim C_FctrCd			
Dim C_FctrCdPop	
Dim C_FctrNm
Dim C_AllocData	

'spread cc
Dim C_ContractFlg2
Dim C_CCCd2				
Dim C_CCPop2		
Dim C_CCNm2			
Dim C_FctrCd2			
Dim C_FctrCdPop2	
Dim C_FctrNm2
Dim C_AllocData2			

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
Dim lgStrPrevKey,lgStrPrevKey2
Dim lgLngCurRows
Dim lgSortKey
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim ihGridCnt                     'hidden Grid Row Count
Dim intItemCnt                    'hidden Grid Row Count
Dim IsOpenPop						' Popup


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

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
	lgBlnFlgChgValue = False
    lgStrPrevKey = ""			'initializes Previous Key
    lgStrPrevKey2 = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
    lgSortKey = 1

End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 

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
Sub InitSpreadSheet(ByVal pvSpdNo)


    
    Call InitSpreadPosVariables(pvSpdNo)
    
	Call AppendNumberPlace("6","11","4")
	IF pvSpdNo="A" THEN
	    With frm1
	           
	    ggoSpread.Source = .vspdData
	    ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
	    
	      
	    
	    .vspdData.ReDraw = False
	    
	    .vspdData.MaxCols = C_AllocData + 1
	    .vspdData.MaxRows = 0
		
		Call GetSpreadColumnPos(pvSpdNo)
		
	    ggoSpread.SSSetCombo		C_ContractFlg,				"사내/외주가공구분",18,0
	    ggoSpread.SetCombo		"사내" & vbTab & "외주가공"  , C_ContractFlg      
	    ggoSpread.SSSetEdit		C_WPCd,				"공정/구매그룹", 15,,,7
	    ggoSpread.SSSetButton C_WPPop
	    ggoSpread.SSSetEdit		C_WPNm,				"공정/구매그룹명", 20
	    ggoSpread.SSSetEdit		C_FctrCd,			"배부요소", 15,,,3
	    ggoSpread.SSSetButton C_FctrCdPop
	    ggoSpread.SSSetEdit		C_FctrNm,			"배부요소명", 15    
	    
		ggoSpread.SSSetFloat		C_AllocData,			"배부요소DATA",25,"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
	   	
	 	Call ggoSpread.MakePairsColumn(C_WPCd,C_WPPop)
	 	Call ggoSpread.MakePairsColumn(C_FctrCd,C_FctrCdPop)
	 	
	 	Call ggoSpread.SSSetColHidden(.vspdData.MaxCols ,.vspdData.MaxCols , True)
			
	    ggoSpread.SSSetSplit2(2) 
		.vspdData.ReDraw = False
		
	    End With
	ELSE


	With frm1
	           
	    ggoSpread.Source = .vspdData2
	    ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
	     
	    .vspdData2.ReDraw = False
	    
	    .vspdData2.MaxCols = C_AllocData2 + 1
	    .vspdData2.MaxRows = 0
		
		Call GetSpreadColumnPos(pvSpdNo)
		ggoSpread.SSSetEdit		C_CCCd2,				"C/C", 10,,,10
	    ggoSpread.SSSetButton C_CCPop2
	    ggoSpread.SSSetEdit		C_CCNm2,				"C/C명", 20
	    ggoSpread.SSSetCombo		C_ContractFlg2,				"사내/외주가공구분",18,0
	    ggoSpread.SetCombo		"사내" & vbTab & "외주가공"  , C_ContractFlg2      

	    ggoSpread.SSSetEdit		C_FctrCd2,			"배부요소", 15,,,3
	    ggoSpread.SSSetButton C_FctrCdPop2
	    ggoSpread.SSSetEdit		C_FctrNm2,			"배부요소명", 15	    
	   	ggoSpread.SSSetFloat		C_AllocData2,			"배부요소DATA",25,"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
	 	Call ggoSpread.MakePairsColumn(C_CCCd2,C_CCPop2)
	 	Call ggoSpread.MakePairsColumn(C_FctrCd2,C_FctrCdPop2)
	 	
	 	Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols ,.vspdData2.MaxCols , True)
	 	Call ggoSpread.SSSetColHidden(C_ContractFlg2 ,C_ContractFlg2 , True)
			
	    ggoSpread.SSSetSplit2(2) 
		.vspdData2.ReDraw = False
		
	    End With


	END IF   
    Call SetSpreadLock(pvSpdNo)
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock(ByVal pvSpdNo)
	If pvSpdNo="A" Then
		ggoSpread.Source = frm1.vspdData    
	Else
		ggoSpread.Source = frm1.vspdData2   	
	End If
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvSpdNo,ByVal pvStartRow, ByVal pvEndRow)
    
    If pvSpdNo="A" Then
		With frm1.vspdData 
    
		.Redraw = False

		ggoSpread.Source = frm1.vspdData    
		ggoSpread.SSSetRequired C_WPCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_WPNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_FctrCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_FctrNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_AllocData, pvStartRow, pvEndRow
    
		.Col = 1
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True
    
		.Redraw = True
    
		End With
	Else
	
			With frm1.vspdData2 
    
		.Redraw = False

		ggoSpread.Source = frm1.vspdData2    
		ggoSpread.SSSetRequired C_CCCd2, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CCNm2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_FctrCd2, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_FctrNm2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_AllocData2, pvStartRow, pvEndRow
    
		.Col = 1
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True
    
		.Redraw = True
    
		End With
	End If
End Sub
'================================== 2.2.5 SetQuerySpreadColor() ==================================================
' Function Name : SetQuerySpreadColor
' Function Desc :  This method set color and protect  in spread sheet celles, after Query
'========================================================================================

Sub SetQuerySpreadColor(ByVal pvSpdNo,ByVal lRow)
    If pvSpdNo="A" Then
		With frm1
			.vspdData.ReDraw = False
		  
			ggoSpread.SSSetProtected C_WPCd, -1, -1
			ggoSpread.SSSetProtected C_WPPop, -1, -1
			ggoSpread.SSSetProtected C_WPNm, -1, -1
			ggoSpread.SSSetProtected C_FctrCd, -1, -1
			ggoSpread.SSSetProtected C_FctrCdPop, -1, -1
			ggoSpread.SSSetProtected C_FctrNm, -1, -1
			ggoSpread.SSSetProtected C_ContractFlg, -1, -1
			ggoSpread.SSSetRequired C_AllocData, -1, -1
			.vspdData.ReDraw = True
		End With
	Else
			With frm1
			.vspdData2.ReDraw = False
		  
			ggoSpread.SSSetProtected C_CCCd2, -1, -1
			ggoSpread.SSSetProtected C_CCPop2, -1, -1
			ggoSpread.SSSetProtected C_CCNm2, -1, -1
			ggoSpread.SSSetProtected C_FctrCd2, -1, -1
			ggoSpread.SSSetProtected C_FctrCdPop2, -1, -1
			ggoSpread.SSSetProtected C_FctrNm2, -1, -1
			ggoSpread.SSSetProtected C_ContractFlg2, -1, -1
			ggoSpread.SSSetRequired C_AllocData2, -1, -1
			.vspdData2.ReDraw = True
		End With
	End If
End Sub
'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo="A" Then	
		C_ContractFlg		=1
		C_WPCd				= 2
		C_WPPop				=3
		C_WPNm				= 4
		C_FctrCd				= 5
		C_FctrCdPop			=6
		C_FctrNm			=7
		C_AllocData				= 8
	Else
		
		C_CCCd2				= 1
		C_CCPop2				=2
		C_CCNm2				= 3
		C_ContractFlg2		=4
		C_FctrCd2				= 5
		C_FctrCdPop2			=6
		C_FctrNm2			=7
		C_AllocData2				= 8
	End If
End Sub

 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		C_ContractFlg		= iCurColumnPos(1)
 		C_WPCd				= iCurColumnPos(2)
 		C_WPPop				= iCurColumnPos(3)
		C_WPNm				= iCurColumnPos(4)
		C_FctrCd				= iCurColumnPos(5)
		C_FctrCdPop			= iCurColumnPos(6)
		C_FctrNm				= iCurColumnPos(7)
		C_AllocData			= iCurColumnPos(8)	
	Case "B"
 		ggoSpread.Source = frm1.vspdData2 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		
 		C_CCCd2				= iCurColumnPos(1)
 		C_CCPop2				= iCurColumnPos(2)
		C_CCNm2				= iCurColumnPos(3)
		C_ContractFlg2		= iCurColumnPos(4)
		C_FctrCd2				= iCurColumnPos(5)
		C_FctrCdPop2		= iCurColumnPos(6)
		C_FctrNm2				= iCurColumnPos(7)
		C_AllocData2			= iCurColumnPos(8)		
	
	
		
 	End Select 
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
    
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

'------------------------------------------  OpenFctrCd()  -------------------------------------------------
'	Name : OpenFctrCd()
'	Description : Condition Dstb_Fctr_cd PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenFctrCd()
	
	Dim arrRet
	Dim strWhere, strFrom
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	

	IsOpenPop = True
	strFrom = "  c_dstb_fctr_s  "
	strWhere = " gen_flag='M' "
			
	arrParam(0) = "배부요소"						' 팝업 명칭 
	arrParam(1) =strFrom					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtFctrCd.Value)	' Code Condition
	arrParam(3) =""										' Name Cindition
	arrParam(4) =strWhere						' Where Condition
	arrParam(5) = "배부요소"							' TextBox 명칭 
	
    arrField(0) ="ED12" & Parent.gColSep &  "dstb_fctr_cd"					' Field명(0)
    arrField(1) = "ED30" & Parent.gColSep & "dbo.ufn_getCodeName('C4000',dstb_fctr_cd) "					' Field명(1)
    
    
    arrHeader(0) = "배부요소"						' Header명(0)
    arrHeader(1) = "배부요소명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetFctrCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtFctrCd.focus
End Function

'------------------------------------------  OpenWPCd()  -------------------------------------------------
'	Name : OpenWPCd()
'	Description : Condition Operation PopUp  OpenCCCd
'---------------------------------------------------------------------------------------------------------
Function OpenWPCd()
	Dim arrRet
	Dim strWhere, strFrom
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	

	IsOpenPop = True
	strFrom = "(select  'M' flag, '사내' AS FLAG_NM , wc_cd as code, wc_nm as cd_nm	 from P_work_center "
	strFrom = strFrom & " union "
	strFrom = strFrom & "	select  'O' flag, '외주가공' AS FLAG_NM, pur_grp as code, pur_grp_nm as cd_nm from b_pur_grp where usage_flg='Y') tmp"
			
	arrParam(0) = "공정/구매그룹팝업"						' 팝업 명칭 
	arrParam(1) =strFrom					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWPCd.Value)	' Code Condition
	arrParam(3) =""										' Name Cindition
	arrParam(4) =""							' Where Condition
	arrParam(5) = "코드"							' TextBox 명칭 
	
    arrField(0) ="HH" & Parent.gColSep &  "code"					' Field명(0)
    arrField(1) = "HH" & Parent.gColSep & "flag"					' Field명(1)
    arrField(2) ="ED17" & Parent.gColSep &  "FLAG_NM"					' Field명(0)
    arrField(3) = "ED8" & Parent.gColSep & "code"					' Field명(1)
    arrField(4) ="ED25" & Parent.gColSep &  "cd_nm"					' Field명(0)
    
    
    arrHeader(0) = "코드"						' Header명(0)
    arrHeader(1) = "사내/외주가공구분"						' Header명(1)
    arrHeader(2) = "사내/외주가공구분"						' Header명(0)    
    arrHeader(3) = "코드"						' Header명(0)
    arrHeader(4) = "코드명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWPCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWPCd.focus
	
End Function
'------------------------------------------  OpenCCCd()  -------------------------------------------------
'	Name : OpenCCCd()
'	Description : Condition Operation PopUp  
'---------------------------------------------------------------------------------------------------------
Function OpenCCCd()
	Dim arrRet
	Dim strWhere, strFrom
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	

	IsOpenPop = True
			
	arrParam(0) = "C/C"						' 팝업 명칭 
	arrParam(1) =" b_cost_center "					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCCCd.Value)	' Code Condition
	arrParam(3) =""										' Name Cindition
	arrParam(4) =""							' Where Condition
	arrParam(5) = "C/C"							' TextBox 명칭 
	
    arrField(0) ="ED10" & Parent.gColSep &  "cost_cd"					' Field명(0)
    arrField(1) = "ED31" & Parent.gColSep & "cost_nm"					' Field명(1)
    
    
    arrHeader(0) = "C/C"						' Header명(0)
    arrHeader(1) = "C/C명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCCCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCCCd.focus
	
End Function
'==========================================  2.4.3 Set Return Value()  =============================================
'	Name : Set Return Value()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetFctrCd()  --------------------------------------------------
'	Name : SetFctrCd()
'	Description : openFctrCdPopup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetFctrCd(byval arrRet)
	frm1.txtFctrCd.Value    = arrRet(0)
	frm1.txtFctrNm.Value    = arrRet(1)				
End Function

'------------------------------------------  SetWPCd()  --------------------------------------------------
'	Name : SetWPCd()
'	Description : openWPCd Popup에서 Return되는 값 setting  
'--------------------------------------------------------------------------------------------------------- 
Function SetWPCd(byval arrRet)
	frm1.txtWPCd.Value    = arrRet(3)		
	frm1.txtWPNm.Value   = arrRet(4)
End Function
'------------------------------------------  SetCCCd()  --------------------------------------------------
'	Name : SetCCCd()
'	Description : openCCCd Popup에서 Return되는 값 setting  
'--------------------------------------------------------------------------------------------------------- 
Function SetCCCd(byval arrRet)
	frm1.txtCCCd.Value    = arrRet(0)		
	frm1.txtCCNm.Value   = arrRet(1)
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
Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

	Call InitSpreadSheet("B")                                        '⊙: Setup the Spread sheet
	Call InitVariables()                                                    '⊙: Initializes local global variables

	'----------  Coding part  -------------------------------------------------------------
	'Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
	Call SetToolBar("11001111001111")		

	Call SetDefault()

End Sub
'========================================== ======================================
'	Name : SetDefault()
'	Description : 
'=========================================================================================================
Sub  SetDefault()

		frm1.txtYYYYMM.Text=LocSvrDate		
		Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)

		frm1.txtYYYYMM.FOCUS

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


'=========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row)
	Dim iIndex
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
       
    with frm1.vspdData
		
		.Col = C_ContractFlg :		.Row = Row
		iIndex=.value
		.Col=Col
		Select Case Col
		Case C_WPCd
			Call checkWPCd(Row,.Text,iIndex)    
		Case C_FctrCd
		    Call checkFctrCd(Row, .Text)
		End Select
	End With
End Sub
'=========================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'==========================================================================================

Sub vspdData2_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row
    
    with frm1.vspdData2
		.Col = Col
		.Row = Row
		Select Case Col
		Case C_CCCd2
			Call checkWPCd(Row,.Text,"")    
		Case C_FctrCd2    
		    Call checkFctrCd(Row, .Text)
		End Select
	End With
    
End Sub
'==========================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'==========================================================================================

Sub vspdData_EditChange(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row        

End Sub
'==========================================================================================
'   Event Name :vspddata2_EditChange
'   Event Desc :
'==========================================================================================

Sub vspdData2_EditChange(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row        

End Sub

'==========================================================================================
 '  Event Name :vspdData_ComboSelChange
'   Event Desc :
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIndex
	
	With frm1.vspdData
		.Col = C_WPCd 
		If Trim(.Text)<>"" Then 
			.Text ="" : .Col =C_WPNm : .Text=""
		End If
		Call vspdData_Change(C_WPCd , Row)

	End With
End Sub

'==========================================================================================
 '  Event Name :vspdData2_ComboSelChange
'   Event Desc :
'==========================================================================================
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
  	
	Call vspdData2_Change(C_ContractFlg2 , Row)
		
End Sub
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows <= 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
	End If

End Sub
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData2_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows <= 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
	End If

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
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'==========================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================

Sub vspdData_DblClick(ByVal Col , ByVal Row )
Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If

End Sub
'==========================================================================================
'   Event Name :vspddata2_DblClick
'   Event Desc :
'==========================================================================================

Sub vspdData2_DblClick(ByVal Col , ByVal Row )
Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If

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
'   Event Name : vspdData2_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData2 
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
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)	Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub
'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop)	Then
		If lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	If frm1.vspdData.MaxRows <0 Then Exit Sub
	
	With frm1.vspdData 		
		ggoSpread.Source = frm1.vspdData
	  
		If  Col = C_WPPop Then
		    .Col = Col :		    .Row = Row
	
			Call OpenSpreadPopup(C_WPPop, Row, .Text)		
			Call SetActiveCell(frm1.vspdData,C_FctrCd,Row,"M","X","X")			
		ElseIf  Col = C_FctrCdPop Then
		    .Col = Col 
		    .Row = Row
			Call OpenSpreadPopup(Col, Row, .Text)		
		    Call SetActiveCell(frm1.vspdData,C_AllocData,Row,"M","X","X")
		End If
    
	End With	
End Sub
'==========================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	If frm1.vspdData2.MaxRows <0 Then Exit Sub
	
	With frm1.vspdData2 		
		ggoSpread.Source = frm1.vspdData2
	  
		If  Col = C_CCPop2 Then
		    .Col = Col :		    .Row = Row
	
			Call OpenSpreadPopup(Col, Row, .Text)		
			Call SetActiveCell(frm1.vspdData2,C_FctrCd2,Row,"M","X","X")			
		ElseIf  Col = C_FctrCdPop2 Then
		    .Col = Col 
		    .Row = Row
			Call OpenSpreadPopup(Col, Row, .Text)	    
		    Call SetActiveCell(frm1.vspdData2,C_AllocData2,Row,"M","X","X")
		End If
    
	End With	
End Sub


'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
  
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 '========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    
    Call InitSpreadSheet(gActiveSpdSheet.ID)
    Call SetQuerySpreadColor(gActiveSpdSheet.ID,1)
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'=======================================================================================================
'   Event Name : txtYYYYMM_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtYYYYMM.Action = 7
        Call SetFocusToDocument("P")
		Frm1.txtYYYYMM.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtYYYYMM_KeyDown
'   Event Desc : 
'=======================================================================================================
Sub  txtYYYYMM_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'=======================================================================================================
'   Event Name : rdoCC_onClick
'   Event Desc : 
'=======================================================================================================
Sub  rdoCC_onClick()
	dim IntRetCD
	ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
	If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			frm1.rdoWP.checked= true
			Exit Sub
		End If
	End If

	with frm1
			.document.all("tmpwp").style.display="none"
			.document.all("divWplbl").style.display="none"
			.document.all("sprdWP").style.display="none"
			
			.document.all("tmpCc").style.display="block"			
			.document.all("divCClbl").style.display="block"
			.document.all("sprdCC").style.display="block"
			.hCode.value=""
			.hFctrCd.value=""
			.hYYYYMM.value=""
		
			Call InitSpreadSheet("B")
			CALL InitVariables()
	End with

End Sub
'=======================================================================================================
'   Event Name :rdoWP_onClick
'   Event Desc : 
'=======================================================================================================
Sub  rdoWP_onClick()
	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData2							'⊙: Preset spreadsheet pointer 
	If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			frm1.rdoCC.checked=true
			Exit Sub
		End If
	End If

	with frm1			
			.document.all("tmpCc").style.display="none"				
			.document.all("divCClbl").style.display="none"
			.document.all("sprdCC").style.display="none"
			
			.document.all("tmpwp").style.display="block"
			.document.all("divWplbl").style.display="block"
			.document.all("sprdWP").style.display="block"
			
			.hCode.value=""
			.hFctrCd.value=""
			.hYYYYMM.value=""
			
			
			Call InitSpreadSheet("A")
			CALL InitVariables()
	End with

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
'********************************************************************************************************* %>
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	If frm1.rdoWP.checked Then
	    ggoSpread.Source = frm1.vspdData										'⊙: Preset spreadsheet pointer 
	    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = True Then									'⊙: Check If data is chaged
	        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'⊙: Display Message
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		
		IF ChkKeyField()=False Then Exit Function 

	    '-----------------------
	    'Erase contents area
	    '-----------------------
	'    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field   
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.ClearSpreadData

	  '  Call InitVariables		
	    '-----------------------
	    'Check condition area
	    '-----------------------
	    If Not chkfield(Document, "1") Then								'⊙: This function check indispensable field
	       Exit Function
	    End If
		Call initVariables()
	    '-----------------------
	    'Query function call area
	    '-----------------------
	    If DbQuery = False Then	
			Call RestoreToolBar()
			Exit Function
		End If															'☜: Query db data

	    FncQuery = True															'⊙: Processing is OK
	Else
		ggoSpread.Source = frm1.vspdData2										'⊙: Preset spreadsheet pointer 
	    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = True Then									'⊙: Check If data is chaged
	        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'⊙: Display Message
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
	    
		
		IF ChkKeyField()=False Then Exit Function 
	    '-----------------------
	    'Erase contents area
	    '-----------------------
	'    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field   
	    ggoSpread.Source = frm1.vspdData2
	    ggoSpread.ClearSpreadData
	    Call InitVariables		

	    '-----------------------
	    'Check condition area
	    '-----------------------
	    If Not chkfield(Document, "1") Then								'⊙: This function check indispensable field
	       Exit Function
	    End If
		Call initVariables()
	    '-----------------------
	    'Query function call area
	    '-----------------------
	    If DbQuery = False Then	
			Call RestoreToolBar()
			Exit Function
		End If															'☜: Query db data
	       
	    FncQuery = True															'⊙: Processing is OK
	
	
	End If
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
    Dim iRow
    Dim iIndex
    FncSave = False                                           '⊙: Processing is NG
    
    Err.Clear                                                 '☜: Protect system from crashing
    If frm1.rdoCC.checked Then 
		ggoSpread.Source = frm1.vspdData2                          '⊙: Preset spreadsheet pointer 

		If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
		    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '⊙: Display Message(There is no changed data.)
		    Exit Function
		End If

		ggoSpread.Source = frm1.vspdData2                          '⊙: Preset spreadsheet pointer 
		If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
		   Exit Function
		End If

		For iRow=1  to frm1.vspdData2.MaxRows			
		    frm1.vspdData2.Row = iRow
		    frm1.vspdData2.Col = 0			
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag				
					frm1.vspdData2.Col = C_CCCd2				
					If  checkWPCd(iRow,frm1.vspdData2.Text,"")=False Then Exit Function 
					
					frm1.vspdData2.Col = C_FctrCd2
					If  checkFctrCd(iRow, frm1.vspdData2.Text)=False Then Exit Function 
			End Select	
		Next
		'-----------------------
		'Save function call area
		'-----------------------
		If DbSave = False Then Exit Function				                                  '☜: Save db data
    
		FncSave = True                                            '⊙: Processing is OK
	Else
		ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 

		If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
		    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '⊙: Display Message(There is no changed data.)
		    Exit Function
		End If

		ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
		If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
		   Exit Function
		End If

		For iRow=1  to frm1.vspdData.MaxRows			
		    frm1.vspdData.Row = iRow
		    frm1.vspdData.Col = C_ContractFlg : iIndex=frm1.vspdData.Value
		    frm1.vspdData.Col = 0			
			Select Case frm1.vspdData.Text
				Case ggoSpread.InsertFlag				
					frm1.vspdData.Col = C_WPCd				
					If  checkWPCd(iRow,frm1.vspdData.Text,iIndex)=False Then Exit Function 
					
					frm1.vspdData.Col = C_FctrCd
					If  checkFctrCd(iRow, frm1.vspdData.Text)=False Then Exit Function 
			End Select	
		Next
		'-----------------------
		'Save function call area
		'-----------------------
		If DbSave = False Then Exit Function				                                  '☜: Save db data
    
		FncSave = True                                            '⊙: Processing is OK
    End If
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG
	IF frm1.rdoWP.checked Then
	    If Frm1.vspdData.MaxRows < 1 Then
	       Exit Function
	    End If
	    
	    ggoSpread.Source = Frm1.vspdData
		
		With Frm1.VspdData
	         .ReDraw = False
			 If .ActiveRow > 0 Then
	            ggoSpread.CopyRow

				SetSpreadColor "A",.ActiveRow, .ActiveRow			

	            .ReDraw = True
			    .Focus
			 End If
		End With
		
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		
		Call SetActiveCell(frm1.vspdData,C_WPCd,frm1.vspdData.ActiveRow,"M","X","X")
		
		'------ Developer Coding part (End )   -------------------------------------------------------------- 
	    If Err.number = 0 Then	
	       FncCopy = True                                                            '☜: Processing is OK
	    End If
	Else
		If Frm1.vspdData2.MaxRows < 1 Then
		      Exit Function
		   End If
		   
		   ggoSpread.Source = Frm1.vspdData2
		
			With Frm1.VspdData2
		        .ReDraw = False
				 If .ActiveRow > 0 Then
		           ggoSpread.CopyRow

					SetSpreadColor "B",.ActiveRow, .ActiveRow			

		           .ReDraw = True
				    .Focus
				 End If		
			'------ Developer Coding part (Start ) -------------------------------------------------------------- 
			End With
				
			Call SetActiveCell(frm1.vspdData2,C_CCCd2,frm1.vspdData2.ActiveRow,"M","X","X")
		
			'------ Developer Coding part (End )   -------------------------------------------------------------- 
		   If Err.number = 0 Then	
		      FncCopy = True                                                            '☜: Processing is OK
		   End If
	   End If
	   
	   Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================

Function FncPaste() 
     ggoSpread.SpreadPaste
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow, i
    Dim iIntIndex
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If
	If frm1.rdoWP.checked Then 
		With frm1
	        .vspdData.ReDraw = False
	        .vspdData.focus
	        ggoSpread.Source = .vspdData
	        ggoSpread.InsertRow ,imRow
			
			.vspdData.Col =C_ContractFlg :				.vspdData.Text ="사내"


	        SetSpreadColor "A",.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1		
			
	        .vspdData.ReDraw = True
	        lgBlnFlgChgValue = True  
	    End With
	Else

		With frm1
	        .vspdData2.ReDraw = False
	        .vspdData2.focus
	        ggoSpread.Source = .vspdData2
	        ggoSpread.InsertRow ,imRow
	        
	        .vspdData2.Col =C_ContractFlg2 :			.vspdData2.Text ="사내"
			

	        SetSpreadColor "B",.vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1		
				
	        .vspdData2.ReDraw = True
	        lgBlnFlgChgValue = True  
			End With

	End If     
	
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
	Dim lDelRows, lDelRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

	If frm1.rdoWP.checked Then 	
	    If Frm1.vspdData.MaxRows < 1 then
	       Exit function
		End if	
		
	    With Frm1.vspdData 
	    	.focus
	    	ggoSpread.Source = frm1.vspdData 
	    	lDelRows = ggoSpread.DeleteRow
	    	
	    End With
	Else
		If Frm1.vspdData2.MaxRows < 1 then
	       Exit function
		End if	
		
	    With Frm1.vspdData2 
	    	.focus
	    	ggoSpread.Source = frm1.vspdData2 
	    	lDelRows = ggoSpread.DeleteRow
	    	
	    End With

	End IF
    lgBlnFlgChgValue = True 
   
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function


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
    Call parent.FncExport(parent.C_SINGLEMULTI)												<%'☜: 화면 유형 %>
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()

    Dim IntRetCD
    
    FncExit = False
    If frm1.rdoWP.checked Then 
		ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
		If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
    Else
		ggoSpread.Source = frm1.vspdData2							'⊙: Preset spreadsheet pointer 
		If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
    End If
    
    FncExit = True
    
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
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================

Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function


'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================

Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function
'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'******************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery()
	Dim strCode, strhCode ,strGubun,strhGubun,strNextKey
    
    DbQuery = False
    
    Call LayerShowHide(1)

    Dim strVal
    Dim sStartDt,sYear,sMon,sDay
    
    Call parent.ExtractDateFromSuper(frm1.txtYYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)	
	sStartDt= (sYear&sMon)
    
    
    If frm1.rdoWP.checked Then 
		strCode= trim(frm1.txtWPCd.value)
		strhCode= trim(frm1.hCode.value)
		strGubun = frm1.rdoWP.Value
		strhGubun= frm1.hGubun.value
		strNextKey = lgStrPrevKey 
	Else
		strCode= trim(frm1.txtCCCd.value)
		strhCode= trim(frm1.hCode.value)
		strGubun = frm1.rdoCC.Value
		strhGubun= frm1.hGubun.value
		strNextKey = lgStrPrevKey2 
	End If

	
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & strNextKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtYYYYMM=" & Trim(frm1.hYYYYMM.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCode=" & strhCode											'☆: 조회 조건 데이타 
		strVal = strVal & "&txtFctrCd=" & Trim(frm1.hFctrCd.value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtGubun=" & strhGubun					'☆: 조회 조건 데이타		

	Else
		strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & strNextKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtYYYYMM=" & Trim(sStartDt)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCode=" & strCode											'☆: 조회 조건 데이타 
		strVal = strVal & "&txtFctrCd=" & Trim(frm1.txtFctrCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtGubun=" & strGubun				'☆: 조회 조건 데이타	

	End If
'msgbox strval & "::"
    Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 
    DbQuery = True                                                          	'⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(ByVal iRow, ByVal LngMaxRow)															'☆: 조회 성공후 실행로직	
	
	Call SetToolBar("11001111001111")											'⊙: 버튼 툴바 제어	
   '-----------------------
    'Reset variables area
    '-----------------------
    If frm1.hGubun.value ="C" Then 
		Call SetQuerySpreadColor("B",iRow)
	Else
		Call SetQuerySpreadColor("A",iRow)

	End If		
		
    lgIntFlgMode = parent.OPMD_UMODE										'⊙: Indicates that current mode is Update mode

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
	Dim lRow        
    Dim lGrpCnt    
    Dim strVal
	Dim	strChangeFlag
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size

	
    DbSave = False                                                          	'⊙: Processing is NG
    
    Call LayerShowHide(1)
    
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1	
	strCUTotalvalLen = 0 : strDTotalvalLen  = 0

    '-----------------------
    'Data manipulate area
    '-----------------------
If frm1.rdoWP.checked Then				'wp  
	frm1.hGubun.value=frm1.rdoWP.value
    For lRow = 1 To .vspdData.MaxRows
		
		strVal = ""
		
        .vspdData.Row = lRow
        .vspdData.Col = 0
			
		Select Case .vspdData.Text
		
			Case ggoSpread.UpdateFlag
				strVal = strVal & "U" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case ggoSpread.InsertFlag
				strVal = strVal & "C" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case ggoSpread.DeleteFlag
				strVal = strVal & "D" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case Else				
				strChangeFlag = "N"
		End Select

		If strChangeFlag = "Y" Then 
			strVal = strVal &lRow & iColSep																					
			.vspdData.Col = C_WPCd
			strVal = strVal & Trim(.vspdData.Text) & iColSep			
			strVal = strVal & "M" & iColSep
			.vspdData.Col = C_FctrCd
			strVal = strVal & Trim(.vspdData.Text) & iColSep			
			.vspdData.Col = C_AllocData
			strVal = strVal & Trim(.vspdData.Text) & iColSep 
			'row count
			strVal = strVal & lRow & parent.gRowSep			

		End If
        
        .vspdData.Col = 0
		Select Case .vspdData.Text
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
				    
		         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		            objTEXTAREA.name = "txtCUSpread"
					objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
				 
		           iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
				       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
				      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
				         
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
				         
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strVal) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
				          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
				       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
				         
		         iTmpDBuffer(iTmpDBufferCount) =  strVal         
		         strDTotalvalLen = strDTotalvalLen + Len(strVal)
				         
		End Select
                
    Next
Else				'c/c
frm1.hGubun.value=frm1.rdoCC.value
	For lRow = 1 To .vspdData2.MaxRows
		
		strVal = ""
		
        .vspdData2.Row = lRow
        .vspdData2.Col = 0
			
		Select Case .vspdData2.Text
		
			Case ggoSpread.UpdateFlag
				strVal = strVal & "U" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case ggoSpread.InsertFlag
				strVal = strVal & "C" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case ggoSpread.DeleteFlag
				strVal = strVal & "D" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case Else				
				strChangeFlag = "N"
		End Select

		If strChangeFlag = "Y" Then 
			strVal = strVal & lRow & iColSep	
			.vspdData2.Col = C_CCCd2
			strVal = strVal & Trim(.vspdData2.Text) & iColSep																				

			strVal = strVal & "M" & iColSep
	
			.vspdData2.Col = C_FctrCd2
			strVal = strVal & Trim(.vspdData2.Text) & iColSep			
			.vspdData2.Col = C_AllocData2
			strVal = strVal & Trim(.vspdData2.Text) & iColSep 
			'row count
			strVal = strVal & lRow & parent.gRowSep			

		End If
        
        .vspdData2.Col = 0
		Select Case .vspdData2.Text
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
				    
		         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		            objTEXTAREA.name = "txtCUSpread"
					objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
				 
		           iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
				       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
				      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
				         
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
				         
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strVal) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
				          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
				       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
				         
		         iTmpDBuffer(iTmpDBufferCount) =  strVal         
		         strDTotalvalLen = strDTotalvalLen + Len(strVal)
				         
		End Select
                
    Next
End If
    
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   	

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'☜: 비지니스 ASP 를 가동 
		
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
	Call InitVariables
	If frm1.hGubun.value="W" Then
		ggoSpread.source = frm1.vspddata
		frm1.vspdData.MaxRows = 0	
    Else
		ggoSpread.source = frm1.vspddata2
		frm1.vspdData2.MaxRows = 0	   
    End If
    lgBlnFlgChgValue = False    
    
    Call RemovedivTextArea
    Call MainQuery()

End Function


'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'----------  Coding part  -------------------------------------------------------------
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lSprdNo,lRow, lCol)
	If lSprdNo="A" Then 
		frm1.vspdData.focus
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = lCol
		frm1.vspdData.Action = 0
		frm1.vspdData.SelStart = 0
		frm1.vspdData.SelLength = len(frm1.vspdData.Text)
	Else
		frm1.vspdData2.focus
		frm1.vspdData2.Row = lRow
		frm1.vspdData2.Col = lCol
		frm1.vspdData2.Action = 0
		frm1.vspdData2.SelStart = 0
		frm1.vspdData2.SelLength = len(frm1.vspdData2.Text)
	
	End If
End Function

'===========================================================================================================
' Description : checkWPCd ;check valid wccd
'===========================================================================================================
Function checkWPCd(ByVal pvLngRow, ByVal pvStrData, ByVal iIndex)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrWcCdInf
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	
	checkWPCd = False
	If pvStrData<>"" Then 
		If frm1.rdoCC.checked Then 	
			iStrSelectList = " COST_NM "
			iStrFromList   = "  b_cost_center  "		
			iStrWhereList =  "  COST_CD =  " & FilterVar(pvStrData , "''", "S") 

			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X",frm1.txtCCCd.alt,"X")
				frm1.vspdData2.Row=pvLngRow :frm1.vspdData2.Col = C_CCNM2 : frm1.vspdData2.Text =""
				Call SetActiveCell(frm1.vspdData2,C_CCCD2,pvLngRow,"M","X","X")			
				checkWPCd = False
				Exit Function
			End If	
			With frm1.vspdData2
				iArrWcCdInf = split(lgF0,chr(11))
				.Row = pvLngRow
				.Col = C_CCNm2	:  .text = Trim(iArrWcCdInf(0))			
			End With
	
		Else	

			iStrSelectList = " code, CD_NM "
			iStrFromList   =  "(select  'M' FLAG,wc_cd as code, wc_nm as cd_nm	 from P_work_center "
			iStrFromList = iStrFromList & " union "
			iStrFromList = iStrFromList & "	select  'O' FLAG,pur_grp as code, pur_grp_nm as cd_nm from b_pur_grp where usage_flg='Y') tmp"

			iStrWhereList =  "   CODE =  " & FilterVar(pvStrData , "''", "S") 
			If iIndex<>0 Then 
				iStrWhereList = iStrWhereList & " AND FLAG ='O' "
			Else
				iStrWhereList = iStrWhereList & " AND FLAG ='M' "
			End If

			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X",frm1.txtWPCd.alt,"X")
				frm1.vspdData.Row=pvLngRow :frm1.vspdData.Col = C_WPNm : frm1.vspdData.Text =""
				Call SetActiveCell(frm1.vspdData,C_WPCd,pvLngRow,"M","X","X")			
				checkWPCd = False
				Exit Function
			End If	
			With frm1.vspdData
				iArrWcCdInf = split(lgF0,chr(11))
				.Row = pvLngRow
				.Col = C_WPNm	:  .text = Trim(iArrWcCdInf(1))			
			End With
	
	
		End IF

		checkWPCd = True
	End IF
End Function

'===========================================================================================================
' Description : checkFctrCd  ; check valid prod order no 
'===========================================================================================================
Function checkFctrCd(ByVal pvLngRow, ByVal pvStrData)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrProdNoInf
		
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	
	checkFctrCd = False	
	
	iStrWhereList  =" dstb_fctr_cd = " & FilterVar(trim(pvStrData), "''", "S") & " "

	Call CommonQueryRs("dstb_fctr_cd, dbo.ufn_getCodeName('C4000',dstb_fctr_cd) as code_nm "," c_dstb_fctr_s " , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If frm1.rdoCC.checked Then
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtFctrCd.alt,"X")
			Call SetActiveCell(frm1.vspdData2,C_FctrCd2,pvLngRow,"M","X","X")			
			checkFctrCd = False
			Exit Function
		End If	
	
		With frm1.vspdData
			iArrProdNoInf = split(lgF0,chr(11))
			.Row = pvLngRow
			.Col = C_FctrCd2	: .text = Trim(iArrProdNoInf(0))
			.Col = C_FctrNm2	: .text = Trim(iArrProdNoInf(1))						
		End With
	Else
			IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtFctrCd.alt,"X")
			Call SetActiveCell(frm1.vspdData,C_FctrCd,pvLngRow,"M","X","X")			
			checkFctrCd = False
			Exit Function
		End If	
	
		With frm1.vspdData
			iArrProdNoInf = split(lgF0,chr(11))
			.Row = pvLngRow
			.Col = C_FctrCd	: .text = Trim(iArrProdNoInf(0))			
			.Col = C_FctrNM	: .text = Trim(iArrProdNoInf(1))		
		End With
	End If
	checkFctrCd = True
End Function

'===========================================================================================================
' Description :spread popup button 
'===========================================================================================================
Function OpenSpreadPopup(ByVal pvLngCol, ByVal pvLngRow, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)
	Dim iIndex
	OpenSpreadPopup = False
	
	If IsOpenPop Then Exit Function

	IsOpenPop = True
	
	frm1.vspdData.Col=C_ContractFlg : frm1.vspdDAta.Row=pvLngRow : iIndex=frm1.vspdData.Value
	
	Select Case pvLngCol
		Case C_WPPop
			iArrParam(1) = "(select 'M' flag, '사내' AS FLAG_NM, wc_cd as code, wc_nm as cd_nm	 from P_work_center "  		<%' TABLE 명칭 %>
			iArrParam(1) = iArrParam(1) & " union "
			iArrParam(1) = iArrParam(1) & "	select 'O' flag, '외주가공' AS FLAG_NM,pur_grp as code, pur_grp_nm as cd_nm from b_pur_grp where usage_flg='Y') tmp"
			
			iArrParam(2) =pvStrData					<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			If iIndex=0 then 
				iArrParam(4) = " FLAG='M' "'AND CODE >= " 	& filtervar(	pvStrData,"''","S")			<%' Where Condition%>
			Else
				iArrParam(4) = " FLAG='O' "'AND CODE >= " 	& filtervar(	pvStrData,"''","S")			<%' Where Condition%>
			End IF
			
			iArrParam(5) = "공정/구매그룹"						<%' TextBox 명칭 %>
			
			iArrField(0) = "HH" & Parent.gColSep & "CODE"	
			iArrField(1) = "HH" & Parent.gColSep & "FLAG"	
			iArrField(2) = "ED18" & Parent.gColSep & "FLAG_NM"	
			iArrField(3) = "ED15" & Parent.gColSep & "CODE"
			iArrField(4) = "ED30" & Parent.gColSep & "CD_NM"
			
			iArrHeader(0) = "공정/구매그룹"    
			iArrHeader(1) = "사내/외주가공구분"
			iArrHeader(2) = "사내/외주가공구분"
			iArrHeader(3) = "공정/구매그룹"
			iArrHeader(4) = "공정/구매그룹명"
			
		Case C_CCPop2
			iArrParam(1) = " b_cost_center "
			
			iArrParam(2) = pvStrData					<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>		
			iArrParam(4) = "" 					<%' Where Condition%>
			
			iArrParam(5) = "C/C"						<%' TextBox 명칭 %>
			
			iArrField(0) = "ED10" & Parent.gColSep & "COST_CD"	
			iArrField(1) = "ED25" & Parent.gColSep & "COST_NM"			
			    
			iArrHeader(0) = "C/C"
			iArrHeader(1) = "C/C명"
			
		Case C_FctrCdPop,C_FctrCdPop2
			
			iArrParam(1) = "c_dstb_fctr_s  "
			
			iArrParam(2) = pvStrData					<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>

			iArrParam(4) = " gen_flag='M'    "

			
			iArrParam(5) = "배부요소"						<%' TextBox 명칭 %>
			
			iArrField(0) ="ED12" & Parent.gColSep &  "dstb_fctr_cd "	
			iArrField(1) ="ED30" & Parent.gColSep & "dbo.ufn_getCodeName('C4000',dstb_fctr_cd) CD_NM "		
						    
			iArrHeader(0) = "배부요소"
			iArrHeader(1) = "배부요소명"

	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=520px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSpreadPopup = SetSpreadPopup(iArrRet,pvLngCol, pvLngRow)
	End If	

End Function
'===========================================================================================================
' Description : set spread popup 
'===========================================================================================================
Function SetSpreadPopup(Byval pvArrRet,ByVal pvLngCol, ByVal pvLngRow)
	SetSpreadPopup = False

	With frm1
	
		If .rdoWP.checked Then 				
		Select Case pvLngCol
			Case C_WPPop
				.vspdData.Row = pvLngRow
				.vspdData.Col = C_WPCd	: .vspdData.Text = pvArrRet(3)
				.vspdData.Col = C_WPNm	: .vspdData.Text = pvArrRet(4)	
			Case C_FctrCdPop
				.vspdData.Row = pvLngRow
				.vspdData.Col = C_FctrCd : .vspdData.Text = PvArrRet(0)
				.vspdData.Col = C_FctrNM : .vspdData.Text = PvArrRet(1)
		End Select
		Else
			Select Case pvLngCol
			Case C_FctrCdPop2
				.vspdData2.Row = pvLngRow
				.vspdData2.Col = C_FctrCd2 : .vspdData2.Text = PvArrRet(0)
				.vspdData2.Col = C_FctrNM2 : .vspdData2.Text = PvArrRet(1)
			CASE C_CCPop2
				 .vspdData2.Row = pvLngRow
				 .vspdData2.Col = C_CCCd2	: .vspdData2.Text = pvArrRet(0)
				 .vspdData2.Col = C_CCNm2	: .vspdData2.Text = pvArrRet(1)	
			End Select
		End If

	End With

	SetSpreadPopup = True
End Function

'========================================================================================================
Sub btnCopyPrev_OnClick()

	If BtnSpreadCheck = False Then Exit Sub

	Err.Clear                                                        

	If  CheckExistData1() Then 
		Call CheckExistData2()
	End If	
	frm1.txtFctrCd.focus()

End Sub
'===========================================================================================================
' Description : CheckExistData ;Check Exist about the previous data 
'===========================================================================================================
Function CheckExistData1()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iTmp
	Dim IntRetCD
	
	Dim PrevDate
	
	CheckExistData1=FALSE
	
	PrevDate	= UNIDateAdd("m", -1, frm1.txtYYYYMM.Text, parent.gDateFormat)
	frm1.txtYYYYMM2.value = replace(left(PrevDate,7),"-","")
		
	iStrSelectList = " top 1 yyyymm "
	If frm1.rdoCC.checked then
		iStrFromList   = " c_mfc_alloc_basis_by_cc_s "
	Else
		iStrFromList   = " c_mfc_alloc_basis_by_opr_s "
	End If
	iStrWhereList  =iStrWhereList & " yyyymm = " & FilterVar(replace(left(PrevDate,7),"-",""), "''", "S")	

	Err.Clear

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		CheckExistData1=TRUE
		Exit Function 
	Else   
		If Err.number = 0 Then   'Data is not exist.
			 Call DisplayMsgBox("236306","X" , "X","X")
			 CheckExistData1=FALSE
		Else								'Err.
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If

End Function
'===========================================================================================================
' Description : CheckExistData2;Check exist about current data
'===========================================================================================================
Function CheckExistData2()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iTmp
	Dim IntRetCD
	
	
	iStrSelectList = " top 1 yyyymm "
	If frm1.rdoCC.checked then
		iStrFromList   = " c_mfc_alloc_basis_by_cc_s "
	Else
		iStrFromList   = " c_mfc_alloc_basis_by_opr_s "
	End If
	iStrWhereList  =iStrWhereList & " yyyymm = " & FilterVar(replace(frm1.txtYYYYMM.Text,"-",""), "''", "S")	

	Err.Clear

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		IntRetCD = DisplayMsgBox("900007", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then 
			Exit Function
		ELSE
			Call CopyPrevData()		
		END IF
	Else   
		If Err.number = 0 Then   'Data is not exist.
			Call CopyPrevData()
		Else								'Err.
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If
	
End Function
'========================================================================================================
' Description : CopyPrevData;Copy data
'===========================================================================================================
Sub CopyPrevData()
	
	Dim iStrVal
	If frm1.rdoCC.checked Then 
		frm1.hGubun.value="C" 
	Else
		frm1.hGubun.value="W" 
	End If

	iStrVal = BIZ_PGM_ID & "?txtMode=" & "btnCopyPrev"					
	iStrVal = iStrVal & "&txtGubun=" & Trim(frm1.hGubun.value)
	iStrVal = iStrVal & "&txtYYYYMM1=" & Trim(frm1.txtYYYYMM.Text)
	iStrVal = iStrVal & "&txtYYYYMM2=" & Trim(frm1.txtYYYYMM2.value)		

	Call RunMyBizASP(MyBizASP, iStrVal)          

End Sub

'========================================================================================================
' Description : BtnSpreadCheck;Check changed data before anyother event
'===========================================================================================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData 

	 '--case multi -- %>
	 'when changed data exist asking what to do  %>
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	 'nothing changed  %>
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function


'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		

	strFrom = "  (select wc_cd as code, wc_nm as cd_nm	 from P_work_center "
	strFrom = strFrom & " union "
	strFrom = strFrom & "	select pur_grp as code, pur_grp_nm as cd_nm from b_pur_grp where usage_flg='Y') tmp  "

'check wc cd	
	If Trim(frm1.txtWPCd.value) <> "" Then
		strWhere = " Code = " & FilterVar(frm1.txtWPCd.value, "''", "S") & " "		
		
		Call CommonQueryRs(" cd_nm  ",strFrom, strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtWPCd.alt,"X")			
			frm1.txtWPNm.value = ""
			ChkKeyField = False
			frm1.txtWPCd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtWPNm.value = strDataNm(0)
	else
		frm1.txtWPNm.value=""
	End If
	'check cc cd	
	If Trim(frm1.txtccCd.value) <> "" Then
		strFrom ="b_cost_center"
		strWhere = " cost_cd = " & FilterVar(frm1.txtccCd.value, "''", "S") & " "		
		
		Call CommonQueryRs(" cost_nm  ",strFrom, strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtCCCd.alt,"X")			
			frm1.txtCCNm.value = ""
			ChkKeyField = False
			frm1.txtCCCd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtCCNm.value = strDataNm(0)
	else
		frm1.txtCCNm.value=""
	End If
	
'check prod order no	
	If Trim(frm1.txtFctrCd.value) <> "" Then
		strFrom = " c_dstb_fctr_s  "
		strWhere = " gen_flag='M' and dstb_fctr_cd= "		& filterVar(Trim(frm1.txtFctrCd.value),"","S")
		
		Call CommonQueryRs(" dbo.ufn_getCodeName('C4000'," & filterVar(Trim(frm1.txtFctrCd.value),"","S") & ")  ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtFctrCd.alt,"X")			
			frm1.txtFctrNM.value = ""
			ChkKeyField = False
			frm1.txtFctrCd.focus 
			Exit Function
		End If	
		strDataNm = split(lgF0,chr(11))
		frm1.txtFctrNm.value = strDataNm(0)
	else
		frm1.txtFctrNm.value=""
	End If
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 BORDER=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>배부요소DATA등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
								<TD CLASS=TD5 NOWRAP>작업년월</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYYYYMM CLASS=FPDTYYYYMM title=FPDATETIME tag="12" ALT="작업년월" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>구분</TD>
									<TD CLASS=TD6 NOWRAP><input type="radio" id="rdoCC" name="rdoGubun" CLASS="RADIO" tag="11" Value="C" checked >C/C
																			<input type="radio" id="rdoWP" name="rdoGubun"  CLASS="RADIO" tag="11" Value="W">공정/구매그룹</TD>
									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>배부요소</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFctrCd" SIZE=10 MAXLENGTH=2 tag="11xxxU" ALT="배부요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFctr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFctrCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFctrNm" SIZE=25 tag="14"></TD></TD>
									<TD CLASS=TD5 NOWRAP><div id=divCClbl>C/C</div>
									<div id=divWplbl style="display:none;">공정/구매그룹</div></TD>
									<TD CLASS=TD6 NOWRAP>
										<div id="tmpCc">
										<INPUT TYPE=TEXT NAME="txtCCCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCCCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCCCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtCCNm" SIZE=25 tag="14">
										</div>
										<div id="tmpWp" style="display:none;">
										<INPUT TYPE=TEXT NAME="txtWPCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="공정/구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWPCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtWPNm" SIZE=25 tag="14">
										</div>
									</td>
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
					<div id="sprdWP"   style="display:none;">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = "A" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>	
								</TD>
							</TR>
						</TABLE>
						</div>
						<div id="sprdCC">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 ID = "B" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD2"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>	
								</TD>
							</TR>
						</TABLE>
						</div>
						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
			  <TD WIDTH=10>&nbsp;</TD>
			  <TD>
			   <BUTTON NAME="btnCopyPrev" CLASS="CLSSBTN">전월COPY</BUTTON>&nbsp;
			   </TD>
			   </TR>
			  </TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hCode" tag="24"><INPUT TYPE=HIDDEN NAME="hFctrCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hGubun" tag="24">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24"><INPUT TYPE=HIDDEN NAME="txtYYYYMM2" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
