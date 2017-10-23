<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1411MA1
'*  4. Program Name         : 전문가 시스템 초기화면 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                                             '☜: indicates that All variables must be declared in advance 

Const BIZ_PGM_QRY_ID = "Q1411MB1.asp"										 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "Q1411MB2.asp"										 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID = "Q1411MB3.asp"
Const BIZ_PGM_JUMP_ID="Q1412MA1.asp"

Const TAB1 = 1
Const TAB2 = 2


Dim lgNextNo					'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo					' ""

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgMpsFirmDate
Dim lgLlcGivenDt								
Dim IsOpenPop          
Dim gSelframeFlg

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE                                               	'⊙: Indicates that current mode is Create mode
	lgIntGrpCount = 0                                                     	  	'⊙: Initializes Group View Size
	'----------  Coding part  -------------------------------------------------------------
	gIsTab= "Y"
	gTabMaxCnt=2

	IsOpenPop = False						'☆: 사용자 변수 초기화 
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()

End Sub

'===========================================  2.3.1 Tab Click 처리  =====================================
'=	Name : Tab Click																					=
'=	Description : Tab Click시 필요한 기능을 수행한다.													=
'========================================================================================================
Function ClickTab1()
	ClickTab1 = false

	If gSelframeFlg = TAB1 Then Exit Function
			
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1

	ClickTab1 = true
End Function

'===========================================  2.3.1 Tab ClickTab2 처리  =====================================
'=	Name : Tab ClickTab2																					=
'=	Description : Tab ClickTab2시 필요한 기능을 수행한다.													=
'========================================================================================================
Function ClickTab2()
	Dim ret
	ClickTab2 = false
	
	If gSelframeFlg = TAB2 Then Exit Function
		
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	
	With frm1

		IF .rdoCase61.Checked= true then
			Q26.Style.display = ""			
		End IF		
				
		IF .rdoCase62.Checked= true then
			Q26.Style.display = "none"
			.rdoCase261.checked = False
			.rdoCase262.checked = False
		End IF		
	
		IF .rdoCase131.Checked= true then
			If .rdoCase82.Checked=true and .rdoCase123.Checked=true Then
				Q23.Style.display = ""
				Q24.Style.display = ""
				Q25.Style.display = ""
			Else
				Q23.Style.display = "none"
				Q24.Style.display = "none"
				Q25.Style.display = "none"
				.rdoCase231.checked = False
				.rdoCase232.checked = False
				.rdoCase241.checked = False
				.rdoCase242.checked = False
				.rdoCase251.checked = False
				.rdoCase252.checked = False
			End If
		End IF		
					
		IF .rdoCase132.Checked= true then
			Q23.Style.display = "none"
			Q24.Style.display = "none"
			Q25.Style.display = "none"
			.rdoCase231.checked = False
			.rdoCase232.checked = False
			.rdoCase241.checked = False
			.rdoCase242.checked = False
			.rdoCase251.checked = False
			.rdoCase252.checked = False
		End IF		
	
		IF .rdoCase133.Checked= true then
			Q23.Style.display = "none"
			Q24.Style.display = "none"
			Q25.Style.display = "none"
			.rdoCase231.checked = False
			.rdoCase232.checked = False
			.rdoCase241.checked = False
			.rdoCase242.checked = False
			.rdoCase251.checked = False
			.rdoCase252.checked = False
		End IF		

		IF .rdoCase261.Checked= true then							
			Q27.Style.display = ""
			Q28.Style.display = ""
			Q29.Style.display = ""
			Q30.Style.display = ""
			Q31.Style.display = ""
			Q32.Style.display = ""
		End IF
					
		IF .rdoCase262.Checked= true then							
			Q27.Style.display = "none"
			Q28.Style.display = "none"
			Q29.Style.display = "none"
			Q30.Style.display = "none"
			Q31.Style.display = "none"
			Q32.Style.display = "none"
			.rdoCase271.checked = False
			.rdoCase272.checked = False
			.rdoCase281.checked = False
			.rdoCase282.checked = False
			.rdoCase291.checked = False
			.rdoCase292.checked = False
			.rdoCase301.checked = False
			.rdoCase302.checked = False
			.rdoCase311.checked = False
			.rdoCase312.checked = False
			.rdoCase321.checked = False
			.rdoCase322.checked = False
		End IF
	End With
	ClickTab2 = true
End Function

'------------------------------------------  OpenPlant1()  -------------------------------------------------
'	Name : OpenPlant()
'	Description :Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant1()
	OpenPlant1 = false
End Function

'------------------------------------------  OpenSavedInspReqNo()  -------------------------------------------------
'	Name : OpenSavedInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 

Function OpenSavedInspReqNo()        
	OpenSavedInspReqNo = false
End Function

'------------------------------------------  OpenNewInspReqNo()  -------------------------------------------------
'	Name : OpenNewInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenNewInspReqNo()        
	OpenNewInspReqNo = false
		
End Function

'------------------------------------------  OpenNewInspReqNo()  -------------------------------------------------
'	Name : OpenNewInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenNewInspReqNo()        
	OpenNewInspReqNo = false
End Function

'------------------------------------------  OpenInspector()  -------------------------------------------------
'	Name : OpenInspector()
'	Description : Inspector PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspector()
	OpenInspector = false

End Function

'------------------------------------------ OpenAct()  -------------------------------------------------
'	Name : OpenAct()
'	Description : Act PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAct()
	OpenAct = false
End Function

'------------------------------------------  SetInspReqNo1()  --------------------------------------------------
'	Name : SetInspReqNo1()
'	Description : Move Type Conf. Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetInspReqNo1(Byval arrRet)
	SetInspReqNo1 = false
End Function

'------------------------------------------  SetInspReqNo2()  --------------------------------------------------
'	Name : SetInspReqNo2()
'	Description : Move Type Conf. Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetInspReqNo2(Byval arrRet)
	SetInspReqNo2 = false
End Function

'------------------------------------------  SetInspector()  --------------------------------------------------
'	Name : SetInspector()
'	Description : Move Type Conf. Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------- 
Function SetInspector(Byval arrRet)
	SetInspector = false
End Function	

'------------------------------------------  SetActCd()  --------------------------------------------------
'	Name : SetActCd()
'	Description : Move Type Conf. Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetAct(Byval arrRet)
	SetAct = false
End Function

'------------------------------------------  SetPlant1()  --------------------------------------------------
'	Name : SetPlant1()
'	Description : Move Type Conf. Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant1(Byval arrRet)
	SetPlant1 = false
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ %>
Function CURadio_Click(ObjRadio)
	CURadio_Click = false
	If ObjRadio.checked Then
	 	ObjRadio.checked = False
	Else
	 	ObjRadio.checked = True
	End If
	CURadio_Click = true
End Function

'------------------------------------------  Condition61()  --------------------------------------------------
'	Name : Condition61()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function Condition61()		
	Condition61 = false
End Function

'------------------------------------------  Condition62()  --------------------------------------------------
'	Name : Condition62()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function Condition62()	
	Condition62 = false
'	Q26.Style.display = "none"
End Function

'------------------------------------------  Condition131()  --------------------------------------------------
'	Name : Condition131()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function Condition131()	
	Condition131 = false
End Function

'------------------------------------------  Condition132()  --------------------------------------------------
'	Name : Condition132()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function Condition132()	
	Condition132 = false
End Function

'------------------------------------------  Condition133()  --------------------------------------------------
'	Name : Condition133()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function Condition133()	
	Condition133 = false
End Function

'------------------------------------------  Condition261()  --------------------------------------------------
'	Name : Condition261()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function Condition261()	
	Condition261 = false
	
	Q27.Style.display = ""
	Q28.Style.display = ""
	Q29.Style.display = ""
	Q30.Style.display = ""
	Q31.Style.display = ""
	Q32.Style.display = ""
	
	Condition261 = true
End Function

'------------------------------------------  Condition262()  --------------------------------------------------
'	Name : Condition262()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function Condition262()	
	Condition262 = false
	
	Q27.Style.display = "none"
	Q28.Style.display = "none"
	Q29.Style.display = "none"
	Q30.Style.display = "none"
	Q31.Style.display = "none"
	Q32.Style.display = "none"
	
	With frm1
		.rdoCase271.checked = False
		.rdoCase272.checked = False
		.rdoCase281.checked = False
		.rdoCase282.checked = False
		.rdoCase291.checked = False
		.rdoCase292.checked = False
		.rdoCase301.checked = False
		.rdoCase302.checked = False
		.rdoCase311.checked = False
		.rdoCase312.checked = False
		.rdoCase321.checked = False
		.rdoCase322.checked = False
	End With
		
	Condition262 = false
End Function

'=============================================  2.5.1 JumpResultReport()  ======================================
'=	Event Name : JumpResultReport
'=	Event Desc :
'========================================================================================================
Function LoadExportCharge()
	Dim intRet
	Dim strChoice
	
	Dim gs_coeff(25, 12)   		'적합도 계수 
	Dim gs_coeff2(7, 2)
	Dim gsel_insp(12)
	Dim gsel_insp2(2)
	Dim gmax_ins(12)			'모든 검사방식 저장 
	
	Dim i 
    Dim j 
    Dim min 
    Dim max_sel 
    Dim max_insp 
    Dim tempsel 
    Dim tempinsp
    	
    Dim InsType(12)			'전문가시스템결과 선정된 검사방식 
	Dim FitnessDegree(12)		'전문가시스템결과 선정된 검사방식의 적합도계수 
	Dim AInsType(5)			'전문가시스템결과 적용가능한 검사방식 
	DIm InsAssureance(5)		'보증방식 
	Dim AFitnessDegree(5)		'전문가시스템결과 적용가능한 검사방식의 적합도계수 
	Dim InsCase			'계량형 계수형 결정 저장 변수 
	
	Dim InsType1			'전문가시스템결과 선정된 검사방식 저장 변수 
	Dim InsType2			'전문가시스템결과 선정된 검사방식 저장 변수 
	Dim InsType3			'전문가시스템결과 선정된 검사방식 저장 변수 
	Dim InsType4			'전문가시스템결과 선정된 검사방식 저장 변수 
	Dim InsType5			'전문가시스템결과 선정된 검사방식 저장 변수 
	
	Dim FitnessDegree1			'전문가시스템결과 선정된 검사방식의 적합도계수 저장 변수 
	Dim FitnessDegree2			'전문가시스템결과 선정된 검사방식의 적합도계수 저장 변수 
	Dim FitnessDegree3			'전문가시스템결과 선정된 검사방식의 적합도계수 저장 변수 
	Dim FitnessDegree4			'전문가시스템결과 선정된 검사방식의 적합도계수 저장 변수 
	Dim FitnessDegree5			'전문가시스템결과 선정된 검사방식의 적합도계수 저장 변수 
		
	DIm InsAssureance1
	DIm InsAssureance2
	DIm InsAssureance3
	DIm InsAssureance4
	DIm InsAssureance5
	Dim IntRetCD
			
	With frm1
		If .rdoCase11.checked = true then
			gs_coeff(0, 0) = 0
		    gs_coeff(0, 1) = 0
		    gs_coeff(0, 2) = 0
		    gs_coeff(0, 3) = 0
		    gs_coeff(0, 4) = (3 / 10)
		    gs_coeff(0, 5) = 0
		    gs_coeff(0, 6) = 0
		    gs_coeff(0, 7) = 0
		    gs_coeff(0, 8) = 0
		    gs_coeff(0, 9) = 0
		    gs_coeff(0, 10) = 0
		    gs_coeff(0, 11) = 0
		ElseIf .rdoCase12.checked = true then
			gs_coeff(0, 0) = 0
			gs_coeff(0, 1) = 0
			gs_coeff(0, 2) = 0
			gs_coeff(0, 3) = 0
			gs_coeff(0, 4) = (3 / 10)
			gs_coeff(0, 5) = 0
			gs_coeff(0, 6) = 0
			gs_coeff(0, 7) = 0
			gs_coeff(0, 8) = 0
			gs_coeff(0, 9) = (3 / 10)
			gs_coeff(0, 10) = 0
			gs_coeff(0, 11) = (-5 / 10)
		ElseIf .rdoCase13.checked = true then
			gs_coeff(0, 0) = 0
		    gs_coeff(0, 1) = 0
		    gs_coeff(0, 2) = 0
		    gs_coeff(0, 3) = 0
		    gs_coeff(0, 4) = (-3 / 10)
		    gs_coeff(0, 5) = (3 / 10)
		    gs_coeff(0, 6) = 0
		    gs_coeff(0, 7) = 0
		    gs_coeff(0, 8) = (-7 / 10)
		    gs_coeff(0, 9) = 0
		    gs_coeff(0, 10) = (-5 / 10)
		    gs_coeff(0, 11) = 0
		ElseIf .rdoCase14.checked = true then
	
		End If
	
		If .rdoCase21.checked = true then
			gs_coeff(1, 0) = 0
		    gs_coeff(1, 1) = (2 / 10)
		    gs_coeff(1, 2) = (2 / 10)
		    gs_coeff(1, 3) = (1 / 10)
		    gs_coeff(1, 4) = 0
		    gs_coeff(1, 5) = 0
		    gs_coeff(1, 6) = 0
		    gs_coeff(1, 7) = 0
		    gs_coeff(1, 8) = 0
		    gs_coeff(1, 9) = 0
		    gs_coeff(1, 10) = (2 / 10)
		    gs_coeff(1, 11) = -1
		ElseIf .rdoCase22.checked = true then
			gs_coeff(1, 0) = 0
		    gs_coeff(1, 1) = 0
		    gs_coeff(1, 2) = 0
		    gs_coeff(1, 3) = 0
		    gs_coeff(1, 4) = 0
		    gs_coeff(1, 5) = 0
		    gs_coeff(1, 6) = 0
		    gs_coeff(1, 7) = 0
		    gs_coeff(1, 8) = 0
		    gs_coeff(1, 9) = 0
		    gs_coeff(1, 10) = 0
		    gs_coeff(1, 11) = -1
		ElseIf .rdoCase23.checked = true then
			gs_coeff(1, 0) = 0
		    gs_coeff(1, 1) = 0
		    gs_coeff(1, 2) = 0
		    gs_coeff(1, 3) = 0
		    gs_coeff(1, 4) = 0
		    gs_coeff(1, 5) = 0
		    gs_coeff(1, 6) = 0
		    gs_coeff(1, 7) = 0
		    gs_coeff(1, 8) = 0
		    gs_coeff(1, 9) = 0
		    gs_coeff(1, 10) = 0
		    gs_coeff(1, 11) = (2 / 10)
		End If
	
		If .rdoCase31.checked = true then
			gs_coeff(9, 0) = (3 / 10)
		    gs_coeff(9, 1) = -1
		    gs_coeff(9, 2) = -1
		    gs_coeff(9, 3) = -1
		    gs_coeff(9, 4) = 0
		    gs_coeff(9, 5) = 0
		    gs_coeff(9, 6) = -1
		    gs_coeff(9, 7) = (2 / 10)
		    gs_coeff(9, 8) = (1 / 10)
		    gs_coeff(9, 9) = (1 / 10)
		    gs_coeff(9, 10) = -1
		    gs_coeff(9, 11) = (2 / 10)	
		ElseIf .rdoCase32.checked = true then
			gs_coeff(9, 0) = 0
		    gs_coeff(9, 1) = (4 / 10)
		    gs_coeff(9, 2) = (4 / 10)
		    gs_coeff(9, 3) = (4 / 10)
		    gs_coeff(9, 4) = (2 / 10)
		    gs_coeff(9, 5) = (3 / 10)
		    gs_coeff(9, 6) = -1
		    gs_coeff(9, 7) = (2 / 10)
		    gs_coeff(9, 8) = (4 / 10)
		    gs_coeff(9, 9) = (2 / 10)
		    gs_coeff(9, 10) = (1 / 10)
		    gs_coeff(9, 11) = (-2 / 10)
		ElseIf .rdoCase33.checked = true then
			gs_coeff(9, 0) = 0
		    gs_coeff(9, 1) = 0
		    gs_coeff(9, 2) = 0
		    gs_coeff(9, 3) = 0
		    gs_coeff(9, 4) = 0
		    gs_coeff(9, 5) = 0
		    gs_coeff(9, 6) = (9 / 10)
		    gs_coeff(9, 7) = 0
		    gs_coeff(9, 8) = 0
		    gs_coeff(9, 9) = 0
		    gs_coeff(9, 10) = 0
		    gs_coeff(9, 11) = 0
		End If
	
		If .rdoCase41.checked = true then
			gs_coeff(10, 0) = 0
		    gs_coeff(10, 1) = -1
		    gs_coeff(10, 2) = -1
		    gs_coeff(10, 3) = (-4 / 10)
		    gs_coeff(10, 4) = 0
		    gs_coeff(10, 5) = 0
		    gs_coeff(10, 6) = (-2 / 10)
		    gs_coeff(10, 7) = 0
		    gs_coeff(10, 8) = 0
		    gs_coeff(10, 9) = (4 / 10)
		    gs_coeff(10, 10) = -1
		    gs_coeff(10, 11) = (5 / 10)
		ElseIf .rdoCase42.checked = true then
			gs_coeff(10, 0) = 0
		    gs_coeff(10, 1) = (-6 / 10)
		    gs_coeff(10, 2) = (-6 / 10)
		    gs_coeff(10, 3) = 0
		    gs_coeff(10, 4) = 0
		    gs_coeff(10, 5) = 0
		    gs_coeff(10, 6) = 0
		    gs_coeff(10, 7) = 0
		    gs_coeff(10, 8) = 0
		    gs_coeff(10, 9) = (1 / 10)
		    gs_coeff(10, 10) = (-7 / 10)
		    gs_coeff(10, 11) = (2 / 10)
		ElseIf .rdoCase43.checked = true then
			gs_coeff(10, 0) = 0
		    gs_coeff(10, 1) = (2 / 10)
		    gs_coeff(10, 2) = (2 / 10)
		    gs_coeff(10, 3) = (3 / 10)
		    gs_coeff(10, 4) = 0
		    gs_coeff(10, 5) = 0
		    gs_coeff(10, 6) = (1 / 10)
		    gs_coeff(10, 7) = 0
		    gs_coeff(10, 8) = 0
		    gs_coeff(10, 9) = (1 / 10)
		    gs_coeff(10, 10) = 0
		    gs_coeff(10, 11) = (-2 / 10)
		ElseIf .rdoCase44.checked = true then
	
		End If
	
		If .rdoCase51.checked = true then
			gs_coeff(2, 0) = -1
		    gs_coeff(2, 1) = 0
		    gs_coeff(2, 2) = 0
		    gs_coeff(2, 3) = 0
		    gs_coeff(2, 4) = 0
		    gs_coeff(2, 5) = 0
		    gs_coeff(2, 6) = 0
		    gs_coeff(2, 7) = 0
		    gs_coeff(2, 8) = 0
		    gs_coeff(2, 9) = 0
		    gs_coeff(2, 10) = 0
		    gs_coeff(2, 11) = 0
		ElseIf .rdoCase52.checked = true then
			gs_coeff(2, 0) = -1
		    gs_coeff(2, 1) = 0
		    gs_coeff(2, 2) = 0
		    gs_coeff(2, 3) = 0
		    gs_coeff(2, 4) = 0
		    gs_coeff(2, 5) = 0
		    gs_coeff(2, 6) = 0
		    gs_coeff(2, 7) = (3 / 10)
		    gs_coeff(2, 8) = 0
		    gs_coeff(2, 9) = 0
		    gs_coeff(2, 10) = 0
		    gs_coeff(2, 11) = 0
		ElseIf .rdoCase53.checked = true then
			gs_coeff(2, 0) = (8 / 10)
		    gs_coeff(2, 1) = 0
		    gs_coeff(2, 2) = 0
		    gs_coeff(2, 3) = 0
		    gs_coeff(2, 4) = 0
		    gs_coeff(2, 5) = 0
		    gs_coeff(2, 6) = 0
		    gs_coeff(2, 7) = (3 / 10)
		    gs_coeff(2, 8) = (3 / 10)
		    gs_coeff(2, 9) = 0
		    gs_coeff(2, 10) = 0
		    gs_coeff(2, 11) = 0
		End If
	
		If .rdoCase61.checked = true then
			gs_coeff(3, 0) = 0
		    gs_coeff(3, 1) = (5 / 10)
		    gs_coeff(3, 2) = (5 / 10)
		    gs_coeff(3, 3) = (3 / 10)
		    gs_coeff(3, 4) = 0
		    gs_coeff(3, 5) = 0
		    gs_coeff(3, 6) = 0
		    gs_coeff(3, 7) = 0
		    gs_coeff(3, 8) = 0
		    gs_coeff(3, 9) = 0
		    gs_coeff(3, 10) = (5 / 10)
		    gs_coeff(3, 11) = 0
		ElseIf .rdoCase62.checked = true then
			gs_coeff(3, 0) = 0
		   	gs_coeff(3, 1) = -1
		   	gs_coeff(3, 2) = -1
		    gs_coeff(3, 3) = -1
		    gs_coeff(3, 4) = 0
		    gs_coeff(3, 5) = 0
		    gs_coeff(3, 6) = 0
		    gs_coeff(3, 7) = 0
		    gs_coeff(3, 8) = 0
		    gs_coeff(3, 9) = 0
		    gs_coeff(3, 10) = -1
		    gs_coeff(3, 11) = (1 / 10)
		End If
	
		If .rdoCase71.checked = true then
			gs_coeff(14, 0) = 0
		    gs_coeff(14, 1) = (2 / 10)
		    gs_coeff(14, 2) = (2 / 10)
		    gs_coeff(14, 3) = (2 / 10)
		    gs_coeff(14, 4) = 0
		    gs_coeff(14, 5) = 0
		    gs_coeff(14, 6) = 0
		    gs_coeff(14, 7) = 0
		    gs_coeff(14, 8) = 0
		    gs_coeff(14, 9) = 0
		    gs_coeff(14, 10) = (2 / 10)
		    gs_coeff(14, 11) = 0
		ElseIf .rdoCase72.checked = true then
			gs_coeff(14, 0) = 0
		    gs_coeff(14, 1) = (-5 / 10)
		    gs_coeff(14, 2) = (-5 / 10)
		    gs_coeff(14, 3) = (-3 / 10)
		    gs_coeff(14, 4) = 0
		    gs_coeff(14, 5) = (5 / 10)
		    gs_coeff(14, 6) = 0
		    gs_coeff(14, 7) = 0
		    gs_coeff(14, 8) = 0
		    gs_coeff(14, 9) = 0
		    gs_coeff(14, 10) = (-6 / 10)
		    gs_coeff(14, 11) = (1 / 10)
		ElseIf .rdoCase73.checked = true then
			gs_coeff(14, 0) = 0
			gs_coeff(14, 1) = (-5 / 10)
			gs_coeff(14, 2) = (-5 / 10)
			gs_coeff(14, 3) = (-3 / 10)
			gs_coeff(14, 4) = (5 / 10)
			gs_coeff(14, 5) = 0
			gs_coeff(14, 6) = (5 / 10)
			gs_coeff(14, 7) = 0
			gs_coeff(14, 8) = 0
			gs_coeff(14, 9) = 0
			gs_coeff(14, 10) = (-6 / 10)
			gs_coeff(14, 11) = (1 / 10)
		End If
	
		If .rdoCase81.checked = true then
			gs_coeff(8, 0) = 0
		    gs_coeff(8, 1) = -1
		    gs_coeff(8, 2) = -1
		    gs_coeff(8, 3) = (-5 / 10)
		    gs_coeff(8, 4) = 0
		    gs_coeff(8, 5) = 0
		    gs_coeff(8, 6) = 0
		    gs_coeff(8, 7) = 0
		    gs_coeff(8, 8) = 0
		    gs_coeff(8, 9) = (7 / 10)
		    gs_coeff(8, 10) = -1
		    gs_coeff(8, 11) = (8 / 10)
		ElseIf .rdoCase82.checked = true then
			gs_coeff(8, 0) = 0
		    gs_coeff(8, 1) = (2 / 10)
		    gs_coeff(8, 2) = (2 / 10)
		    gs_coeff(8, 3) = (1 / 10)
		    gs_coeff(8, 4) = 0
		    gs_coeff(8, 5) = 0
		    gs_coeff(8, 6) = 0
		    gs_coeff(8, 7) = 0
		    gs_coeff(8, 8) = 0
		    gs_coeff(8, 9) = 0
		    gs_coeff(8, 10) = (2 / 10)
		    gs_coeff(8, 11) = (-4 / 10)
		End If
	
		If .rdoCase91.checked = true then
			gs_coeff(20, 0) = 0
		    gs_coeff(20, 1) = 0
		    gs_coeff(20, 2) = 0
		    gs_coeff(20, 3) = 0
		    gs_coeff(20, 4) = 0
		    gs_coeff(20, 5) = 0
		    gs_coeff(20, 6) = 0
		    gs_coeff(20, 7) = 0
		    gs_coeff(20, 8) = 0
		    gs_coeff(20, 9) = 0
		    gs_coeff(20, 10) = 0
		    gs_coeff(20, 11) = 0
		ElseIf .rdoCase92.checked = true then
			gs_coeff(20, 0) = 0
		    gs_coeff(20, 1) = 0
		    gs_coeff(20, 2) = 0
		    gs_coeff(20, 3) = 0
		    gs_coeff(20, 4) = 0
		    gs_coeff(20, 5) = 0
		    gs_coeff(20, 6) = 0
		    gs_coeff(20, 7) = 0
		    gs_coeff(20, 8) = 0
		    gs_coeff(20, 9) = 0
		    gs_coeff(20, 10) = 0
		    gs_coeff(20, 11) = 0
		ElseIf .rdoCase93.checked = true then
			gs_coeff(20, 0) = 0
		    gs_coeff(20, 1) = 0
		    gs_coeff(20, 2) = 0
		    gs_coeff(20, 3) = 0
		    gs_coeff(20, 4) = 0
		    gs_coeff(20, 5) = 0
		    gs_coeff(20, 6) = 0
		    gs_coeff(20, 7) = 0
		    gs_coeff(20, 8) = 0
		    gs_coeff(20, 9) = 0
		    gs_coeff(20, 10) = 0
		    gs_coeff(20, 11) = 0
		End If
	
		If .rdoCase101.checked = true then
			gs_coeff(21, 0) = (1 / 10)
		    gs_coeff(21, 1) = (1 / 10)
		    gs_coeff(21, 2) = (1 / 10)
		    gs_coeff(21, 3) = (1 / 10)
		    gs_coeff(21, 4) = 0
		    gs_coeff(21, 5) = 0
		    gs_coeff(21, 6) = 0
		    gs_coeff(21, 7) = 0
		    gs_coeff(21, 8) = 0
		    gs_coeff(21, 9) = 0
		    gs_coeff(21, 10) = (1 / 10)
		    gs_coeff(21, 11) = 0
		ElseIf .rdoCase102.checked = true then
			gs_coeff(21, 0) = 0
		    gs_coeff(21, 1) = 0
		    gs_coeff(21, 2) = 0
		    gs_coeff(21, 3) = 0
		    gs_coeff(21, 4) = 0
		    gs_coeff(21, 5) = 0
		    gs_coeff(21, 6) = 0
		    gs_coeff(21, 7) = 0
		    gs_coeff(21, 8) = 0
		    gs_coeff(21, 9) = 0
		    gs_coeff(21, 10) = 0
		    gs_coeff(21, 11) = 0
		ElseIf .rdoCase103.checked = true then
			gs_coeff(21, 0) = (-2 / 10)
			gs_coeff(21, 1) = (-4 / 10)
			gs_coeff(21, 2) = (-4 / 10)
			gs_coeff(21, 3) = (-3 / 10)
			gs_coeff(21, 4) = (3 / 10)
			gs_coeff(21, 5) = (3 / 10)
			gs_coeff(21, 6) = (2 / 10)
			gs_coeff(21, 7) = (4 / 10)
			gs_coeff(21, 8) = (3 / 10)
			gs_coeff(21, 9) = (-3 / 10)
			gs_coeff(21, 10) = 0
			gs_coeff(21, 11) = 0
		End If
	
		If .rdoCase111.checked = true then
			gs_coeff(4, 0) = 0
		    gs_coeff(4, 1) = -1
		    gs_coeff(4, 2) = -1
		    gs_coeff(4, 3) = -1
		    gs_coeff(4, 4) = 0
		    gs_coeff(4, 5) = 0
		    gs_coeff(4, 6) = 0
		    gs_coeff(4, 7) = (9 / 10)
		    gs_coeff(4, 8) = 0
		    gs_coeff(4, 9) = 0
		    gs_coeff(4, 10) = -1
		    gs_coeff(4, 11) = (6 / 10)
		ElseIf .rdoCase112.checked = true then
			gs_coeff(4, 0) = 0
		    gs_coeff(4, 1) = -1
		    gs_coeff(4, 2) = -1
		    gs_coeff(4, 3) = (-9 / 10)
		    gs_coeff(4, 4) = 0
		    gs_coeff(4, 5) = 0
		    gs_coeff(4, 6) = 0
		    gs_coeff(4, 7) = 0
		    gs_coeff(4, 8) = 0
		    gs_coeff(4, 9) = 0
		    gs_coeff(4, 10) = -1
		    gs_coeff(4, 11) = 0
		ElseIf .rdoCase113.checked = true then
			gs_coeff(4, 0) = 0
		    gs_coeff(4, 1) = (3 / 10)
		    gs_coeff(4, 2) = (3 / 10)
		    gs_coeff(4, 3) = (3 / 10)
		    gs_coeff(4, 4) = 0
		    gs_coeff(4, 5) = 0
		    gs_coeff(4, 6) = 0
		    gs_coeff(4, 7) = 0
		    gs_coeff(4, 8) = 0
		    gs_coeff(4, 9) = 0
		    gs_coeff(4, 10) = (4 / 10)
		    gs_coeff(4, 11) = 0
		End If
	
		If .rdoCase121.checked = true then
			gs_coeff(6, 0) = 0
			gs_coeff(6, 1) = -1
			gs_coeff(6, 2) = -1
			gs_coeff(6, 3) = -1
			gs_coeff(6, 4) = 0
			gs_coeff(6, 5) = 0
			gs_coeff(6, 6) = 0
			gs_coeff(6, 7) = (9 / 10)
			gs_coeff(6, 8) = 0
			gs_coeff(6, 9) = 0
			gs_coeff(6, 10) = -1
			gs_coeff(6, 11) = (6 / 10)
		ElseIf .rdoCase122.checked = true then
		    gs_coeff(6, 0) = 0
		    gs_coeff(6, 1) = -1
		    gs_coeff(6, 2) = -1
		    gs_coeff(6, 3) = (-9 / 10)
		    gs_coeff(6, 4) = 0
		    gs_coeff(6, 5) = 0
		    gs_coeff(6, 6) = 0
		    gs_coeff(6, 7) = 0
		    gs_coeff(6, 8) = 0
		    gs_coeff(6, 9) = 0
		    gs_coeff(6, 10) = -1
		    gs_coeff(6, 11) = 0
		ElseIf .rdoCase123.checked = true then
			gs_coeff(6, 0) = 0
		    gs_coeff(6, 1) = (3 / 10)
		    gs_coeff(6, 2) = (3 / 10)
		    gs_coeff(6, 3) = (3 / 10)
		    gs_coeff(6, 4) = 0
		    gs_coeff(6, 5) = 0
		    gs_coeff(6, 6) = 0
		    gs_coeff(6, 7) = 0
		    gs_coeff(6, 8) = 0
		    gs_coeff(6, 9) = 0
		    gs_coeff(6, 10) = (4 / 10)
		    gs_coeff(6, 11) = 0
		End If
	
		If .rdoCase131.checked = true then
	      	gs_coeff(7, 0) = 0
		    gs_coeff(7, 1) = (5 / 10)
		    gs_coeff(7, 2) = (5 / 10)
		    gs_coeff(7, 3) = (4 / 10)
		    gs_coeff(7, 4) = 0
		    gs_coeff(7, 5) = 0
		    gs_coeff(7, 6) = 0
		    gs_coeff(7, 7) = 0
		    gs_coeff(7, 8) = 0
		    gs_coeff(7, 9) = 0
		    gs_coeff(7, 10) = (5 / 10)
		    gs_coeff(7, 11) = (-8 / 10)
		ElseIf .rdoCase132.checked = true then
			gs_coeff(7, 0) = 0
		    gs_coeff(7, 1) = (-7 / 10)
		    gs_coeff(7, 2) = (-7 / 10)
		    gs_coeff(7, 3) = (-4 / 10)
		    gs_coeff(7, 4) = 0
		    gs_coeff(7, 5) = 0
		    gs_coeff(7, 6) = 0
		    gs_coeff(7, 7) = 0
		    gs_coeff(7, 8) = 0
		    gs_coeff(7, 9) = 0
		    gs_coeff(7, 10) = (-6 / 10)
		    gs_coeff(7, 11) = 0
		ElseIf .rdoCase133.checked = true then
			gs_coeff(7, 0) = 0
		    gs_coeff(7, 1) = -1
		    gs_coeff(7, 2) = -1
		    gs_coeff(7, 3) = (-6 / 10)
		    gs_coeff(7, 4) = 0
		    gs_coeff(7, 5) = 0
		    gs_coeff(7, 6) = 0
		    gs_coeff(7, 7) = (3 / 10)
		    gs_coeff(7, 8) = 0
		    gs_coeff(7, 9) = 0
		    gs_coeff(7, 10) = -1
		    gs_coeff(7, 11) = (5 / 10)
		End If
	
		If .rdoCase141.checked = true then
			gs_coeff(18, 0) = 0
		    gs_coeff(18, 1) = 0
		    gs_coeff(18, 2) = 0
		    gs_coeff(18, 3) = 0
		    gs_coeff(18, 4) = 0
		    gs_coeff(18, 5) = 0
		    gs_coeff(18, 6) = 0
		    gs_coeff(18, 7) = 0
		    gs_coeff(18, 8) = (3 / 10)
		    gs_coeff(18, 9) = 0
		    gs_coeff(18, 10) = 0
		    gs_coeff(18, 11) = 0
		ElseIf .rdoCase142.checked = true then
			gs_coeff(18, 0) = 0
		    gs_coeff(18, 1) = 0
		    gs_coeff(18, 2) = 0
		    gs_coeff(18, 3) = 0
		    gs_coeff(18, 4) = (3 / 10)
		    gs_coeff(18, 5) = (3 / 10)
		    gs_coeff(18, 6) = 0
		    gs_coeff(18, 7) = 0
		    gs_coeff(18, 8) = (3 / 10)
		    gs_coeff(18, 9) = (5 / 10)
		    gs_coeff(18, 10) = 0
		    gs_coeff(18, 11) = 0
		End If
	
		If .rdoCase151.checked = true then
			gs_coeff(5, 0) = 0
		    gs_coeff(5, 1) = (5 / 10)
		    gs_coeff(5, 2) = (5 / 10)
		    gs_coeff(5, 3) = (4 / 10)
		    gs_coeff(5, 4) = 0
		    gs_coeff(5, 5) = 0
		    gs_coeff(5, 6) = 0
		    gs_coeff(5, 7) = 0
		    gs_coeff(5, 8) = 0
		    gs_coeff(5, 9) = 0
		    gs_coeff(5, 10) = (5 / 10)
		    gs_coeff(5, 11) = (-8 / 10)
		ElseIf .rdoCase152.checked = true then
			gs_coeff(5, 0) = 0
		    gs_coeff(5, 1) = (-7 / 10)
		    gs_coeff(5, 2) = (-7 / 10)
		    gs_coeff(5, 3) = (-4 / 10)
		    gs_coeff(5, 4) = 0
		    gs_coeff(5, 5) = 0
		    gs_coeff(5, 6) = 0
		    gs_coeff(5, 7) = 0
		    gs_coeff(5, 8) = 0
		    gs_coeff(5, 9) = 0
		    gs_coeff(5, 10) = (-6 / 10)
		    gs_coeff(5, 11) = 0
		ElseIf .rdoCase153.checked = true then
			gs_coeff(5, 0) = 0
		    gs_coeff(5, 1) = -1
		    gs_coeff(5, 2) = -1
		    gs_coeff(5, 3) = (-6 / 10)
		    gs_coeff(5, 4) = 0
		    gs_coeff(5, 5) = 0
		    gs_coeff(5, 6) = 0
		    gs_coeff(5, 7) = 0
		    gs_coeff(5, 8) = 0
		    gs_coeff(5, 9) = 0
		    gs_coeff(5, 10) = -1
		    gs_coeff(5, 11) = (5 / 10)
		End If
	
		If .rdoCase161.checked = true then
			gs_coeff(11, 0) = 0
		    gs_coeff(11, 1) = (5 / 10)
		    gs_coeff(11, 2) = (5 / 10)
		    gs_coeff(11, 3) = (3 / 10)
		    gs_coeff(11, 4) = (-2 / 10)
		    gs_coeff(11, 5) = (-2 / 10)
		    gs_coeff(11, 6) = 0
		    gs_coeff(11, 7) = 0
		    gs_coeff(11, 8) = 0
		    gs_coeff(11, 9) = 0
		    gs_coeff(11, 10) = (7 / 10)
		    gs_coeff(11, 11) = (-5 / 10)
		ElseIf .rdoCase162.checked = true then
			gs_coeff(11, 0) = (2 / 10)
		    gs_coeff(11, 1) = 0
		    gs_coeff(11, 2) = 0
		    gs_coeff(11, 3) = (-7 / 10)
		    gs_coeff(11, 4) = (3 / 10)
		    gs_coeff(11, 5) = (3 / 10)
		    gs_coeff(11, 6) = (3 / 10)
		    gs_coeff(11, 7) = (3 / 10)
		    gs_coeff(11, 8) = (3 / 10)
		    gs_coeff(11, 9) = (3 / 10)
		    gs_coeff(11, 10) = 0
		    gs_coeff(11, 11) = 0
		End If
	
		If .rdoCase171.checked = true then
			gs_coeff(16, 0) = 0
		    gs_coeff(16, 1) = (2 / 10)
		    gs_coeff(16, 2) = (2 / 10)
		    gs_coeff(16, 3) = (2 / 10)
		    gs_coeff(16, 4) = 0
		    gs_coeff(16, 5) = 0
		    gs_coeff(16, 6) = 0
		    gs_coeff(16, 7) = 0
		    gs_coeff(16, 8) = 0
		    gs_coeff(16, 9) = 0
		    gs_coeff(16, 10) = (2 / 10)
		    gs_coeff(16, 11) = 0
		ElseIf .rdoCase172.checked = true then
			gs_coeff(16, 0) = 0
		    gs_coeff(16, 1) = (-5 / 10)
		    gs_coeff(16, 2) = (-5 / 10)
		    gs_coeff(16, 3) = (-3 / 10)
		    gs_coeff(16, 4) = 0
		    gs_coeff(16, 5) = (5 / 10)
		    gs_coeff(16, 6) = 0
		    gs_coeff(16, 7) = 0
		    gs_coeff(16, 8) = 0
		    gs_coeff(16, 9) = 0
		    gs_coeff(16, 10) = (-6 / 10)
		    gs_coeff(16, 11) = (1 / 10)
		ElseIf .rdoCase173.checked = true then
			gs_coeff(16, 0) = 0
			gs_coeff(16, 1) = (-5 / 10)
			gs_coeff(16, 2) = (-5 / 10)
			gs_coeff(16, 3) = (-3 / 10)
			gs_coeff(16, 4) = (5 / 10)
			gs_coeff(16, 5) = 0
			gs_coeff(16, 6) = (5 / 10)
			gs_coeff(16, 7) = 0
			gs_coeff(16, 8) = 0
			gs_coeff(16, 9) = 0
			gs_coeff(16, 10) = (-6 / 10)
			gs_coeff(16, 11) = (1 / 10)
		End If
	
		If .rdoCase181.checked = true then
			gs_coeff(17, 0) = 0
		    gs_coeff(17, 1) = 0
		    gs_coeff(17, 2) = 0
		    gs_coeff(17, 3) = 0
		    gs_coeff(17, 4) = (3 / 10)
		    gs_coeff(17, 5) = (-3 / 10)
		    gs_coeff(17, 6) = (3 / 10)
		    gs_coeff(17, 7) = (-1 / 10)
		    gs_coeff(17, 8) = (3 / 10)
		    gs_coeff(17, 9) = 0
		    gs_coeff(17, 10) = 0
		    gs_coeff(17, 11) = 0
		ElseIf .rdoCase182.checked = true then
			gs_coeff(17, 0) = 0
		    gs_coeff(17, 1) = 0
		    gs_coeff(17, 2) = 0
		    gs_coeff(17, 3) = 0
		    gs_coeff(17, 4) = (-3 / 10)
		    gs_coeff(17, 5) = (3 / 10)
		    gs_coeff(17, 6) = (-3 / 10)
		    gs_coeff(17, 7) = (1 / 10)
		    gs_coeff(17, 8) = (-3 / 10)
		    gs_coeff(17, 9) = 0
		    gs_coeff(17, 10) = 0
		    gs_coeff(17, 11) = 0
		End If
	
		If .rdoCase191.checked = true then
			gs_coeff(19, 0) = 0
		    gs_coeff(19, 1) = 0
		    gs_coeff(19, 2) = 0
		    gs_coeff(19, 3) = 0
		    gs_coeff(19, 4) = (3 / 10)
		    gs_coeff(19, 5) = (3 / 10)
		    gs_coeff(19, 6) = 0
		    gs_coeff(19, 7) = 0
		    gs_coeff(19, 8) = (3 / 10)
		    gs_coeff(19, 9) = (5 / 10)
		    gs_coeff(19, 10) = 0
		    gs_coeff(19, 11) = 0
		ElseIf .rdoCase192.checked = true then
			gs_coeff(19, 0) = 0
		    gs_coeff(19, 1) = 0
		    gs_coeff(19, 2) = 0
		    gs_coeff(19, 3) = 0
		    gs_coeff(19, 4) = 0
		    gs_coeff(19, 5) = 0
		    gs_coeff(19, 6) = 0
		    gs_coeff(19, 7) = 0
		    gs_coeff(19, 8) = 0
		    gs_coeff(19, 9) = 0
		    gs_coeff(19, 10) = (-5 / 10)
		    gs_coeff(19, 11) = 0
		End If
	
		If .rdoCase201.checked = true then
			gs_coeff(15, 0) = 0
		    gs_coeff(15, 1) = 0
		    gs_coeff(15, 2) = 0
		    gs_coeff(15, 3) = 0
		    gs_coeff(15, 4) = 0
		    gs_coeff(15, 5) = 0
		    gs_coeff(15, 6) = 0
		    gs_coeff(15, 7) = 0
		    gs_coeff(15, 8) = (3 / 10)
		    gs_coeff(15, 9) = 0
		    gs_coeff(15, 10) = 0
		    gs_coeff(15, 11) = 0
		ElseIf .rdoCase202.checked = true then
			gs_coeff(15, 0) = 0
		    gs_coeff(15, 1) = (-5 / 10)
		    gs_coeff(15, 2) = (-5 / 10)
		    gs_coeff(15, 3) = 0
		    gs_coeff(15, 4) = (8 / 10)
		    gs_coeff(15, 5) = 0
		    gs_coeff(15, 6) = 0
		    gs_coeff(15, 7) = 0
		    gs_coeff(15, 8) = (-3 / 10)
		    gs_coeff(15, 9) = 0
		    gs_coeff(15, 10) = 0
		    gs_coeff(15, 11) = (1 / 10)
		ElseIf .rdoCase203.checked = true then
			gs_coeff(15, 0) = 0
		    gs_coeff(15, 1) = 0
		    gs_coeff(15, 2) = 0
		    gs_coeff(15, 3) = 0
		    gs_coeff(15, 4) = 0
		    gs_coeff(15, 5) = 0
		    gs_coeff(15, 6) = 0
		    gs_coeff(15, 7) = 0
		    gs_coeff(15, 8) = 0
		    gs_coeff(15, 9) = 0
		    gs_coeff(15, 10) = 0
		    gs_coeff(15, 11) = 0
		ElseIf .rdoCase204.checked = true then
			gs_coeff(15, 0) = 0
		    gs_coeff(15, 1) = 0
		    gs_coeff(15, 2) = 0
		    gs_coeff(15, 3) = 0
		    gs_coeff(15, 4) = (-3 / 10)
		    gs_coeff(15, 5) = (-3 / 10)
		    gs_coeff(15, 6) = 0
		    gs_coeff(15, 7) = (5 / 10)
		    gs_coeff(15, 8) = 0
		    gs_coeff(15, 9) = 0
		    gs_coeff(15, 10) = 0
		    gs_coeff(15, 11) = 0
		End If
	
		If .rdoCase211.checked = true then
			gs_coeff(13, 0) = 0
		    gs_coeff(13, 1) = (-2 / 10)
		    gs_coeff(13, 2) = (-2 / 10)
		    gs_coeff(13, 3) = (-3 / 10)
		    gs_coeff(13, 4) = 0
		    gs_coeff(13, 5) = 0
		    gs_coeff(13, 6) = 0
		    gs_coeff(13, 7) = 0
		    gs_coeff(13, 8) = 0
		    gs_coeff(13, 9) = 0
		    gs_coeff(13, 10) = 0
		    gs_coeff(13, 11) = 0
		ElseIf .rdoCase212.checked = true then
			gs_coeff(13, 0) = 0
		    gs_coeff(13, 1) = (3 / 10)
		    gs_coeff(13, 2) = (3 / 10)
		    gs_coeff(13, 3) = (3 / 10)
		    gs_coeff(13, 4) = 0
		    gs_coeff(13, 5) = 0
		    gs_coeff(13, 6) = 0
		    gs_coeff(13, 7) = 0
		    gs_coeff(13, 8) = 0
		    gs_coeff(13, 9) = 0
		    gs_coeff(13, 10) = 0
		    gs_coeff(13, 11) = 0
		End If
	
		If .rdoCase221.checked = true then
			gs_coeff(12, 0) = 0
		    gs_coeff(12, 1) = 0
		    gs_coeff(12, 2) = 0
		    gs_coeff(12, 3) = 0
		    gs_coeff(12, 4) = 0
		    gs_coeff(12, 5) = 0
		    gs_coeff(12, 6) = 0
		    gs_coeff(12, 7) = (5 / 10)
		    gs_coeff(12, 8) = 0
		    gs_coeff(12, 9) = 0
		    gs_coeff(12, 10) = 0
		    gs_coeff(12, 11) = 0
		ElseIf .rdoCase222.checked = true then
			gs_coeff(12, 0) = 0
		    gs_coeff(12, 1) = 0
		    gs_coeff(12, 2) = 0
		    gs_coeff(12, 3) = 0
		    gs_coeff(12, 4) = 0
		    gs_coeff(12, 5) = 0
		    gs_coeff(12, 6) = 0
		    gs_coeff(12, 7) = 0
		    gs_coeff(12, 8) = (5 / 10)
		    gs_coeff(12, 9) = 0
		    gs_coeff(12, 10) = 0
		    gs_coeff(12, 11) = 0
		ElseIf .rdoCase223.checked = true then
			gs_coeff(12, 0) = 0
		    gs_coeff(12, 1) = (-3 / 10)
		    gs_coeff(12, 2) = (-3 / 10)
		    gs_coeff(12, 3) = (-2 / 10)
		    gs_coeff(12, 4) = (-1 / 10)
		    gs_coeff(12, 5) = (-1 / 10)
		    gs_coeff(12, 6) = 0
		    gs_coeff(12, 7) = (7 / 10)
		    gs_coeff(12, 8) = (-1 / 10)
		    gs_coeff(12, 9) = (-1 / 10)
		    gs_coeff(12, 10) = (-4 / 10)
		    gs_coeff(12, 11) = (1 / 10)
		End If
	
		If .rdoCase231.checked = true then
			gs_coeff(22, 0) = 0
		    gs_coeff(22, 1) = (-3 / 10)
		    gs_coeff(22, 2) = (-3 / 10)
		    gs_coeff(22, 3) = 0
		    gs_coeff(22, 4) = 0
		    gs_coeff(22, 5) = 0
		    gs_coeff(22, 6) = 0
		    gs_coeff(22, 7) = 0
		    gs_coeff(22, 8) = 0
		    gs_coeff(22, 9) = 0
		    gs_coeff(22, 10) = 0
		    gs_coeff(22, 11) = 0
		ElseIf .rdoCase232.checked = true then
			gs_coeff(22, 0) = 0
		    gs_coeff(22, 1) = (3 / 10)
		    gs_coeff(22, 2) = (3 / 10)
		    gs_coeff(22, 3) = 0
		    gs_coeff(22, 4) = 0
		    gs_coeff(22, 5) = 0
		    gs_coeff(22, 6) = 0
		    gs_coeff(22, 7) = 0
		    gs_coeff(22, 8) = 0
		    gs_coeff(22, 9) = 0
		    gs_coeff(22, 10) = 0
		    gs_coeff(22, 11) = 0
		End If
	
		If .rdoCase241.checked = true then
		'	Option3.Enabled = True
		    	'	Option7.Enabled = True
			gs_coeff(23, 0) = 0
			gs_coeff(23, 1) = (1 / 10)
			gs_coeff(23, 2) = (1 / 10)
			gs_coeff(23, 3) = 0
			gs_coeff(23, 4) = 0
			gs_coeff(23, 5) = 0
			gs_coeff(23, 6) = 0
			gs_coeff(23, 7) = 0
			gs_coeff(23, 8) = 0
			gs_coeff(23, 9) = 0
			gs_coeff(23, 10) = 0
			gs_coeff(23, 11) = 0
		ElseIf .rdoCase242.checked = true then
			gs_coeff(23, 0) = 0
		    gs_coeff(23, 1) = -1
		    gs_coeff(23, 2) = -1
		    gs_coeff(23, 3) = 0
		    gs_coeff(23, 4) = 0
		    gs_coeff(23, 5) = 0
		    gs_coeff(23, 6) = 0
		    gs_coeff(23, 7) = 0
		    gs_coeff(23, 8) = (5 / 10)
		    gs_coeff(23, 9) = 0
		    gs_coeff(23, 10) = 0
		    gs_coeff(23, 11) = 0
		End If
	
		If .rdoCase251.checked = true then
			gs_coeff(24, 0) = 0
		    gs_coeff(24, 1) = -1
		    gs_coeff(24, 2) = 0
		    gs_coeff(24, 3) = 0
		    gs_coeff(24, 4) = 0
		    gs_coeff(24, 5) = 0
		    gs_coeff(24, 6) = 0
	   		gs_coeff(24, 7) = 0
		    gs_coeff(24, 8) = 0
		    gs_coeff(24, 9) = 0
		    gs_coeff(24, 10) = 0
		    gs_coeff(24, 11) = 0
		ElseIf .rdoCase252.checked = true then
			gs_coeff(24, 0) = 0
		    gs_coeff(24, 1) = 1
		    gs_coeff(24, 2) = 0
		    gs_coeff(24, 3) = 0
		    gs_coeff(24, 4) = 0
		    gs_coeff(24, 5) = 0
		    gs_coeff(24, 6) = 0
		    gs_coeff(24, 7) = 0
		    gs_coeff(24, 8) = 0
		    gs_coeff(24, 9) = 0
		    gs_coeff(24, 10) = 0
		    gs_coeff(24, 11) = 0
		End If
	
		If .rdoCase261.checked = true then
			gs_coeff2(0, 0) = 0
		    gs_coeff2(0, 1) = (3 / 10)
		ElseIf .rdoCase262.checked = true then
		   	gs_coeff2(0, 0) = 0
		   	gs_coeff2(0, 1) = -1
		End If
	
		If .rdoCase271.checked = true then
			gs_coeff2(1, 0) = 0
		    gs_coeff2(1, 1) = (5 / 10)
		ElseIf .rdoCase272.checked = true then
			gs_coeff2(1, 0) = 0
		   	gs_coeff2(1, 1) = (-5 / 10)
		End If
	
		If .rdoCase281.checked = true then
			gs_coeff2(2, 0) = 0
		   	gs_coeff2(2, 1) = (3 / 10)
		ElseIf .rdoCase282.checked = true then
			gs_coeff2(2, 0) = 0
		   	gs_coeff2(2, 1) = (-6 / 10)
		End If
	
		If .rdoCase291.checked = true then
			gs_coeff2(3, 0) = 0
		    gs_coeff2(3, 1) = (3 / 10)
		ElseIf .rdoCase292.checked = true then
			gs_coeff2(3, 0) = (3 / 10)
		   	gs_coeff2(3, 1) = 0
		End If
	
		If .rdoCase301.checked = true then
			gs_coeff2(4, 0) = 0
		    gs_coeff2(4, 1) = (-8 / 10)
		ElseIf .rdoCase302.checked = true then
			gs_coeff2(4, 0) = 0
		   	gs_coeff2(4, 1) = (2 / 10)
		End If
	
		If .rdoCase311.checked = true then
			gs_coeff2(5, 0) = 0
		   	gs_coeff2(5, 1) = (1 / 10)
		ElseIf .rdoCase312.checked = true then
			gs_coeff2(5, 0) = (1 / 10)
		   	gs_coeff2(5, 1) = 0
		End If
	
		If .rdoCase321.checked = true then
			gs_coeff2(6, 0) = 0
		    gs_coeff2(6, 1) = (5 / 10)
		ElseIf .rdoCase322.checked = true then
			gs_coeff2(6, 0) = 0
		    gs_coeff2(6, 1) = (-5 / 10)
		ElseIf .rdoCase323.checked = true then
			gs_coeff2(6, 0) = 0
		    gs_coeff2(6, 1) = (-5 / 10)
		ElseIf .rdoCase324.checked = true then
			gs_coeff2(6, 0) = 0
		    gs_coeff2(6, 1) = (-7 / 10)
		End If
	End With
	
	For i = 0 To 11
    		gsel_insp(i) = gs_coeff(0, i)
	Next 
    
	For i = 1 To 24
		For j = 0 To 11
			If ((gsel_insp(j) = -1) Or (gs_coeff(i, j) = -1)) Then
				gsel_insp(j) = -1
			ElseIf ((gsel_insp(j) >= 0) And (gs_coeff(i, j) >= 0)) Then
				gsel_insp(j) = gsel_insp(j) + gs_coeff(i, j) - gsel_insp(j) * gs_coeff(i, j)
			ElseIf ((gsel_insp(j) < 0) And (gs_coeff(i, j) < 0)) Then
				gsel_insp(j) = gsel_insp(j) + gs_coeff(i, j) + gsel_insp(j) * gs_coeff(i, j)
			Else
				If (Abs(gsel_insp(j)) > (gs_coeff(i, j))) Then
			       min = Abs(gs_coeff(i, j))
				Else
			       min = Abs(gsel_insp(j))
				End If
			            		
			gsel_insp(j) = (gsel_insp(j) + gs_coeff(i, j)) / (1 - min)
			End If
		Next 
	Next 
           
	For i = 0 To 11
    		gmax_ins(i) = i
	Next 
    
	For i = 0 To 11
    		For j = 0 To i
        			If gsel_insp(i) > gsel_insp(j) Then
            			tempsel = gsel_insp(j)
            			tempinsp = gmax_ins(j)
            			gsel_insp(j) = gsel_insp(i)
            			gmax_ins(j) = gmax_ins(i)
            			gsel_insp(i) = tempsel
            			gmax_ins(i) = tempinsp
        			End If
    		Next 
	Next 
    
	For i = 0 To 11
    	Select Case gmax_ins(i)
    		Case 0:
   	             InsType(i) = "0500"
                 FitnessDegree(i) = gsel_insp(i)
    		Case 1:
    			InsType(i) = "0600"
    			FitnessDegree(i) = gsel_insp(i)
    		Case 2:
    		    InsType(i) = "0700"
    		    FitnessDegree(i) = gsel_insp(i)
    		Case 3:
    	        InsType(i) = "0800"
    			FitnessDegree(i) = gsel_insp(i)
    		Case 4:
    		    InsType(i) = "0201"
    		 	FitnessDegree(i) = gsel_insp(i)
    	    Case 5:
        		InsType(i) = "0202"
        		FitnessDegree(i) = gsel_insp(i)
  			Case 6:                		   
        		InsType(i) = "0400"
        		FitnessDegree(i) = gsel_insp(i)
    		Case 7:                		                     		    
        		InsType(i) = "0100"
        		FitnessDegree(i) = gsel_insp(i)
    		Case 8:                		                    		    
        		InsType(i) = "0300"
        		FitnessDegree(i) = gsel_insp(i)
    		Case 9:            		 	                		 	     
      	        InsType(i) = "0900"
      			FitnessDegree(i) = gsel_insp(i)
      		Case 10:              		                 		     
      	        InsType(i) = "1000"
      			FitnessDegree(i) = gsel_insp(i)
    		Case 11:          			    
    		    InsType(i) = "1100"
    			FitnessDegree(i) = gsel_insp(i)
    	End Select
    Next 
    	
	j = 0
	For i=0 To 11
		If InsType(i) <> "" and j < 5 Then 
			AInsType(j) = InsType(i)
			AFitnessDegree(j) = FormatNumber(FitnessDegree(i), 2)
			j=j+1
		End If
	Next
    	  		
	For i=0 To 4
    	Select Case AInsType(i)	    	
    		Case "0201"
    			AInsType(i)="02"
    			InsAssureance(i)="01"
		Case "0202"
    			AInsType(i)="02"
    			InsAssureance(i)="02"
	End Select
	Next
    	
  	    
	For i = 0 To 1
    		gsel_insp2(i) = gs_coeff2(0, i)
	Next 
    
	For i = 1 To 6
    	For j = 0 To 1
        	If ((gsel_insp2(j) = -1) Or (gs_coeff2(i, j) = -1)) Then
        		gsel_insp2(j) = -1
        	ElseIf ((gsel_insp2(j) >= 0) And (gs_coeff2(i, j) >= 0)) Then
          		gsel_insp2(j) = gsel_insp2(j) + gs_coeff2(i, j) - gsel_insp2(j) * gs_coeff2(i, j)
        	ElseIf ((gsel_insp2(j) < 0) And (gs_coeff2(i, j) < 0)) Then
          		gsel_insp2(j) = gsel_insp2(j) + gs_coeff2(i, j) + gsel_insp2(j) * gs_coeff2(i, j)
        	End If

        	If (Abs(gsel_insp2(j)) > (gs_coeff2(i, j))) Then
           		min = Abs(gs_coeff2(i, j))
        	Else
           		min = Abs(gsel_insp2(j))
           	End If
                	
        	gsel_insp2(j) = (gsel_insp2(j) + gs_coeff2(i, j)) / (1 - min)
            			
    	Next 
	Next 
    
	If (gsel_insp2(0) = -1) Then
    		max_insp = 1
	ElseIf (gsel_insp2(1) = -1) Then
    		max_insp = 0
	Else
   	     If (gsel_insp2(0) < gsel_insp2(1)) Then
            max_sel = gsel_insp2(1)
            max_insp = 1
   	     Else
            max_sel = gsel_insp2(0)
            max_insp = 0
   	    End If
	End If
	
	For i=0 To 4
		Select Case max_insp
	    	Case 0:
	    			InsCase = "01"
		        				
	    	Case 1:
	    		InsCase = "02"
		       				
	    		If AInsType(i) = "02" Then 
	    			AInsType(i)=""
	    			InsAssureance(i)=""
	    			AFitnessDegree(i)=""
	    		End If
	    		If AInsType(i) = "04" Then 
	    			AInsType(i)= ""
	    			InsAssureance(i)=""
	    			AFitnessDegree(i)=""
	    		End If
	    		If AInsType(i) = "07" Then 
	    			AInsType(i)=""
	    			InsAssureance(i)=""
	    			AFitnessDegree(i)=""
	    		End If
	    End Select	
	Next
	
	Dim F_InsType(4)
	Dim F_InsAssureance(4)
	Dim F_FitnessDegree(4)
	Dim k

	k = 0
	For i = 0 To 4
		If AInsType(i) <> "" Then
			F_InsType(k) = AInsType(i)
			F_InsAssureance(k) = InsAssureance(i)
			F_FitnessDegree(k) = AFitnessDegree(i)
			k = k + 1
		End If
	Next
	
	WriteCookie "txtInsVA", (max_insp)
	WriteCookie "txtInsCase", Trim(InsCase)
	
	WriteCookie "txtInsType1", Trim(F_InsType(0))
	WriteCookie "txtInsType2", Trim(F_InsType(1))
	WriteCookie "txtInsType3", Trim(F_InsType(2))
	WriteCookie "txtInsType4", Trim(F_InsType(3))
	WriteCookie "txtInsType5", Trim(F_InsType(4))
	
	WriteCookie "txtInsAssureance1", Trim(F_InsAssureance(0))
	WriteCookie "txtInsAssureance2", Trim(F_InsAssureance(1))
	WriteCookie "txtInsAssureance3", Trim(F_InsAssureance(2))
	WriteCookie "txtInsAssureance4", Trim(F_InsAssureance(3))
	WriteCookie "txtInsAssureance5", Trim(F_InsAssureance(4))
	
	WriteCookie "txtFitnessDegree1", Trim(F_FitnessDegree(0))
	WriteCookie "txtFitnessDegree2", Trim(F_FitnessDegree(1))
	WriteCookie "txtFitnessDegree3", Trim(F_FitnessDegree(2))
	WriteCookie "txtFitnessDegree4", Trim(F_FitnessDegree(3))
	WriteCookie "txtFitnessDegree5", Trim(F_FitnessDegree(4))
	
	Navigate BIZ_PGM_JUMP_ID
	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	gIsTab = "Y" 
    gTabMaxCnt = 2   ' Tab의 갯수 

	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	
	Call SetDefaultVal
    Call InitVariables                                                      									'⊙: Initializes local global variables
    Call SetToolbar("10000000000111")
  	
  	Q23.Style.display = "none"
 	Q24.Style.display = "none"
 	Q25.Style.display = "none"
 	Q26.Style.display = "none"
 	Q27.Style.display = "none"
 	Q28.Style.display = "none"
 	Q29.Style.display = "none"
 	Q30.Style.display = "none"
 	Q31.Style.display = "none"
 	Q32.Style.display = "none" 	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	FncQuery = false
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	FncNew = false
End Function

'========================================================================================
' Function Name : Fnc
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	FncDelete = false
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	FncSave = false
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = false
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	FncInsertRow = false
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = false
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
	FncPrev = false
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	FncNext = false
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 Call parent.FncExport(Parent.C_SINGLE)					'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	Call parent.FncFind(Parent.C_SINGLE, False)     
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	DbDelete= false
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()									'☆: 삭제 성공후 실행 로직 
	DbDeleteOk = false
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	DbQuery =false
End Function

'========================================================================================
' Function Name : DbQueryOkOPEN
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()									'☆: 조회 성공후 실행로직 
	DbQueryOk = false
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQueryOkOPEN
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()	
	DbQueryOk = false							'☆: 조회 성공후 실행로직 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															'☆: 저장 성공후 실행 로직 
	DbSaveOk = false
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> BORDER=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>검사방식 선정 질의1</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>검사방식 선정 질의2</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR height=*>
		<TD  VALIGN="TOP" WIDTH=100% CLASS="Tab11">
			<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
				<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 VALIGN="top"  WIDTH="100%">
							<FIELDSET CLASS="CLSFLD">
								<LEGEND>검사방식 선정 질의내용</LEGEND>
								<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_40%> >
								<TR>
									<TD CLASS=TD5   HEIGHT=10 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q1>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>1.검사의 종류?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion1" TAG="2X" ID="rdoCase11"><LABEL FOR="rdoCase11">수입검사</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion1" TAG="2X" ID="rdoCase12"><LABEL FOR="rdoCase12">공정검사</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion1" TAG="2X" ID="rdoCase13"><LABEL FOR="rdoCase13">최종검사</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion1" TAG="2X" ID="rdoCase14"><LABEL FOR="rdoCase14">출하검사</LABEL></TD>
								</TR>
								<TR ID=Q2>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>2.검사의 성격?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion2" TAG="2X" ID="rdoCase21"><LABEL FOR="rdoCase21">완전파괴</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion2" TAG="2X" ID="rdoCase22"><LABEL FOR="rdoCase22">반파괴</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion2" TAG="2X" ID="rdoCase23"><LABEL FOR="rdoCase23">비파괴</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q3>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>3.로트의 형태?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion3" TAG="2X" ID="rdoCase31"><LABEL FOR="rdoCase31">단일 또는 고립로트</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion3" TAG="2X" ID="rdoCase32"><LABEL FOR="rdoCase32">연속로트</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion3" TAG="2X" ID="rdoCase33"><LABEL FOR="rdoCase33">로트 구분 안됨</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q4>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>4.검사항목의 결점수준?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion4" TAG="2X" ID="rdoCase41"><LABEL FOR="rdoCase41">치명결점</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion4" TAG="2X" ID="rdoCase42"><LABEL FOR="rdoCase42">중결점</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion4" TAG="2X" ID="rdoCase43"><LABEL FOR="rdoCase43">경결점</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion4" TAG="2X" ID="rdoCase44"><LABEL FOR="rdoCase44">미결점</LABEL></TD>	
								</TR>
								<TR ID=Q5>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>5.검사의 대상물?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion5" TAG="2X" ID="rdoCase51"><LABEL FOR="rdoCase51">개수</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion5" TAG="2X" ID="rdoCase52"><LABEL FOR="rdoCase52">부피</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion5" TAG="2X" ID="rdoCase53"><LABEL FOR="rdoCase53">수명</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q6>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>6.공정의 안정성?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion6" TAG="2X" ID="rdoCase61" ONCLICK="vbscript:Condition61()"><LABEL FOR="rdoCase61">안정</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion6" TAG="2X" ID="rdoCase62" ONCLICK="vbscript:Condition62()"><LABEL FOR="rdoCase62">불안정</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q7>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>7.생산품의 적재면적?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion7" TAG="2X" ID="rdoCase71"><LABEL FOR="rdoCase71">충분함</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion7" TAG="2X" ID="rdoCase72"><LABEL FOR="rdoCase72">어느 정도 있음</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion7" TAG="2X" ID="rdoCase73"><LABEL FOR="rdoCase73">거의 없음</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q8>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>8.부품의 안정성 유무요구?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion8" TAG="2X" ID="rdoCase81"><LABEL FOR="rdoCase81">요구됨</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion8" TAG="2X" ID="rdoCase82"><LABEL FOR="rdoCase82">요구되지 않음</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q9>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>9.검사원의 학력?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion9" TAG="2X" ID="rdoCase91"><LABEL FOR="rdoCase91">대졸이상</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion9" TAG="2X" ID="rdoCase92"><LABEL FOR="rdoCase92">고졸이상</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion9" TAG="2X" ID="rdoCase93"><LABEL FOR="rdoCase93">고졸이하</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>	
								</TR>
								<TR ID=Q10>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>10.검사원의 경력?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion10" TAG="2X" ID="rdoCase101"><LABEL FOR="rdoCase101">2년 이상</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion10" TAG="2X" ID="rdoCase102"><LABEL FOR="rdoCase102">1~2년</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion10" TAG="2X" ID="rdoCase103"><LABEL FOR="rdoCase103">1년 미만</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q11>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>11.공급자의 수?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion11" TAG="2X" ID="rdoCase111"><LABEL FOR="rdoCase111">독점</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion11" TAG="2X" ID="rdoCase112"><LABEL FOR="rdoCase112">2~5인</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion11" TAG="2X" ID="rdoCase113"><LABEL FOR="rdoCase113">6인 이상</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q12>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>12.업체와 거래빈도?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion12" TAG="2X" ID="rdoCase121"><LABEL FOR="rdoCase121">단 한번의 거래</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion12" TAG="2X" ID="rdoCase122"><LABEL FOR="rdoCase122">보통</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion12" TAG="2X" ID="rdoCase123"><LABEL FOR="rdoCase123">빈번한 편</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q13>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>13.업체에 대한 등급?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion13" TAG="2X" ID="rdoCase131" ONCLICK="vbscript:Condition131()"><LABEL FOR="rdoCase131">1등급</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion13" TAG="2X" ID="rdoCase132" ONCLICK="vbscript:Condition132()"><LABEL FOR="rdoCase132">2등급</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion13" TAG="2X" ID="rdoCase133" ONCLICK="vbscript:Condition133()"><LABEL FOR="rdoCase133">3등급</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q14>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>14.품질향상의 자극여부?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion14" TAG="2X" ID="rdoCase141"><LABEL FOR="rdoCase141">자극이 필요함</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion14" TAG="2X" ID="rdoCase142"><LABEL FOR="rdoCase142">자극이 필요하지 않음</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q15>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>15.불량 공급자의 대체 용이성?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion15" TAG="2X" ID="rdoCase151"><LABEL FOR="rdoCase151">쉽다</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion15" TAG="2X" ID="rdoCase152"><LABEL FOR="rdoCase152">어느 정도 가능</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion15" TAG="2X" ID="rdoCase153"><LABEL FOR="rdoCase153">어렵다</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR ID=Q16>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>16.우선 고려사항?</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion16" TAG="2X" ID="rdoCase161"><LABEL FOR="rdoCase161">검사량의 최소화</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion16" TAG="2X" ID="rdoCase162"><LABEL FOR="rdoCase162">사용의 편리함</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
				</TABLE>
			</DIV>
			<DIV ID="TabDiv" STYLE="DISPLAY: none " SCROLL=no>
				<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 VALIGN="top"  WIDTH="100%">
							<FIELDSET CLASS="CLSFLD">
								<LEGEND>검사방식 선정 질의내용</LEGEND>
								<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_40%> >
									<TR>
										<TD CLASS=TD5   HEIGHT=10 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q17>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>17.최종 품질 보증유무?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion17" TAG="2X" ID="rdoCase171"><LABEL FOR="rdoCase171">원하지 않음</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion17" TAG="2X" ID="rdoCase172"><LABEL FOR="rdoCase172">LTPD보증</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion17" TAG="2X" ID="rdoCase173"><LABEL FOR="rdoCase173">AOQL보증</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q18>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>18.품질보증기간은?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion18" TAG="2X" ID="rdoCase181"><LABEL FOR="rdoCase181">평균(장기)</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion18" TAG="2X" ID="rdoCase182"><LABEL FOR="rdoCase182">단기</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q19>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>19.시료수와 로트크기 관계?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion19" TAG="2X" ID="rdoCase191"><LABEL FOR="rdoCase191">로트크기에 따라 시료수 변함</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion19" TAG="2X" ID="rdoCase192"><LABEL FOR="rdoCase192">관계 없음</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q20>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>20.불합격된 로트의 처리?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion20" TAG="2X" ID="rdoCase201"><LABEL FOR="rdoCase201">반품</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion20" TAG="2X" ID="rdoCase202"><LABEL FOR="rdoCase202">전수선별</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion20" TAG="2X" ID="rdoCase203"><LABEL FOR="rdoCase203">파괴</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion20" TAG="2X" ID="rdoCase204"><LABEL FOR="rdoCase204">규칙없음</LABEL></TD>	
									</TR>
									<TR ID=Q21>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>21.바람직한 로트합격여부 결정기준?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion21" TAG="2X" ID="rdoCase211"><LABEL FOR="rdoCase211">현 로트에서만 판단하는 것이 좋다</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion21" TAG="2X" ID="rdoCase212"><LABEL FOR="rdoCase212">과거 또는 미래의 결과 이용 가능</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q22>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>22.보호대상?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion22" TAG="2X" ID="rdoCase221"><LABEL FOR="rdoCase221">생산자, 소비자</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion22" TAG="2X" ID="rdoCase222"><LABEL FOR="rdoCase222">생산자</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion22" TAG="2X" ID="rdoCase223"><LABEL FOR="rdoCase223">소비자</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q23>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>23.바람직한 검사시점은?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion23" TAG="2X" ID="rdoCase231"><LABEL FOR="rdoCase231">비정기적</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion23" TAG="2X" ID="rdoCase232"><LABEL FOR="rdoCase232">정기적</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q24>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>24.주기적인 샘플채취 가능 여부?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion24" TAG="2X" ID="rdoCase241"><LABEL FOR="rdoCase241">가능</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion24" TAG="2X" ID="rdoCase242"><LABEL FOR="rdoCase242">불가능</LABEL></TD>	
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q25>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>25.기본목표?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion25" TAG="2X" ID="rdoCase251"><LABEL FOR="rdoCase251">로트의 합격, 불합격 판정</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion25" TAG="2X" ID="rdoCase252"><LABEL FOR="rdoCase252">전에 수행된 검사의 신뢰도 점검</LABEL></TD>	
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q26>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>26.계량데이터 획득가능 여부?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion26" TAG="2X" ID="rdoCase261" ONCLICK="vbscript:Condition261()"><LABEL FOR="rdoCase261">가능</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion26" TAG="2X" ID="rdoCase262" ONCLICK="vbscript:Condition262()"><LABEL FOR="rdoCase262">불가능</LABEL></TD>	
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q27>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>27.계측기 보유여부?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion27" TAG="2X" ID="rdoCase271"><LABEL FOR="rdoCase271">있다</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion27" TAG="2X" ID="rdoCase272"><LABEL FOR="rdoCase272">없다</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q28>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>28.계측기 유지비용?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion28" TAG="2X" ID="rdoCase281"><LABEL FOR="rdoCase281">비싸다</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion28" TAG="2X" ID="rdoCase282"><LABEL FOR="rdoCase282">싸다</LABEL></TD>	
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>								
									</TR>
									<TR ID=Q29>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>29.불량률 이외의 정보 필요유무?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion29" TAG="2X" ID="rdoCase291"><LABEL FOR="rdoCase291">제시해야 한다</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion29" TAG="2X" ID="rdoCase292"><LABEL FOR="rdoCase292">제시할 필요 없다</LABEL></TD>	
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q30>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>30.물적 증거 제시여부?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion30" TAG="2X" ID="rdoCase301"><LABEL FOR="rdoCase301">제시해야 한다</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion30" TAG="2X" ID="rdoCase302"><LABEL FOR="rdoCase302">제시할 필요 없다</LABEL></TD>	
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q31>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>31.검사소요기간?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion31" TAG="2X" ID="rdoCase311"><LABEL FOR="rdoCase311">대체적으로 길다</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion31" TAG="2X" ID="rdoCase312"><LABEL FOR="rdoCase312">대체적으로 짧다</LABEL></TD>	
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR ID=Q32>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>32.측정치 분포?</TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion32" TAG="2X" ID="rdoCase321"><LABEL FOR="rdoCase321">정규분포</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion32" TAG="2X" ID="rdoCase322"><LABEL FOR="rdoCase322">지수분포</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion32" TAG="2X" ID="rdoCase323"><LABEL FOR="rdoCase323">와이블분포</LABEL></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQuestion32" TAG="2X" ID="rdoCase324"><LABEL FOR="rdoCase324">분포형태가 명확치 않음</LABEL></TD>		
									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
				</TABLE>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
	    	<TD WIDTH="100%">
	    		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
	    			<TR>
	    				<TD WIDTH=10>&nbsp;</TD>
	    				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadExportCharge">결과보기</A></TD>
						<TD WIDTH=10>&nbsp;</TD>
	    			</TR>
	    		</TABLE>
	    	</TD>
	</TR>
    	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

