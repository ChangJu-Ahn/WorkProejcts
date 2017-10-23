<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m2111ra1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open Po Ref Popup ASP														*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2001/05/08																*
'*                            2002/04/30
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : Min, HJ															*	
'*                            Kim Jae Soon
'* 11. Comment              :																			*
'* 12. Common Coding Guide  :																			*
'* 13. History              :																			*
'********************************************************************************************************
Response.Expires = -1													'☜ : ASP가 캐쉬되지 않도록 한다.
%>
<HTML>
<HEAD>
<!--<TITLE>구매요청참조</TITLE>-->
<TITLE></TITLE>
<%
'########################################################################################################
'#						1. 선 언 부																		#
'########################################################################################################
%>
<%
'********************************************  1.1 Inc 선언  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!-- #Include file="../../inc/IncSvrVariables.inc" -->
<%
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<%
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================
%>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBS">

Option Explicit					<% '☜: indicates that All variables must be declared in advance %>
	
	
<%
'********************************************  1.2 Global 변수/상수 선언  *******************************
'*	Description : 1. Constant는 반드시 대문자 표기														*
'********************************************************************************************************
%>

<%
'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================
%>
	'상단 스프레스 
	Const C_PurGrp 			= 1
	Const C_PurGrpNm 		= 2															'☆: Spread Sheet의 Column별 상수 
	Const C_BpCd	 		= 3
	Const C_BpCdNm 			= 4
	Const C_ProCType		= 5
	Const C_ProCTypeNm		= 6
	
	'하단 스프레드 
	Const C_ReqNo 			= 1
	Const C_PlantCd 		= 2															'☆: Spread Sheet의 Column별 상수 
	Const C_PlantNm 		= 3
	Const C_ItemCd 			= 4
	Const C_ItemNm			= 5
	Const C_Spec			= 6
	Const C_Qty 			= 7
	Const C_Unit 			= 8
	Const C_DlvyDt 			= 9
	Const C_PlantDt 		= 10	
	Const C_ReqType			= 11
	Const C_ReqTypeNm		= 12
	Const C_SoNo			= 13	
	Const C_SoSeqNo			= 14
	Const C_TrackingNo		= 15	
	Const C_SLCd 			= 16
	Const C_SLNm 			= 17
	Const C_HSCd			= 18
	Const C_HSNm 			= 19
	
	'이성룡 추가 
	Const C_hUnderTot		= 20
	Const C_hOverTot		= 21	
	
	
    Const BIZ_PGM_ID 		= "m2111rb1_1.asp"                              '☆: Biz Logic ASP Name
     
<%
'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
%>
	Const C_SHEETMAXROWS_D  = 100                                          '☆: Fetch max count at once
	Const C_MaxKey_1        = 6                                           '☆: key count of SpreadSheet
	'이성룡 수정 
	Const C_MaxKey			= 21
	'Const C_MaxKey          = 19                                           '☆: key count of SpreadSheet
<%
'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
%>
<!-- #Include file="../../inc/lgvariables.inc" -->	
<%
'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================
%>


Dim lgStrPrevKey_1			'두번째 그리드에서 사용되는 변수 
Dim lgPageNo_1				'두번째 그리드에서 사용되는 변수 
		
Dim lgSelectList                                            '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgSelectListDT                                          '☜: SpreadSheet의 초기  위치정보관련 변수 

Dim lgSortFieldNm                                           '☜: Orderby popup용 데이타(필드설명)      
Dim lgSortFieldCD                                           '☜: Orderby popup용 데이타(필드코드)      

Dim lgPopUpR                                                '☜: Orderby default 값                    

Dim lgKeyPos                                                '☜: Key위치                               
Dim lgKeyPosVal                                             '☜: Key위치 Value                         
Dim IscookieSplit 

Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam
Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

<%
'########################################################################################################
'#						2. Function 부																	#
'#																										#
'#	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기술					#
'#	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.							#
'#						 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함)			#
'########################################################################################################
%>
<% 
'*******************************************  2.1 변수 초기화 함수  *************************************
'*	기능: 변수초기화																					*
'*	Description : Global변수 처리, 변수초기화 등의 작업을 한다.											*
'********************************************************************************************************
%>
<%
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
%>
Function InitVariables()
		lgStrPrevKey     = ""								   'initializes Previous Key
		lgPageNo         = ""
		
		lgStrPrevKey_1     = ""								   'initializes Previous Key
		lgPageNo_1         = ""
        
        lgBlnFlgChgValue = False	                           'Indicates that no value changed
        
        lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
        frm1.vspdData.OperationMode  = 5
        frm1.vspdData1.OperationMode = 3
        
        lgSortKey        = 1   
        
        lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>

        gblnWinEvent = False
       
        Redim arrReturn(0,0)        
        Self.Returnvalue = arrReturn     
End Function

<%'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 

		<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>                                '☆: 

		'------ Developer Coding part (End )   -------------------------------------------------------------- 
	End Sub
<%
'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
%>
 Sub InitComboBox()
	'-----------------------------------------------------------------------------------------------------
	' List Minor code for Procurement Type(조달구분)
	'-----------------------------------------------------------------------------------------------------
	if frm1.hdnSubcontraflg.value  = "N" then
			Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1003' AND MINOR_CD = 'P' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
			Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
	Elseif  frm1.hdnSubcontraflg.value ="Y" then
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1003' AND MINOR_CD != 'P' ORDER BY MINOR_CD DESC", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
	Else
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1003' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
	End if
End Sub
<%
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
%>
	Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("M2111RA1_1_2","S","B","V20030303",PopupParent.C_SORT_DBAGENT,frm1.vspdData1, _
									C_MaxKey_1, "X","X")

	
		Call SetZAdoSpreadSheet("M2111RA1_1_1","S","A","V20030303",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MAXKEY , "X","X")
		
		Call SetSpreadLock 
		'Call SetSpreadLock("A")
	End Sub


<%
'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
%>
	Sub SetSpreadLock()
		ggoSpread.Source = frm1.vspdData1
  	    ggoSpread.SpreadLockWithOddEvenRowColor()

	End Sub	

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

	Sub SetSpreadLock_1()
	
		ggoSpread.Source = frm1.vspdData
  	    ggoSpread.SpreadLockWithOddEvenRowColor()
	End Sub
<%
'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
%>	
	
	Function OKClick()
	
		Dim intColCnt, intRowCnt, intInsRow
		
		with frm1
		If .vspdData.SelModeSelCount > 0 Then 
			
			intInsRow = 0

			'Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols-2)
			Redim arrReturn(frm1.vspdData.SelModeSelCount, frm1.vspdData.MaxCols-2)
			For intRowCnt = 1 To frm1.vspdData.MaxRows

				frm1.vspdData.Row = intRowCnt
				
				If frm1.vspdData.SelModeSelected Then
					For intColCnt = 0 To frm1.vspdData.MaxCols - 2
						frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
						arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					Next
										
					intInsRow = intInsRow + 1
				End IF								
			Next
			arrReturn(intInsRow, 0) = frm1.hdnSupplierCd.value
			arrReturn(intInsRow, 1) = frm1.hdnGroupCd.value
			arrReturn(intInsRow, 2) = frm1.hdnSubcontraflg.value
			arrReturn(intInsRow, 3) = frm1.hdnGroupNm.value
			
			
		End if		
		
		end with
		
		Self.Returnvalue = arrReturn
		Self.Close()
		
	End Function
<%
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
%>
	Function CancelClick()
		Self.Close()
	End Function
<%
'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
'========================================================================================================
%>
	Function MousePointer(pstr1)
	      Select case UCase(pstr1)
	            case "PON"
					window.document.search.style.cursor = "wait"
	            case "POFF"
					window.document.search.style.cursor = ""
	      End Select
	End Function
	
<% 
'*******************************************  2.4 POP-UP 처리함수  **************************************
'*	기능: POP-UP																						*
'*	Description : POP-UP Call하는 함수 및 Return Value setting 처리										*
'********************************************************************************************************
%>

'===========================================  2.4.1 POP-UP Open 함수()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
'------------------------------------------  OpenGroup()  -------------------------------------------------
'	Name : OpenGroup()
'	Description : 
'--------------------------------------------------------------------------------------------------------- %>
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Or UCase(frm1.txtGroupCd.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGroup(arrRet)
	End If	

End Function 


Function SetGroup(byval arrRet)
	frm1.txtGroupCd.Value= arrRet(0)		
	frm1.txtGroupNm.Value= arrRet(1)	
	'frm1.txtGroupCd.focus	
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet,lgIsOpenPop
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else	
		frm1.txtPlantCd.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function	

'------------------------------------------  OpenSupplier()-------------------------------------------------
'	Name : OpenSupplier()
'	Description : Supplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"	
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSupplier(arrRet)
	End If	
	
End Function

Function SetSupplier(byval arrRet)
	
	frm1.txtSupplierCd.Value    = arrRet(0)		
	frm1.txtSupplierNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
	
End Function

'===========================================================================
' Function Name : OpenSoNo
' Function Desc : OpenSoNo Reference Popup
'===========================================================================
 Function OpenSoNo()

	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True
		
'	strRet = window.showModalDialog("../s3/s3111pa1.asp", "", _
'		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	iCalledAspName = AskPRAspName("S3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtSoNo.value = strRet
	End If	

End Function
<%
'===========================================================================
' Function Name : OpenTrackingNo
' Function Desc : OpenTrackingNo Reference Popup
'===========================================================================
%>

Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
    arrParam(2) = ""	'공장 
    arrParam(3) = ""	'모품목 
    arrParam(4) = ""	'수주번호 
    arrParam(5) = ""	'추가 Where절 
    
'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		lgBlnFlgChgValue = True
	End If	

End Function
 
<%
'=======================================  2.4.2 POP-UP Return값 설정 함수  ==============================
'=	Name : Set???()																						=
'=	Description : Reference 및 POP-UP의 Return값을 받는 부분											=
'========================================================================================================
%>
'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderBy()
	Dim arrRet
	
	On Error Resume Next
	
	'If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


<% '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ %>
<% '------------------------------------------  SetSorgCode()  --------------------------------------------------
'	Name : SetBPCd()
'	Description : SetSorgCode Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- %>

<%
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개별 프로그램마다 필요한 개발자 정의 Procedure(Sub, Function, Validation & Calulation 관련 함수)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>

<%
'########################################################################################################
'#						3. Event 부																		#
'#	기능: Event 함수에 관한 처리																		#
'#	설명: Window처리, Single처리, Grid처리 작업.														#
'#		  여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.								#
'#		  각 Object단위로 Grouping한다.																	#
'########################################################################################################
%>
<%
'********************************************  3.1 Window처리  ******************************************
'*	Window에 발생 하는 모든 Even 처리																	*
'********************************************************************************************************
%>
<%
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
%>
Sub Form_Load()

'parent.msgbox "aaa"

    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
'    ReDim lgPopUpR(C_MaxSelList - 1,1)
    
	'Call GetAdoFieldInf("M2111RA1_1","S","A")			              '☆: spread sheet 필드정보 query
	'
                                                                  ' 1. Program id
                                                                  ' 2. G is for Qroup , S is for Sort     
                                                                  ' 3. Spreadsheet no     
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
'    Call MakePopData(gDefaultT,gFieldNM,gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,C_MaxSelList)    ' You must not this line    
    Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitComboBox
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

Sub SetDefaultVal()
		Dim arrParam
		
		arrParam = arrParent(1)
		
		frm1.vspdData1.OperationMode = 3 
		frm1.vspdData.OperationMode = 5
		
		frm1.txtSupplierCd.value 	= arrParam(0)
		frm1.txtSupplierNm.value 	= arrParam(1)
		frm1.txtGroupCd.value 		= arrParam(2)
	'	msgbox PopupParent.gPurGrp
		If arrParam(2) = "" then
			frm1.txtGroupCd.value = PopupParent.gPurGrp
		End if

		frm1.txtGroupNm.value 		= arrParam(3)
		
		frm1.hdnSubcontraflg.value 	= arrParam(4)
		'frm1.hdnSubcontraflg.value 	= arrParam(4)
		
		
	'	If ubound(arrParam) = 5 then		'2002-12-04(LJT)
	'		frm1.hdnSTOflg.value = arrParam(5)
	'	Else 
	'		frm1.hdnSTOflg.value = "N"
	'	End If
		
		If arrParam(0) <> "" then		'2002-12-04(LJT)
			ggoOper.SetReqAttr		frm1.txtGroupCd, "Q"
			ggoOper.SetReqAttr		frm1.txtGroupNm, "Q"
		End if
		
		if  arrParam(2) <> "" then
			ggoOper.SetReqAttr		frm1.txtSupplierCd, "Q"
			ggoOper.SetReqAttr		frm1.txtSupplierNm, "Q"
		End if
		'	ggoOper.SetReqAttr		frm1.cboProcType, "Q"
			'ggoOper.SetReqAttr		frm1.txtGroupCd, "Q"
		
		
		frm1.txtFrPoDt.text 	= UnIDateAdd("d", -15, EndDate, PopupParent.gDateFormat)
		frm1.txtToPoDt.text 	= UnIDateAdd("d", +15, EndDate, PopupParent.gDateFormat)
		
		frm1.txtFrDlvyDt.text 	= EndDate
		frm1.txtToDlvyDt.text 	= UnIDateAdd("m", +1, EndDate, PopupParent.gDateFormat)
		
		' Tracker No.9743 공장코드 세팅 - 2005.07.22 =========================================
		frm1.txtPlantCd.value=PopupParent.gPlant
		frm1.txtPlantNm.value=PopupParent.gPlantNm
		' Tracker No.9743 공장코드 세팅 - 2005.07.22 =========================================		
		
End Sub

<%
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
%>
	Sub Form_QueryUnload(Cancel, UnloadMode)
	   
	End Sub
<%
'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
'********************************************************************************************************
%>



<%
'==========================================================================================
'   Event Name : OCX_Keypress()
'   Event Desc : 
'==========================================================================================
%>
	Sub txtFrPoDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtToPoDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtFrDlvyDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtToDlvyDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub
<%
'==========================================================================================
'   Event Name : txtFrPoDt
'   Event Desc :
'==========================================================================================
%>
Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
	End if
End Sub

<%
'==========================================================================================
'   Event Name : txtToPoDt
'   Event Desc :
'==========================================================================================
%>
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
	End if
End Sub

<%
'==========================================================================================
'   Event Name : txtFrDlvyDt
'   Event Desc :
'==========================================================================================
%>
Sub txtFrDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDlvyDt.Action = 7
	End if
End Sub

<%
'==========================================================================================
'   Event Name : txtToDlvyDt
'   Event Desc :
'==========================================================================================
%>
Sub txtToDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDlvyDt.Action = 7
	End if
End Sub

<%
'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
%>
	Function vspdData_DblClick(ByVal Col, ByVal Row)
	
	 If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
     End If
	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
	End Function
'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	
	Dim iPrevRows
	Dim strPurGrp, strPurNm, strBpCd, strProcureType, strVal
	ggoSpread.Source = frm1.vspdData1
	gMouseClickStatus = "SPC"   
	
	frm1.vspdData.MaxRows = 0
	
	Set gActiveSpdSheet = frm1.vspdData1
	Call SetPopupMenuItemInf("0000111111")

	If frm1.vspdData1.MaxRows = 0 Then Exit Sub
'	If Row = IgPrevRow Then Exit Sub
	
	If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
	Else
 		'------ Developer Coding part (Start)
 		
		frm1.vspdData1.Row = row
		
		frm1.vspdData1.Col = C_PurGrp 	
		strPurGrp = frm1.vspdData1.text
		frm1.hdnGroupCd.value = strPurGrp
		
		frm1.vspdData1.Col = C_PurGrpNm 	
		strPurNm = frm1.vspdData1.text
		frm1.hdnGroupNm.value = strPurNm
		
		frm1.vspdData1.Col = C_BpCd 	
		strBpCd = frm1.vspdData1.text
		frm1.hdnSupplierCd.value = strBpCd
		
		frm1.vspdData1.Col = C_ProCType 	
		strProcureType = frm1.vspdData1.text
		frm1.hdnProcuType.value = strProcureType
		
		
		'이성룡 추가 
		lgPageNo = ""
			
		If DbQuery2(strPurGrp,strBpCd,strProcureType) = False Then
			'	Call ResetToolBar(lgOldRow)
				Exit Sub 
		End If	
	 	'------ Developer Coding part (End)
 	End If
End Sub


Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
	Sub vspdData1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If		

		If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	    '☜: 재쿼리 체크	
			If lgPageNo_1 <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				If DbQuery = False Then
					Exit Sub
				End if
			End If
		End If		 
	End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
		Dim purGrp, bpCd, procType

		purGrp  = frm1.hdnGroupCd.value
		bpCd	= frm1.hdnSupplierCd.value
		procType = frm1.hdnProcuType.value
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If		
		

		If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
			If lgPageNo <> "" Then                '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				If DbQuery2(purGrp,bpCd,procType) = False Then
					Exit Sub
				End if
			End If
		End If		 
	End Sub
<% '#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'######################################################################################################### %>

<% '#########################################################################################################
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
'######################################################################################################### %>

<% '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* %>
<%
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
%>
Function FncQuery() 

	Dim strPlant
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtFrPoDt, frm1.txtToPoDt) = False Then Exit Function
	If ValidDateCheck(frm1.txtFrDlvyDt, frm1.txtToDlvyDt) = False Then Exit Function
   
    '-----------------------
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    
	'이성룡 추가 PLANT
	strPlant = frm1.txtPlantCd.value	
	frm1.hdnPlantCd.value = strPlant    
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function
	
<%
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
%>
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'☜: 비지니스 처리 ASP의 상태	
			strVal = strVal & "&txtFrPoDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToPoDt=" & .hdnToDt.value
			strVal = strVal & "&txtFrDlvyDt=" & .hdnFrDt2.value
			strVal = strVal & "&txtToDlvyDt=" & .hdnToDt2.value		
			strVal = strVal & "&txtSoNo=" & .hdnSoNo.value
			strVal = strVal & "&txtTrackingNo=" & .hdnTrackingNo.value		
			strVal = strVal & "&txtSupplier=" & .hdnSupplierCd.value
			strVal = strVal & "&txtGroup=" & .hdnGroupCd.value
			strVal = strVal & "&txtProcure=" & .hdnProcuType.value 
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			'이성룡 
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
						
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey   
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)
			strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)
			strVal = strVal & "&txtFrDlvyDt=" & Trim(.txtFrDlvyDt.text)
			strVal = strVal & "&txtToDlvyDt=" & Trim(.txtToDlvyDt.text)
			strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
			strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
			strVal = strVal & "&txtSupplier=" & Trim(.txtSupplierCd.value)
			strVal = strVal & "&txtGroup=" & Trim(.txtGroupCd.value)
			strVal = strVal & "&txtProcure=" & Trim(.cboProcType.value )
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			'이성룡 
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If				
	    strVal = strVal & "&lgPageNo="		 & lgPageNo_1						'☜: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '☜: 한번에 가져올수 있는 데이타 건수  
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
		strVal = strVal & "&txtGridNum="	 & "B"
		
		Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
		
    End With
    
    DbQuery = True    

End Function

<%
'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=========================================================================================================
%>
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	Dim lRow, i, strPurGrp, strBpCd, strProcuType 
		
	'lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Focus
		frm1.vspdData1.Row = 1	
		
		frm1.vspdData1.col = C_PurGrp
		strPurGrp = frm1.vspdData1.value
		frm1.vspdData1.col = C_BpCd 
		strBpCd = frm1.vspdData1.value
		frm1.vspdData1.col = C_ProCType 
		strProcuType = frm1.vspdData1.value
		
		
	
		frm1.hdnGroupCd.value = strPurGrp
		frm1.hdnSupplierCd.value = strBpCd
		frm1.hdnProcuType.value = strProcuType
		
		If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			Call DbQuery2(strPurGrp,strBpCd,strProcuType)
			lgIntFlgMode = PopupParent.OPMD_UMODE
		End If
		
		frm1.vspdData1.SelModeSelected = True		
	Else
	'	frm1.txtDnType.focus
	End If
	
	call SetSpreadLock

End Function
'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal purGrp, ByVal bpCd, ByVal procType)
	Err.Clear														'☜: Protect system from crashing
	DbQuery2 = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
	'frm1.vspdData.MaxRows = 0
	
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'☜: 비지니스 처리 ASP의 상태	
			strVal = strVal & "&txtFrPoDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToPoDt=" & .hdnToDt.value
			strVal = strVal & "&txtFrDlvyDt=" & .hdnFrDt2.value
			strVal = strVal & "&txtToDlvyDt=" & .hdnToDt2.value		
			strVal = strVal & "&txtSoNo=" & .hdnSoNo.value
			strVal = strVal & "&txtTrackingNo=" & .hdnTrackingNo.value		
			strVal = strVal & "&txtSupplier=" & bpCd
			strVal = strVal & "&txtGroup=" & purGrp
			strVal = strVal & "&txtProcure=" & procType
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			'이성룡 
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey   
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)
			strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)
			strVal = strVal & "&txtFrDlvyDt=" & Trim(.txtFrDlvyDt.text)
			strVal = strVal & "&txtToDlvyDt=" & Trim(.txtToDlvyDt.text)
			strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
			strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
			strVal = strVal & "&txtSupplier=" & bpCd
			strVal = strVal & "&txtGroup=" & purGrp
			strVal = strVal & "&txtProcure=" & procType
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			'이성룡 
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If				
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '☜: 한번에 가져올수 있는 데이타 건수  
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&txtGridNum="	 & "A"
		
		Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery2 = True    
End Function

Function DbQuery2Ok()
	DbQuery2Ok = False
	call SetSpreadLock_1
	DbQuery2Ok = true
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>조달구분</TD>
						<TD CLASS="TD6"><SELECT NAME="cboProcType" ALT="조달구분" STYLE="Width: 168px;" ></SELECT></TD>
						<TD CLASS="TD5" NOWRAP>구매그룹</TD>
						<TD CLASS="TD6">
						<INPUT TYPE=TEXT AlT="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
						<INPUT TYPE=TEXT AlT="구매그룹" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>공급처</TD>
						<TD CLASS="TD6">
						<INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 ALT="공급처" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
						<INPUT TYPE=TEXT AlT="공급처" ID="txtSupplierNm" tag="14X">
						</TD>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6">
						<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
						<INPUT TYPE=TEXT AlT="공장" ID="txtPlantNm" tag="14X">
						</TD>
					</TR>	
					<TR>
						<TD CLASS="TD5" NOWRAP>발주예정일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellpadding=0 cellspacing=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m2111ra1_1_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
									   <script language =javascript src='./js/m2111ra1_1_fpDateTime1_txtToPoDt.js'></script>
									</td>
								</tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>필요일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m2111ra1_1_fpDateTime2_txtFrDlvyDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m2111ra1_1_fpDateTime2_txtToDlvyDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>						
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>수주번호</TD>
						<TD CLASS="TD6"><INPUT NAME="txtSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=26 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo"></TD>
						<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
						<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No." TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=60% valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/m2111ra1_1_vaSpread1_vspdData1.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=40% valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/m2111ra1_1_vaSpread1_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
					<IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderBy()"></IMG></TD>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" WIDTH=100% SRC="../../blank.htm" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrDt2" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt2" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnTrackingNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnProcuType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSTOflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">


</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     