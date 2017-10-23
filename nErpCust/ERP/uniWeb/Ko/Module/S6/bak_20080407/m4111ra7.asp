<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4111ra7.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Purchase Goods Movement Reference ASP For C/C Dtl							*
'*  6. Comproxy List        : + M41118ListGrForExCcDtlSvr												*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2002/07/10																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son Bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'********************************************************************************************************
Response.Expires = -1							'☜ : ASP가 캐쉬되지 않도록 한다.
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!--
'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit		

'########################################################################################################
'#                       1.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       1.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "m4111rb7.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       1.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '☆: Fetch max count at once
Const C_SHEETMAXROWS    = 30
Const C_MaxKey       = 13                                           '☆: key count of SpreadSheet

	
'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================

Dim lgCookValue 
Dim gblnWinEvent

Dim arrReturn										<% '--- Return Parameter Group %>
Dim arrParam	
Dim arrParent
Dim lgIsOpenPop

Dim iDBSYSDate
Dim EndDate, StartDate

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'########################################################################################################
'#						2. Function 부																	#
'#																										#
'#	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기술					#
'#	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.							#
'#						 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함)			#
'########################################################################################################
'*******************************************  2.1 변수 초기화 함수  *************************************
'*	기능: 변수초기화																					*
'*	Description : Global변수 처리, 변수초기화 등의 작업을 한다.											*
'********************************************************************************************************
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
	Function InitVariables()
	    lgStrPrevKey     = ""
	    lgPageNo         = ""
		lgBlnFlgChgValue = False
	    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	    lgSortKey        = 1
			
		gblnWinEvent = False
		ReDim arrReturn(0,0)
		Self.Returnvalue = arrReturn
	End Function

'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
	Sub SetDefaultVal()


		arrParam = arrParent(1)

		frm1.txtApplicant.value = arrParam(0)
		frm1.txtApplicantNm.value = arrParam(1)
		
		frm1.txtFromDt.text = StartDate
		frm1.txtToDt.text = EndDate
	End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		<% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "RA") %>
	End Sub
'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
	Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("M4111RA7","S","A","V20021213",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
										C_MaxKey, "X","X")
		Call SetSpreadLock       
			      
	End Sub

'========================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
	Sub SetSpreadLock()
		frm1.vspdData.OperationMode = 5
		ggoSpread.SpreadLockWithOddEvenRowColor
	End Sub
		
'==========================================  2.2.6 InitComboBox()  ======================================
'=	Name : InitComboBox()																				=
'=	Description : Combo Display																			=
'========================================================================================================
'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================	
	Function OKClick()
	
		Dim intColCnt, intRowCnt, intInsRow

		If frm1.vspdData.SelModeSelCount > 0 Then 

			intInsRow = 0

			Redim arrReturn(frm1.vspdData.SelModeSelCount - 1, frm1.vspdData.MaxCols - 1)

			For intRowCnt = 0 To frm1.vspdData.MaxRows - 1

				frm1.vspdData.Row = intRowCnt + 1

				If frm1.vspdData.SelModeSelected Then
					For intColCnt = 0 To frm1.vspdData.MaxCols - 2
						frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)
						arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					Next

					intInsRow = intInsRow + 1

				End IF
			Next
		End if			
		
		Self.Returnvalue = arrReturn
		Self.Close()
	End Function	

		
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
	Function CancelClick()
		Redim arrReturn(1,1)
		arrReturn(0,0) = ""
		Self.Returnvalue = arrReturn
		Self.Close()
	End Function
	
'*******************************************  2.4 POP-UP 처리함수  **************************************
'*	기능: POP-UP																						*
'*	Description : POP-UP Call하는 함수 및 Return Value setting 처리										*
'********************************************************************************************************
'===========================================  2.4.1 POP-UP Open 함수()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
<%
'+++++++++++++++++++++++++++++++++++++++++++++  OpenItem()  +++++++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenItem()																					+
'+	Description : Sales Group PopUp Window Call															+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function OpenItem()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "품목"					
		arrParam(1) = "B_ITEM"						
		arrParam(2) = Trim(frm1.txtItem.value)		
		arrParam(3) = ""							
		arrParam(4) = ""							
		arrParam(5) = "품목"					

		arrField(0) = "item_cd"						
		arrField(1) = "item_nm"						
		arrField(2) = "Spec"						

		arrHeader(0) = "품목"					
		arrHeader(1) = "품목명"						

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetItem(arrRet)
		End If
	End Function
<%
'++++++++++++++++++++++++++++++++++++++++++++  OpenPurGroup()  ++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenPurGroup()																				+
'+	Description : Business Partner PopUp Window Call													+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function OpenPurGroup()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "구매그룹"						
		arrParam(1) = "B_PURCHASE_GROUP"					
		arrParam(2) = Trim(frm1.txtPurGroup.value)			
		arrParam(3) = ""									
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
			Call SetPurGroup(arrRet)
		End If
	End Function
<%
'++++++++++++++++++++++++++++++++++++++++++++  OpenPlant()  +++++++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenPlant()																					+
'+	Description : Business Partner PopUp Window Call													+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function OpenPlant()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "공장"						
		arrParam(1) = "B_PLANT"							
		arrParam(2) = Trim(frm1.txtPlant.value)			
		arrParam(3) = ""								
		arrParam(4) = ""							
		arrParam(5) = "공장"					

		arrField(0) = "PLANT_CD"					
		arrField(1) = "PLANT_NM"					

		arrHeader(0) = "공장"					
		arrHeader(1) = "공장명"					

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetPlant(arrRet)
		End If
	End Function

<%
'=======================================  2.4.2 POP-UP Return값 설정 함수  ==============================
'=	Name : Set???()																						=
'=	Description : Reference 및 POP-UP의 Return값을 받는 부분											=
'========================================================================================================
%>
<%
'+++++++++++++++++++++++++++++++++++++++++++  SetPlant()  +++++++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetPlant()																					+
'+	Description : Set Return array from Business Partner PopUp Window									+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function SetPlant(arrRet)
		frm1.txtPlant.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	End Function

<%
'+++++++++++++++++++++++++++++++++++++++++++  SetItem()  +++++++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetItem()																					+
'+	Description : Set Return array from Business Partner PopUp Window									+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function SetItem(arrRet)
		frm1.txtItem.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
	End Function



'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개별 프로그램마다 필요한 개발자 정의 Procedure(Sub, Function, Validation & Calulation 관련 함수)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'########################################################################################################
'#						3. Event 부																		#
'#	기능: Event 함수에 관한 처리																		#
'#	설명: Window처리, Single처리, Grid처리 작업.														#
'#		  여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.								#
'#		  각 Object단위로 Grouping한다.																	#
'########################################################################################################
'********************************************  3.1 Window처리  ******************************************
'*	Window에 발생 하는 모든 Even 처리																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
	Sub Form_Load()
		Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format		
		
		Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
		Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)				
		Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>	

		Call InitVariables														    '⊙: Initializes local global variables		
		Call SetDefaultVal		
		Call InitSpreadSheet()
		Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
		Call FncQuery()
	
	End Sub	
	

'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
'********************************************************************************************************
'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
	Function vspdData_DblClick(ByVal Col, ByVal Row)

		If Row = 0 Then Exit Function

		If frm1.vspdData.MaxRows = 0 Then Exit Function

		If Row > 0 Then Call OKClick()

	End Function

'========================================  3.3.2 vspdData_KeyPress()  ===================================
'=	Event Name : vspdData_KeyPress																		=
'=	Event Desc :																						=
'========================================================================================================
    Function vspdData_KeyPress(KeyAscii)
         On Error Resume Next
         If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
            Call OKClick()
         ElseIf KeyAscii = 27 Then
            Call CancelClick()
         End If
    End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
'=	Event Name : vspdData_TopLeftChange																	=
'=	Event Desc :																						=
'========================================================================================================
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If
    

		If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
			If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				If DbQuery = False Then
					Exit Sub
				End if
			End If
		End If
	End Sub

'==========================================================================================
'   Event Name : OCX_DbClick()
'   Event Desc : OCX_DbClick() 시 Calendar Popup
'==========================================================================================

	Sub txtFromDt_DblClick(Button)
		If Button = 1 Then
			frm1.txtFromDt.Action = 7
		End If
	End Sub

	Sub txtToDt_DblClick(Button)
		If Button = 1 Then
			frm1.txtToDt.Action = 7
		End If
	End Sub

'==================================== 3.2.23 txtDt_KeyPress()  ========================================
'   Event Name : txtDt_KeyPress
'   Event Desc : keyboard Operation
'=======================================================================================================

	Sub txtFromDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtToDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub
'########################################################################################################
'#					     4. Common Function부															#
'########################################################################################################
'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        

	
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function

	Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field        
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

	Call InitVariables												

	If Not chkField(Document, "1") Then				
		Exit Function
	End If

    If DbQuery = False Then Exit Function							

    FncQuery = True									
        
End Function


'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
	Function DbQuery()
		Dim strVal
		
		Err.Clear															<%'☜: Protect system from crashing%>

		DbQuery = False														<%'⊙: Processing is NG%>

		If LayerShowHide(1) = False Then
			Exit Function
		End If		



		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtItem=" & Trim(frm1.txtHItem.value)				<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtPlant=" & Trim(frm1.txtHPlant.value)
			strVal = strVal & "&txtFromDt=" & Trim(frm1.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(frm1.txtHToDt.value)
			strVal = strVal & "&txtApplicant=" & Trim(frm1.txtApplicant.value)					
	
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.value)				<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlant.value)
			strVal = strVal & "&txtFromDt=" & Trim(frm1.txtFromDt.text)
			strVal = strVal & "&txtToDt=" & Trim(frm1.txtToDt.text)
			strVal = strVal & "&txtApplicant=" & Trim(frm1.txtApplicant.value)
			
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If


    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
        strVal =     strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D                  '☜: 한번에 가져올수 있는 데이타 건수 
		
		strVal =	 strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal =	 strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal =	 strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

		DbQuery = True														<%'⊙: Processing is NG%>
	End Function
	
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtItem.focus
	End If
End Function

'===========================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================

Function OpenSortPopup()
	
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
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
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5 NOWRAP>품목</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItem" SIZE=10 MAXLENGTH=18 TAG="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON" OnClick="vbscript:OpenItem">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>공장</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlant" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" Onclick="vbscript:OpenPlant">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 TAG="14"></TD>
						</TR>		
						<TR>
							<TD CLASS=TD5 NOWRAP>출고일</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/m4111ra7_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/m4111ra7_fpDateTime2_txtToDt.js'></script>
							</TD>
							<TD CLASS=TD5>수입자</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="수입자">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
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
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/m4111ra7_vaSpread_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR HEIGHT="20">
			<TD WIDTH="100%">
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
							<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>
							<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>						
						<TD WIDTH=30% ALIGN=RIGHT>
							<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX=-1 ></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHItem" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHPlant" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHSONo" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHCurrency" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHIncoTerms" TAG="24" TABINDEX=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>