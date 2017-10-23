<%@ LANGUAGE="VBSCRIPT" %>

<!--
'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Reference Popup Business Part												*
'*  3. Program ID           : 																			*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Reference Popup															*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2000/03/29																*
'*  9. Modifier (First)     : Kang Tae Bum																*
'* 10. Modifier (Last)      : Kang Tae Bum																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :																			*
'*                            																			*
'********************************************************************************************************
 -->
<HTML>
<HEAD>
<!--
'########################################################################################################
'						1. 선 언 부																		
'########################################################################################################
 -->
<!--
'********************************************  1.1 Inc 선언  ********************************************
'	Description : Inc. Include																			
'********************************************************************************************************
 -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
 -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================
 -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit

'********************************************  1.2 Global 변수/상수 선언  *******************************
'	Description : 1. Constant는 반드시 대문자 표기														
'********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================

'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================

Const BIZ_PGM_ID = "a3125rb2.ASP"							 '☆: 비지니스 로직 ASP명 
Const C_SHEETMAXROWS_D = 30
Const C_MaxKey          = 15	
'*********************************************  1.3 변 수 선 언  ****************************************
'*	설명: Constant는 반드시 대문자 표기.																*
'********************************************************************************************************


Dim  arrReturn
Dim  arrParent
Dim  arrParam					
		
  
Dim  IsBpPop  
Dim  IsDocPop  
Dim  IsSpreadMode

Dim lgIsOpenPop                                          
Dim IsOpenPop   

'------ Set Parameters from Parent ASP ------ 
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)
	
top.document.title = "채권발생정보"

Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

dtToday = "<%=GetSvrDate%>"
Call PopupParent.ExtractDateFrom(dtToday, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

EndDate = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
StartDate = UNIDateAdd("M", -1, EndDate, PopupParent.gDateFormat)

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
	Redim arrReturn(0, 0)
	Self.Returnvalue = arrReturn
    frm1.lgPageNo.value         = ""
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
End Function

'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub  LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "A","NOCOOKIE","RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("I", "A","NOCOOKIE","RA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================

Sub SetDefaultVal()
	Dim lsMode 
	
	
	frm1.txtBpCd.value = arrParam(0)
	frm1.txtBpNm.value = arrParam(1)
	if  Trim(arrParam(2))	<>"" then
		frm1.txtDocCur.value = arrParam(2)	
	Else
		frm1.txtDocCur.value= PopupParent.gcurrency
	ENd if
	lsMode = arrParam(3)	
	frm1.txtBizCd.value = arrParam(4)
	frm1.txtBizNm.value = arrParam(5)
	frm1.htxtParentArno.value = arrParam(7)
	'txtRcptAmt.text = arrParam(8)
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
	If frm1.txtBpCd.value <> "" Then				
		Call ggoOper.SetReqAttr(frm1.txtBpCd,   "Q")		
	Else		
		Call ggoOper.SetReqAttr(frm1.txtBpCd,   "N")		
	End If
	
	If  Trim(arrParam(2)) <> "" Then				
		Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")		
	Else		
		Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")		
	End If	
	
	If  frm1.txtBizCd.value <> "" Then				
		Call ggoOper.SetReqAttr(frm1.txtBizCd,   "Q")		
	Else	
		IF lsMode = "Q" Then
			Call ggoOper.SetReqAttr(frm1.txtBizCd,   "N")		
		Else	
			Call ggoOper.SetReqAttr(frm1.txtBizCd,   "D")		
		END IF	
	End If	
	
	frm1.txtArDt.text	= StartDate
	frm1.txtToArDt.text	=  EndDate   

End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================

Sub InitSpreadSheet()
    frm1.vspdData.operationmode = 5
    Call SetZAdoSpreadSheet("a3125RA2","S","A","V20020310",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")

	Call SetSpreadLock()

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub  SetSpreadLock()
    
    frm1.vspdData.ReDraw = False
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()

    ggoSpread.SpreadUNLock GetKeyPos("A",10),-1, GetKeyPos("A",10)    
    frm1.vspdData.ReDraw = True


End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow, ByVal lRow2)

	
End Sub
'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
	
	Function OKClick()
		
		Dim intColCnt, intRowCnt, intInsRow,arrCnt
		

		if frm1.vspdData.SelModeSelCount > 0 Then 			
			intInsRow = 0
		
			Redim arrReturn(frm1.vspdData.SelModeSelCount - 1, C_MaxKey+2)
			
			For intRowCnt = 0 To frm1.vspdData.MaxRows - 1
			
				frm1.vspdData.Row = intRowCnt + 1
			
				If frm1.vspdData.SelModeSelected Then
					For intColCnt = 0 To C_MaxKey - 1 
						frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)
						
						arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					Next					
					
					'arrReturn(intInsRow, C_MaxKey  ) = txtBpCd.value
					'arrReturn(intInsRow, C_MaxKey + 1 ) = txtBpNm.value
					'arrReturn(intInsRow, C_MaxKey + 2 ) = txtDocCur.value	
					
																									
					intInsRow = intInsRow + 1
					
				End IF
			Next
		ENd if
		Self.Returnvalue = arrReturn
		Self.Close()
					
	End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================

	Function CancelClick()
		Self.Close()
	End Function

 '******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
'======================================================================================================
'   Event Name : OpenCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function  OpenCurrencyInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDocCur.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "거래통화팝업"					' 팝업 명칭 
	arrParam(1) = "b_currency"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtDocCur.value)							 	    ' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "거래통화" 			
	
    arrField(0) = "CURRENCY"							' Field명(0)
    arrField(1) = "CURRENCY_DESC"						' Field명(1)
    
    
    arrHeader(0) = "거래통화"						' Header명(0)
    arrHeader(1) = "거래통화명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    frm1.txtDocCur.focus
		Exit Function
	Else
		Call SetCurrencyInfo(arrRet)
	End If	

End Function

'======================================================================================================
'   Event Name : SetCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function SetCurrencyInfo(Byval arrRet)'	
	
		frm1.txtDocCur.value = arrRet(0)
		frm1.txtDocCur.focus
	    lgBlnFlgChgValue = True
		
End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	If iWhere = 1 Then
		if UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = "A_OPEN_AR"							' 채권과 연계(거래처 유무)
	arrParam(2) = frm1.txtARDt.Text								'FrDt
	arrParam(3) = frm1.txtToARDt.Text									'ToDt
	arrParam(4) = "B"							'B :매출 S: 매입 T: 전체 
	Select Case iWhere
		Case 1
			arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
		Case 2
			arrParam(5) = "SOL"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	End Select
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.TxtBpcd.focus
			Case 2
				frm1.TxtDealBpcd.focus
		ENd Select
		Exit Function
	Else
		Select Case iWhere
			Case 1
				Call SetBpCd(arrRet)
			Case 2
				Call SetDealBpCd(arrRet)
		ENd Select
	End If	
End Function
 '------------------------------------------  OpendealBpCd()  -------------------------------------------------
'	Name : OpendealBpCd()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenDealBpCd()'
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDEALBpCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "주문처팝업"
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.txtDealBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "주문처"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "주문처"		
    arrHeader(1) = "주문처명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	IF 	arrRet(0) <> "" then			
		Call SetDealBpCd(arrRet)
	Else
		frm1.TxtDealBpcd.focus
	end if
End Function


 
 '------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetBpCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetDealBpCd(Byval arrRet)'
	
	frm1.txtDealBpCd.value = arrRet(0)		
	frm1.txtDealBpNm.value = arrRet(1)
	frm1.TxtDealBpcd.focus
	lgBlnFlgChgValue = True
	
End Function

 '------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd()'
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBpCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "수금처팝업"
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "수금처"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "수금처"		
    arrHeader(1) = "수금처명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	IF 	arrRet(0) <> "" then			
		Call SetBpCd(arrRet)
	Else
		frm1.txtBpCd.focus
	end if
End Function
 '------------------------------------------  OpenBizCd()  -------------------------------------------------
'	Name : OpenBizCd()
'	Description : Cost PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizCd()'
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBizCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "사업장팝업"			' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizCd.Value)				' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "사업장"			
	
    arrField(0) = "BIZ_AREA_CD"							' Field명(0)
    arrField(1) = "BIZ_AREA_NM"							' Field명(1)
    
    arrHeader(0) = "사업장"					' Header명(0)
    arrHeader(1) = "사업장명"				' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	IF 	arrRet(0) <> "" then		
		Call SetBizCd(arrRet)
	Else
		frm1.txtBizCd.focus
	end if
	
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetBpCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBpCd(Byval arrRet)'
	
	frm1.txtBpCd.value = arrRet(0)		
	frm1.txtBpNm.value = arrRet(1)
	frm1.txtBpCd.focus				
	lgBlnFlgChgValue = True
	
End Function
 '------------------------------------------  SetBizCd()  --------------------------------------------------
'	Name : SetBizCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizCd(Byval arrRet)'
	
	frm1.txtBizCd.value = arrRet(0)		
	frm1.txtBizNm.value = arrRet(1)
	frm1.txtBizCd.focus				
	lgBlnFlgChgValue = True
	
End Function

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
	
	Call LoadInfTB19029  

	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                                     
	Call InitVariables		
	Call InitSpreadSheet()
	Call SetDefaultVal()
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================

Sub Form_QueryUnload(Cancel, UnloadMode)
		
End Sub

'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
'********************************************************************************************************


'==========================================  3.2.1 Call Fncquery() =======================================
'========================================================================================================

Function FncQuery()
	
	Dim IntRetCD
		
	'-----------------------
	'Check condition area
	'-----------------------
	Call InitVariables 	
	If Not chkField(Document, "1") Then									'This function check indispensable field
		Exit Function
	End If
		
	If CompareDateByFormat(frm1.txtArDt.text,frm1.txtToArDt.text,frm1.txtArDt.Alt,frm1.txtToArDt.Alt, _
               "970025",frm1.txtArDt.UserDefinedFormat,popupparent.gComDateType, true) = False Then
	   txtArDt.focus
	   Exit Function
	End If
		
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

		
	Call DbQuery()
End Function


'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
Sub txtArDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub txtToArDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub txtRcptAmt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    
	If frm1.vspdData.MaxRows < NewTop + PopupParent.VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	
    	If frm1.lgPageNo.value <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================

Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 and Col <> GetKeyPos("A",10)  Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	ElseIf Col = GetKeyPos("A",10)  Then		'//반제금액 editing할 수 있게 
		 frm1.vspdData.OperationMode = 0
	End If
End Function
'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Function vspdData_Click(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0  Then
		 frm1.vspdData.OperationMode = 5		'//다시 cell selecting 할 수 있는 모드 
	End If
End Function

'########################################################################################################
'#					     4. Common Function부															#
'########################################################################################################

'=======================================================================================================
'   Event Name : txtArDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtArDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtArDt.Action = 7    
        Call SetFocusToDocument("P")
		frm1.txtArDt.Focus                         
    End If
End Sub

'=======================================================================================================
'   Event Name : txtArDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtArDt_Change()
    
End Sub
'=======================================================================================================
'   Event Name : txtToArDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtToArDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToArDt.Action = 7     
        Call SetFocusToDocument("P")
		frm1.txtToArDt.Focus                         
    End If
End Sub
'=======================================================================================================
'   Event Name : txtToArDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtToArDt_Change()
    
End Sub


Sub txtBpCd_OnChange()
    
   
    Call CommonQueryRs("bp_cd, bp_nm", "b_biz_partner", "bp_cd =  " & FilterVar(frm1.txtBpCD.value , "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    
    if lgF0 <> "" and lgF0 <> "X" then
       frm1.txtBPCD.value = Left(lgF0, Len(lgF0)-1)
       frm1.txtBpNm.value = Left(lgF1, Len(lgF1)-1)
    else
       frm1.txtBPCD.value = ""
       frm1.txtBpNm.value = ""
    End if

End Sub


Sub txtDealBpCd_OnChange()
    
   
    Call CommonQueryRs("bp_cd, bp_nm", "b_biz_partner", "bp_cd =  " & FilterVar(frm1.txtDealBpCd.value , "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    
    if lgF0 <> "" and lgF0 <> "X" then
       frm1.txtDealBpCd.value = Left(lgF0, Len(lgF0)-1)
       frm1.txtDealBpNM.value = Left(lgF1, Len(lgF1)-1)
    else
       frm1.txtDealBpCd.value = ""
       frm1.txtDealBpNm.value = ""
    End if

End Sub

'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################


'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************

Function DbQuery()
	    Err.Clear                                                               '☜: Protect system from crashing
	
	    DbQuery = False                                                         '⊙: Processing is NG
	    
		Call LayerShowHide(1)


		With frm1		
	    	
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
		    .htxtBizCd.value		=Trim(frm1.txtBizCd.value)				'☜:     	
			.htxtBpCd.value			=Trim(frm1.txtBpCd.value)					'☆: 조회 조건 데이타 
			.htxtArDt.value			=Trim(frm1.txtArDt.text)					'☆: 조회 조건 데이타 
			.htxtToArDt.value		= Trim(frm1.txtToArDt.text)					'☆: 조회 조건 데이타 
			.htxtDocCur.value		=Trim(frm1.txtDocCur.value)					'☆: 조회 조건 데이타 
			.htxtDealBpCd.value		=Trim(frm1.txtDealBpCd.value)					'☆: 조회 조건 데이타 
			.htxtArNo.value			=Trim(frm1.txtArNo.value)					'☆: 조회 조건 데이타 
			.htxtParentArNo.value	=Trim(frm1.htxtParentArNo.value)
			.htxtDealBpCd_alt.value	=Trim(frm1.txtDealBpCd.alt)
			.htxtBpCd_alt.value		=Trim(frm1.txtBpCd.alt)
			
			
		'--------- Developer Coding Part (End) ------------------------------------------------------------									
			        
			.lgMaxCount.value		= C_SHEETMAXROWS_D
			.lgSelectListDT.value	= GetSQLSelectListDataType("A")
			.lgTailList.value		= MakeSQLGroupOrderByList("A")
			.lgSelectList.value		= EnCoding(GetSQLSelectList("A"))
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
               
		
	End with
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
    lgBlnFlgChgValue = True                                                 'Indicates that no value changed

End Function
Function DetailConditionClick()
	If DetailCondition.style.display = "none" Then
		DetailCondition.style.display = ""
		Call ggoOper.SetReqAttr(frm1.txtBpCd,   "D")
	Else
		DetailCondition.style.display = "none"
		If arrParam(0) <> "" Then
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "Q")
		Else
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "N")
		End If
	End If
End Function

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenOrderBy()
	Dim arrRet
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Popupparent.SORTW_WIDTH & "px; dialogHeight=" & Popupparent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables()
       Call InitSpreadSheet()       
   End If
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	

</HEAD>
<!--
'########################################################################################################
'						6. Tag 부																		
'########################################################################################################
 -->
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>				
						<TD CLASS=TD5 NOWRAP>채권일자</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a3125ra2_I648343729_txtArDt.js'></script>								
						&nbsp;~&nbsp;<script language =javascript src='./js/a3125ra2_I646284812_txtToArDt.js'></script></TD>												
						<TD CLASS=TD5 NOWRAP>거래통화</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="12NXXU"><IMG align=top name=btnCalType onclick="vbscript:OpenCurrencyInfo()" src="../../image/btnPopup.gif"  TYPE="BUTTON"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>수금처</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="12NXXU" ALT="수금처"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="수금처명"></TD>
						<TD CLASS=TD5 NOWRAP>사업장</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="사업장"><IMG SRC="../../image/btnPopup.gif" NAME="btnCostCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizCd()"> <INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14" ALT="사업장명">
						<IMG SRC="../../image/icon/QualityC.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="DetailConditionClick()" ></IMG></TD>					
					</TR>
					
					<TR ID="DetailCondition" style="display: none">
						<TD CLASS=TD5 NOWRAP>주문처</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtDealBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11NXXU" ALT="주문처"><IMG SRC="../../image/btnPopup.gif" NAME="btnDealBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtDealBpCd.Value,2)"> <INPUT TYPE=TEXT NAME="txtDealBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="주문처"></TD>
						<TD CLASS=TD5 NOWRAP>채권번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtArNo" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="채권번호"></TD>					
					</TR> 
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=80%>
	<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<script language =javascript src='./js/a3125ra2_vspdData_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=10>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenOrderBy()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					</TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
									 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=  <%=BizSize%> FRAMEBORDER=0 SCROLLING= no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>


<INPUT TYPE=hidden NAME="htxtBizCd" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtBpCd" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtArDt" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtToArDt" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtDocCur" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtDealBpCd" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtArNo" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtParentArNo" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtDealBpCd_alt" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="htxtBpCd_alt" tag="14" tabindex=-1>

<INPUT TYPE=hidden NAME="lgPageNo" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgMaxCount" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgSelectListDT" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgTailList" tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgSelectList" tag="14" tabindex=-1>


</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
