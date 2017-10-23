<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

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
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "a7127rb2.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const C_MaxKey          = 15					                          '☆: SpreadSheet의 키의 갯수

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          
Dim lgPopUpR                                              

Dim  lgQueryFlag
Dim  lgCode		
Dim  arrReturn
Dim  arrParent
Dim  arrParam		
		
Dim  IsOpenPop     
Dim  IsBpPop  
Dim  IsDocPop  

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인

 '------ Set Parameters from Parent ASP ------ 
arrParent        = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)

'top.document.title = "부서별자산정보"
top.document.title = PopupParent.gActivePRAspName


Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
'##
dtToday = "<%=GetSvrDate%>"
Call PopupParent.ExtractDateFrom(dtToday, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

EndDate = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
'StartDate = UNIDateAdd("M", -1, EndDate, PopupParent.gDateFormat)
StartDate = UNIDateClientFormat(PopupParent.gFiscStart)
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
	Redim arrReturn(0,0)
    
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    
	Self.Returnvalue = arrReturn
	
	' 권한관리 추가
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If	
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "RA") %>  
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
			
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	
	frm1.hORGCHANGEID.value = arrParam(1)
	
	frm1.txtFrAcqDt.text = StartDate
	frm1.txtToAcqDt.text = EndDate
	frm1.txtSoldyyyymm.text = arrParam(2)
	Call ggoOper.FormatDate(frm1.txtSoldyyyymm, PopupParent.gDateFormat, 2)
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    frm1.vspdData.operationmode = 5
    Call SetZAdoSpreadSheet("A7127RA2","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock()
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
    End With
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
	Dim ii,jj,kk
	Dim strtemp
	Dim aa
	If frm1.vspdData.SelModeSelCount > 0 Then 			

		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1,C_MaxKey)
		kk = 0
		For ii = 0 To frm1.vspdData.MaxRows - 1
			frm1.vspdData.Row = ii + 1
			If frm1.vspdData.SelModeSelected Then
				Call SetSpreadColumnValue("A", frm1.vspdData, frm1.vspdData.Col, frm1.vspdData.Row)
				For jj = 0 To C_MaxKey - 1
					frm1.vspdData.Col = jj + 1
					arrReturn(kk,jj)  = GetKeyPosVal("A", jj + 1)
					'aa = aa & "/" & arrReturn(kk,jj)
				Next			
				kk = kk + 1
			End If
		Next

	End If
	'msgbox aa
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

 '------------------------------------------  OpenBizCd()  -------------------------------------------------
'	Name : OpenAcctPopup()
'	Description : OpenAcctPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcctPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정코드팝업"			' 팝업 명칭 
	arrParam(1) = "A_ASSET_ACCT A, A_ACCT B"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtAcctCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "A.ACCT_CD = B.ACCT_CD "								' Where Condition
	arrParam(5) = "계정코드"			
	
    arrField(0) = "A.ACCT_CD"							' Field명(0)
    arrField(1) = "B.ACCT_NM"							' Field명(1)
    
    arrHeader(0) = "계정코드"					' Header명(0)
    arrHeader(1) = "계정명"				' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	IF 	arrRet(0) <> "" then		
		Call SetAcctCd(arrRet)
	end if
	
End Function


'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetBizCd()  --------------------------------------------------
'	Name : SetAcctCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetAcctCd(Byval arrRet)

	frm1.txtAcctCd.value = arrRet(0)		
	frm1.txtAcctNm.value = arrRet(1)
				
	'lgBlnFlgChgValue = True
	
End Function

'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = frm1.txtFrAcqDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToAcqDt.Text
	'arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition 
	
	' 권한관리 추가
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID	 
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(PopupParent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	frm1.txtDeptCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetDept(Byval arrRet)
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtFrAcqDt.text = arrRet(4)
		frm1.txtToAcqDt.text = arrRet(5)
End Function





Function OpenPopUp(Byval PopFg,strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	SELECT CASE UCase(PopFg)
	CASE "FR"		
			arrParam(0) = "자산마스터팝업"				' 팝업 명칭 
			arrParam(1) = "A_ASSET_MASTER"    				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			If lgInternalCd <> "" Then
				arrParam(4) = " INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = " INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			Else
				arrParam(4) = ""
			End If
			arrParam(5) = "자산"				' 조건필드의 라벨 명칭 

			arrField(0) = "ASST_NO"	     				' Field명(0)
			arrField(1) = "ASST_NM"			    		' Field명(1)
    
			arrHeader(0) = "자산번호"				' Header명(0)
			arrHeader(1) = "자산명"		  			' Header명(1)    	
	CASE "TO"
			arrParam(0) = "자산마스터팝업"				' 팝업 명칭 
			arrParam(1) = "A_ASSET_MASTER"    				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			If lgInternalCd <> "" Then
				arrParam(4) = " INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If
			
			If lgSubInternalCd <> "" Then
				arrParam(4) = " INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			Else
				arrParam(4) = ""
			End If			
			arrParam(5) = "자산"				' 조건필드의 라벨 명칭 

			arrField(0) = "ASST_NO"	     				' Field명(0)
			arrField(1) = "ASST_NM"			    		' Field명(1)
    
    
			arrHeader(0) = "자산번호"				' Header명(0)
			arrHeader(1) = "자산명"		  			' Header명(1)    
	END SELECT
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		With frm1
		select case PopFg
			case "FR"
				.txtFrAsstNo.focus
			case "TO"	
				.txtToAsstNo.focus
			end select 
		End With
		Exit Function
	Else
		Call SetPopUp(PopFg,arrRet)
	End If	

End Function

Function SetPopUp(Byval PopupFg,Byval arrRet)
	
	With frm1
	select case PopupFg
		case "FR"
			.txtFrAsstNo.focus
			.txtFrAsstNo.value = arrRet(0)
'			.txtFrAsstNm.value = arrRet(1)
		case "TO"	
			.txtToAsstNo.focus
			.txtToAsstNo.value = arrRet(0)
'			.txtToAsstNm.value = arrRet(1)
		end select 
	End With

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
	Call LoadInfTB19029()
	Call ggoOper.FormatField(Document, "1",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables()
	Call SetDefaultVal()
	Call InitSpreadSheet()
	Call CookiePage(0)

	frm1.txtFrAcqDt.focus
	Call ggoOper.SetReqAttr(frm1.txtSoldyyyymm, "Q")
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")		
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
	Dim strOrgChangeId
		
	FncQuery = False                                            
    
	Err.Clear            
	'-----------------------
	'Check condition area
	'-----------------------
	strOrgChangeId = frm1.hORGCHANGEID.value
	Call ggoOper.ClearField(Document, "2")						
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

	frm1.hORGCHANGEID.value = strOrgChangeId
	
	Call InitVariables() 											
		
	If Not chkField(Document, "1") Then									'This function check indispensable field
		Exit Function
	End If
	If Trim(frm1.txtSoldyyyymm.text) = "" Then 
		Exit Function
	End If
	
	If frm1.txtToAcqDt.text <> "" Then
		If CompareDateByFormat(frm1.txtFrAcqDt.text,frm1.txtToAcqDt.text,frm1.txtFrAcqDt.Alt,frm1.txtToAcqDt.Alt, _
							   "970025",frm1.txtFrAcqDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
		   frm1.txtFrAcqDt.focus
		   Exit Function
		End If
	End If

	lgQueryFlag = "1"	
	lgCode = ""
		
	 If DbQuery = False Then Exit Function
		 
	 FncQuery = True	
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
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
														
'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
Sub txtFrAcqDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub txtSoldyyyymm_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub  txtToAcqDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
    
End Sub
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SPC"   
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
'            ggoSpread.SSSort, lgSortKey
			ggoSpread.SSSort Col 
            lgSortKey = 2
        Else
'            ggoSpread.SSSort, lgSortKey
			ggoSpread.SSSort Col,lgSortKey 
            lgSortKey = 1
        End If    
    End If
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub  vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              		' Title cell을 dblclick했거나....
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows = 0 Then  	'NO Data
		Exit Sub
	End If
	Call OKClick()
End Sub


'########################################################################################################
'#					     4. Common Function부															#
'########################################################################################################
'=======================================================================================================
'   Event Name : txtFrAcqDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtFrAcqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrAcqDt.Action = 7                        
    End If
End Sub

'=======================================================================================================
'   Event Name : txtApDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtApDt_Change()
    
End Sub
'=======================================================================================================
'   Event Name : txtToAcqDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtToAcqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToAcqDt.Action = 7
    End If
End Sub

'Sub  txtSoldyyyymm_DblClick(Button)
'    If Button = 1 Then
'        frm1.txtSoldyyyymm.Action = 7
'    End If
'End Sub

'=======================================================================================================
'   Event Name : txtToApDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtToApDt_Change()
    
End Sub


'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################


'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************

Function DbQuery()
	Dim strVal
	Dim strFrYear,strFrMonth,strFrDay
	Dim strSoldyyyymm
    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

        strVal = BIZ_PGM_ID
	    Call ExtractDateFrom(frm1.txtSoldyyyymm.Text,frm1.txtSoldyyyymm.UserDefinedFormat,PopupParent.gComDateType,strFrYear,strFrMonth,strFrDay)    
        strSoldyyyymm = strFrYear & strFrMonth
        
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> PopupParent.OPMD_UMODE Then   ' This means that it is first search

			strVal = strVal & "?txtSoldyyyymm="	& Trim(strSoldyyyymm)		
			strVal = strVal & "&txtFrAcqDt="	& Trim(.txtFrAcqDt.Text)
			strVal = strVal & "&txtToAcqDt="	& Trim(.txtToAcqDt.Text)					'☆: 조회 조건 데이타
			strVal = strVal & "&txtFrAsstNo="	& Trim(.txtFrAsstNo.value)					'☆: 조회 조건 데이타
			strVal = strVal & "&txtToAsstNo="	& Trim(.txtToAsstNo.value)					'☆: 조회 조건 데이타
			strVal = strVal & "&txtAcctCd="		& Trim(.txtAcctCd.value)					'☆: 조회 조건 데이타
			strVal = strVal & "&txtDeptCd="		& Trim(.txtDeptCd.value)
			strVal = strVal & "&txtDeptCd_Alt="		& Trim(.txtDeptCd.alt)
			strVal = strVal & "&txtAcctCd_Alt="		& Trim(.txtAcctCd.alt)
			strVal = strVal & "&txtOrgChangeId="		& Trim(.hORGCHANGEID.value)
        Else
			strVal = strVal & "?txtSoldyyyymm="	& Trim(strSoldyyyymm)		
			strVal = strVal & "&txtFrAcqDt="	& Trim(.hFrAcqDt.value)
			strVal = strVal & "&txtToAcqDt="	& Trim(.hToAcqDt.value)					'☆: 조회 조건 데이타
			strVal = strVal & "&txtFrAsstNo="	& Trim(.hFrAsstNo.value)					'☆: 조회 조건 데이타
			strVal = strVal & "&txtToAsstNo="	& Trim(.hToAsstNo.value)					'☆: 조회 조건 데이타
			strVal = strVal & "&txtAcctCd="		& Trim(.hAcctCd.value)					'☆: 조회 조건 데이타
			strVal = strVal & "&txtDeptCd="		& Trim(.hDeptCd.value)
			strVal = strVal & "&txtDeptCd_Alt="		& Trim(.txtDeptCd.alt)
			strVal = strVal & "&txtAcctCd_Alt="		& Trim(.txtAcctCd.alt)
			strVal = strVal & "&txtOrgChangeId="		& Trim(.hORGCHANGEID.value)
        End If   

    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        'strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
	    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' 권한관리 추가
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
	
        Call RunMyBizASP(MyBizASP, strVal)							
        
    End With
    
    DbQuery = True
    
End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
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
       Call InitVariables
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
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>				
						<TD CLASS=TD5 NOWRAP>매각월</TD>
						<TD CLASS=TD6 NOWRAP>
							<!--<INPUT TYPE="Text" NAME="txtSoldyyyymm" SIZE=12 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="매각월">-->

							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSoldyyyymm CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="매각월" tag="14" id=FSoldyyyymm></OBJECT>');</SCRIPT>
							</OBJECT>
						</TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP>
						</TD>
					<TR>				
						<TD CLASS=TD5 NOWRAP>취득일자</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtFrAcqDt CLASSID=<%=gCLSIDFPDT%> ALT="시작취득일자" tag="11"> </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtToAcqDt CLASSID=<%=gCLSIDFPDT%> ALT="종료취득일자" tag="11"> </OBJECT>');</SCRIPT>
						</TD>
						<TD CLASS=TD5 NOWRAP>자산번호</TD>
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE="Text" NAME="txtFrAsstNo" SIZE=12 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="시작자산번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnFrAsstCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('FR',frm1.txtFrAsstNo.Value)">&nbsp;~&nbsp;
						<INPUT TYPE="Text" NAME="txtToAsstNo" SIZE=12 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="종료자산번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnToAsstCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('TO',frm1.txtToAsstNo.Value)">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>부서코드</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
											 <INPUT NAME="txtDeptNm" ALT="부서명"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
						<TD CLASS=TD5 NOWRAP>계정코드</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="계정코드" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenAcctPopup()">
											 <INPUT NAME="txtAcctNm" ALT="계정명"   MAXLENGTH="30" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR height=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% tag="2" HEIGHT=100% id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
						<IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Call FncQuery()">	</IMG>
					</TD>
					<TD ALIGN=RIGHT>
						<IMG SRC="../../image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hSoldyyyymm"  tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hFrAcqDt"  tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hToAcqDt"  tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hFrAsstNo" tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hToAsstNo" tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hAcctCd"	tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hDeptCd"   tag="34" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hORGCHANGEID"   tag="24" TABINDEX = "-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
