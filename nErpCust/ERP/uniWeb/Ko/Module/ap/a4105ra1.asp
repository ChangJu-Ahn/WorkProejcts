
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :  Ado query Sample with DBAgent(Sort)
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

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
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "a4105rb1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 14					                      '☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  lgIsOpenPop                                          
Dim  lgPopUpR                                              
Dim  lgQueryFlag
Dim  lgCode		
		
Dim  IsOpenPop     
Dim  IsBpPop  
Dim  IsDocPop  
     	
Dim  arrReturn
Dim  arrParent
Dim  arrParam

	 '------ Set Parameters from Parent ASP ------ 
arrParent        = window.dialogArguments
Set PopupParent = arrParent(0)	 
arrParam		= arrParent(1)




top.document.title = PopupParent.gActivePRAspName	
'top.document.title = "채무발생정보"

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
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "A","NOCOOKIE", "RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("I", "A", "NOCOOKIE", "RA") %>
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	If arrParam(0) <> "" Then
		frm1.txtBpCd.value = arrParam(0)
		frm1.txtBpNm.value = arrParam(1)			
	End If
		
	If arrParam(2) <> "" Then	
		frm1.txtDocCur.value = arrParam(2)				
	End If	
	
	frm1.htxtAllcDt.value	= arrParam(3) 
    frm1.htxtAllcAlt.value	= arrParam(4) 		
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
	If frm1.txtBpCd.value <> "" Then		
		Call ggoOper.SetReqAttr(frm1.txtBpCd,   "Q")		
	Else
		Call ggoOper.SetReqAttr(frm1.txtBpCd,   "N")		
	End If
	
	If frm1.txtDocCur.value <> "" Then		
		Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")		
	Else
		Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")		
	End If	

	If arrParam(3) =  "Q" Then					
		Call ggoOper.SetReqAttr(frm1.txtBizCd,   "N")		
	Else			
		Call ggoOper.SetReqAttr(frm1.txtBizCd,   "D")		
	End If
	
	frm1.txtApDt.text	= UNIDateAdd("M", -1, arrParam(3),PopupParent.gDateFormat)
	frm1.txtToApDt.text	= arrParam(3)
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	frm1.vspdData.OperationMode = 5
	Call AppendNumberPlace("6","15","2")
    Call SetZAdoSpreadSheet("a4105ra1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock() 
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.Source = frm1.vspdData
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

	If frm1.vspdData.SelModeSelCount > 0 Then 			
		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1,C_MaxKey)
		kk = 0
		For ii = 0 To frm1.vspdData.MaxRows - 1
			frm1.vspdData.Row = ii + 1			
			If frm1.vspdData.SelModeSelected Then
				For jj = 0 To C_MaxKey - 1
					frm1.vspdData.Col	 = GetKeyPos("A",jj + 1)		
					arrReturn(kk,jj) = frm1.vspdData.Text
				Next			
				kk = kk + 1
			End If
		Next	
	End If		
		
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
 '------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBpCd.className = "protected" Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = "A_OPEN_AP"							' 채권과 연계(거래처 유무)
	arrParam(2) = frm1.txtApDt.text								'FrDt
	arrParam(3) = frm1.txtToApDt.text									'ToDt
	arrParam(4) = "S"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = "PAYTO"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If 	arrRet(0) <> "" then			
		Call SetBpCd(arrRet)
	Else
		frm1.txtBpCd.focus
	End If	
End Function
 '------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd()'
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	IF IsBpPop = False Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "지급처팝업"
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "지급처"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "지급처"		
    arrHeader(1) = "지급처명"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If 	arrRet(0) <> "" then			
		Call SetBpCd(arrRet)
	Else
		frm1.txtBpCd.focus
	End If
End Function

 '------------------------------------------  OpenBizCd()  -------------------------------------------------
'	Name : OpenBizCd()
'	Description : Cost PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBizCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "사업장팝업"			' 팝업 명칭 
	arrParam(1) = "B_Biz_Area"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "사업장"			
	
    arrField(0) = "BIZ_AREA_CD"							' Field명(0)
    arrField(1) = "BIZ_AREA_NM"							' Field명(1)
    
    arrHeader(0) = "사업장"					' Header명(0)
    arrHeader(1) = "사업장명"				' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If 	arrRet(0) <> "" then		
		Call SetBizCd(arrRet)
	Else
		frm1.txtBizCd.focus
	End if
End Function

'======================================================================================================
'   Event Name : OpenCurrencyInfo
'   Event Desc : 
'=======================================================================================================
Function  OpenCurrencyInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDocCur.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "거래통화팝업"					' 팝업 명칭 
	arrParam(1) = "b_currency"							' TABLE 명칭 
	arrParam(2) = strCode						 	    ' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "거래통화" 			
	
    arrField(0) = "CURRENCY"							' Field명(0)
    arrField(1) = "CURRENCY_DESC"						' Field명(1)
        
    arrHeader(0) = "거래통화"						' Header명(0)
    arrHeader(1) = "거래통화명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
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
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetBpCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBpCd(Byval arrRet)	
	frm1.txtBpCd.value = arrRet(0)		
	frm1.txtBpNm.value = arrRet(1)
	frm1.txtBpCd.focus
End Function
 '------------------------------------------  SetBizCd()  --------------------------------------------------
'	Name : SetBizCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizCd(Byval arrRet)
	frm1.txtBizCd.value = arrRet(0)		
	frm1.txtBizNm.value = arrRet(1)
	frm1.txtBizCd.focus
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


'==========================================  3.2.1 FncQuery =======================================
'========================================================================================================
Function FncQuery()
	FncQuery = False                                            
    
    Err.Clear                                                   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
	
	lgQueryFlag = "1"	
	lgCode = ""		
	
	If Not ChkQueryDate Then
		Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------

    If DbQuery = False Then Exit Function

    FncQuery = True													
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 

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
Sub  txtApDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub  txtToApDt_KeyPress(KeyAscii)
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
    
	If frm1.vspdData.MaxRows < NewTop + PopupParent.VisibleRowCnt(frm1.vspdData,NewTop) Then	    
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
            ggoSpread.SSSort, lgSortKey
			ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
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
	Call OKClick
End Sub

'########################################################################################################
'#					     4. Common Function부															#
'########################################################################################################

'=======================================================================================================
'   Event Name : txtApDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtApDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtApDt.Action = 7                        
        Call SetFocusToDocument("P")
		Frm1.txtApDt.Focus
		
    End If
End Sub

'=======================================================================================================
'   Event Name : txtApDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtApDt_Change()
    
End Sub
'=======================================================================================================
'   Event Name : txtToApDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtToApDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToApDt.Action = 7     
        Call SetFocusToDocument("P")
		Frm1.txtToApDt.Focus
		                   
    End If
End Sub

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

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> PopupParent.OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtBizCd=" & Trim(.txtBizCd.Value)
			strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtApDt=" & Trim(.txtApDt.Text)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtToApDt=" & Trim(.txtToApDt.Text)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtDocCur=" & Trim(.txtDocCur.value)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtBpcd_Alt=" & Trim(.txtBpCd.alt)
			strVal = strVal & "&txtBizCd_Alt=" & Trim(.txtBizCd.alt)  	
        Else
			strVal = strVal & "?txtBizCd=" & Trim(.htxtBizCd.value)
			strVal = strVal & "&txtBpCd=" & Trim(.htxtBpCd.value)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtApDt=" & Trim(.htxtApDt.value)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtToApDt=" & Trim(.htxtToApDt.value)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtDocCur=" & Trim(.htxtDocCur.value)					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtBpcd_Alt=" & Trim(.txtBpCd.alt)
			strVal = strVal & "&txtBizCd_Alt=" & Trim(.txtBizCd.alt) 
        End If   
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&txtAllcDt="	     & Trim(.htxtAllcDt.value)
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

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

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================


'========================================================================================================
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'========================================================================================================
'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function  OpenOrderByPopup()

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

'=======================================================================================================
'   Function Name : ChkQueryDate
'   Function Desc : 
'=======================================================================================================
Function ChkQueryDate()
	chkQueryDate= True
	
	If PopupParent.CompareDateByFormat(frm1.txtApDt.text,frm1.txtToApDt.text,frm1.txtApDt.Alt,frm1.txtToApDt.Alt, _
	               "970025",frm1.txtApDt.UserDefinedFormat,PopupParent.gComDateType, True) = False Then
		chkQueryDate= False
		frm1.txtApDt.focus
		Exit Function
	End If

	
	If CompareDateByFormat(frm1.txtApDt.text,frm1.htxtAllcDt.Value,frm1.txtApDt.Alt,frm1.htxtAllcAlt.value, _
   	           "970025",frm1.txtApDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   frm1.txtApDt.focus
	   Exit Function
	End If
	
	If CompareDateByFormat(frm1.txtToApDt.text,frm1.htxtAllcDt.Value,frm1.txtToApDt.Alt, frm1.htxtAllcAlt.value,_
   	           "970025",frm1.txtToApDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   chkQueryDate= False
	   frm1.txtToApDt.focus
	   Exit Function
	End If

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
						<TD CLASS=TD5 NOWRAP>채무일자</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a4105ra1_fpDateTime_txtApDt.js'></script>								
							    &nbsp;~&nbsp;<script language =javascript src='./js/a4105ra1_fpDateTime1_txtToApDt.js'></script></TD>												
						<TD CLASS=TD5 NOWRAP>거래통화</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: left" tag ="12NXXU"><IMG align=top name=btnCalType onclick="vbscript:OpenCurrencyInfo(txtDocCur.Value)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>지급처</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12NXXU" ALT="지급처"><IMG align=top name=btnBpcd onclick="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14" ALT="거래처명"></TD>
						<TD CLASS=TD5 NOWRAP>사업장</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag=11NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizCd()"> <INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14" ALT="사업장명"></TD>					
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<script language =javascript src='./js/a4105ra1_vspdData_vspdData.js'></script>
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
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Call FncQuery()">	</IMG>&nbsp;
					<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME=Config ONMOUSEOUT="javascript:MM_swapImgRestore()" ONMOUSEOVER="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ONCLICK="OpenOrderByPopup()"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtBizCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBpCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtApDt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToApDt"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDocCur"    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtAllcDt"	tag="14">
<INPUT TYPE=HIDDEN NAME="htxtAllcAlt"      tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

