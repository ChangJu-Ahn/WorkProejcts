<%@ LANGUAGE="VBSCRIPT" %>

<!--======================================================================================================
'*  1. Module Name          : Template
'*  2. Function Name        : 
'*  3. Program ID           : A6114RA1
'*  4. Program Name         : 
'*  5. Program Desc         :  Ado query Sample with DBAgent(Sort)
'*  6. Component List       :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2001/04/18
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
========================================================================================================
=                          3.3 Client Side Script
========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "AcctCtrl.vbs">							</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit   

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "a6114rb1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 1					                          '☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop  
Dim lgMaxFieldCount                   
Dim lgCookValue 
Dim lgSaveRow 
Dim arrReturn
Dim arrParent
Dim IsOpenPop
Dim lgAuthorityFlag

'------ Set Parameters from Parent ASP -----------------------------------------------------------------------
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)

top.document.title = "계산서수정팝업"

'<% 
'	Call	ExtractDateFrom(GetSvrDate, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)
'
'	FirstDate= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, "01")		'☆: 초기화면에 뿌려지는 시작 날짜 
'	LastDate= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)		'☆: 초기화면에 뿌려지는 마지막 날짜 
'%>


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
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

    Redim arrReturn(0)
    Self.Returnvalue = arrReturn
    

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

	frm1.txtFrIssuedDt.Text	= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, "01")
	frm1.txtToIssuedDt.Text	= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
	frm1.txtFrVatLocAmt.Text=""
	frm1.txtToValLocAmt.Text = "" '"9999999999"
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
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","RA") %>
End Sub


'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
'Call SetCombo(frm1.cboPrcFlg, "T", "진단가")

'Call SetCombo(frm1.cboPrcFlg, "F", "가단가")
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
			
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    
    frm1.vspdData.OperationMode = 3
    Call SetZAdoSpreadSheet("A6114RA1Q01", "S", "A", "V20021108", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
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
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub

'===========================================================================
' Function Name : OpenOrderBy
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

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029															
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    FncQuery = False                                            
    
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
	if frm1.txtBpCd.value="" then 
		frm1.txtBPNm.value = ""
	end if 
	'$$
	If CompareDateByFormat(frm1.txtFrIssuedDt.text,frm1.txtToIssuedDt.text,frm1.txtFrIssuedDt.Alt,frm1.txtToIssuedDt.Alt, _
                        "970025",frm1.txtFrIssuedDt.UserDefinedFormat,PopupParent.gComDateType,True) = False Then			
		Exit Function
    End If
    '-----------------------
    'Query function call area
    '-----------------------
	frm1.vspdData.MaxRows = 0
    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function


'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

		strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        
        strVal = strVal & "?txtFrIssuedDt="       & Trim(.txtFrIssuedDt.text)
        strVal = strVal & "&txtToIssuedDt="   & Trim(.txtToIssuedDt.Text)
        if .rdo(0).checked = true then
			strVal = strVal & "&rdoInOut=I"           
        elseif .rdo(1).checked = true then
			strVal = strVal & "&rdoInOut=O"
        else
			strVal = strVal & "&rdoInOut="
        end if
        strVal = strVal & "&txtBPCd="            & Trim(.txtBPCd.value)
        strVal = strVal & "&txtFrVatLocAmt="     & Trim(.txtFrVatLocAmt.text) 
        strVal = strVal & "&txtToValLocAmt="     & Trim(.txtToValLocAmt.text) 
        strVal = strVal & "&txtTempGlNo="        & Trim(.txtTempGlNo.value)
		strVal = strVal & "&txtGlNo="			 & Trim(.txtGlNo.value)           
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
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
    lgSaveRow        = 1
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement

End Function



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
 '------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBPCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet)
		lgBlnFlgChgValue = True
	End If


End Function
 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

			arrParam(0) = "거래처 팝업"					' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "거래처"						' 조건필드의 라벨 명칭 

			arrField(0) = "BP_CD"						' Field명(0)
			arrField(1) = "BP_NM"						' Field명(1)
    
			arrHeader(0) = "거래처코드"					' Header명(0)
			arrHeader(1) = "거래처명"					' Header명(1)

    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBPCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet)
	End If	

End Function


 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet)
	With frm1
		' 거래처 
		.txtBPCd.focus
		.txtBPCd.value = UCase(Trim(arrRet(0)))
		.txtBPNM.value = arrRet(1)
				
					'.txtBPCd.focus
	End With
End Function

Function OpenReftempgl()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(4)	                           '권한관리 추가 (3 -> 4)

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5101ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "a5101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	lgIsOpenPop = True
	
'	Call CookiePage("TEMP_GL_POPUP")
	
	arrParam(4)	= lgAuthorityFlag              '권한관리 추가	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> ""  Then	
		Call SetRefTempGl(arrRet)
	End If
	frm1.txttempGlNo.focus
	
End Function

Function SetRefTempGl(ByVal arrRet)	
	With frm1
		.txttempGlNo.value = UCase(Trim(arrRet(0)))
    End With    
   
End Function


Function OpenRefGL()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(4)	                           '권한관리 추가 (3 -> 4)
	
	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5104ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "a5104ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	IsOpenPop = True
	
'	Call CookiePage("GL_POPUP")
	
	arrParam(4)	= lgAuthorityFlag 
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> ""  Then			
		Call SetRefGL(arrRet)
	End If
	frm1.txtGLNo.focus 
	
End Function

Function SetRefGL(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	
	With frm1
		.txtGlNo.Value = UCase(Trim(arrRet(0)))
    End With    
   
End Function


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim ii
	
    gMouseClickStatus = "SPC"   
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
	If Row < 1 Then Exit Sub
	
	lgCookValue = ""
	

    
End Sub
	
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
'   Event Name : vspdData_KeyPress
'   Event Desc : 
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function


'========================================================================================================
'   Event Name : fpdtFromEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtFromEnterDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpdtFromEnterDt.Focus
	End If
End Sub
'========================================================================================================
'   Event Name : fpdtToEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtToEnterDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpdtToEnterDt.Focus
	End If
End Sub



'==========================================================================================
'   Event Name : txtFrArDt
'   Event Desc :
'==========================================================================================

Sub  txtFrIssuedDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrIssuedDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFrIssuedDt.Focus
	End if
End Sub

Sub  txtFrIssuedDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		frm1.txtToIssuedDt.focus
		Call Fncquery()
	End IF
End Sub


'==========================================================================================
'   Event Name : txtToArDt
'   Event Desc :
'==========================================================================================

Sub  txtToIssuedDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToIssuedDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToIssuedDt.Focus
	End if
End Sub

Sub  txtToIssuedDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		frm1.txtFrIssuedDt.focus
		Call Fncquery()
	End IF
End Sub


Sub  txtFrVatLocAmt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

Sub  txtToValLocAmt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub



'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================

	Function CancelClick()
		Self.Close()			
	End Function

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
	
Function OKClick()
	
	If frm1.vspdData.ActiveRow > 0 Then 				
		Redim arrReturn(1)
		frm1.vspdData.row	= frm1.vspdData.ActiveRow
		frm1.vspdData.Col	= GetKeyPos("A",1)		
		arrReturn(0)		= frm1.vspdData.Text
	End if			
		
	Self.Returnvalue = arrReturn
	Self.Close()
					
End Function


Sub  vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub



'========================================================================================================



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
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
						<TD CLASS=TD5 NOWRAP>계산서일</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/a6114ra1_OBJECT1_txtFrIssuedDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/a6114ra1_OBJECT2_txtToIssuedDt.js'></script></TD>												
						<TD CLASS=TD5 NOWRAP>매입/매출 구분</TD>				
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=radio ALT="매입" class="radio" NAME="rdo" id="rdoIn"  Value="IN"  tag="11" CHECKED><label for="rdoIn">매입</label>
							<INPUT TYPE=radio ALT="매출" class="radio" NAME="rdo" id="rdoOut" Value="OUT" tag="11"><label for="rdoOut">매출</label></TD>
							
					</TR>			
					<TR>
						<TD CLASS=TD5 NOWRAP>거래처</TD>
						<TD CLASS="TD6">
							<INPUT TYPE=TEXT ID="txtBPCd" NAME="txtBPCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="거래처" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBPCd.Value, 1)">&nbsp;
							<INPUT TYPE=TEXT ID="txtBPNm" NAME="txtBPNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="거래처" tag="14X" ></TD>									
						<TD CLASS="TD5" NOWRAP>세금액</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/a6114ra1_fpDoubleSingle8_txtFrVatLocAmt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/a6114ra1_fpDoubleSingle8_txtToValLocAmt.js'></script></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>결의번호</TD>
						<TD CLASS="TD6">
							<INPUT TYPE=TEXT ID="txtTempGlNo" NAME="txtTempGlNo" SIZE=18 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" ALT="결의번호" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenReftempgl()">&nbsp;
							</TD>									
						<TD CLASS=TD5 NOWRAP>전표번호</TD>
						<TD CLASS="TD6">
							<INPUT TYPE=TEXT ID="txtGlNo" NAME="txtGlNo" SIZE=18 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" ALT="전표번호" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefGL()">&nbsp;
							</TD>																						
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%= HEIGHT_TYPE_03 %> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%= LR_SPACE_TYPE_20 %>>
				<TR HEIGHT=100%>
					<TD>
						<script language =javascript src='./js/a6114ra1_OBJECT3_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
									 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
