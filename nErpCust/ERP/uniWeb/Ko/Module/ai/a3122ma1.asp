
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a3122ma1
'*  4. Program Name         : 기초가수금정보 등록 
'*  5. Program Desc         : 기초가수금정보 등록 수정 삭제 조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2003/02/07
'*  8. Modified date(Last)  : 2003/02/11
'*  9. Modifier (First)     : lee nam yo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'            1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'*********************************************************************************************************
'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
Const BIZ_PGM_QUERY_ID = "a3122mb1.asp"							'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID  = "a3122mb2.asp"							'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID   = "a3122mb3.asp"							'☆: Head Query 비지니스 로직 ASP명 


Const RcptJnlType = "SR"

Dim IsOpenPop						' Popup
Dim	lgFormLoad
Dim	lgQueryOk						' Queryok여부 (loc_amt =0 check)
Dim lgQueryState					' 조회후 상태 flag
Dim lgstartfnc

<%
Dim dtToday
dtToday = GetSvrDate
%> 




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.1 Common Group -1
' Description : This part declares 1st common function group
'=======================================================================================================
'*******************************************************************************************************





'======================================================================================================
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'=======================================================================================================
Sub initSpreadPosVariables()

End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE						'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False								'Indicates that no value changed
    lgIntGrpCount = 0										'initializes Group View Size
    lgStrPrevKey = ""										'initializes Previous Key
    lgLngCurRows = 0										'initializes Deleted Rows Count

	lgstartfnc = False
	lgFormLoad = True
	lgQueryOk  = False
End Sub
 
'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtRcptDt.Text = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDocCur.value = parent.gCurrency
	frm1.htxtGlDt.value=frm1.txtRcptDt.Text
	frm1.txtXchRate.text = 1
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	
	frm1.txtRcptNo.focus
 
	lgBlnFlgChgValue = False
	lgQueryOk = False	
	lgQueryState = False
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>    
End Sub


'========================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
   
End Sub

'========================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
   
End Sub

'========================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
 
End Sub

'======================================================================================================
' Function Name : GetSpreadColumnPos()
' Function Desc : This method Call saved columnorder
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	
End Sub

'======================================================================================================
' Function Name : OpenPopupGL
' Function Desc : This method Open The Popup window for GL
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1) 
	Dim IntRetCD
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If Trim(frm1.txtGlNo.value) = "" Then
		IntRetCD = DisplayMsgBox("970000","X" , frm1.txtGlNo.Alt, "X") 	
		IsOpenPop = False
		Exit Function
	End If
	


	If IsOpenPop = True Then Exit Function
 
	arrParam(0) = Trim(frm1.txtGlNo.value)					'회계전표번호 
	arrParam(1) = ""										'Reference번호 

	IsOpenPop = True
	  
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Or UCase(frm1.txtDept.className) = "PROTECTED" Then Exit Function 
	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtRcptDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDept.focus
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
				.txtDept.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtRcptDt.text = arrRet(3)
				call txtDept_OnBlur()  
				.txtDept.focus
        End Select
	End With
End Function     
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBpCd.className = parent.UCN_PROTECTED Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscReturnVal(iwhere)
		Exit Function
	Else  
		Call SetReturnVal(arrRet,iWhere)
		lgBlnFlgChgValue = True
	End If 

End Function
'=========================================================================================================
' Name : Open???()
' Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'      ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
Function OpenPopup(Byval strCode, Byval iWhere )
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strNoteFg,IntRetCD

	If IsOpenPop = True Then Exit Function
 
	Select Case iWhere
		Case 0
		Case 2
			If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function 

			arrParam(0) = "거래처팝업" 
			arrParam(1) = "B_BIZ_PARTNER"
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래처코드"
 
			arrField(0) = "BP_CD" 
			arrField(1) = "BP_NM"
			 
			arrHeader(0) = "거래처코드"  
			arrHeader(1) = "거래처명" 
		Case 3    
			If IsOpenPop = True Or UCase(frm1.txtDocCur.className) = "PROTECTED" Then Exit Function
		 
			arrParam(0) = "거래통화팝업" 
			arrParam(1) = "B_CURRENCY"    
			arrParam(2) = Trim(frm1.txtDocCur.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래통화"
 
			arrField(0) = "CURRENCY" 
			arrField(1) = "CURRENCY_DESC" 
			 
			arrHeader(0) = "거래통화"  
			arrHeader(1) = "거래통화명" 
		Case 4
			If frm1.txtRcptType.className = parent.UCN_PROTECTED Then Exit Function
		 
			arrParam(0) = frm1.txtRcptType.Alt							' 팝업 명칭 
			arrParam(1) = "a_jnl_item"									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtRcptType.Value)					' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "jnl_type =  " & FilterVar(RcptJnlType  , "''", "S") & ""			' Where Condition
			arrParam(5) = frm1.txtRcptType.Alt							' 조건필드의 라벨 명칭 

			arrField(0) = "JNL_CD"										' Field명(0)
			arrField(1) = "JNL_NM"										' Field명(1)
			 
			arrHeader(0) = frm1.txtRcptType.Alt							' Header명(0)
			arrHeader(1) = frm1.txtRcptTypeNm.Alt						' Header명(1)
	End Select
 
	IsOpenPop = True
 
	If iWhere = 0 Then
		Dim iCalledAspName
		iCalledAspName = AskPRAspName("a3122ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3122ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
	
		arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam), _
	       "dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")     
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")   
	End If 

	IsOpenPop = False
 
	If arrRet(0) = "" Then
		Call EscReturnVal(iwhere)
		Exit Function
	Else  
		Call SetReturnVal(arrRet,iWhere)
	End If 
End Function
'===========================================================================================
' Name : SetReturnVal()
' Description : Plant Popup에서 Return되는 값 setting
'===========================================================================================
Function EscReturnVal(byval iWhere)
	With frm1 
		Select Case iWhere   
			Case 0
				.txtRcptNo.focus  
			Case 2 'OpenBpCd
				.txtBpCd.focus
			Case 3 'OpenCurrency
				.txtDocCur.focus
			Case 4 'OpenBpCd
				.txtRcptType.focus
		End Select 

		
	End With
End Function
'===========================================================================================
' Name : SetReturnVal()
' Description : Plant Popup에서 Return되는 값 setting
'===========================================================================================
Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1 
		Select Case field_fg   
			Case 0
				.txtRcptNo.Value     = arrRet(0)   
				.txtRcptNo.focus  
			Case 2 'OpenBpCd
				.txtBpCd.Value       = arrRet(0)
				.txtBpNm.Value       = arrRet(1)
				.txtBpCd.focus
			Case 3 'OpenCurrency
				.txtDocCur.Value     = arrRet(0)
				Call txtDocCur_OnChange()
				.txtDocCur.focus
			Case 4 'OpenBpCd
				.txtRcptType.Value   = arrRet(0)
				.txtRcptTypeNm.Value = arrRet(1)
				.txtRcptType.focus
		End Select 

		If field_fg <> 0 Then lgBlnFlgChgValue = True
	End With
End Function

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************



'=====================================================================================================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=====================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029()														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field

	Call InitVariables()   
    Call SetDefaultVal()
    Call SetToolbar("1110100100001111")    
	Call chgBtnDisable(1)
 
	
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
    Dim IntRetCD 
    
    FncQuery = False                                                        
    lgstartfnc = True
    Err.Clear                                                               
	'-----------------------
    'Check previous data area
    '-----------------------
   
    If lgBlnFlgChgValue = True  Then  
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")     
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables()														'⊙: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then											'⊙: This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Query function Call area
    '-----------------------
    Call DbQuery()																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    lgstartfnc = False	 
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
	   
    FncNew = False                                                          
    lgstartfnc = True 
	lgQueryState = False
	    
  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True  Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                               '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                               '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                '⊙: Lock  Suitable  Field
	
    Call InitVariables()   
    Call txtDocCur_OnChange()
    Call SetDefaultVal()    
    Call chgBtnDisable(1)
    Call ggoOper.SetReqAttr(frm1.chkConfFg,"D")	
    
    FncNew = True																'⊙: Processing is OK
    lgFormLoad = True							' tempgldt read
    lgstartfnc = False
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False															'⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                   'Check if there is retrived data
        intRetCD = DisplayMsgBox("900002","x","x","x")					'☜ 바뀐부분 
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")			'☜ 바뀐부분 
    If IntRetCD = vbNo Then
        Exit Function
    End If
	'-----------------------
    'Delete function Call area
    '-----------------------
    Call DbDelete()																'☜: Delete db data
    
    FncDelete = True															'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
	
    FncSave = False                                                         
    
    Err.Clear                                                               

    
    If lgBlnFlgChgValue = False Then							'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")					'⊙: Display Message(There is no changed data.)
		Exit Function
    End If
    
    If Not chkField(Document, "2") Then											'⊙: Check required field(Single area)
		Exit Function
    End If
	
	frm1.txtGlFlag.value = ""
  	frm1.htxtGlDt.value =frm1.txtgldt.text
  	
  	If frm1.chkConfFg.checked= True Then
		frm1.txtConfFg.value = "C"
	Else
		frm1.txtConfFg.value = "U"	
	End If
	
  	'-----------------------
    'Save function Call area
    '-----------------------
    Call DbSave()																	'☜: Save db data
    
    FncSave = True  
 End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy() 
	
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
   
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	parent.FncPrint()    
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
    Call FncExport(parent.C_SINGLEMULTI)            '☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                                                    
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
	
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False

	If lgBlnFlgChgValue = True Then											'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")					'데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.3 Common Group - 3
' Description : This part declares 3rd common function group
'=======================================================================================================
'*******************************************************************************************************





'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    DbDelete = False																		'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtRcptNo=" & Trim(frm1.txtRcptNo.value)							'☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True																			'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()																		'☆: 삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "2")													'Clear Condition Field
	Call ggoOper.LockField(Document, "N")													'Lock  Suitable  Field    
	Call InitVariables()																	'Initializes local global variables
	Call SetDefaultVal()
	Call chgBtnDisable(1)
			       
	frm1.txtRcptNo.Value = ""
	frm1.txtRcptNo.focus
	Call SetToolbar("1110110000001111")   
	
	lgBlnFlgChgValue = False    
End Function
 
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         '⊙: Processing is NG
	Err.Clear
     
	Call LayerShowHide(1)
    Dim strVal
	If lgIntFlgMode = parent.OPMD_UMODE Then                                           
		strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & parent.UID_M0001			'☜: 
		strVal = strVal & "&txtRcptNo=" & frm1.txtRcptNo.value						'Hidden의 검색조건으로 Query
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & parent.UID_M0001			'☜: 
		strVal = strVal & "&txtRcptNo=" & Trim(frm1.txtRcptNo.value)		'현재 검색조건으로 Query
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
 
    DbQuery = True                                                          '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()  
	Dim varData
	
	lgIntFlgMode = parent.OPMD_UMODE										'⊙: Indicates that current mode is Update mode

    lgQueryOk= True
	lgQueryState = True
	
	Call SetToolbar("1111100000011111")										'⊙: 버튼 툴바 제어 
	 
	Call txtDocCur_OnChange()
	Call txtDept_OnBlur()
	
	'//Button 조작 및 PROTECT
	
	If frm1.chkConfFg.checked = True Then
		
		Call chgBtnDisable(1)
		
		IF Trim(frm1.txtGlNo.value) = ""  Then
		
			if Trim(frm1.htxttempGlNo.value) <> "" then
				Call ggoOper.SetReqAttr(frm1.txtRcptType,"Q")
				Call ggoOper.SetReqAttr(frm1.txtDept,"Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptDt,"Q")
				Call ggoOper.SetReqAttr(frm1.txtRefNo,"Q")
				Call ggoOper.SetReqAttr(frm1.txtBpCd,"Q")
				Call ggoOper.SetReqAttr(frm1.txtDocCur,"Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptAmt,"Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptLocAmt,"Q")
				Call ggoOper.SetReqAttr(frm1.txtDesc,"Q")
				Call ggoOper.SetReqAttr(frm1.chkConfFg,"Q")		'//
				Call chgBtnDisable(3)
			frm1.txtGlDt.text=frm1.htxtGlDt.value
			else
			
				Call ggoOper.SetReqAttr(frm1.txtRcptType,"N")
				Call ggoOper.SetReqAttr(frm1.txtDept,"N")
				Call ggoOper.SetReqAttr(frm1.txtRcptDt,"N")
				Call ggoOper.SetReqAttr(frm1.txtRefNo,"D")		'//
				Call ggoOper.SetReqAttr(frm1.txtBpCd,"D")		'//
				Call ggoOper.SetReqAttr(frm1.txtDocCur,"N")
				Call ggoOper.SetReqAttr(frm1.txtRcptAmt,"N")
				Call ggoOper.SetReqAttr(frm1.txtRcptLocAmt,"D")	'//
		  		Call ggoOper.SetReqAttr(frm1.txtDesc,"D")		'//
				Call ggoOper.SetReqAttr(frm1.chkConfFg,"D")		'//
			End if
		Else
		
			IF Trim(frm1.txtRcptType.value) <> ""  Then
			
				Call ggoOper.SetReqAttr(frm1.txtRcptType,"Q")
				Call ggoOper.SetReqAttr(frm1.txtDept,"Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptDt,"Q")
				Call ggoOper.SetReqAttr(frm1.txtRefNo,"Q")
				Call ggoOper.SetReqAttr(frm1.txtBpCd,"Q")
				Call ggoOper.SetReqAttr(frm1.txtDocCur,"Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptAmt,"Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptLocAmt,"Q")
				Call ggoOper.SetReqAttr(frm1.txtDesc,"Q")
				Call ggoOper.SetReqAttr(frm1.chkConfFg,"Q")		'//
				Call chgBtnDisable(3)
			frm1.txtGlDt.text=frm1.htxtGlDt.value
			End If	
		End If	
		
		
		
	Else
		IF Trim(frm1.txtGlNo.value) = "" Then
	
			Call ggoOper.SetReqAttr(frm1.txtRcptType,"N")
		    Call ggoOper.SetReqAttr(frm1.txtDept,"N")
		    Call ggoOper.SetReqAttr(frm1.txtRcptDt,"N")
			Call ggoOper.SetReqAttr(frm1.txtRefNo,"D")		'//
		    Call ggoOper.SetReqAttr(frm1.txtBpCd,"D")		'//
		    Call ggoOper.SetReqAttr(frm1.txtDocCur,"N")
		    Call ggoOper.SetReqAttr(frm1.txtRcptAmt,"N")
		    Call ggoOper.SetReqAttr(frm1.txtRcptLocAmt,"D")	'//
		  	Call ggoOper.SetReqAttr(frm1.txtDesc,"D")		'//
			Call ggoOper.SetReqAttr(frm1.chkConfFg,"D")		'//
	
			If Trim(frm1.txtRcptType.value) = "" Then
				Call chgBtnDisable(1)
			Else
				Call chgBtnDisable(2)
			End If	
		End If	
		frm1.txtGlDt.text=frm1.htxtGlDt.value
	End If	
	
	frm1.txtRcptNo.focus
	lgBlnFlgChgValue = False    
	lgQueryOk= False	
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim pAr0061 
    Dim IntRows 
    Dim IntCols 
    Dim vbIntRet 
    Dim lStartRow 
    Dim lEndRow 
    Dim boolCheck 
    Dim lGrpcnt 
	Dim strVal, strDel
	Dim ApAmt, PayAmt
 
    DbSave = False                                                          '⊙: Processing is NG
    
    On Error Resume Next													'☜: Protect system from crashing
	
	Call LayerShowHide(1)
	 
	With frm1
		.txtMode.value = parent.UID_M0002									'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태   
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 
	        
	DbSave = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()	
													'☆: 저장 성공후 실행 로직 
    Call ggoOper.ClearField(Document, "2")							'⊙: Clear Contents  Field
    
    Call InitVariables()													'⊙: Initializes local global variables
    Call DbQuery()
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************
    
'===================================== CurFormatNumericOCX()  =======================================
' Name : CurFormatNumericOCX()
' Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 입금액 
		ggoOper.FormatFieldByObjectOfCur .txtRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoOper.FormatFieldByObjectOfCur .txtClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 청산금액 
		ggoOper.FormatFieldByObjectOfCur .txtSttlAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		
	End With
End Sub

'====================================================================================================
'	Name : XchLocRate()
'	Description : 환율이 변경되는 Factor 가 변했을 때 수정되는 Local Amt. Setting
'====================================================================================================
Sub XchLocRate()
	Dim ii

	With frm1
		.txtRcptLocAmt.text = "0"
	End With
End Sub


'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************





'===================================== PopSaveSpreadColumnInf()  ======================================
' Name : PopSaveSpreadColumnInf()
' Description : 이동한 컬럼의 정보를 저장 
'====================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'===================================== PopRestoreSpreadColumnInf()  ======================================
' Name : PopRestoreSpreadColumnInf()
' Description : 컬럼의 순서정보를 복원함 
'====================================================================================================
Sub  PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************



'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
			.vspdData.Row = NewRow
			.vspdData.Col = 0
			If .vspddata.Text = ggoSpread.DeleteFlag Then
				Exit Sub       
			End if
		End With
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
End Sub

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	
End Sub

'======================================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata_DblClick(ByVal Col,ByVal Row)
 
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 가수금구분이 예적금,어음일 경우에만 어음번호,계좌번호 Enabled 되게.
'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
	
End Sub





'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************





'=======================================================================================================
' Name : txtDocCur_onblur()
' Description : 
'=======================================================================================================
Function txtDocCur_onblur()
  
End Function

'========================================================================================
' Function Name :txtXchRate_onblur
' Function Desc : 
'========================================================================================
Function txtXchRate_onblur()
	lgBlnFlgChgValue = True
End Function



'=======================================================================================================
'   Event Name : txtRcptDt_DblClick(Button)
'   Event Desc : 입금일관련 달력을 호출한다.
'=======================================================================================================
Sub txtRcptDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtRcptDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtRcptDt.Focus
    End If
    Call txtRcptDt_OnBlur()  
End Sub

'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGlDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtGlDt.Focus
    End If
End Sub


'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		Call CurFormatNumericOCX()
	End If    
	
	If lgQueryOk<> True Then
		Call XchLocRate()
	End If	
End Sub


'==========================================================================================
'   Event Name : txtRcp_Change
'   Event Desc : 
'==========================================================================================
Sub txtRcptAmt_Change()
	lgBlnFlgChgValue = True
	frm1.txtRcptLocAmt.text = "0"
End Sub

'==========================================================================================
'   Event Name : txtRcptLocAmt_Change
'   Event Desc : 
'==========================================================================================
Sub txtRcptLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtDept_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtDept_OnBlur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtRcptDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDept.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDept.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtRcptDt.Text, gDateFormat,""), "''", "S") & "))"			
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDept.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDept.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------

End Sub

'==========================================================================================
'   Event Name : txtRcptDt_onBlur
'   Event Desc : 
'==========================================================================================
Sub txtRcptDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
   If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
	
				If LTrim(RTrim(.txtDept.value)) <> "" and Trim(.txtRcptDt.Text <> "") Then
					'----------------------------------------------------------------------------------------
						strSelect	=			 " Distinct org_change_id "    		
						strFrom		=			 " b_acct_dept(NOLOCK) "		
						strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDept.value)), "''", "S") 
						strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
						strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
						strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtRcptDt.Text, gDateFormat,""), "''", "S") & "))"			
	
					IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
					If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
							.txtDept.value = ""
							.txtDeptNm.value = ""
							.hOrgChangeId.value = ""
							.txtDept.focus
					End if
				End If
			End With
		'----------------------------------------------------------------------------------------
		End If
	End IF
	
	Call XchLocRate()
End Sub

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.8 HTML Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************

'========================================================================================
' Function Name :txtBpCd_onBlur
' Function Desc : 
'========================================================================================
Function txtBpCd_onBlur()
	If frm1.txtBpCd.value = "" then
	 frm1.txtBpNm.value = ""
	End if
End Function

'========================================================================================================
' Name : chgBtnDisable
' Desc : 버튼변경 
'========================================================================================================
Sub chgBtnDisable(Gubun)
	Select Case Gubun
		Case 1		'//버튼 둘다 비활성 
				frm1.btnConf.disabled	=	True
				frm1.btnUnCon.disabled	=	True	
				frm1.txtgldt.text=""
				Call ggoOper.SetReqAttr(frm1.txtGlDt,"Q")		
		Case 2		'//확정버튼 활성화, 취소버튼비활성화 
				frm1.btnConf.disabled	=	False
				frm1.btnUnCon.disabled	=	True	
				Call ggoOper.SetReqAttr(frm1.txtGlDt,"N")
				
		Case 3		'//확정버튼 비활성화, 취소버튼 활성화 
				frm1.btnConf.disabled	=	True
				frm1.btnUnCon.disabled	=	False	
				frm1.txtgldt.text=""
				Call ggoOper.SetReqAttr(frm1.txtGlDt,"Q")
					
	End Select
	

End Sub

'========================================================================================================
' Name : fnBttnConf
' Desc : 전표작업 
'========================================================================================================
Sub fnBttnConf(Gubun)
	Dim IntRetCD
	Dim strVal
	Dim strYear,strMonth,strDay, txtGlDt
	Err.Clear                                                                    '☜: Clear err status
	   	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("189217", parent.VB_INFORMATION,"X","X") '☜ 데이타가 변경되었습니다. 저장후 진행하십시오.
		Exit Sub
	End If	
	
	frm1.txtGlFlag.value = Gubun
	if Gubun="D" then 
		frm1.htxtGlDt.Value= frm1.txtrcptdt.text
	Else
		frm1.htxtGlDt.value =frm1.txtgldt.text			
	
	end if
	
	Call Dbsave
	
	Set gActiveElement = document.ActiveElement   
End Sub


Sub chkConfFg_onchange()
	
	If frm1.chkConfFg.checked = True Then
		frm1.txtConfFg.value = "C"
		Call chgBtnDisable(1)	
		
	Else
		frm1.txtConfFg.value = "U"	
		
		IF Trim(frm1.txtGlNo.value) = "" Then
			If lgIntFlgMode = Parent.OPMD_CMODE	 Then
				Call chgBtnDisable(1)
			Else
				Call chgBtnDisable(2)
			
			End If	
			frm1.txtgldT.text=frm1.txtrcptdt.text
		Else
			IF lgIntFlgMode = Parent.OPMD_UMODE	 Then
				Call chgBtnDisable(3)
				frm1.txtgldt.text=frm1.htxtGlDt.Value
			End If	
		End If	
	
		
	End If
	lgBlnFlgChgValue = True	
	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>
<!-- '#########################################################################################################
'            6. Tag부 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기초가수금등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>   
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- 본문내용  -->
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
							<TD CLASS="TD5" NOWRAP>가수금번호</TD>
							<TD CLASS="TD6"><INPUT NAME="txtRcptNo" TYPE="Text" MAXLENGTH=18 tag="12XXXU" ALT="가수금번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtRcptNo.value, 0)"></TD>
							<TD CLASS="TDT"></TD>
							<TD CLASS="TD6"></TD>
						</TR>
					</TABLE>
					</FIELDSET>
				</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH="100%">     
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>가수금유형</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRcptType" SIZE=10 MAXLENGTH=20  tag="22XXXU" ALT="가수금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRcptType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup('',4)">&nbsp;<INPUT TYPE=TEXT NAME="txtRcptTypeNm" SIZE=25 tag="24" ALT="가수금유형명"></TD>
							<TD CLASS=TD5 NOWRAP>프로젝트</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="프로젝트" MAXLENGTH=25 SIZE=25 tag="2X"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>부서</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDept" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDept.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=25 tag="24" ></TD>            
							<TD CLASS="TD5" NOWRAP>입금일자</TD>                           
							<TD CLASS="TD6" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtRcptDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="입금일자"> </OBJECT>');</SCRIPT>               
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>수금처</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBpCd" ALT="수금처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 2)"> <INPUT NAME="txtBpNm" TYPE="Text" SIZE=25 tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>참조번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="참조번호" MAXLENGTH="30" STYLE="TEXT-ALIGN: left" tag="21XXXU">&nbsp;</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>거래통화</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" SIZE=10 tag="22XXXU" MAXLENGTH="3"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.value, 3)"></TD>
							<TD CLASS=TD5 NOWRAP>환율</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name="txtXchRate" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="환율" tag="24x5"> </OBJECT>');</SCRIPT>&nbsp;
							</TD>         
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>입금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtRcptAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="입금액" tag="22x2"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
							<TD CLASS=TD5 NOWRAP>입금액(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtRcptLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="입금액(자국)" tag="21x2"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
						</TR>
						<TR>                    
							<TD CLASS=TD5 NOWRAP>반제금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtClsAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="반제금액" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
							<TD CLASS=TD5 NOWRAP>반제금액(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtClsLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="반제금액(자국)" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
						</TR>        
						<TR>                      
							<TD CLASS=TD5 NOWRAP>청산금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtSttlAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="청산금액" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
							<TD CLASS=TD5 NOWRAP>청산금액(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 name=txtSttlLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="청산금액(자국)" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>                                 
						</TR>
						<TR>                      
							<TD CLASS=TD5 NOWRAP>잔액</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtBalAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="잔액" tag="24"> </OBJECT>');</SCRIPT>&nbsp;
						    </TD>
							<TD CLASS=TD5 NOWRAP>잔액(자국)</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtBalLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="잔액(자국)" tag="24"> </OBJECT>');</SCRIPT> &nbsp;
						    </TD>
						</TR>       
						<TR>
							<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGLNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="회계전표번호"></TD>
							<TD CLASS="TD5" NOWRAP><LABEL FOR=chkConfFg>기회계처리</LABEL></TD>
							<TD CLASS="TD6" NOWRAP><INPUT type="checkbox" CLASS="STYLE CHECK"  NAME=chkConfFg ID=chkConfFg tag="1" onclick=chkConfFg_onchange()></TD>
						
						</TR>        
						<TR>
							<TD CLASS="TD5" NOWRAP>전표일자</TD>
							<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtGlDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="전표일자" tag="22X1" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
						
							<TD CLASS=TD5 NOWRAP>비고</TD>
							<TD CLASS=TD656 NOWRAP ><INPUT TYPE=TEXT NAME="txtDesc" SIZE=50 MAXLENGTH=100 tag="2X" ALT="비고"></TD>        						
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
<TR HEIGHT="20">
	<TD HEIGHT=20 WIDTH="100%">
		<FIELDSET CLASS="CLSFLD">
			<TABLE <%=LR_SPACE_TYPE_40%>>
				<TR HEIGHT=20>
					<TD CLASS=TD6 NOWRAP></TD>
					<TD CLASS=TD6 NOWRAP></TD>
					<TD CLASS=TD6 NOWRAP><BUTTON NAME="btnConf" CLASS="CLSMBTN" OnClick="VBScript:Call fnBttnConf('C')">전표확정</BUTTON>&nbsp<BUTTON NAME="btnUnCon" CLASS="CLSMBTN" OnClick="VBScript:Call fnBttnConf('D')">전표취소</BUTTON></TD>
					<TD CLASS=TD6 NOWRAP></TD>									
				</TR>						
			</TABLE>
		</FIELDSET>
	</TD>
</TR>
 <TR>
	<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
	</TD>
 </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA><% '업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
	<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
	<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">
	<INPUT TYPE=hidden NAME="txtGlFlag"   tag="24" TABINDEX="-1">
	<INPUT TYPE=hidden NAME="txtConfFg"   tag="24" TABINDEX="-1">
	<INPUT TYPE=hidden NAME="htxtGlDt"   tag="34" TABINDEX="-1">
	<INPUT TYPE=hidden NAME="htxtTempGlNO"   tag="34" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

