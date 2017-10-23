
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

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit																	'☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID1 = "a4103mb2.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "a4103mb1.asp"	
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
												'☆: Server에서 한번에 fetch할 최대 데이타 건수 
										'☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          
Dim lgQueryFlag																	' 신규조회 및 추가조회 구분 Flag

Dim  IsOpenPop          
<% 
Dim dtToday
	dtToday = GetSvrDate                                                 
%>

Dim C_CHECK_FG
Dim C_Ap_DT 
Dim C_GL_DT 
Dim C_Ap_NO 
Dim C_BP_NM 
Dim C_DOC_CUR 
Dim C_Ap_AMT 
Dim C_Ap_LOC_AMT 
Dim C_DEPT_CD 
Dim C_TEMP_GL_NO 
Dim C_GL_NO 
Dim C_CONF_FG

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================
'======================================================================================================
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'=======================================================================================================
Sub initSpreadPosVariables()
	C_CHECK_FG   = 1
	C_AP_DT      = 2
	C_GL_DT      = 3
	C_AP_NO      = 4 
	C_BP_NM      = 5 
	C_DOC_CUR    = 6
	C_AP_AMT     = 7
	C_AP_LOC_AMT = 8
	C_DEPT_CD    = 9
	C_TEMP_GL_NO = 10
	C_GL_NO      = 11
	C_CONF_FG    = 12
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgPageNo  = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtFromReqDt.text  =  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtToReqDt.text    =  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.GIDate.text		=  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtFromReqDt.focus	
	frm1.cboConfFg.value	=	"U"  
	frm1.hOrgChangeId.value = parent.gChangeOrgId	  
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "*","NOCOOKIE" , "MA") %>     
	<% Call LoadBNumericFormatA("Q", "*", "NOCOOKIE", "MA") %>                           
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Function

'========================================================================================
'                       InitComboBox()
' ========================================================================================  
Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
End Sub
'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet()
    Call initSpreadPosVariables()
    
    With frm1.vspdData
    
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

		.Redraw = False

		.MaxCols = C_CONF_FG + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols															'공통콘트롤 사용 Hidden Column
		.ColHidden = True    
		.MaxRows = 0
		    
		Call AppendNumberPlace("6","3","0")
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck  C_CHECK_FG   , ""			       ,5 , ,"", True,-1
		ggoSpread.SSSetDate   C_AP_DT      , "채무일자"    ,10, 2, parent.gDateFormat  
		ggoSpread.SSSetDate   C_GL_DT      , "전표일자"    ,10, 2, parent.gDateFormat  
		ggoSpread.SSSetEdit   C_AP_NO      , "채무번호"    ,20,3
		ggoSpread.SSSetEdit   C_BP_NM      , "거래처"      ,12,3        
		ggoSpread.SSSetEdit   C_DOC_CUR    , "통화"        ,12,3        
		ggoSpread.SSSetFloat  C_AP_AMT     , "채무액"      ,15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_AP_LOC_AMT , "채무액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_DEPT_CD    , "부서코드"    ,10, ,,10,2
		ggoSpread.SSSetEdit   C_TEMP_GL_NO , "결의전표번호",15,3        
		ggoSpread.SSSetEdit   C_GL_NO      , "회계전표번호",15,3        
		ggoSpread.SSSetCheck  C_CONF_FG    , ""			       ,3 , ,"", True,-1

		Call ggoSpread.MakePairsColumn(C_AP_AMT,C_AP_LOC_AMT)
		
		Call ggoSpread.SSSetColHidden(C_CONF_FG,C_CONF_FG,True)

		.Redraw = True 
    End With

	Call SetSpreadLock()
End Sub

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromReqDt.Focus     
    End If
End Sub
Sub txtToReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToReqDt.Focus     
    End If
End Sub

'=======================================================================================================
'	Name : OpenPopUpGL()
'	Description : 
'======================================================================================================= 
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim arrField
	Dim intFieldCount
	Dim i	
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .MaxRows > 0 Then
		
			.Row = .ActiveRow
			.Col =C_GL_NO		
		
			arrParam(0) = Trim(.Text)	'회계전표번호 
			arrParam(1) = ""			'Reference번호 
		End If
	End With

	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'=======================================================================================================
'	Name : OpenPopuptempGL()
'	Description : 
'=======================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim arrField
	Dim intFieldCount
	Dim i
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .MaxRows > 0 Then

			.Row = .ActiveRow
			.Col = C_TEMP_GL_NO
		
			arrParam(0) = Trim(.Text)	'Temp_gl_no
			arrParam(1) = ""			'Reference번호 
		End If
	End With

	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function
'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtFromReqDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToReqDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
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
		frm1.txtFromReqDt.text = arrRet(4)
		frm1.txtToReqDt.text = arrRet(5)
		frm1.txtDeptCd.focus
End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'Code Condition
   	arrParam(1) = ""									'채무와 연계(거래처 유무)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "S"									'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function
 '========================================== 2.4.2 Open???()  =============================================
'	Name : OpenPopUp()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 1
			arrParam(0) = "거래처 팝업"  				' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER"	 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "거래처"	    			' 조건필드의 라벨 명칭 

			arrField(0) = "BP_CD"						' Field명(0)
			arrField(1) = "BP_NM"						' Field명(1)
    
			arrHeader(0) = "거래처"	     			' Header명(0)
			arrHeader(1) = "거래처명"				' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				frm1.txtBpCd.value  = arrRet(0)
				frm1.txtBpNm.value  = arrRet(1)			    
				frm1.txtBpCd.focus
		End Select

	End With
End Function

Sub txtBpCd_onBlur()
	If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = ""
	End If
End Sub	

'==========================================================================================
'   Event Name : txtDeptCd_Onchange
'   Event Desc : 
'==========================================================================================
Sub txtDeptCD_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromReqDt.Text, gDateFormat,""), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtToReqDt.Text, gDateFormat,""), "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
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


'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	Dim arrField
	Dim intFieldCount
	Dim i,j	

	With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SpreadUnLock C_CHECK_FG	,-1 , C_CHECK_FG							'승인체크박스 UnLoking
		ggoSpread.SpreadLock   C_AP_DT		,-1 ,C_AP_DT							'승인체크박스부터 전표일전까지 Locking
		ggoSpread.SpreadUnLock C_GL_DT		,-1 , C_GL_DT							'전표일 UnLocking
		ggoSpread.SpreadLock   C_AP_NO		,-1									'전표일다음부터 끝까지 Locking
				
		.vspdData.ReDraw = True
    End With
End Sub





'================================== SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim arrField
	Dim intFieldCount
	Dim i,j
	
     With frm1
	    .vspdData.ReDraw = False

		.vspdData.Col = C_CONF_FG

		If .vspdData.Text = "C" Then
			ggoSpread.SSSetProtected	C_GL_DT,  pvStartRow, pvEndRow
		Else
			ggoSpread.SSSetRequired		C_GL_DT,  pvStartRow, pvEndRow
		End If	
		
		.vspdData.ReDraw = True
    End With
End Sub
'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method call saved columnorder
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		
				C_CHECK_FG   = 1
				C_AP_DT      = 2
				C_GL_DT      = 3
				C_AP_NO      = 4 
				C_BP_NM      = 5 
				C_DOC_CUR    = 6
				C_AP_AMT     = 7
				C_AP_LOC_AMT = 8
				C_DEPT_CD    = 9
				C_TEMP_GL_NO = 10
				C_GL_NO      = 11
				C_CONF_FG    = 12
	End select
End Sub
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
    Call LoadInfTB19029()														

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)    
    Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables()														
	Call InitSpreadSheet()
	Call InitComboBox()	
	Call SetDefaultVal()	
	
    Call SetToolbar("1100000000000111")		

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
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Sub txtFromReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then
		frm1.txtToReqDt.focus
		Call FncQuery
	End if
End Sub

Sub txtToReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt.focus
		Call FncQuery
	end if
End Sub

Sub GIDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If                                                  
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables() 											
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
    
	If CompareDateByFormat(frm1.txtFromReqDt.text,frm1.txtToReqDt.text,frm1.txtFromReqDt.Alt,frm1.txtToReqDt.Alt, _
        	               "970025",frm1.txtFromReqDt.UserDefinedFormat,parent.gComDateType, true) = False Then
		frm1.txtFromReqDt.focus
		Exit Function
	End If    

	IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("800600","X",frm1.txtFromReqDt.alt,"X")            '⊙: Display Message(There is no changed data.)
		Exit Function
	End if

	lgQueryFlag = "New"		' 신규조회 및 추가조회 구분 Flag (현재는 신규임)	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																	'☜: Query db data

    FncQuery = True													
	Set gActiveElement = document.ActiveElement 
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																	'☜: Processing is NG
    Err.Clear																		'☜: Clear err status

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X") '☜ 바뀐부분    
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    Call ggoOper.ClearField(Document, "1")											'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")											'⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    Call InitVariables()															'⊙: Initializes local global variables
    Call SetDefaultVal()
    
    FncNew = True				
    Set gActiveElement = document.ActiveElement                                                     '⊙: Processing is OK													'⊙: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False																'☜: Processing is NG
    Err.Clear																		'☜: Clear err status

    Set gActiveElement = document.ActiveElement       
    FncDelete = True																'☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim lRow
	Dim Fg
	Dim GLDT
	
    FncSave = False																	'☜: Processing is NG
    Err.Clear																		'☜: Clear err status    
    On Error Resume Next															'☜: Protect system from crashing    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then			'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    If Not chkField(Document, "1") Then												'⊙: Check required field(Single area)
       Exit Function
    End If
	
	With frm1
		For lRow = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow
				.vspdData.Col = C_CHECK_FG
			If Trim(.vspdData.Text) = "1" Then
				.vspdData.Col = C_CONF_FG
				Fg= .vspdData.text
				.vspdData.Col = C_GL_DT
				If Trim(.vspdData.text)= ""  and  Trim(Fg)="U"Then
					Call DisplayMsgBox("117523","X","X","X") 
				    Exit Function				
				End if	
				GLDT = Trim(.vspdData.text)
				.vspdData.Col = C_AP_DT
				If CompareDateByFormat(.vspdData.text,GLDT,"채무일자","전표일", _
        	               "970023",.txtFromReqDt.UserDefinedFormat ,parent.gComDateType, true) = False Then
					Exit Function
				End If	
			End if
		Next
	End With
	'-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then											'⊙: Check contents area
       Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																	'☜: Save db data

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    Set gActiveElement = document.ActiveElement      	
    
    FncSave = True 
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    Set gActiveElement = document.ActiveElement   
    
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    if frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo															
    
    Set gActiveElement = document.ActiveElement   

    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Set gActiveElement = document.ActiveElement   
    
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Set gActiveElement = document.ActiveElement   
    
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
	
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Set gActiveElement = document.ActiveElement   	
    
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Set gActiveElement = document.ActiveElement   	
    
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = Frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
	    Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
End Function

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False																	'☜: Processing is NG
    Err.Clear																		'☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")				'데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
    
    FncExit = True																	'☜: Processing is OK
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

    DbQuery = False
    Call LayerShowHide(1)
    
    Err.Clear																				'☜: Protect system from crashing
    
    With frm1
		strVal = BIZ_PGM_ID

		strVal = strVal & "?txtMode="      & parent.UID_M0001							'☜:조회표시 
		strVal = strVal & "&txtBpCd="     & Trim(.txtBpCd.value)	 			    '☆: 조회 조건 데이타 
		strVal = strVal & "&txtDeptCd="    & Trim(.txtDeptcd.value)
		strVal = strVal & "&cboConfFg="    & Trim(.cboConfFg.value)
		strVal = strVal & "&txtFromReqDt=" & UNIConvDate(Trim(.txtFromReqDt.Text))
		strVal = strVal & "&txtToReqDt="   & UNIConvDate(Trim(.txtToReqDt.Text))
		strVal = strVal & "&OrgChangeId=" & Trim(.hOrgChangeId.Value)
		strVal = strVal & "&txtMaxRows="   & frm1.vspdData.MaxRows
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgPageNo="     & lgPageNo         		

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

		Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 
        
    End With

    DbQuery = True
End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()		
	Dim arrField
	Dim intFieldCount
	Dim i	

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE											'⊙: Indicates that current mode is Update mode
       
    Call ggoOper.LockField(Document, "Q")											'⊙: This function lock the suitable field
    Call LayerShowHide(0)
    
    Call SetToolbar("11001000000111")    

	Call SetSpreadColor(1, frm1.vspddata.MaxRows)
	
	If UCase(Trim(frm1.cboConfFg.value)) = "C" Then
		frm1.vspdData.ReDraw = False
		ggoSpread.source = frm1.vspdData
		ggoSpread.SpreadLock C_GL_DT, 1, C_GL_DT, frm1.vspdData.MaxRows

		frm1.vspdData.ReDraw = True
	
	End If
	
End Function

Function SetGridFocus()
	With frm1 
		.vspdData.Col = 1
		.vspdData.Row = 1
		.vspdData.Action = 1
	End With 
End Function 

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	
	Dim arrField_S
	Dim intFieldCount
	Dim i,j,k	
	
    DbSave = False																		'⊙: Processing is NG
    Call LayerShowHide(1)
    
    On Error Resume Next																'☜: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 1
			If Trim(.vspdData.Text) = "1" Then
				strVal = strVal & "U" & parent.gColSep										'☜: U=Update
 				.vspdData.Col =  C_AP_NO 
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep						'AP_NO


				.vspdData.Col = C_GL_DT
				strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep		'Journal Date
				
				
				.vspdData.Col =  C_CONF_FG 														'CONF_FG

				If .vspdData.Text = "U" Then
					strVal = strVal & "C" & parent.gRowSep
				Else
					strVal = strVal & "U" & parent.gRowSep
				End If	

				lGrpCnt = lGrpCnt + 1
			End if
		Next
		
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal	

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID1)												'☜: 비지니스 ASP 를 가동 
	End With

    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
    Call LayerShowHide(0)

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	ggoSpread.SSDeleteFlag 1 , frm1.vspdData.MaxRows
	ggoSpread.ClearSpreadData
	
    Call SetSpreadLock()
	frm1.vspdData.ReDraw = True
	lgQueryFlag = "Save"		' 신규조회 및 추가조회 구분 Flag (현재는 추가조회(저장후)임)	
	
	IF frm1.cboConfFg.value = "C" Then
		frm1.cboConfFg.value = "U"
	Else
		frm1.cboConfFg.value = "C"
	End If
	
	Call InitVariables()	
	Call DBQuery()		
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 

End Function




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
End Sub



'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0001111111")

    gMouseClickStatus = "SPC"									'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col							'Ascending Sort
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey				'Descending Sort
			lgSortKey = 1
		End If										
		Exit Sub
	End If		

	
End Sub	

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
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
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
	lgBlnFlgChgValue = true
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : This event is spread sheet data Button Clicked
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData 
	    If Row >= NewRow Then
			Exit Sub
		End If
    End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspdData_KeyPress(index , KeyAscii )
    lgBinFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
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
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
					Call RestoreToolBar()
					Exit Sub
			End if
    	End If
    End If
End Sub

'========================================================================================================
'   Event Name : fpdtFromEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtFromEnterDt.Action = 7
		Call SetFocusToDocument("M")	
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
		Call SetFocusToDocument("M")	
		frm1.fpdtToEnterDt.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtFromEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtToEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		Call MainQuery()
	End If   
End Sub

'=======================================================================================================
'   Event Name : GIDate_Change()
'   Event Desc :  
'=======================================================================================================
Sub GIDate_DblClick(Button)
    If Button = 1 Then
        frm1.GIDate.Action = 7
        Call SetFocusToDocument("M")	
		frm1.GIDate.Focus
    End If
End Sub

Sub GIDate_Change()
	Dim gDate
	Dim IRow

	If UCase(Trim(frm1.cboConfFg.value)) = "C" Then
		Exit Sub
	End If

	gDate = frm1.GIDate.Text

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If

	frm1.vspdData.Col = C_GL_DT

	For IRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row  = IRow
		frm1.vspdData.Text = gDate
	Next
	lgBlnFlgChgValue = True
End Sub
'==========================================================================================
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtFromReqDt.Text, gDateFormat,""), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtToReqDt.Text, gDateFormat,""), "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With
		'----------------------------------------------------------------------------------------

End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD  <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5"NOWRAP>채무일자</TD>
									<TD CLASS="TD6"NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFromReqDt" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT="시작일자" id=fpDateTime1 ></OBJECT>');</SCRIPT>~ 
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToReqDt" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT="종료일자" id=fpDateTime2></OBJECT>');</SCRIPT>										
									</TD>
									<TD CLASS="TD5"NOWRAP>부서</TD>
									<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDeptCd" SIZE=10  MAXLENGTH=10 tag="11XXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
										 <INPUT TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 tag="14X" ALT="부서명">
									</TD>
						
								</TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>승인상태</TD>
									<TD CLASS="TD6"NOWRAP><SELECT NAME="cboConfFg" tag="12" STYLE="WIDTH:82px:" Alt="승인상태"></OPTION></SELECT>
									<TD CLASS="TD5"NOWRAP>거래처</TD>
									<TD CLASS="TD6"NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)">
										 <INPUT TYPE=TEXT ID="txtBpNm" NAME="txtBpNm" SIZE=20 tag="14X" ALT="거래처명">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>전표일자</TD>
									<TD CLASS="TD6"NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=GIDate name=GIDate CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="전표일자"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5"NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6"NOWRAP>&nbsp;</TD>
									
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>				
							<TR>
								<TD HEIGHT="100%" COLSPAN="4"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD class=TDT NOWRAP></TD>
								<TD class=TD6 NOWRAP></TD>
								<TD class=TD5 NOWRAP>채무액(자국)</TD>
								<TD class=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무액(자국)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
							</TR>							
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE01%>></TD>
	</TR>		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hDeptCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBizCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hcboConfFg" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtWorkFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromReqDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hToReqDt" tag="24">
<INPUT		TYPE=hidden	 NAME="hOrgChangeId"	tag="14" TABINDEX="-1">
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

