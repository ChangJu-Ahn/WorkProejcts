<!--
======================================================================================================
*  1. Module Name          : Account
*  2. Function Name        : 
*  3. Program ID           : A4105ma1 
*  4. Program Name         : 일괄출금등록 
*  5. Program Desc         : Multi Allocation AP
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/11/18
*  8. Modified date(Last)  : 2003/10/13
*  9. Modifier (First)     :
* 10. Modifier (Last)      : Jeong Yong Kyun
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

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID  = "A4105MB1.asp"                                      'Biz Logic ASP for Mmulti #1
Const BIZ_PGM_ID2 = "A4105MB3.asp"                                      'Biz Logic ASP for Save, Updat, Delete
Const JUMP_PGM_ID_NOTE_INF = "f5101ma1"									'어음정보등록 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey            = 2                                    '☆☆☆☆: Max key value
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 
Dim lgIsOpenPop        
Dim IsOpenPop 
Dim lgOpenApCondFg										'채무조건 팝업을 열었는지 여부 
Dim lgRefApCondFg										'채무조건 팝업에서 레퍼런스를 정상적으로 했는지 
Dim lgIsQuery

'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
Dim C_Checked_1
Dim C_pay_bp_cd_1
Dim C_bp_nm_1
'Dim C_ap_due_dt_1
Dim C_doc_cur_1
Dim C_ap_amt_1
Dim	C_bal_amt_1     
Dim C_Cls_amt_1
Dim C_Ap_No_1
Dim C_Note_No_1

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim C_Checked_2
Dim C_pay_bp_cd_2
Dim C_bp_nm_2
Dim C_ap_no_2
Dim C_ap_dt_2
Dim C_ap_due_dt_2
Dim C_doc_cur_2
Dim C_ap_amt_2
Dim C_bal_amt_2
Dim C_Cls_amt_2
Dim C_over_due_fg_2
Dim C_ACCT_CD_2
Dim C_ap_desc_2

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgDueDtFg
Dim lsCurrentClickRow
Dim vspdData1ButtonClicked
Dim vspdData2ButtonClicked
Dim lgstartfnc
Dim lgFormLoad

Const C_SHEETMAXROWS = 50

<%
Dim BaseDate
BaseDate = GetSvrDate
%>

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

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False										'Indicates that no value changed
    lgIntFlgMode     = parent.OPMD_CMODE							'Indicates that current mode is Create mode

	lgSortKey         = 1     
	lsCurrentClickRow = ""
	
	vspdData2ButtonClicked = False
	vspdData1ButtonClicked = False
	lgstartfnc = False												'deptCd orgchangeid check하기 위해 
	lgFormLoad = True												'DeptCd orgchangeid Check하기 위해 
	lgIsQuery = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	With frm1
		.txtAllcDt.text     = UniConvDateAToB("<%=BaseDate%>",parent.gServerDateFormat,parent.gDateFormat)
	
		If .txtinputtype.Value = "" Then
			.txtNoteDueDt.text = ""
		Else			
			.txtNoteDueDt.text = UniConvDateAToB("<%=BaseDate%>",parent.gServerDateFormat,parent.gDateFormat)
		End If
				
		.txtDocCur.value	= parent.gCurrency
		.hOrgChangeId.value = parent.gChangeOrgId
		.txtXchRate.text	= 1

		lgDueDtFg	= True
		.txtDueDt.value	    = UniConvDateAToB("<%=BaseDate%>", parent.gServerDateFormat,parent.gDateFormat)

		Call ggoOper.SetReqAttr(.txtAllcDt,   "Q")
		Call ggoOper.SetReqAttr(.txtDeptCd,   "Q")
		Call ggoOper.SetReqAttr(.txtInputType,"Q")
		Call ggoOper.SetReqAttr(.txtAcctCd,   "Q")
	End With	
	lgBlnFlgChgValue = False											'Indicates that no value changed 	
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================= 
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'========================================================================================================= 
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			C_Checked_1		=  1
			C_pay_bp_cd_1   =  2
			C_bp_nm_1       =  3
			C_doc_cur_1     =  4
			C_ap_amt_1      =  5
			C_bal_amt_1     =  6
			C_Cls_amt_1		=  7
			C_Ap_No_1       =  8
			C_Note_No_1     =  9
		Case "B"
			C_Checked_2		=  1
			C_pay_bp_cd_2   =  2
			C_bp_nm_2       =  3
			C_ap_no_2       =  4
			C_ap_dt_2       =  5
			C_ap_due_dt_2   =  6
			C_doc_cur_2     =  7
			C_ap_amt_2      =  8
			C_bal_amt_2     =  9
			C_Cls_amt_2		=  10
			C_over_due_fg_2 =  11
			C_ACCT_CD_2     =  12
			C_ap_desc_2     =  13
	End Select
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub  InitSpreadSheet(ByVal pvSpdNo)
	Call initSpreadPosVariables(pvSpdNo)

	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData
				ggoSpread.Source = frm1.vspdData
				ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

				.ReDraw  = False
				.MaxCols = C_Note_No_1 + 1														' ☜:☜: Add 1 to Maxcols
				.Col     = .MaxCols        : .ColHidden = True
				.MaxRows = 0																		' ☜: Clear spreadsheet data

				Call GetSpreadColumnPos(pvSpdNo)
										'Tag(From "6")  'Integeral  'Decimal
				Call AppendNumberPlace("6"             ,"4"        ,"0")

									'ColumnPosition     Header      Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
				ggoSpread.SSSetCheck   C_Checked_1	 ,""			 ,3 , ,"", True,-1
				ggoSpread.SSSetEdit    C_pay_bp_cd_1 ,"지급처"	 ,15, ,  ,     , 2
				ggoSpread.SSSetEdit    C_bp_nm_1	 ,"지급처명" ,20, ,  ,     , 2
									'ColumnPosition     Header      Width  Align(0:L,1:R,2:C)  Format        Row
'				ggoSpread.SSSetDate    C_ap_due_dt_1 ,"만기일"   ,10,2,parent.gDateFormat ,-1
				ggoSpread.SSSetEdit    C_doc_cur_1   ,"거래통화" ,10, ,  ,     ,2
				ggoSpread.SSSetFloat   C_ap_amt_1	 ,"채무금액" ,20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec													
				ggoSpread.SSSetFloat   C_Cls_amt_1	 ,"반제금액" ,20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec													
				ggoSpread.SSSetFloat   C_bal_amt_1	 ,"채무잔액" ,20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec													
									'ColumnPosition     Header      Width  Align(0:L,1:R,2:C)  ComboEditable   Row
				ggoSpread.SSSetCombo   C_Ap_No_1	 ,"채무번호" ,20,2,True,-1
				ggoSpread.SSSetEdit    C_Note_No_1	 ,"어음번호" ,25, ,    ,  ,2
				
				Call ggoSpread.SSSetColHidden(C_Ap_No_1,C_Ap_No_1,True)
				
				frm1.vspddata.col = C_Checked_1
				frm1.vspddata.UserResizeCol = 2 
				       
				.ReDraw = True
					
				Call SetSpreadLock(pvSpdNo)
			End With
		Case "B"
			With frm1.vspdData2
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
	
				.ReDraw  = False
				.MaxCols = C_ap_desc_2 + 1                                                  ' ☜:☜: Add 1 to Maxcols
				.Col     = .MaxCols        : .ColHidden = True
				.MaxRows = 0

				Call GetSpreadColumnPos(pvSpdNo)
				Call AppendNumberPlace("6"             ,"4"        ,"0")
			                          'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
				ggoSpread.SSSetCheck   C_Checked_2		 ,""			 ,3, , "", True, -1
				ggoSpread.SSSetEdit    C_pay_bp_cd_2     ,"지급처"   ,10, ,     ,     ,2
				ggoSpread.SSSetEdit    C_bp_nm_2         ,"지급처명" ,10, ,     ,     ,2
				ggoSpread.SSSetEdit    C_ap_no_2         ,"채무번호" ,15, ,     ,     ,2
				ggoSpread.SSSetDate    C_ap_dt_2		 ,"채무일자" ,10,2,parent.gDateFormat,-1
				ggoSpread.SSSetDate    C_ap_due_dt_2     ,"만기일"	 ,10,2,parent.gDateFormat,-1
				ggoSpread.SSSetEdit    C_doc_cur_2       ,"거래통화" ,10, ,     ,     ,2
				ggoSpread.SSSetFloat   C_ap_amt_2        ,"채무금액" ,15,parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat   C_Cls_amt_2		 ,"반제금액" ,15,parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat   C_bal_amt_2       ,"채무잔액" ,15,parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit    C_over_due_fg_2   ,"만기"     ,10, ,     ,     ,2
				ggoSpread.SSSetEdit    C_ACCT_CD_2       ,"계정코드" ,18, ,     ,     ,2
				ggoSpread.SSSetEdit    C_ap_desc_2       ,"비고"     ,18, ,     ,     ,2

				Call ggoSpread.SSSetColHidden(C_pay_bp_cd_2,C_pay_bp_cd_2,True)
				Call ggoSpread.SSSetColHidden(C_bp_nm_2,C_bp_nm_2,True)

			   .ReDraw = True	

			   Call SetSpreadLock(pvSpdNo)
			End With
	End Select			
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadLock(Byval pSpreadSheetNo )
    Select Case pSpreadSheetNo
        Case "A"
			With frm1
			   .vspdData.ReDraw = False
			   ggoSpread.Source = .vspdData 
			   ggoSpread.spreadUnlock	C_Checked_1	 , -1
			   ggoSpread.SpreadLock		C_pay_bp_cd_1, -1
			   ggoSpread.spreadUnlock	C_ap_No_1	 , -1
			   ggoSpread.SpreadLock		C_note_no_1, -1			   
			   .vspdData.ReDraw = True
			End With
		Case "B"
			With frm1
				.vspdData2.ReDraw = False
				ggoSpread.Source = .vspdData2 
				ggoSpread.spreadUnlock	C_Checked_2	 , -1
			    ggoSpread.SpreadLock	C_pay_bp_cd_2, -1
			    .vspdData2.ReDraw = True
			End With
	End Select   
End Sub

'========================================================================================================= 
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================================= 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_checked_1		= iCurColumnPos(1)
            C_pay_bp_cd_1	= iCurColumnPos(2)
            C_bp_nm_1		= iCurColumnPos(3)
            C_doc_cur_1		= iCurColumnPos(4)
            C_ap_amt_1		= iCurColumnPos(5)
            C_bal_amt_1		= iCurColumnPos(6)
            C_cls_amt_1		= iCurColumnPos(7)
            C_ap_No_1		= iCurColumnPos(8)
            C_note_No_1		= iCurColumnPos(9)
	   Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
					
			C_checked_2		= iCurColumnPos(1)
			C_pay_bp_cd_2	= iCurColumnPos(2)
			C_bp_nm_2		= iCurColumnPos(3)
			C_ap_no_2		= iCurColumnPos(4)
			C_ap_dt_2		= iCurColumnPos(5)
			C_ap_due_dt_2   = iCurColumnPos(6)
			C_doc_cur_2		= iCurColumnPos(7)
			C_ap_amt_2		= iCurColumnPos(8)
			C_bal_amt_2		= iCurColumnPos(9)
			C_Cls_amt_2		= iCurColumnPos(10)
			C_over_due_fg_2	= iCurColumnPos(11)
			C_acct_cd_2		= iCurColumnPos(12)
			C_ap_desc_2		= iCurColumnPos(13)
    End Select    
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'=======================================================================================================
Sub MakeKeyStream(Col, Row)
	ReDim lgKeyStream(3)

	With frm1.vspdData
		.Row=Row
		.Col=C_pay_bp_cd_1	:		lgKeyStream(0) = .text       'You Must append one character(gColSep)
		.Col=C_doc_cur_1	:		lgKeyStream(1) = .text	
		lgKeyStream(2) = Getcombolist(C_Ap_No_1, Row)	
	End With       
End Sub        

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox(byval strApNo)
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo strApNo, C_Ap_No_1, Frm1.vspdData.ActiveRow
End Sub

'========================================== 2.4.2 OpenRefOpenPaymCon()  =============================================
'	Name : OpenRefOpenPaymCon
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefOpenPaymCon()
	Dim arrRet
	Dim arrParam(11,1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4115RA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4115RA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	With frm1
		arrParam(0,0) = Trim(.txtDueDt.value)
		arrParam(0,1) = Trim(.txtDocCur.value)	
		arrParam(6,0) = Trim(lgDueDtFg)	
	End With	

	' 권한관리 추가 
	arrParam(8,0) = lgAuthBizAreaCd
	arrParam(9,0) = lgInternalCd
	arrParam(10,0) = lgSubInternalCd
	arrParam(11,0) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=420px; dialogHeight=220px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0,0) = "" Then
		Exit Function
	Else
		lgOpenApCondFg = True
		Call SetRefPaymCon(arrRet)
		Call DBQuery()
		Call SetvspdDataCombo()
	End If
End Function

'==========================================  SetRefPaymCon  ============================================
'	Name : SetRefPaymCon
'	Description : SetRefPaymCon
'=======================================================================================================
Sub SetRefPaymCon(Byval arrRet)
	With frm1
		lgDueDtFg				= arrRet(6,0)
		.txtDueDt.value			= arrRet(0,0)
		.txtDocCur.value		= arrRet(0,1) 
		.txtPayBpCd.value		= arrRet(1,0)
		.txtPaymType.value		= arrRet(2,0)
		.txtBizAreaCd.value		= arrRet(3,0)
		.txtBizAreaCd1.value	= arrRet(4,0)
		
		Call txtDocCur_OnChange	()
	End With

	lgBlnFlgChgValue = True
	lgIntFlgMode     = parent.OPMD_CMODE 
End Sub

'=======================================================================================================
' Name : OpenPopupGL()
' Description : 회계전표POP-UP
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1) 
	Dim iCalledAspName
	Dim IntRetCD
	
	iCalledAspName = AskPRAspName("A5120RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
 
	arrParam(0) = Trim(frm1.txtGlNo.value)												'회계전표번호 
	arrParam(1) = ""																	'Reference번호 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
End Function

'========================================== 2.4.2 OpenPopuptempGL()  =====================================
'	Name : OpenPopuptempGL()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD
	
	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtTempGlNo.value)														'회계전표번호 
	arrParam(1) = ""																				'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName,  Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'========================================================================================================
'	Name : CookiePage()
'	Description : 
'========================================================================================================
Function CookiePage(ByVal Kubun)
	Dim strTemp

	Select Case Kubun		
		Case "FORM_LOAD"
			strTemp = ReadCookie("NOTE_NO")
			Call WriteCookie("NOTE_NO", "")
			
			If strTemp = "" then Exit Function

			frm1.txtNoteNoQry.value = strTemp
	
			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("NOTE_NO", "")
				Exit Function 
			End If
				
			Call MainQuery()
		Case JUMP_PGM_ID_NOTE_INF	'어음정보등록 
			With frm1.vspddata 
				If .activeRow = 0 Then
					strTemp = ""
				Else					
					.row = .activeRow
					.col = C_Note_No_1
					strTemp = Trim(.text) 
				End If					
			End With				

			Call WriteCookie("NOTE_NO", strTemp)
		Case Else
			Exit Function
	End Select
End Function	

'========================================================================================================
'	Desc : 화면이동 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")									'데이타가 변경되었습니다. 계속하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call CookiePage(strPgmId)
    Call PgmJump(strPgmId)
End Function

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iArrParam(8)
	Dim strCd	
	Dim iCalledAspName
	Dim IntRetCD	
	
	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0
			If frm1.txtBatchAllcNo.className = "protected" Then Exit Function
		Case 1
			If frm1.txtDeptCd.className = "protected" Then Exit Function
			
			arrParam(0)  = "부서팝업"																' 팝업 명칭 
			arrParam(1)  = "B_ACCT_DEPT A"																' TABLE 명칭 
			arrParam(2)  = Trim(frm1.txtDeptCd.Value)													' Code Condition
			arrParam(3)  = ""																			' Name Cindition
			arrParam(4)  = "A.ORG_CHANGE_ID = " & FilterVar(gChangeOrgId, "''", "S")
			
			arrParam(5)  = "부서"			
	
			arrField(0)  = "A.Dept_CD"																	' Field명(0)
			arrField(1)  = "A.Dept_NM"																	' Field명(1)
			    
			arrHeader(0) = "부서"																	' Header명(0)
			arrHeader(1) = "부서명"																	' Header명(1)   			    		
		Case 2
			If frm1.txtAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "계정코드팝업"															' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"									' TABLE 명칭 
			arrParam(2) = ""																			' Code Condition
			arrParam(3) = ""																			' Name Cindition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
							" and C.trans_type = " & FilterVar("ap001", "''", "S") & "  and C.jnl_cd = " & FilterVar(frm1.txtInputType.Value, "''", "S")	' Where Condition
			arrParam(5) = "계정코드"																' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"																	' Field명(0)
			arrField(1) = "A.Acct_NM"																	' Field명(1)
    		arrField(2) = "B.GP_CD"																		' Field명(2)
			arrField(3) = "B.GP_NM"																		' Field명(3)
					
			arrHeader(0) = "계정코드"																' Header명(0)
			arrHeader(1) = "계정코드명"																' Header명(1)
			arrHeader(2) = "그룹코드"																' Header명(2)
			arrHeader(3) = "그룹명"																	' Header명(3)		
		Case 3			
			If frm1.txtInputType.className = "protected" Then Exit Function    
			
			If frm1.txtDocCur.value <> "" Then
				If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
					arrParam(0) = "지급유형"														' 팝업 명칭						
					arrParam(1) = "B_MINOR,B_CONFIGURATION "
					arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
					arrParam(3) = ""																	' Name Cindition
					arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
								& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " " _ 
								& " AND B_MINOR.MINOR_CD IN (" & FilterVar("CS", "''", "S") & " ," & FilterVar("DP", "''", "S") & " ," & FilterVar("NP", "''", "S") & " ," & FilterVar("CP", "''", "S") & " ) "						' Where Condition								
					arrParam(5) = "지급유형"														' TextBox 명칭 
		
					arrField(0) = "B_MINOR.MINOR_CD"													' Field명(0)
					arrField(1) = "B_MINOR.MINOR_NM"													' Field명(1)
	    
					arrHeader(0) = "지급유형"														' Header명(0)
					arrHeader(1) = "지급유형명"														' Header명(1)		
				Else
					arrParam(0) = "지급유형"														' 팝업 명칭						
					arrParam(1) = "B_MINOR,B_CONFIGURATION "
					arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
					arrParam(3) = ""																	' Name Cindition
					arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD and B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
								& " and B_CONFIGURATION.SEQ_NO = 2 and B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " " _
								& " AND B_MINOR.MINOR_CD IN (" & FilterVar("CS", "''", "S") & " ," & FilterVar("DP", "''", "S") & " ," & FilterVar("NP", "''", "S") & " ," & FilterVar("CP", "''", "S") & " ) " _                     
								& " And B_minor.minor_cd Not in ( Select  minor_cd  from b_configuration " _ 
								& " where major_cd=" & FilterVar("a1006", "''", "S") & "  and seq_no=4 and reference=" & FilterVar("NO", "''", "S") & " ) "			' Where Condition								
					arrParam(5) = "지급유형"														' TextBox 명칭 
		
					arrField(0) = "B_MINOR.MINOR_CD"													' Field명(0)
					arrField(1) = "B_MINOR.MINOR_NM"													' Field명(1)
	    
					arrHeader(0) = "지급유형"														' Header명(0)
					arrHeader(1) = "지급유형명"														' Header명(1)		
				End If
			Else
				arrParam(0) = "지급유형"															' 팝업 명칭						
				arrParam(1) = "B_MINOR,B_CONFIGURATION "
				arrParam(2) = Trim(frm1.txtInputType.value)												' Code Condition
				arrParam(3) = ""																		' Name Cindition
				arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
							& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "		' Where Condition								
				arrParam(5) = "지급유형"															' TextBox 명칭 
		
				arrField(0) = "B_MINOR.MINOR_CD"														' Field명(0)
				arrField(1) = "B_MINOR.MINOR_NM"														' Field명(1)
	    
				arrHeader(0) = "지급유형"															' Header명(0)
				arrHeader(1) = "지급유형명"															' Header명(1)									
			End If							
		Case 4					
			If frm1.txtBankCd.className = "protected" Then Exit Function
			
			Dim strWhere 
			Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
			
			strWhere = "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND B_CONFIGURATION.SEQ_NO = 2 AND  B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & "  "
			strWhere = strWhere & "AND  MINOR_CD= " & FilterVar(UCase(frm1.txtInputType.value), "''", "S") & ""


			If CommonQueryRs( "MINOR_CD" , "B_CONFIGURATION" , strWhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
			
				Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
					Case "NP" 
					
						arrParam(0)  = "은행팝업"
						arrParam(1)  = "f_note_no, B_BANK"				
						arrParam(2)  = Trim(frm1.txtBankCd.Value)
						arrParam(3)  = ""
						arrParam(4)  = "f_note_no.BANK_CD = B_BANK.BANK_CD"
						arrParam(5)  = "은행"			
	
						arrField(0)  = "f_note_no.BANK_CD "	
						arrField(1)  = "B_BANK.BANK_NM"	
    
						arrHeader(0) = "은행"		
						arrHeader(1) = "은행명"	
					Case "DP"	
					
						arrParam(0)  = "은행팝업"
						arrParam(1)  = "F_DPST, B_BANK"				
						arrParam(2)  = Trim(frm1.txtBankCd.Value)
						arrParam(3)  = ""
						arrParam(4)  = "F_DPST.BANK_CD = B_BANK.BANK_CD"
						arrParam(5)  = "은행"			
	
						arrField(0)  = "F_DPST.BANK_CD"	
						arrField(1)  = "B_BANK.BANK_NM"	
    
						arrHeader(0) = "은행"		
						arrHeader(1) = "은행명"	
				END Select 
			End If				
		Case 5
			If frm1.txtBankCd.className = "protected" Then Exit Function
			
			arrParam(0) = "계좌번호팝업"
			arrParam(1) = "F_DPST, B_BANK"				
			arrParam(2) = Trim(frm1.txtBankAcct.Value)
			arrParam(3) = ""
			
			If Trim(frm1.txtBankCd.Value) = "" Then
				strCd = "F_DPST.BANK_CD = B_BANK.BANK_CD "
			Else
				strCd = "F_DPST.BANK_CD = B_BANK.BANK_CD AND  F_DPST.BANK_CD =  " & FilterVar(frm1.txtBankCd.Value, "''", "S") & " "	
			End If		

			arrParam(4) = strCd
			arrParam(5) = "계좌번호"			
			
		    arrField(0) = "F_DPST.BANK_ACCT_NO"	
		    arrField(1) = "F_DPST.BANK_CD"	
		    arrField(2) = "B_BANK.BANK_NM"	
		    
		    arrHeader(0) = "계좌번호"		
		    arrHeader(1) = "은행"	
		    arrHeader(2) = "은행명"						
		Case 6
			If frm1.txtCardCoCd.className = "protected" Then Exit Function
			
			arrParam(0) = "카드사팝업"
			arrParam(1) = "B_card_co "				
			arrParam(2) = Trim(frm1.txtCardCoCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "카드사"			
			
		    arrField(0) = "card_co_cd"	
		    arrField(1) = "card_co_nm"	
		    
		    arrHeader(0) = "카드사"		
		    arrHeader(1) = "카드사명"	
	End Select				

	IsOpenPop = True

	iCalledAspName = AskPRAspName("A4115RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4115RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	' 권한관리 추가 
	iArrParam(5) = lgAuthBizAreaCd
	iArrParam(6) = lgInternalCd
	iArrParam(7) = lgSubInternalCd
	iArrParam(8) = lgAuthUsrID

	If iwhere = 0 Then
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, iArrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;") 
	Else
		arrRet = window.showModalDialog("../../comasp/adoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	End If
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function

'==========================================  OpenDept()  ===============================================
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'=======================================================================================================
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtAllcDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If	
End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtBatchAllcNo.value = arrRet(0)
				lgIntFlgMode = parent.OPMD_UMODE
				Call PayConRefView()
				.txtBatchAllcNo.focus
			Case 1	
				.txtDeptCd.value = arrRet(0)		
				.txtDeptNm.value = arrRet(1)
				.txtAllcDt.text = arrRet(3)
				Call txtDeptCd_OnChange()  
			Case 2
				.txtAcctCd.value = arrRet(0)		
				.txtAcctnm.value = arrRet(1)
			Case 3
				.txtInputType.value = arrRet(0)		 	
				.txtInputTypeNm.value = arrRet(1)		 	

				Call txtInputType_OnChange()				
			Case 4
				.txtBankCd.value = arrRet(0)		
				.txtBankNm.value = arrRet(1)			    		
			Case 5
				.txtBankAcct.value = arrRet(0)		
				.txtBankCd.value = arrRet(1)		
				.txtBankNm.value = arrRet(2)
			Case 6
				.txtCardCoCd.value = arrRet(0)		
				.txtCardCoNm.value = arrRet(1)			    							
		End Select				
	End With
	
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If	
End Function

'=======================================================================================================
'   Event Name : txtInputType_OnChange()
'   Event Desc :  
'=======================================================================================================
Sub txtInputType_OnChange()
	Dim IntRetCD

    lgBlnFlgChgValue = True
	
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
	If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD = " & FilterVar(frm1.txtInputType.value, "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
			Case "CS" 
				frm1.txtBankCd.value   = "" : frm1.txtBankNm.value=""
				frm1.txtBankAcct.value = ""
				frm1.txtAcctCd.value   = "" : frm1.txtAcctNm.value = "" 			
				frm1.txtCardCoCd.value = "" : frm1.txtCardCoNm.value = ""	
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
'				Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "D")	
				Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "Q")				
				frm1.txtNoteDueDt.text = ""				
			Case "DP" 			' 예적금 
				frm1.txtBankCd.value   = "" : frm1.txtBankNm.value=""
				frm1.txtBankAcct.value   = ""
				frm1.txtAcctCd.value   = "" : frm1.txtAcctNm.value = "" 
				frm1.txtCardCoCd.value = "" : frm1.txtCardCoNm.value = ""																
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
'				Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "D")
				Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "Q")
				Call ggoOper.SetReqAttr(frm1.txtCardCoCd,   "Q")
				frm1.txtNoteDueDt.text = ""				
			Case "NO" 		'어음 
				If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
					If UCase(Trim(frm1.txtInputType.value)) = "NP" Then     '지급어음 
						If lgIntFlgMode = parent.OPMD_UMODE Then			
							frm1.txtBankCd.value   = frm1.hBankCd.value
							frm1.txtBankNm.value   = frm1.hBankNm.value
							Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")					
						Else
							frm1.txtBankCd.value   = "" : frm1.txtBankNm.value=""
							frm1.txtBankAcct.value   = ""				
							frm1.txtCardCoCd.value = "" : frm1.txtCardCoNm.value = ""
							frm1.txtAcctCd.value   = "" : frm1.txtAcctNm.value = "" 								
							Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
						End If					
						Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
						Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "N")
						Call ggoOper.SetReqAttr(frm1.txtCardCoCd,   "Q")
					ElseIf UCase(Trim(frm1.txtInputType.value)) = "CP" Then
						If lgIntFlgMode = parent.OPMD_UMODE Then	'지불구매카드		
							frm1.txtCardCoCd.value = frm1.hCardCoCd.value
							frm1.txtBankNm.value   = frm1.hCardCoNm.value
							Call ggoOper.SetReqAttr(frm1.txtCardCoCd,   "Q")					
						Else
							frm1.txtBankCd.value   = "" : frm1.txtBankNm.value=""
							frm1.txtBankAcct.value   = ""				
							frm1.txtCardCoCd.value = "" : frm1.txtCardCoNm.value = ""
							frm1.txtAcctCd.value   = "" : frm1.txtAcctNm.value = "" 								
							Call ggoOper.SetReqAttr(frm1.txtCardCoCd,   "N")
						End If
						Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")				
						Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
						Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "N")					
					End If
				Else
					IntRetCD = DisplayMsgBox("111524","X","X","X")  
					frm1.txtInputType.value = ""
					frm1.txtInputTypeNm.value = ""					
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")				
					Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
					Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "Q")								
					Exit Sub
				End If
			Case Else
				IntRetCD = DisplayMsgBox("141140","X","X","X")
				
				frm1.txtInputType.value = ""
				frm1.txtInputTypeNm.value = ""
				Exit Sub
		End Select
	End If
End Sub

'=======================================================================================================
'   Event Name : txtInputType_OnChange2()
'   Event Desc :  
'=======================================================================================================
Sub txtInputType_OnChange2()
	Dim IntRetCD

    lgBlnFlgChgValue = True
	' SetReqAttr(Object, Option) ; N : Required, Q : Protect, D : Default
	
	If CommonQueryRs("REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD = " & FilterVar(frm1.txtInputType.value, "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
			Case "CS" 
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
'				Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "D")	
				Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "Q")	
				frm1.txtNoteDueDt.text = ""
			Case "DP" 		' 예적금 
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
'				Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "D")	
				Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "Q")				
				frm1.txtNoteDueDt.text = ""
			Case "NO" 	'어음 
			Case "NO" 		'어음 
				If Trim(frm1.txtInputType.value) = "NP" Then     '지급어음 
					If lgIntFlgMode = parent.OPMD_UMODE Then			
						frm1.txtBankCd.value   = frm1.hBankCd.value
						frm1.txtBankNm.value   = frm1.hBankNm.value
						Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")					
					Else
						frm1.txtBankCd.value   = "" : frm1.txtBankNm.value=""
						frm1.txtBankAcct.value   = ""				
						frm1.txtAcctCd.value   = "" : frm1.txtAcctNm.value = "" 								
						Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
					End If					
					Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
					Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "N")
				ElseIf Trim(frm1.txtInputType.value) = "CP" Then
					If lgIntFlgMode = parent.OPMD_UMODE Then	'지불구매카드		
						frm1.txtCardCoCd.value = frm1.hCardCoCd.value
						frm1.txtBankNm.value   = frm1.hCardCoNm.value
						Call ggoOper.SetReqAttr(frm1.txtCardCoCd,   "Q")					
					Else
						frm1.txtCardCoCd.value = "" : frm1.txtCardCoNm.value = ""
						frm1.txtAcctCd.value   = "" : frm1.txtAcctNm.value = "" 								
						Call ggoOper.SetReqAttr(frm1.txtCardCoCd,   "N")
					End If					
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")				
					Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
					Call ggoOper.SetReqAttr(frm1.txtNoteDueDt,  "N")					
				End If								
			Case Else
				IntRetCD = DisplayMsgBox("141140","X","X","X")
				
				frm1.txtInputType.value = ""
				frm1.txtInputTypeNm.value = ""
				Exit Sub
		End Select
	End If
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
   	Call LoadInfTB19029()																	'⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")													'⊙: Lock  Suitable  Field

   	Call SetDefaultVal()
	Call InitVariables() 																	'⊙: Initializes local global variables
	Call InitSpreadSheet("A")																'Setup the Spread sheet
	Call InitSpreadSheet("B")	
	Call SetToolbar("1110000000000011")														'⊙: 버튼 툴바 제어 
	
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
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
	
	FncQuery = False																		'☜: Processing is NG
	lgstartfnc = True 
		
    On Error Resume Next																	'☜: Protect system from crashing    
    Err.Clear		
    
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")						'☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    If Not chkField(Document, "1") Then														'This function check indispensable field
		Exit Function
    End If

    Call ggoOper.ClearField(Document, "2")													'☜: Clear Contents  Field
	Call ggoOper.LockField(Document, "Q")	
    Call InitVariables()																	'⊙: Initializes local global variables

	lgIsQuery = True
	
	lgIntFlgMode = parent.OPMD_UMODE

    If DBQuery = False Then															'☜: Query db data
		lgIsQuery = False		
		Exit Function
    End If

	lgIntFlgMode = parent.OPMD_UMODE

    FncQuery = True																			'☜: Processing is OK
    lgstartfnc = False	

	Set gActiveElement = document.ActiveElement    
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    Dim var
    
    FncNew = False                                                          
    lgstartfnc = True  
    
    On Error Resume Next																	'☜: Protect system from crashing    
    Err.Clear		
	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspddata
    var = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var= True Then
		IntRetCD = DisplayMsgBox("900015",parent.VB_YES_NO,"X","X")   
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
    
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")													'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")													'Clear Condition Field
    Call ggoOper.LockField(Document, "N")													'Lock  Suitable  Field
'	Call ggoOper.SetReqAttr(frm1.txtNotedueDt, "N")    
    Call SetDefaultVal()
    Call InitVariables()
	Call PayConRefView()
	Call SetToolbar("1110000000000011")	
   	ggoSpread.Source = frm1.vspdData	:	Call ggoSpread.ClearSpreadData()

	lgBlnFlgChgValue = False																'Indicates that no value changed 
    
    frm1.txtBatchAllcNo.focus

    FncNew = True
    lgstartfnc = False
    lgFormLoad = True	
    
    Set gActiveElement = document.ActiveElement    
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
	Dim IntRetCD
																							'☜: Processing is OK
    FncDelete = False                                                      
   
    On Error Resume Next																	'☜: Protect system from crashing    
    Err.Clear			
        
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")							'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then																'☜: Delete db data
		Exit Function																		'☜:
    End If					
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    FncDelete = True             
    Set gActiveElement = document.ActiveElement                                                                                               
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    On Error Resume Next
    
    FncSave = False																			'☜: Processing is NG
    Err.Clear																				'☜: Clear err status
  
    If Not chkField(Document, "2") Then														'Check contents area
       Exit Function
    End if 
      
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = False or ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
          IntRetCD = DisplayMsgBox("900001","X","X","X")									'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
	IF UniCDBL(frm1.txtPaymAmt.text)=0 then
        IntRetCD = DisplayMsgBox("111516","X","X","X")										'⊙: Display Message(There is no changed data.)
        Exit Function
	End if

    If chkInputType= False Then
		Exit Function
    End If      

    If DbSave = False Then																	'☜: Query db data
		Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True																			'☜: Processing is OK
End Function

'=======================================================================================================
' Function Name : chkInputType
' Function Desc : 
'========================================================================================================
Function chkInputType()
	Dim intI
	Dim IntRetCD
	
	chkInputType = True

	If CommonQueryRs("REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD = " & FilterVar(frm1.txtInputType.value, "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
			Case "NO" 	
				If UCase(Trim(frm1.txtDocCur.value)) <> UCase(parent.gCurrency) Then		
					IntRetCD = DisplayMsgBox("111524","X","X","X")
					frm1.txtInputType.value = ""
					frm1.txtInputTypeNm.value = ""					
					frm1.txtAcctCd.value = ""
					frm1.txtAcctNm.value = ""										
					frm1.txtInputType.focus
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")				
					Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")

					chkInputType = False
				End If					
			Case Else
		End Select
	End If	
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next																	'☜: If process fails
    Err.Clear																				'☜: Clear error status

    FncPrint = False																		'☜: Processing is NG

	Call Parent.FncPrint()																	'☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True																		'☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function  FncExcel() 
    On Error Resume Next                                                                 '☜: If process fails
    Err.Clear                                                                            '☜: Clear error status

    FncExcel = False                                                                     '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then
       FncExcel = True                                                                   '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function  FncFind() 
    On Error Resume Next                                                                 '☜: If process fails
    Err.Clear                                                                            '☜: Clear error status

    FncFind = False                                                                      '☜: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then
       FncFind = True                                                                    '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                        
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

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
	If Err.number = 0 Then
		FncExit = True
	End If

	Set gActiveElement = document.ActiveElement											'☜: Processing is OK
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

	On Error Resume Next
    Err.Clear																			'☜: Protect system from crashing

    DbQuery = False
    Call DisableToolBar(TBC_QUERY)														'☜: Disable Query Button Of ToolBar
	Call LayerShowHide(1)																'☜: Protect system from crashing

    With frm1
		.txtTempAPno.value = ""

		strVal = BIZ_PGM_ID & "?txtFlgMode= "	& lgIntFlgMode
		strVal = strVal & "&txtDueDt="			& .txtDueDt.value
		strVal = strVal & "&txtDocCur="		& Trim(.txtDocCur.value)
		strVal = strVal & "&txtDueDtFg="		& lgDueDtFg
		strVal = strVal & "&txtPayBpCd="		& Trim(.txtPayBpCd.value)
		strVal = strVal & "&txtBatchAllcNo="	& Trim(.txtBatchAllcNo.value)
		strVal = strVal & "&txtAllcDt="		& Trim(.txtAllcDt.text)
		strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		strVal = strVal & "&txtPaymType="		& Trim(.txtPaymType.value)
		strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
		strVal = strVal & "&txtBizAreaCd1="	& Trim(.txtBizAreaCd1.value)

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
		
		Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동 
    End With

    If Err.number = 0 Then
		DbQuery = True
	End If

    Set gActiveElement = document.ActiveElement  
End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim GetApList

    On Error Resume Next
    DbSave = False																		'☜: Processing is NG
    Err.Clear																			'☜: Clear err status

    Call DisableToolBar(TBC_SAVE)														'☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)																'☜: Show Processing Message
		
    ggoSpread.Source = frm1.vspdData

	lGrpCnt = 1
    strVal = ""
	With Frm1
		Redim iRes_Note(1,.vspdData.MaxRows)
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C_Checked_1
			 If .vspdData.text="1" Then
			 	.vspdData.Col = C_Ap_No_1
			 	GetApList = frm1.vspddata.TypeComboBoxList
			 	strVal = strVal & Replace(GetApList,vbTab,parent.gColSep) 
			 	lGrpCnt = lGrpCnt + 1
			 End If
		Next 
		
		If Trim(strval) = "" then 
			strval = ""
		Else			
			strVal = strVal & parent.gRowSep
		End If			

		.txtFlgMode.value = lgIntFlgMode

		.txtMaxRows.value = lGrpCnt-1													'Spread Sheet의 변경된 최대갯수 
		.txtSpread.value  = strVal

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID2)

    DbSave = True																		'☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    
    Err.Clear																			'☜: Clear err status
    DbDelete = False																	'☜: Processing is NG

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID2 & "?txtFlgMode=" & parent.UID_M0003
    strVal = strVal & "&txtBatchAllcNo=" & Trim(frm1.txtBatchAllcNo.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtInputType=" & Trim(frm1.txtInputType.value)					'☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
    
	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 
                                                      
    If Err.number = 0 Then
		DbDelete = True
	End If

	Set gActiveElement = document.ActiveElement											'☜: Processing is OK
End Function

'========================================================================================================
' Name : DbQueryOk1
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk1()
	Dim Row, arrVal
	Dim ii
    Dim TempAllApNo

	With frm1	
		If Trim(.txtTempAPno.value) <> "" Then
			arrVal=Split(.txtTempAPno.value, Chr(12)) 
			For Row=1 To .vspdData.maxRows	
				TempAllApNo = Replace(arrVal(Row-1),Chr(11), vbTab)
				.vspdData.row = Row
				ggoSpread.Source = .vspdData
				ggoSpread.SetCombo TempAllapNo, C_Ap_No_1, .vspdData.row
			Next
		End If

		If .vspdData.MaxRows > 0 Then								
			lsCurrentClickRow = ""
			
			Call MakeKeyStream(1, 1)
			Call DbQueryDetail()			
			.vspdData.focus
		Else
			Call ggoOper.LockField(Document, "Q") 			
		End If

		If Trim(.txtBatchAllcNo.value) <> "" Then
			Call txtDeptCd_OnChange()		
		End If
	End With

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetToolbar("1111100000001111")
	Else
		Call SetToolbar("1110100000001111")
	End If	

	lgBlnFlgChgValue = False

	Call PayConRefView()
	Call ggoOper.LockField(Document, "Q")	
End Sub
		
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk(txtBatchAllcNo)
    lgBlnFlgChgValue = false

	Call ggoOper.ClearField(Document, "2")													'Clear Contents  Field
    Call ggoOper.LockField(Document, "Q") 
    Call InitVariables	
    
    lgIntFlgMode = parent.OPMD_UMODE
    Call PayConRefView()																	'Initializes local global variables
    
    With frm1
		ggoSpread.Source = .vspdData	:	Call ggoSpread.ClearSpreadData()
		ggoSpread.Source = .vspdData2	:	Call ggoSpread.ClearSpreadData()    
		.txtBatchAllcNo.focus
    End With
    
    lgIsQuery = True

    Call DBQuery() 

    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	lgBlnFlgChgValue = False
	
    Call ggoOper.ClearField(Document, "1")													'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")													'Clear Condition Field
    Call ggoOper.LockField(Document, "N")													'Lock  Suitable  Field
    Call SetDefaultVal    
	Call InitVariables    
	Call PayConRefView()																	'Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData	:	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData2	:	Call ggoSpread.ClearSpreadData()    
   
    Call SetToolbar("1110000000000011")	
    frm1.txtBatchAllcNo.Value = ""
    frm1.txtBatchAllcNo.focus
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'=======================================================================================================
'   Event Name : txtAllcDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAllcDt.Action = 7                        
    End If
End Sub

'=======================================================================================================
'   Event Name : txtAllcDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_Change()
    Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
    
    lgBlnFlgChgValue = True
    frm1.txtXchRate.Text = 0

	If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
				If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtAllcDt.Text <> "") Then
					strSelect	=			 " Distinct org_change_id "    		
					strFrom		=			 " b_acct_dept(NOLOCK) "		
					strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
					strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
					strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
					strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtAllcDt.Text, gDateFormat,""), "''", "S") & "))"			
	
					IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 			
					If IntRetCD = False  Or Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
						.txtDeptCd.value = ""
						.txtDeptNm.value = ""
						.hOrgChangeId.value = ""
					End If
				End If
			End With
		End If
	End If
End Sub
	
'=======================================================================================================
'   Event Name : txtAllcDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtNoteDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtNoteDueDt.Action = 7                        
    End If
End Sub

'=======================================================================================================
'   Event Name : txtAllcDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtNoteDueDt_Change()
	lgBlnFlgChgValue = True
End Sub	




'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
function Getcombolist(Col, Row)
	frm1.vspddata.col=C_Ap_No_1
	frm1.vspddata.row=Row
	Getcombolist=Trim(frm1.vspddata.TypeComboBoxList)
End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
	Call SetPopupMenuItemInf("0000111111")
    
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.Maxrows = 0 Then Exit Sub

	If Row <= 0 Then
		Exit Sub
	Else
		If lsCurrentClickRow <> Row  Then
			lsCurrentClickRow = Row
  		End If
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
   If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
 			ggoSpread.Source = .vspddata2
			ggoSpread.ClearSpreadData
        End With

		Call MakeKeyStream(Col, NewRow)
		Call DbQueryDetail()
    End If    
End Sub

'==========================================================================================
'   Event Name : DbQueryDetail
'   Event Desc : 
'==========================================================================================
Function DbQueryDetail()
	Dim strVal	
	Dim lngRows

	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strCond,strCond2
	Dim strApno 

	Dim ii
	Dim arrVal
	Dim arrTemp
	Dim Indx1

	DbQueryDetail = False

	With frm1
	    .vspdData2.redraw = False

		Call LayerShowHide(1)

		arrVal=Split(lgKeyStream(2), vbTab) 
	
		If UBound(arrVal) > 1 Then	
			For ii = 1 To UBound(arrVal) - 1
				If ii <> UBound(arrVal) - 1 Then
					strApno= strApno & FilterVar(arrVal(ii),null,"S") & ","
				Else
					strApno= strApno & FilterVar(arrVal(ii),null,"S") 
				End If
			Next
		Else
			strApno = "''"
		End If

		If lgDueDtFg Then
			strCond = " and a.ap_due_dt = " & FilterVar(.txtDueDt.value , "''", "S")
		Else
			strCond = " and a.ap_due_dt <= " & FilterVar(.txtDueDt.value , "''", "S")
		End If
	
		strCond = strCond & " and a.doc_cur = " & FilterVar(lgKeyStream(1), "''", "S")
		strCond = strCond & " and a.pay_bp_cd = " & FilterVar(lgKeyStream(0), "''", "S")

		If Trim(.txtPaymType.value) <> "" Then 
			strCond = strCond & " and a.paym_type = " & FilterVar(.txtPaymType.value, "''", "S")	
		End If		

		strcond2= strcond
		strcond = strCond  & " and a.ap_no in (" & strApno & ")"
		strcond2= strcond2 & " and a.ap_no not in (" & strApno & ")"

		If Cstr(lgIntFlgMode) = Cstr(parent.OPMD_UMODE) Then
			strSelect =             " is_checked,pay_bp_cd,bp_nm,ap_no,ap_dt,ap_due_dt,doc_cur,ap_amt,bal_amt,cls_amt,over_due_fg, "
			strSelect = strSelect & " acct_cd,ap_desc "
			
			strFrom   =             " ( " 
			strFrom   = strFrom   & " SELECT " & FilterVar("1", "''", "S") & "  is_checked, A.pay_bp_cd, B.bp_nm, A.ap_no, A.ap_dt, A.ap_due_dt, A.doc_cur, A.ap_amt,A.cls_amt, A.bal_amt, "
			strFrom   = strFrom   & " CASE WHEN A.ap_due_dt >= " & FilterVar(.txtAllcDt.Text , "''", "S") & " THEN " & FilterVar("IN", "''", "S") & "  ELSE " & FilterVar("OVER", "''", "S") & "  END over_due_fg, A.acct_cd, A.ap_desc "
			strFrom   = strFrom   & " FROM ufn_A_ClsApByBatchPayment(" & FilterVar(Trim(.txtBatchAllcNo.value),null,"S") & ") A"
			strFrom   = strFrom   & "	LEFT JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd "
			strFrom   = strFrom   & " WHERE A.ap_no IN (" & strApno & ") "
			strFrom   = strFrom   & " UNION ALL"
			strFrom   = strFrom   & " SELECT " & FilterVar("0", "''", "S") & "  is_checked, A.pay_bp_cd, B.bp_nm, A.ap_no, A.ap_dt, A.ap_due_dt, A.doc_cur, A.ap_amt, " & FilterVar("0", "''", "S") & "  cls_amt, A.cls_amt bal_amt,"
			strFrom   = strFrom   & " CASE WHEN A.ap_due_dt >=" & FilterVar(.txtAllcDt.Text , "''", "S") & " THEN " & FilterVar("IN", "''", "S") & "  ELSE " & FilterVar("OVER", "''", "S") & "  END over_due_fg, A.acct_cd, A.ap_desc"
			strFrom   = strFrom   & " FROM ufn_A_ClsApByBatchPayment(" & FilterVar(Trim(.txtBatchAllcNo.value),null,"S") & ") A"
			strFrom   = strFrom   & "	LEFT JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd"
			strFrom   = strFrom   & " WHERE A.ap_no NOT IN (" & strApno & ")) TMP "

			strWhere  = strWhere  & " doc_cur = " & FilterVar(lgKeyStream(1), "''", "S")
			strWhere  = strWhere  & " And pay_bp_cd= "	& FilterVar(lgKeyStream(0), "''", "S")			
			strWhere  = strWhere  & " ORDER BY ap_due_dt asc , ap_no asc "			
		Else
			strSelect =             " is_checked,pay_bp_cd,bp_nm,ap_no,ap_dt,ap_due_dt,doc_cur,ap_amt,bal_amt,cls_amt,over_due_fg, "
			strSelect = strSelect & " acct_cd,ap_desc "
			
			strFrom   =             " ( "
			strFrom   = strFrom   & " SELECT " & FilterVar("1", "''", "S") & "  is_checked, A.PAY_BP_CD, B.bp_nm, A.ap_no, A.ap_dt, A.ap_due_dt, A.doc_cur, A.ap_amt, A.bal_amt cls_amt, A.bal_amt,"
			strFrom   = strFrom   & " CASE WHEN A.ap_due_dt >= " & FilterVar(.txtAllcDt.Text , "''", "S") & " THEN " & FilterVar("IN", "''", "S") & "  ELSE " & FilterVar("OVER", "''", "S") & "  END over_due_fg, A.acct_cd, A.ap_desc "
			strFrom   = strFrom   & " FROM a_open_ap A LEFT JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd "
			strFrom   = strFrom   & " WHERE A.ap_amt > 0  AND A.conf_fg = " & FilterVar("C", "''", "S") & "  "
			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				strFrom   = strFrom   & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
			End If
	
			If lgInternalCd <> "" Then
				strFrom   = strFrom   & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
			End If
	
			If lgSubInternalCd <> "" Then
				strFrom   = strFrom   & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
			End If
	
			If lgAuthUsrID <> "" Then
				strFrom   = strFrom   & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
			End If			
			strFrom   = strFrom   & " and A.gl_no <> ''" & strcond 
			strFrom   = strFrom   & " UNION ALL"
			strFrom   = strFrom   & " SELECT " & FilterVar("0", "''", "S") & "  is_checked, A.PAY_BP_CD, B.bp_nm, A.ap_no, A.ap_dt, A.ap_due_dt, A.doc_cur, A.ap_amt, " & FilterVar("0", "''", "S") & "  cls_amt, a.bal_amt,"
			strFrom   = strFrom   & " CASE WHEN A.ap_due_dt >= " & FilterVar(.txtAllcDt.Text , "''", "S") & " THEN " & FilterVar("IN", "''", "S") & "  ELSE " & FilterVar("OVER", "''", "S") & "  END over_due_fg, A.acct_cd, A.ap_desc "
			strFrom   = strFrom   & " FROM a_open_ap A LEFT JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd "
			strFrom   = strFrom   & " WHERE A.ap_sts = " & FilterVar("O", "''", "S") & "  AND A.ap_amt > 0 AND A.conf_fg = " & FilterVar("C", "''", "S") & "  "
			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				strFrom   = strFrom   & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
			End If
	
			If lgInternalCd <> "" Then
				strFrom   = strFrom   & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
			End If
	
			If lgSubInternalCd <> "" Then
				strFrom   = strFrom   & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
			End If
	
			If lgAuthUsrID <> "" Then
				strFrom   = strFrom   & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
			End If
			strFrom   = strFrom   & " and A.gl_no <> ''" & strcond2 & ") AP "
			strFrom   = strFrom   & " ORDER BY ap_due_dt asc , ap_no asc "
		End If	

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = frm1.vspdData2
			arrTemp =  Split(lgF2By2,Chr(12))

			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next

			lgF2By2 = Join(arrTemp,Chr(12))
			ggoSpread.SSShowData lgF2By2
		End If

		.vspdData2.ReDraw = True
	End With

	Call LayerShowHide(0)

	DbQueryDetail = True
'	lgQueryOk = True
End Function


'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData2_Click(Col, Row)
	Call SetPopupMenuItemInf("0000111111")
    
    gMouseClickStatus = "SP2C"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData2
    
	If frm1.vspdData2.Maxrows = 0 Then Exit Sub

	If Row <= 0 Then
		Exit Sub
	End If  		
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim indx1 	
	Dim tempSum
	Dim dblTotPaymAmt
	Dim iChangeAmt,iChangeAmt2

	vspdData1ButtonClicked = True

	lgBlnFlgChgValue = True
	With frm1
		If .vspdData2.MaxRows > 0 Then
   			If Col = C_Checked_1 And vspdData2ButtonClicked = False Then
   				.vspddata.ReDraw = False
				.vspddata2.ReDraw = False

				If Trim(ButtonDown) = "1"  Then										'체크되어 있지 않은 것을 체크할때 
					For indx1 = 1 To .vspddata2.MaxRows					
						.vspdData2.Col  = C_Checked_2
						.vspdData2.Row  = indx1
						.vspdData2.Text = "1"					
					Next				
				
					If lgIntFlgMode = parent.OPMD_UMODE Then						'상단그리드는 수정모드일때 잔액의 금액을 반제액으로 합산 
						.vspddata.Row  = Row				
						.vspddata.Col  = C_bal_amt_1
						iChangeAmt     = .vspddata.text
						.vspddata.text = UNIConvNumPCToCompanyByCurrency(0, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")
						.vspddata.Col  = C_cls_amt_1
						iChangeAmt     = UNICdbl(iChangeAmt) + UNICdbl(.vspddata.text)
						.vspddata.text = UNIConvNumPCToCompanyByCurrency(iChangeAmt, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")
					Else															'상단그리드는 입력모드일때는 잔액의 금액을 반제액으로 복사 
						.vspddata.Row  = Row				
						.vspddata.Col  = C_bal_amt_1
						iChangeAmt     = .vspddata.text
						.vspddata.Col  = C_cls_amt_1
						.vspddata.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(iChangeAmt), .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")					
					End If
				Else																'체크되어 있는 것을 체크할때 
					For indx1 = 1 To .vspddata2.MaxRows					
						.vspdData2.Col  = C_Checked_2
						.vspdData2.Row  = indx1
						.vspdData2.Text = "0"					
					Next

					If lgIntFlgMode = parent.OPMD_UMODE Then						'상단그리드는 수정모드일때 잔액을 반제액에 합산 
						.vspddata.Row  = Row				
						.vspddata.Col  = C_cls_amt_1
						iChangeAmt     = .vspddata.text
						.vspddata.text = UNIConvNumPCToCompanyByCurrency(0, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")
						.vspddata.Col  = C_bal_amt_1
						iChangeAmt = UNICdbl(iChangeAmt) + UNICdbl(.vspddata.text)
						.vspddata.text = UNIConvNumPCToCompanyByCurrency(iChangeAmt, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")					
					Else
						.vspddata.Row  = Row										'상단그리드는 입력모드일때 반제액을 0으로 
						.vspddata.Col  = C_cls_amt_1
						.vspddata.text = UNIConvNumPCToCompanyByCurrency(0, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")						
					End If
				End If
				
				.vspddata2.ReDraw = True
				.vspddata.ReDraw = True

				If vspdData1ButtonClicked = True Then
					Call DoSum()
				End If			
			End If

			Call SetvspdDataCombo()			
		End If	
	End With
	
	vspdData1ButtonClicked = False
	lgIsQuery = False
End sub

'========================================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim indx
	Dim ClickedFg
	Dim tempSum
	Dim tempAmt
	Dim iChangeAmt
	dim sCurrency

	lgBlnFlgChgValue = True

	ClickedFg = False																	'vspdData2의 전체 Check Button이 "0" 인지 알아본다.
	vspdData2ButtonClicked = True														'vspdData2에서 Click Event가 일어났는지 알아본다.
	
	With frm1
		.vspdData.Col = C_Checked_1
		.vspdData.Row = .vspddata.ActiveRow
		
		Select Case Trim(ButtonDown)
			Case "1"																		'vspdData2가 체크되어 있지 않은것을 체크할경우 
				If lgIntFlgMode = parent.OPMD_UMODE Then									'하단그리드는 수정모드일때는 잔액과 반제액을 서로 바꾼다			
					.vspddata2.Row  = Row
					.vspddata2.Col  = C_bal_amt_2
					iChangeAmt      = .vspddata2.text

					.vspddata2.Col  = C_cls_amt_2
					.vspddata2.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(iChangeAmt), .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")				

					.vspddata2.Col  = C_bal_amt_2
					.vspddata2.text = UNIConvNumPCToCompanyByCurrency(0, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")				

					.vspddata.Row   = .vspddata.ActiveRow			
					.vspddata.Col   = C_bal_amt_1											'상단그리드는 수정모드일때 반제액에서 차감하고 
					TempAmt = UNICdbl(.vspddata.text) - UNICdbl(iChangeAmt)
					.vspddata.text  = UNIConvNumPCToCompanyByCurrency(TempAmt, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")
					.vspddata.Col   = C_cls_amt_1											'채무잔액에 합산 
					TempAmt = UNICdbl(.vspddata.text) + UNICdbl(iChangeAmt)
					.vspddata.text  = UNIConvNumPCToCompanyByCurrency(TempAmt, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")
				Else																		'하단그리드는 입력모드일때에는 잔액의 금액을 반제액으로 복사 
					.vspddata2.Row  = Row
					.vspddata2.Col  = C_bal_amt_2
					iChangeAmt = .vspddata2.text

					.vspddata2.Col  = C_cls_amt_2
					.vspddata2.text = UNIConvNumPCToCompanyByCurrency(UNICdbl(iChangeAmt), .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X") 

					.vspddata.Row   = .vspddata.ActiveRow			
					.vspddata.Col   = C_cls_amt_1
					TempAmt = UNICdbl(.vspddata.text) + UNICdbl(iChangeAmt)
					.vspddata.text  = UNIConvNumPCToCompanyByCurrency(TempAmt, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X") 				
				End If					
				
				.vspdData.Col = C_Checked_1												
				If Trim(.vspdData.Text) = "0" Or Trim(.vspdData.Text) = "" Then
					.vspdData.Text = "1"
				End If
				
				Call SetvspdDataCombo()				
			Case "0"																	'vspdData2가 체크된것을 체크할경우 
				If lgIntFlgMode = parent.OPMD_UMODE Then								'하단그리드는 수정모드일때는 반제액과 잔액을 서로 바꾼다						
					.vspddata2.Row = Row
					.vspddata2.Col = C_cls_amt_2
					iChangeAmt = .vspddata2.text

					.vspddata2.Col = C_cls_amt_2
					.vspddata2.text = UNIConvNumPCToCompanyByCurrency(0, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")					

					.vspddata2.Col = C_bal_amt_2
					.vspddata2.text = UNIConvNumPCToCompanyByCurrency(UNICdbl(iChangeAmt), .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")

					.vspddata.Row = .vspddata.ActiveRow			
					.vspddata.Col = C_bal_amt_1											'상단그리드는 수정모드일때에는 채무잔액에 합산하고 
					TempAmt = UNICdbl(.vspddata.text) + UNICdbl(iChangeAmt)
					.vspddata.text = UNIConvNumPCToCompanyByCurrency(TempAmt, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")
					.vspddata.Col = C_cls_amt_1											'채무반제액에서 차감 
					TempAmt = UNICdbl(.vspddata.text) - UNICdbl(iChangeAmt)
					.vspddata.text = UNIConvNumPCToCompanyByCurrency(TempAmt, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X")
				Else																	'하단그리드는 입력모드일때 반제액을 0으로 
					.vspddata2.Row = Row
					.vspddata2.Col = C_bal_amt_2
					iChangeAmt = .vspddata2.text

					.vspddata2.Col = C_cls_amt_2
					.vspddata2.text = UNIConvNumPCToCompanyByCurrency(0, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X") 

					.vspddata.Row = .vspddata.ActiveRow									'상단그리드는 입력모드일때 반제액에서 차감 
					.vspddata.Col = C_cls_amt_1
					TempAmt = UNICdbl(.vspddata.text) - UNICdbl(iChangeAmt)
					.vspddata.text = UNIConvNumPCToCompanyByCurrency(TempAmt, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X") 
				End If
				
				For indx = 1 To .vspdData2.MaxRows
					.vspdData2.Col = C_Checked_2
					.vspdData2.Row = indx
					If .vspdData2.Text = "1" Then
						ClickedFg = True
						Exit For
					End If
				Next

				.vspdData.Col = C_Checked_1				
				If Trim(.vspdData.Text) = "1" And ClickedFg = False Then				'vspdData2가 하나도 체크된것이 없으면 vspdData도 체크되지 안되도록 
					.vspdData.Text = "0"
				End If					

				Call SetvspdDataCombo()
		End Select

		If vspdData1ButtonClicked = False Then
			Call DoSum()
		End If			
	End With
	
	vspdData2ButtonClicked = False
End Sub

'==========================================================================================
'   Event Name : vspdData2_LeaveCell
'   Event Desc : This event is spread sheet data Button Clicked
'==========================================================================================
Sub vspdData2_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData2 
	    If Row >= NewRow Then
			Exit Sub
		End If
    End With
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
'   Event Name : vspdData2_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : SetvspdDataCombo
'   Event Desc : 
'========================================================================================================
Sub SetvspdDataCombo()
	Dim indx
	Dim TempApNo

	TempApNo = ""

	With frm1
		For indx = 1 To .vspdData2.MaxRows
		   .vspdData2.Col = C_Checked_2
		   .vspdData2.Row = indx
		   If .vspdData2.text = "1" Then						'⊙: 변경된값이 있으면 combo박스 다시 셋팅 
				.vspddata.Row = .vspddata.ActiveRow
				.vspddata.Col = C_checked_1
				If .vspddata.text = "0" or .vspddata.text = "" Then
					.vspddata.text = "1"
				End If					
				.vspdData2.Col = C_ap_no_2
				TempApNo=TempApNo & vbTab & .vspdData2.text				
		   End If
		Next
	End With

	Call InitComboBox(TempApNo)

	vspdData2ButtonClicked = False
End Sub

'========================================================================================================
'   Event Name : PayConRefView
'   Event Desc : 
'========================================================================================================
Sub PayConRefView()
	If lgIntFlgMode = parent.OPMD_CMODE then
		allcNoFG.innerHTML = "<A href=""vbscript:OpenRefOpenPaymCon()"">채무조건</A>"
    Else 
		allcNoFG.innerHTML = "<font color=""#777777"">채무조건</font>"
	End if
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	
End Sub
'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
'   Event Name :vspddata2_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtAllcDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtAllcDt.Text, gDateFormat,""), "''", "S") & "))"
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hOrgChangeId.value = ""
	Else 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
					
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.hOrgChangeId.value = Trim(arrVal2(2))
		Next	
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
'Sub  vspdData2_Click(ByVal Col, ByVal Row)
'	Call SetPopupMenuItemInf("0000111111")
'   
'    gMouseClickStatus = "SP2C"	'Split 상태코드 
'    Set gActiveSpdSheet = frm1.vspdData2
'      
'    If Row <= 0 Then
'        ggoSpread.Source = frm1.vspdData2
'        If lgSortKey = 1 Then
'            ggoSpread.SSSort Col
'            lgSortKey = 2
'        Else
'            ggoSpread.SSSort Col ,lgSortKey
'            lgSortKey = 1
'        End If
'   End If    
'End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True

    If UCase(Trim(frm1.txtDocCur.Value)) = UCase(parent.gCurrency) Then
		frm1.txtXchRate.Text = 1
	Else
		frm1.txtXchRate.Text = 0
	End If

	If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 출금액 
		ggoOper.FormatFieldByObjectOfCur .txtPaymAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' 채무금액 
		ggoSpread.SSSetFloatByCellOfCur C_ap_amt_1,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_Cls_amt_1,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 채무잔액 
		ggoSpread.SSSetFloatByCellOfCur C_bal_amt_1,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		
		ggoSpread.Source = frm1.vspdData2
		' 채무액 
		ggoSpread.SSSetFloatByCellOfCur C_ap_amt_2,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_Cls_amt_2,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 채무잔액 
		ggoSpread.SSSetFloatByCellOfCur C_bal_amt_2,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
	End With
End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()
		Case "VSPDDATA2"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")			
			Call ggoSpread.ReOrderingSpreadData()
	End Select
End Sub

'========================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'========================================================================================================
Sub DoSum()
	Dim indx
	Dim TempSum

	With frm1
		For indx = 1 To .vspdData.MaxRows
			.vspdData.Col = C_Checked_1
			.vspdData.Row = indx
			If .vspdData.Text = "1" Then
				.vspddata.col = C_cls_amt_1
				TempSum = TempSum + UNICdbl(.vspddata.Text)
			End If		
		Next	

		.txtPaymAmt.text = UNIConvNumPCToCompanyByCurrency(TempSum, .txtDocCur.value ,parent.ggAmtOfMoneyNo, "X" , "X") 		
	End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>일괄출금등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right >
					<Span id="allcNoFG"><A href="vbscript:OpenRefOpenPaymCon()">채무조건</A></Span>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
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
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>일괄출금번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtBatchAllcNo" ALT="일괄출금번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript: Call OpenPopup(frm1.txtBatchAllcNo.value,0)"></TD>								
								</TR>						
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>		
							<TABLE <%=LR_SPACE_TYPE_60%> border=0>																				
								<TR>
									<TD CLASS=TD5 NOWRAP>출금일</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtAllcDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="출금일" id=fpDateTime1></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>부서</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="22NXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.value,1)">
										<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="부서명">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>지급유형</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtInputType" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtInputType.value,3)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
													   <INPUT TYPE=TEXT NAME="txtInputTypeNm" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>																	   
									<TD CLASS=TD5 NOWRAP>은행</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="Text" NAME="txtBankCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="24NXXU" ALT="은행"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtBankCd.value,4)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtBankNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="은행명">
									</TD>

								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출금계정코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="출금계정코드" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtAcctCd.value,2)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
													 <INPUT NAME="txtAcctnm" ALT="계정코드명" MAXLENGTH="20"  tag  ="24"></TD>
									<TD CLASS=TD5 NOWRAP>계좌번호</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT  TYPE=TEXT NAME="txtBankAcct" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" tag="24XXXU" ALT="계좌번호"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtBankAcct.value,5)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									</TD>
												
								</TR>
								<TR>																		
									<TD CLASS=TD5 NOWRAP>어음만기일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtNoteDueDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="어음만기일" id=fpDateTime1></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>카드사</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="Text" NAME="txtCardCoCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="24NXXU" ALT="카드사"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtCardCoCd.value,6)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtCardCoNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="카드사명">
									</TD>
<!--									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD> -->
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtPaymAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="출금액" tag="24X2" id=OBJECT3></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>거래통화
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="24NXXU" STYLE="TEXT-ALIGN: left" ALT="거래통화">									
									&nbsp;&nbsp;환율<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtXchRate" align="top" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="환율" tag="24X5Z" id=OBJECT5></OBJECT>');</SCRIPT></TD>       
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> </TD>																						
									<TD CLASS="TD5" NOWRAP>전표번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="전표번호"></TD>								
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>비고</TD>
									<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtPaymDesc" SIZE=90 MAXLENGTH=128 tag="21XXX" ALT="비고"></TD>
								</TR>						
								<TR>
									<TD WIDTH=100% HEIGHT=100% valign=top COLSPAN=4>
										<TABLE <%=LR_SPACE_TYPE_30%>>
											<TR>
												<TD HEIGHT="50%">
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
												</TD>
											</TR>
											<TR>
												<TD HEIGHT="50%" >
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT5> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
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
	<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_NOTE_INF)">어음정보등록</A>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"   tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtPayBpCd"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtPaymType"       tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtDueDt"		    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtBizAreaCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtBizAreaCd1"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"      tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtTempAPno" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="hstrVal"           tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hBankCd"           tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hBankNm"           tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hCardCoCd"         tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hCardCoNm"         tag="24" TABINDEX="-1">
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


