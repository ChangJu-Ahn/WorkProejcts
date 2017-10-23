<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Account Receivable
'*  3. Program ID           : a3101ma1
'*  4. Program Name         : 채권등록 
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/01/03
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
 -->
 
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'            1. 선 언 부 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc 선언   
' 기능: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>	
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit                 '☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
' .Constant는 반드시 대문자 표기.
' .변수 표준에 따름. prefix로 g를 사용함.
' .Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_QRY_ID  = "a3101mb1.asp"							'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a3101mb2.asp"							'☆: Save 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID  = "a3101mb3.asp"

Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"			'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1													'☜: Tab의 위치 
Const TAB2 = 2

Const MENU_NEW  = "1110100000011111"
Const MENU_SIN_CRT = "1110100000011111"       
Const MENU_MUL_CRT = "1110111100111111"
Const MENU_SIN_UPD = "1111100000011111"
Const MENU_MUL_UPD = "1111111100111111"

Dim C_ItemSeq													
Dim C_AcctCd 
Dim C_AcctPB 
Dim C_AcctNm 
Dim C_DeptCd 
Dim C_DeptPB 
Dim C_DeptNm 
Dim C_VatType 
Dim C_VatPB 
Dim C_VatNm 
Dim C_NetAmt 
Dim C_NetLocAmt 
Dim C_ItemDesc 
Dim C_DrCRFG

Dim  lgStrPrevKeyDtl
Dim  IsOpenPop													'Popup
Dim  gSelframeFlg
Dim  lgQueryOk					' Queryok여부 (loc_amt =0 check)
Dim	 lgstartfnc
Dim	 lgFormLoad

Dim  lgCurrRow
Dim  lgArrAcctForVat
Dim  lgBlnGetAcctForVat

<%
Dim dtToday
dtToday = GetSvrDate
%>

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

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
	C_ItemSeq   = 1
	C_AcctCd    = 2
	C_AcctPB    = 3
	C_AcctNm    = 4 
	C_DeptCd    = 5 
	C_DeptPB    = 6
	C_DeptNm    = 7
	C_VatType   = 8
	C_VatPB     = 9
	C_VatNm     = 10
	C_NetAmt    = 11
	C_NetLocAmt = 12
	C_ItemDesc  = 13
	C_DrCRFG    = 14
End Sub

'======================================================================================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode     = parent.OPMD_CMODE						'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False									'Indicates that no value changed
    lgIntGrpCount    = 0										'initializes Group View Size
    lgStrPrevKey     = 0										'initializes Previous Key
    lgStrPrevKeyDtl  = 0										'initializes Previous Key
    lgLngCurRows     = 0										'initializes Deleted Rows Count
    lgQueryOk        = False
	lgstartfnc		 = False
	lgFormLoad		 = True
End Sub

'======================================================================================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtArDt.text = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,gDateFormat)
	frm1.txtDueDt.text = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,gDateFormat)
'	frm1.cboArType.value = "NR" 
	frm1.txtDocCur.value = parent.gCurrency
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	
	frm1.txtXchRate.text = 1
	
	Call ggoOper.SetReqAttr(frm1.txtPayBpCd, "N")  
	Call ggoOper.SetReqAttr(frm1.txtDocCur,  "N") 

	lgBlnFlgChgValue = False									'Indicates that no value changed 
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
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

		.MaxCols = C_ItemDesc + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols															'공통콘트롤 사용 Hidden Column
		.ColHidden = True    
		.MaxRows = 0
		    
		Call AppendNumberPlace("6","3","0")
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetFloat  C_ItemSeq  , "NO"            ,6 ,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"
		ggoSpread.SSSetEdit   C_AcctCd   , "계정코드"      ,20,,,20,2
		ggoSpread.SSSetButton C_AcctPB
		ggoSpread.SSSetEdit   C_AcctNm   , "계정코드명"    ,20,3
		ggoSpread.SSSetEdit   C_DeptCd   , "부서"          ,17, ,,10,2
		ggoSpread.SSSetButton C_DeptPB
		ggoSpread.SSSetEdit   C_DeptNm   , "부서명"        ,20,3    
		ggoSpread.SSSetEdit   C_VatType  , "부가세"        ,10,3,,10,2        
		ggoSpread.SSSetButton C_VatPB
		ggoSpread.SSSetEdit   C_VatNm    , "부가세유형명"  ,12,3        
		ggoSpread.SSSetFloat  C_NetAmt   , "순매출액"      ,15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_NetLocAmt, "순매출액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_ItemDesc , "비고"          ,50, ,,128           
		ggoSpread.SSSetEdit   C_DrCRFG   , "차대"          ,2, ,,2           

		Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPB)
		Call ggoSpread.MakePairsColumn(C_DeptCd,C_DeptPB)
		Call ggoSpread.MakePairsColumn(C_VatType,C_VatPB)
		Call ggoSpread.SSSetColHidden(C_DrCRFG,C_DrCRFG,True)
		
		.Redraw = True 
    End With

	Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock()
    Dim objSpread
    Dim C_MAX
    
    C_MAX = frm1.vspdData.MaxCols

    With frm1
		ggoSpread.Source = .vspdData
		Set objSpread = .vspdData

		objSpread.Redraw = False
		    
		ggoSpread.SpreadLock C_ItemSeq, -1, C_ItemSeq, -1
		ggoSpread.SpreadLock C_AcctCd , -1, C_AcctCd, -1
		ggoSpread.SpreadLock C_AcctPB , -1, C_AcctPB, -1
		ggoSpread.SpreadLock C_AcctNm , -1, C_AcctNm, -1
		ggoSpread.SpreadLock C_DeptNm , -1, C_DeptNm, -1                            
		ggoSpread.SpreadLock C_VatType  , -1, C_VatNm, -1                            

		ggoSpread.SSSetRequired  C_DeptCd, -1, -1 
		ggoSpread.SSSetRequired  C_NetAmt, -1, -1
		ggoSpread.SSSetProtected C_MAX   , -1, -1
		
		
		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		.Redraw = False
				
		ggoSpread.SSSetProtected C_ItemSeq, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_AcctCd , pvStartRow, pvEndRow				
 
		ggoSpread.SSSetProtected C_AcctNm , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DeptCd , pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_DeptNm , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_NetAmt , pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_VatNm  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VatType  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VatPB  , pvStartRow, pvEndRow
		    
		.Col = 1							'절대 위치로 이동 
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True
		.Redraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpread2ColorAr
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpread2ColorAr()
	Dim i
	dim j
    With frm1
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False  
		For i = 1 to .vspdData2.MaxRows
			ggoSpread.SSSetProtected C_DtlSeq   , i, i
			ggoSpread.SSSetProtected C_CtrlCd   , i, i
			ggoSpread.SSSetProtected C_CtrlNm   , i, i
			ggoSpread.SSSetProtected C_CtrlValNm, i, i   
			.vspddata2.Row = i
			.vspddata2.Col = C_DrFg

			If (.vspddata2.text = "Y") OR (.vspddata2.text = "C") OR (.vspddata2.text = "DC") Then
				ggoSpread.SSSetRequired C_CtrlVal, i, i ' 
			End if
		Next
		.vspdData2.ReDraw = True
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

			C_ItemSeq   = iCurColumnPos(1)
			C_AcctCd    = iCurColumnPos(2)
			C_AcctPB    = iCurColumnPos(3)
			C_AcctNm    = iCurColumnPos(4) 
			C_DeptCd    = iCurColumnPos(5)
			C_DeptPB    = iCurColumnPos(6)
			C_DeptNm    = iCurColumnPos(7)
			C_VatType   = iCurColumnPos(8)
			C_VatPB     = iCurColumnPos(9)
			C_VatNm     = iCurColumnPos(10)
			C_NetAmt    = iCurColumnPos(11)
			C_NetLocAmt = iCurColumnPos(12)
			C_ItemDesc  = iCurColumnPos(13)
			C_DrCRFG    = iCurColumnPos(14)
	End select
End Sub

'======================================================================================================
' Function Name : OpenPopupGL
' Function Desc : This method Open The Popup window for GL
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8) 
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)							'회계전표번호 
	arrParam(1) = ""												'Reference번호	

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
	  
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
End Function

 '========================================== 2.4.2 OpenPopuptempGL()  =============================================
'	Name : OpenPopuptempGL()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	arrParam(0) = Trim(frm1.txtTempGlNo.value)						'결의전표번호 
	arrParam(1) = ""												'Reference번호 

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _	
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

 '========================================== 2.4.2 OpenRefPreRcptNo()  =============================================
' Name : OpenRefPreRcptNo()
' Description : 선수금(Pre Receipt) Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefPreRcptNo()
	Dim arrRet
	Dim arrParam(9)
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a3109ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3109ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPayBpCd.value))    ' 검색조건이 있을경우 파라미터 
	arrParam(1) = UCase(Trim(frm1.txtpaybpnm.value))
	arrParam(2) = UCase(Trim(frm1.txtDocCur.value))
	arrParam(3) = "S"
	arrParam(4) = UCase(Trim(frm1.txtArDt.text))
    arrParam(5) = ""

	' 권한관리 추가 
	arrParam(6) = lgAuthBizAreaCd
	arrParam(7) = lgInternalCd
	arrParam(8) = lgSubInternalCd
	arrParam(9) = lgAuthUsrID
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
 
	If arrRet(0) = "" Then  
		frm1.txtPrPaymNo.focus
		Exit Function
	Else  
		Call SetRefPreRcptNo(arrRet)
	End If
End Function

 '------------------------------------------  SetRefPreRcptNo()  ---------------------------------------
' Name : SetRefPreRcptNo()
' Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function  SetRefPreRcptNo(Byval arrRet)
	lgBlnFlgChgValue = True
	
	With frm1
		.txtPrPaymNo.Value    = arrRet(0)  'C_PpNo = 1
		.txtPayBpCd.Value     = arrRet(3)  'C_BpCd = 4
		.txtpaybpnm.Value     = arrRet(4)  'C_BpNm = 5
		.txtDocCur.value      = arrRet(8)  'C_DocCur = 9 
		.txtPrRcptAmt.Text    = arrRet(11)
		.txtPrRcptLocAmt.Text = arrRet(12)

		If UCase(Trim(frm1.txtDocCur.value)) <> parent.gCurrency Then
			frm1.txtXchRate.Text = 0
		Else
			frm1.txtXchRate.Text = 1
		End If	

		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
		 
		If frm1.txtPayBpCd.value <> "" Then     
		   Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "Q")  
		Else   
		   Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "N")  
		End If
		
		If frm1.txtPayBpCd.value <> "" Then     
		   Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")  
		Else   
		   Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")  
		End If    

		frm1.txtPrPaymNo.focus
	End With
End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = frm1.txtArDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscDeptPopup(iWhere)
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
End Function
'------------------------------------------  EscPopup()  ------------------------------------------------
'	Name : EscPopup()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function EscDeptPopup(Byval iWhere)
	With frm1
	
		Select Case iWhere
		     Case "0"
				.txtDeptCd.focus
			 Case "1"
				Call SetActiveCell(.vspdData,C_deptcd,.vspdData.ActiveRow ,"M","X","X")
	    End Select
	End With
End Function     
'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		     Case "0"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtArDt.text = arrRet(3)
				call txtDeptCd_OnBlur()  
				.txtDeptCd.focus
			 Case "1"
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
				
				.vspdData.Col  = C_deptcd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_deptnm
				.vspdData.Text = arrRet(1)
				Call deptCd_underChange(arrRet(0))
				Call SetActiveCell(frm1.vspdData,C_deptcd,frm1.vspdData.ActiveRow ,"M","X","X")

	    End Select
	End With
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

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "B"							'B :매출 S: 매입 T: 전체 
	Select Case iWhere
		Case 3
			arrParam(5) = "SOL"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
		Case 9
			arrParam(5) = "INV"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
		Case 4
			arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	End Select
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then     
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function 
 
	Select Case iWhere
		Case 0

		Case 1
			arrParam(0) = "계정코드팝업"							' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"	' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
			    " and C.trans_type = " & FilterVar("AR005", "''", "S") & "  and C.jnl_cd = " & FilterVar("AR", "''", "S") & " "       ' Where Condition
			arrParam(5) = "계정코드"								' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"									' Field명(0)
			arrField(1) = "A.Acct_NM"									' Field명(1)
			   arrField(2) = "B.GP_CD"									' Field명(2)
			arrField(3) = "B.GP_NM"										' Field명(3)
		 
			arrHeader(0) = "계정코드"								' Header명(0)
			arrHeader(1) = "계정코드명"								' Header명(1)
			arrHeader(2) = "그룹코드"								' Header명(2)
			arrHeader(3) = "그룹명"									' Header명(3)
		Case 2
			arrParam(0) = "계정코드팝업"							' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B"						' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & " "			' Where Condition
			arrParam(5) = "계정코드"								' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"									' Field명(0)
			arrField(1) = "A.Acct_NM"									' Field명(1)
		    arrField(2) = "B.GP_CD"										' Field명(2)
			arrField(3) = "B.GP_NM"										' Field명(3)
		 
			arrHeader(0) = "계정코드"								' Header명(0)
			arrHeader(1) = "계정코드명"								' Header명(1)
			arrHeader(2) = "그룹코드"								' Header명(2)
			arrHeader(3) = "그룹명"   
		Case 3
			arrParam(0) = "주문처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "주문처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    
			arrHeader(0) = "주문처"							' Header명(0)
			arrHeader(1) = "주문처명"						' Header명(1)
		Case 4
			If UCase(frm1.txtPayBpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
			arrParam(0) = "수금처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "수금처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    		arrHeader(0) = "수금처"							' Header명(0)
			arrHeader(1) = "수금처명"						' Header명(1)
		Case 5       
			arrParam(0) = "사업장팝업"								' 팝업 명칭 
			arrParam(1) = "B_Biz_AREA"									' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(4) = ""											' Where Condition
			arrParam(5) = "사업장"   
 
			arrField(0) = "Biz_AREA_CD"									' Field명(0)
			arrField(1) = "Biz_AREA_NM"									' Field명(1)    
			 
			arrHeader(0) = "사업장"									' Header명(0)
			arrHeader(1) = "사업장명"								' Header명(1)
		Case 8
			If UCase(frm1.txtDocCur.className) = UCase(parent.UCN_PROTECTED) Then Exit Function 
		 
			arrParam(0) = "거래통화팝업"							' 팝업 명칭 
			arrParam(1) = "b_currency"									' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(4) = ""											' Where Condition
			arrParam(5) = "거래통화"    

			arrField(0) = "CURRENCY"									' Field명(0)
			arrField(1) = "CURRENCY_DESC"								' Field명(1)

			arrHeader(0) = "거래통화"								' Header명(0)
			arrHeader(1) = "거래통화명"								' Header명(1)    
		Case 9
			arrParam(0) = "세금계산서발행처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "세금계산서발행처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    
			arrHeader(0) = "세금계산서발행처"							' Header명(0)
			arrHeader(1) = "세금계산서발행처명"						' Header명(1)

		Case 10
			If  UCase(frm1.txtPayMethCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
 
			arrHeader(0) = "결제방법"								' Header명(0)
			arrHeader(1) = "결제방법명"								' Header명(1)
			arrHeader(2) = "Reference"
			 
			arrField(0) = "B_Minor.MINOR_CD"							' Field명(0)
			arrField(1) = "B_Minor.MINOR_NM"							' Field명(1)
			arrField(2) = "b_configuration.REFERENCE"
			 
			arrParam(0) = "결제방법"								' 팝업 명칭 
			arrParam(1) = "B_Minor,b_configuration"						' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPayMethCd.Value)					' Code Condition
		 
			arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9004", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
			              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"  
			arrParam(5) = "결제방법"								' TextBox 명칭 
		Case 11  
			If UCase(frm1.txtPayTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
 
			arrParam(0) = "입금유형"								' 팝업 명칭 
			If Trim(frm1.txtPayMethCd.Value) = "" then
				arrParam(1) = "B_MINOR,B_CONFIGURATION "
				arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
				  & "AND B_CONFIGURATION.SEQ_NO = 3 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PR", "''", "S") & " "   ' Where Condition     
			Else   
				arrParam(1) = "B_MINOR,B_CONFIGURATION," _
				& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & " "_
				 & "And MINOR_CD= " & FilterVar(frm1.txtPayMethCd.value, "''", "S") & " And SEQ_NO >= 2) C "   ' TABLE 명칭 
				arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
				  & " AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & " ," & FilterVar("R", "''", "S") & " ) "		' Where Condition
			End If  
			arrParam(2) = Trim(frm1.txtPayTypeCd.value)					' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(5) = "입금유형"								' TextBox 명칭 
	 
			arrField(0) = "B_MINOR.MINOR_CD"							' Field명(0)
			arrField(1) = "B_MINOR.MINOR_NM"							' Field명(1)
			  
			arrHeader(0) = "입금유형"								' Header명(0)
			arrHeader(1) = "입금유형명"								' Header명(1)  
		Case 12
			arrHeader(0) = "부가세유형"								' Header명(0)
			arrHeader(1) = "부가세명"								' Header명(1)
			arrHeader(2) = "부가세Rate"
			 
			arrField(0) = "B_Minor.MINOR_CD"							' Field명(0)
			arrField(1) = "B_Minor.MINOR_NM"							' Field명(1)
			arrField(2) = "b_configuration.REFERENCE"
			 
			arrParam(0) = "부가세유형"								' 팝업 명칭 
			arrParam(1) = "B_Minor,b_configuration"						' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			 
			arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9001", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
			              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"  
			arrParam(5) = "부가세유형"								' TextBox 명칭 
	End Select    
 
	IsOpenPop = True
	 
	If iwhere = 0 Then  
		Dim iCalledAspName

		iCalledAspName = AskPRAspName("a3101ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
			IsOpenPop = False
			Exit Function
		End If

		Dim arrParam_1(8)

		' 권한관리 추가 
		arrParam_1(5)		= lgAuthBizAreaCd
		arrParam_1(6)		= lgInternalCd
		arrParam_1(7)		= lgSubInternalCd
		arrParam_1(8)		= lgAuthUsrID
	   
		arrRet = window.showModalDialog(iCalledAspName,array(window.parent,arrParam_1), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")   
	End If
	 
	IsOpenPop = False
 
	If arrRet(0) = "" Then     
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : EscPopup(Byval iWhere)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtArNo.focus
			Case 1 
				.txtAcctCd.focus
			Case 2
				Call SetActiveCell(.vspdData,C_AcctCD,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				.txtDealBpCd.focus
			Case 4
				.txtPayBpCd.focus
			Case 5   
				.txtReportBizCd.focus
			Case 8
			    .txtDocCur.focus
			Case 9
				.txtReportBpCd.focus
			Case 10
				.txtPayMethCd.focus
			Case 11 
			    .txtPayTypeCd.focus            
			Case 12 
				Call SetActiveCell(.vspdData,C_VatType,.vspdData.ActiveRow ,"M","X","X")

		End Select    
	End With

End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtArNo.value = arrRet(0)
				.txtArNo.focus
			Case 1 
				.txtAcctCd.value = arrRet(0)
				.txtAcctNm.value = arrRet(1)
				.txtAcctCd.focus
			Case 2
				.vspdData.Col = C_AcctCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_AcctNm
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow )  ' 변경이 일어났다고 알려줌        
				Call SetActiveCell(.vspdData,C_AcctCD,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				.txtDealBpCd.value = arrRet(0)
				.txtDealBpNm.value = arrRet(1)
				Call txtDealBpCd_onChange()
				.txtDealBpCd.focus
			Case 4
				.txtPayBpCd.value = arrRet(0)
				.txtPayBpNm.value = arrRet(1)
				.txtPayBpCd.focus
			Case 5   
				.txtReportBizCd.value = arrRet(0)
				.txtReportBizNm.value = arrRet(1)
				.txtReportBizCd.focus
			Case 8
			    .txtDocCur.value = arrRet(0)
			    Call txtDocCur_OnChange()
			    .txtDocCur.focus
			Case 9
			    .txtReportBpCd.value = arrRet(0)
				.txtReportBpNm.value = arrRet(1)
				.txtReportBpCd.focus
			Case 10
			    .txtPayMethCd.Value = arrRet(0)
				.txtPayMethNm.Value = arrRet(1)
				.txtPayMethCd.focus
			Case 11 
				.txtPayTypeCd.value = arrRet(0)
			    .txtPayTypeNm.value = arrRet(1)   
			    .txtPayTypeCd.focus            
			Case 12 
			    .vspdData.Col = C_VatType
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_VatNm
				.vspdData.Text = arrRet(1) 
				Call vspdData_Change(.vspdData.Col, .vspdData.Row )  ' 변경이 일어났다고 알려줌       
				Call SetActiveCell(.vspdData,C_VatType,.vspdData.ActiveRow ,"M","X","X")

		End Select    
	End With
 
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If 
End Function

'======================================================================================================
' 기능: Tab Click
' 설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar(MENU_SIN_CRT)				'⊙: 버튼 툴바 제어 
	Else     
	    Call SetToolbar(MENU_SIN_UPD)				'⊙: 버튼 툴바 제어 
	End If
 
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)									'~~~ 첫번째 Tab 
	gSelframeFlg = TAB1  
End Function

Function ClickTab2()
	frm1.vspdData.ReDraw = True  
	frm1.vspdData2.ReDraw = True  
	 
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)									'~~~ 두번째 Tab 
	gSelframeFlg = TAB2
 
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolBar(MENU_MUL_CRT)
	Else     
		Call SetToolBar(MENU_MUL_UPD)
	End if 
End Function



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************


'======================================================================================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub  Form_Load()
    Call LoadInfTB19029()																	'Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
       gDateFormat, parent.gComNum1000, parent.gComNumDec)    
                         
    Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field    
	Call InitSpreadSheet()																	'Setup the Spread sheet
	Call InitCtrlSpread()														'관리항목 그리드 초기화	
	Call InitCtrlHSpread()														'관리항목 그리드 초기화	
    Call InitVariables()																	'Initializes local global variables    
    
    Call SetToolbar(MENU_NEW)													'버튼 툴바 제어 
	Call SetDefaultVal()
	Call GetAcctForVat()  
 
	gIsTab     = "Y" 
	gTabMaxCnt = 2   

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
	 
	frm1.txtArNo.focus
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    Dim var1, var2

    FncQuery = False                                                        
	lgstartfnc = True

    Err.Clear                                                               
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then						'This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then  
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")     
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					'Clear Contents  Field
    Call ClickTab1()
    Call InitVariables()									'Initializes local global variables
    ggoSpread.Source = frm1.vspdData        				
    ggoSpread.ClearSpreadData						
    ggoSpread.Source = frm1.vspdData2        				
    ggoSpread.ClearSpreadData						
    ggoSpread.Source = frm1.vspdData3        				
    ggoSpread.ClearSpreadData						
    '-----------------------
    'Query function call area
    '-----------------------                  
    Call DbQuery()											'☜: Query db data    
    FncQuery = True  
    lgstartfnc = False	 
    
	Set gActiveElement = document.activeElement
       
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
	Dim var1, var2
     
    FncNew = False                                                          
	lgstartfnc = True 

    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field    
    Call InitVariables()																'Initializes local global variables
    call SetDefaultVal()
    
    ggoSpread.Source = frm1.vspdData        				
    ggoSpread.ClearSpreadData						
    ggoSpread.Source = frm1.vspdData2        				
    ggoSpread.ClearSpreadData						
    ggoSpread.Source = frm1.vspdData3        				
    ggoSpread.ClearSpreadData						
    
    frm1.txtArNo.Value = ""
    frm1.txtArNo.focus
    
    Call txtDocCur_OnChange()
    
	lgBlnFlgChgValue = False    
    FncNew = True      
    lgFormLoad = True							' tempgldt read
    lgstartfnc = False
    Set gActiveElement = document.activeElement

End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncDelete() 
    Dim IntRetCD
    
    FncDelete = False                                                      
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")              'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete()               '☜: Delete db data
    
    FncDelete = True   
    Set gActiveElement = document.activeElement
                                                     
End Function

'=======================================================================================================
' Function Name : `
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1,var2
 
    FncSave = False                                                         
    
    Err.Clear                                                               
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False And var2 = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
		Exit Function
    End If
    
    If Not chkField(Document, "2") Then               '⊙: Check required field(Single area)
		Exit Function
    End If

    '================================================================================================
    '일자관계 체크 : LC발행일(txtLcDt)<=송장일(txtInvDt)<=선하증권일(txtBlDt)<=채권/채무일(txtArDt)
    '================================================================================================
    If frm1.txtBlDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtBlDt.Text,frm1.txtArDt.Text,frm1.txtBlDt.Alt,frm1.txtArDt.Alt, _
		                      "970025",frm1.txtBlDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			frm1.txtBlDt.focus
			Exit Function
		End If
    End If
    
    If frm1.txtInvDt.Text <> "" Then
		If frm1.txtBlDt.Text = "" Then
			If CompareDateByFormat(frm1.txtInvDt.Text,frm1.txtArDt.Text,frm1.txtInvDt.Alt,frm1.txtArDt.Alt, _
			                    "970025",frm1.txtInvDt.UserDefinedFormat,parent.gComDateType, true) = False Then
				frm1.txtInvDt.focus
				Exit Function
			End If
		Else
			If CompareDateByFormat(frm1.txtInvDt.Text,frm1.txtBlDt.Text,frm1.txtInvDt.Alt,frm1.txtBlDt.Alt, _
			                     "970025",frm1.txtInvDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			    frm1.txtInvDt.focus
			    Exit Function
			End If
		End If
    End If
    '================================================================================================
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Call ClickTab2()
		Exit Function
    End If

    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Call ClickTab2()
		Exit Function
    End If
	
	If CheckSpread3 = False then
		IntRetCD = DisplayMsgBox("110420", "X", "X", "X")
        Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()                                                 '☜: Save db data
    FncSave = True  

    Set gActiveElement = document.activeElement
                                                     
End Function

'=======================================================================================================
' Function Name : CheckSpread3
' Function Desc : 저장시에  관리항목 필수여부 check 하기위해 호출되는 Function
'=======================================================================================================
Function CheckSpread3()
	Dim indx, jj
	Dim tmpDrCrFG,tmpItemSeq

	CheckSpread3 = False

	With frm1
		For jj = 1 To .vspdData.MaxRows
			.vspdData.row = jj
			tmpDrCrFG = "C"
			.vspdData.col = C_ItemSeq
			tmpItemSeq = .vspddata.Text

	 		For indx = 1 to .vspdData3.MaxRows
			    .vspdData3.Row = indx
	 			.vspdData3.Col = 8

	 			If tmpItemSeq = .vspddata3.Text Then
					.vspdData3.Col = 14

					If (tmpDrCrFG = .vspddata3.Text) Or .vspddata3.Text = "DC" Then
  						.vspdData3.Col = 5
						If Trim(.vspdData3.Text) = "" Then
							Exit Function
			  			End If
					End If
				End If	
			Next
		Next	
	End With

	CheckSpread3 = True
End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
	Dim  IntRetCD
	 
	If frm1.vspdData.Maxrows < 1 Then Exit Function 
	frm1.vspdData.ReDraw = False
	If lgBlnFlgChgValue = True Then
		' IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X") '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	 
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.Source = frm1.vspdData 
		ggoSpread.CopyRow
		Call SetSpreadColor(frm1.vspdData.ActiveRow,  frm1.vspdData.ActiveRow)
		MaxSpreadVal frm1.vspdData, C_ItemSeq, frm1.vspdData.ActiveRow
		Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)
		.vspdData.ReDraw = True
	End With
	Call DoSum()
	
	Set gActiveElement = document.activeElement

End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then exit Function
    
    With frm1.vspdData
		.Row = .ActiveRow
		.Col = 0
		If .Text = ggoSpread.InsertFlag Then
		    .Col = C_AcctCd
			If Len(Trim(.text)) > 0 Then  
				.Col = C_ItemSeq          
				Call DeleteHSheet(.Text)
			End If
		End If
   
		ggoSpread.Source = frm1.vspdData 
		ggoSpread.EditUndo
		Call DoSum()   
		If frm1.vspdData.MaxRows < 1 Then exit Function
		.Row = .ActiveRow
		.Col = 0
  
		IF .Row = 0 Then
			 Exit Function
		Else 
			If .Text = ggoSpread.InsertFlag Then
				.Col = C_AcctCd
				If len(Trim(.text)) > 0 Then 
					.Col = C_ItemSeq
					frm1.hItemSeq.value = .Text
					ggoSpread.Source = frm1.vspdData2        				
					ggoSpread.ClearSpreadData						
					Call DbQuery3(.ActiveRow)
					Call SetSpread2ColorAr()
				End IF 
			Else
			    .Col = C_ItemSeq
			    frm1.hItemSeq.value = .Text
				ggoSpread.Source = frm1.vspdData2        				
				ggoSpread.ClearSpreadData						
			    Call DbQuery2(.ActiveRow)
			End if
		End IF        
	End With
		
	Set gActiveElement = document.activeElement

End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
	
    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error stat	

    If gSelframeFlg <> TAB2 Then
		Call ClickTab2()																'sstData.Tab = 1
    End If
    
	FncInsertRow = False															'☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		

	With frm1.vspdData
		iCurRowPos = .ActiveRow
        .ReDraw = False
        ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow ,imRow
		
		For ii = .ActiveRow To  .ActiveRow + imRow - 1
			Call MaxSpreadVal(frm1.vspdData, C_ItemSeq, ii)
			.Col = C_DeptCd
			.Row = ii
			.Text = frm1.txtDeptCd.Value
		
			.Col = C_DeptNm
			.Row = ii
			.Text = frm1.txtDeptNm.Value

			.Col = C_DrCRFG
			.Row = ii
			.Text = "CR"
		Next

		.Col = 2																	' 컬럼의 절대 위치로 이동      
		.Row = 	ii - 1
		.Action = 0		
		
        Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)        
        .ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
   
	ggoSpread.Source = frm1.vspdData2        				
	ggoSpread.ClearSpreadData		
		
	Set gActiveElement = document.activeElement
				
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows
	Dim iDelRowCnt, i
    Dim DelItemSeq

	If frm1.vspdData.MaxRows < 1 Then exit Function
 
    With frm1.vspdData 
		.Row = .ActiveRow
		.Col = C_ItemSeq 
		DelItemSeq = .Text
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
    End With

	ggoSpread.Source = frm1.vspdData2        				
	ggoSpread.ClearSpreadData						
    Call DeleteHsheet(DelItemSeq)
    
    Call DoSum()
    	
	Set gActiveElement = document.activeElement

End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next    
	Call parent.FncPrint()  
		
	Set gActiveElement = document.activeElement
                                         
End Function

'=======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function  FncPrev() 
    On Error Resume Next         
    	
	Set gActiveElement = document.activeElement
                                      
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function  FncNext() 
    On Error Resume Next     
    	
	Set gActiveElement = document.activeElement
                                          
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
		
	Set gActiveElement = document.activeElement

End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)   
    	
	Set gActiveElement = document.activeElement
                       
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

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1,var2
 
	FncExit = False

	ggoSpread.Source = frm1.vspdData
	var1 = ggoSpread.SSCheckChange

	ggoSpread.Source = frm1.vspdData2
	var2 = ggoSpread.SSCheckChange
	   
	If lgBlnFlgChgValue = True or var1 = True or var2 = True Then  '⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	   
	FncExit = True
		
	Set gActiveElement = document.activeElement

End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.3 Common Group - 3
' Description : This part declares 3rd common function group
'=======================================================================================================
'*******************************************************************************************************


'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbQuery() 
    Dim strVal
   
    DbQuery = False                                                             
    Call LayerShowHide(1)

    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001     '☜: 
			strVal = strVal & "&txtArNo=" & Trim(.htxtArNo.value)    '조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001     '☜: 
			strVal = strVal & "&txtArNo=" & Trim(.txtArNo.value)    '조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    End With

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
	
	Call RunMyBizASP(MyBizASP, strVal)              '☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk()
	Dim  row, strAcctCd, arrVal

	With frm1
		If .vspdData.MaxRows > 0 Then
			Call SetSpreadLock()  
		End if 
		'---I-------------------
		'Reset variables area
		'-----------------------  
		Call ggoOper.LockField(Document, "Q")				'This function lock the suitable field        
		Call SetToolbar(MENU_SIN_UPD) 
		Call InitVariables()
		lgIntFlgMode = parent.OPMD_UMODE							'Indicates that current mode is Update mode
		       
		If .vspdData.MaxRows > 0 Then
		    .vspdData.Row = 1
		    .vspdData.Col = C_ItemSeq 
		    .hItemSeq.Value = .vspdData.Text 
		    Call DbQuery2(1)
		End If
		If Trim(frm1.txtInvNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtInvDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtInvDt, "N")		'N:Required, Q:Protected, D:Default
		End If
		If Trim(frm1.txtBlNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtBlDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtBlDt, "N")		'N:Required, Q:Protected, D:Default
		End If
	
		for row= 0 to .vspdData.MaxRows	'부가세 체크 
			.vspdData.Col = C_AcctCd
			.vspdData.Row = row  
			strAcctCd = Trim(.vspdData.text)
			IF CommonQueryRs( "ACCT_TYPE" , "A_ACCT" ,  " ACCT_CD =  " & FilterVar(strAcctCd , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
				arrVal = Split(lgF0, Chr(11))  
				ggoSpread.Source = frm1.vspdData
				If TRim(arrVal(0))="VR" OR Trim(arrVal(0)) = "VP"Then
					ggoSpread.SpreadunLock C_VatType, Row, C_VatType, Row 
					ggoSpread.SSSetRequired C_VatType, Row, Row '
					ggoSpread.SpreadunLock C_VatPB, Row, C_VatPB, Row 
				End If
			End If
		Next 
	
	End With 

	lgQueryOk = True
	
	Call DoSum()
	Call txtDocCur_OnChange()
	Call txtDeptCd_OnBlur
	
	frm1.txtArNo.focus
	
	lgBlnFlgChgValue = False
	lgQueryOk= False
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function  DbDelete() 
    Dim strVal

    DbDelete = False              

    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtArNo=" & Trim(frm1.txtArNo.value)    '☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
	    
	Call RunMyBizASP(MyBizASP, strVal)          '☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()																'삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "2")									'Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field    
    
    Call InitVariables()																'Initializes local global variables
    Call ClickTab1()
    Call SetDefaultVal()
    
	ggoSpread.Source = frm1.vspdData        				
	ggoSpread.ClearSpreadData						
	ggoSpread.Source = frm1.vspdData2        				
	ggoSpread.ClearSpreadData						
	ggoSpread.Source = frm1.vspdData3        				
	ggoSpread.ClearSpreadData							    
    
    frm1.txtArNo.Value = ""
    frm1.txtArNo.focus
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    Dim lngRows 
    Dim lGrpcnt
    DIM strVal 
    Dim strDel

    DbSave = False                                                          
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   
	Err.Clear 
	
	With frm1
		.txtFlgMode.value = lgIntFlgMode         
	End With
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 
    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep
			    
			    .Col = C_ItemSeq '1
			    strVal = strVal & Trim(.Text) & parent.gColSep
			    .Col = C_AcctCd  '2
			    strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_DeptCd  '3
			    strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_VatType  '4
			    strVal = strVal & Trim(.Text) & parent.gColSep
			    .Col = C_NetAmt  '5
			    strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
			    .Col = C_NetLocAmt '6
			    strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
			    .Col = C_ItemDesc '7
			    strVal = strVal & Trim(.Text) & parent.gRowSep
			          
			    lGrpCnt = lGrpCnt + 1          
			End If
		Next
	End With
 
	frm1.txtMaxRows.value = lGrpCnt-1						'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value =  strDel & strVal					'Spread Sheet 내용을 저장 

	lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    ggoSpread.Source = frm1.vspdData3
    With frm1.vspdData3
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			If .Text <> ggoSpread.DeleteFlag Then  
				strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep
				.Col =  1
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col =  2 
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col =  3
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col =  5
				strVal = strVal & Trim(.Text) & parent.gRowSep

				lGrpCnt = lGrpCnt + 1
			End If
		Next
	End With
 
    frm1.txtMaxRows3.value = lGrpCnt-1							'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread3.value  =  strDel & strVal					'Spread Sheet 내용을 저장 


	'권한관리추가 start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'권한관리추가 end
		
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'저장 비지니스 ASP 를 가동 
    
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk(ByVal ArNo)								'☆: 저장 성공후 실행 로직 
    ggoSpread.Source = frm1.vspdData						'☆: frm1.vspdData move
    ggoSpread.SSDeleteFlag 1
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtArNo.value = ArNo
	End If   
 
	Call ggoOper.ClearField(Document, "2")					'Clear Contents  Field
	frm1.txtArNo.focus
    Call ClickTab1()
    Call InitVariables()									'Initializes local global variables
	ggoSpread.Source = frm1.vspdData        				
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2        				
	ggoSpread.ClearSpreadData						
	ggoSpread.Source = frm1.vspdData3        				
	ggoSpread.ClearSpreadData														
	Call DBquery()     
End Function



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************



'=======================================================================================================
' Function Name : DbQuery2                    
' Function Desc : This function is data query and display            
'=======================================================================================================
Function DbQuery2(ByVal Row)
	Dim strVal 
	Dim lngRows
	 
	Dim strSelect
	Dim strFrom
	Dim strWhere  
 
	Dim strTableid
	Dim strColid
	Dim strColNm 
	Dim strMajorCd 
	Dim strNmwhere
	Dim i,indx1
	Dim arrVal,arrTemp

	'Err.Clear
	With frm1
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq 
		.hItemSeq.Value = .vspdData.Text

		If Trim(.hItemSeq.Value) = "" Then
		    Exit Function
		End If
		 
		If CopyFromData(.hItemSeq.Value) = True Then
		   	Call SetSpread2ColorAr()
		    Exit Function
		End If
		 
		Call LayerShowHide(1)
 
		DbQuery2 = False
  
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
  
		strSelect =    " C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END , " & .hItemSeq.Value & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
		    
		strFrom = " A_CTRL_ITEM A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_OPEN_AR_DTL C (NOLOCK), A_OPEN_AR_ITEM D (NOLOCK) "
  
		strWhere =     " D.AR_NO = " & FilterVar(UCase(.txtARNo.value), "''", "S") 
		strWhere = strWhere & " AND D.ITEM_SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.AR_NO  =  C.AR_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
		 
		frm1.vspdData2.ReDraw = False
		 
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
		    ggoSpread.Source = frm1.vspdData2
		    
		    arrTemp =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp,Chr(12))
  		    ggoSpread.SSShowData lgF2By2       
   
			For lngRows = 1 To frm1.vspdData2.Maxrows
				frm1.vspddata2.Row = lngRows 
				frm1.vspdData2.Col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspdData2.Col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspdData2.Col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspdData2.Col = C_ColNm
					strColNm = frm1.vspddata2.text 
					frm1.vspdData2.Col = C_MajorCd     
					strMajorCd = frm1.vspddata2.text 
						 
					frm1.vspdData2.Col = C_CtrlVal
						 
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text , "''", "S") & " " 
						 
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") & " "
					End If     
						 
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspdData2.Col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
					End If
				End If
				strVal = strVal & Chr(11) & .hItemSeq.Value 

				frm1.vspdData2.Col = C_DtlSeq  
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlCd   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlNm   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlVal 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlPB   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlValNm 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Seq 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Tableid 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Colid 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_ColNm 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Datatype 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_DataLen 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_DRFg 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_MajorCd 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_MajorCd+1 				
				.vspdData2.Text = lngRows
				strVal = strVal & Chr(11) & .vspdData2.Text
				
				strVal = strVal & Chr(11) & Chr(12)         
			Next     
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal 
		End If   
  
'		Call CopyFromData(.hItemSeq.value)
		Call SetSpread2ColorAr()
	End With
 
	Call LayerShowHide(0)
	frm1.vspdData2.ReDraw = True
	DbQuery2 = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk2()
	lgQueryOk= True
	'-----------------------
	'Reset variables area
	'-----------------------
	Call SetSpread2ColorAr()
	Call txtDocCur_OnChange()
	
	lgBlnFlgChgValue = False        
	lgQueryOk= False
End Function

'=======================================================================================================
'   Event Name : GetAcctForVat
'   Event Desc :
'======================================================================================================= 
Sub GetAcctForVat() 
	Dim ii
	lgBlnGetAcctForVat = False
	If CommonQueryRs("acct_cd", "a_acct(nolock)", "acct_type LIKE " & FilterVar("V_", "''", "S") & "  and del_fg <> " & FilterVar("Y", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then   
		lgArrAcctForVat = Split(Mid(lgF0, 1, Len(lgF0) - 1), Chr(11))     
		lgBlnGetAcctForVat = True
	End If
End Sub

'=======================================================================================================
'   Event Name : InputCtrlVal
'   Event Desc :
'=======================================================================================================  
Sub InputCtrlVal(ByVal Row)
	Dim strAcctCd  
	Dim ii
	DIm arrVal
		 
	lgBlnFlgChgValue = True
  
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Col = C_AcctCd
	frm1.vspdData.Row = Row  
	strAcctCd = Trim(frm1.vspdData.text)
	IF CommonQueryRs( "ACCT_TYPE" , "A_ACCT" ,  " ACCT_CD =  " & FilterVar(strAcctCd , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		arrVal = Split(lgF0, Chr(11))  
		If TRim(arrVal(0))="VR" OR Trim(arrVal(0)) = "VP"Then
			ggoSpread.SpreadunLock C_VatType, Row, C_VatType, Row 
			ggoSpread.SSSetRequired C_VatType, Row, Row ' 
			ggoSpread.SpreadunLock C_VatPB, Row, C_VatPB, Row  
		ELSE
			frm1.vspdData.Col = C_VatType
			frm1.vspdData.text=""
			frm1.vspdData.Col = C_VatNm
			frm1.vspdData.text=""
			
			ggoSpread.SSSetProtected C_VatType, Row, Row  
			ggoSpread.SpreadLock C_VatPB, Row, C_VatType, Row  
		End if
		
	End If

	frm1.vspdData.Col = C_deptcd
	frm1.vspdData.Row = Row   
  
	Call AutoInputDetail(strAcctCd, Trim(frm1.vspdData.text), frm1.txtArDt.text, Row)
	
	For ii = 1 To frm1.vspdData2.MaxRows
		frm1.vspdData2.Col = C_CtrlVal
		frm1.vspddata2.Row = ii
			  
		If Trim(frm1.vspddata2.text) <> "" Then
			Call CopyToHSheet2(frm1.vspdData.ActiveRow,ii)       
		End if
	Next
End Sub


'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'=======================================================================================================
Sub DoSum()
	Dim dblTotNetAmt
	Dim dblTotNetLocAmt

	With frm1.vspdData
		dblTotNetAmt = FncSumSheet1(frm1.vspdData,C_NetAmt, 1, .MaxRows, False, -1, -1, "V")
		dblTotNetLocAmt = FncSumSheet1(frm1.vspdData,C_NetLocAmt, 1, .MaxRows, False, -1, -1, "V")
	 
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then                     
			frm1.txtTotNetAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotNetAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End if 
		frm1.txtTotNetLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotNetLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")  
	End With 
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True

    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value, "''", "S") & " " , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
    	If lgQueryOk <> True Then
			If UCase(Trim(frm1.txtDocCur.value)) <> parent.gCurrency Then
				frm1.txtXchRate.Text = 0
			Else
				frm1.txtXchRate.Text = 1
			End If
		End If
				    
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
		Call DoSum()
	End If   

	If lgQueryOk <> True Then
		Call XchLocRate()
	End If 
End Sub
'===================================== CurFormatNumericOCX()  =======================================
' Name : CurFormatNumericOCX()
' Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		 ' 부가세금액 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 순매출액 
		ggoOper.FormatFieldByObjectOfCur .txtNetAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 현금매출액 
		ggoOper.FormatFieldByObjectOfCur .txtCashAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 선수금매출액 
		ggoOper.FormatFieldByObjectOfCur .txtPrRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 매출총액 
		ggoOper.FormatFieldByObjectOfCur .txtArTotAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 외상매출액 
		ggoOper.FormatFieldByObjectOfCur .txtArAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec  
		' 채권잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec    
		' 순매출액(탭)
		ggoOper.FormatFieldByObjectOfCur .txtTotNetAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec    
	End With
End Sub
'===================================== CurFormatNumSprSheet()  ======================================
' Name : CurFormatNumSprSheet()
' Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' 순매출액(탭)
		ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
'		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,1,-1,.txtDocCur.value,C_NetAmt,"A" ,"I","X","X")         		
	End With
End Sub

'====================================================================================================
'	Name : XchLocRate()
'	Description : 환율이 변경되는 Factor 가 변했을 때 수정되는 Local Amt. Setting
'====================================================================================================
Sub XchLocRate()
	Dim ii

	With frm1
		For ii = 1 To .vspdData.MaxRows 
			.vspdData.Row = ii	
			.vspdData.Col = C_NetLocAmt	
			.vspdData.Text = ""    
			ggoSpread.Source = .vspdData
			ggoSpread.UpdateRow ii						
		Next	
		.txtTotNetLocAmt.text="0"
		.txtPrRcptLocAmt.text=""
		.txtCashLocAmt.text=""
	End With
End Sub


'=======================================================================================================
'   Event Name : txtDealBpCd_onChange()
'   Event Desc :  
'=======================================================================================================
Sub  txtDealBpCd_onChange()
    lgBlnFlgChgValue = True
		
	If lgIntFlgMode <> parent.OPMD_UMODE Then 		
		Call CommonQueryRs("A.PARTNER_BP_CD,B.BP_NM", "B_BIZ_PARTNER_FTN A,B_BIZ_PARTNER B", "A.PARTNER_BP_CD = B.BP_CD AND A.PARTNER_FTN  = 'SPA' and DEFAULT_FLAG = " & FilterVar("Y", "''", "S") & "  and A.BP_CD = " & FilterVar(frm1.txtDealBpCd.value, "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) <> "" Then 
			frm1.txtPayBpCd.value = REPLACE(lgF0,Chr(11),"")
			frm1.txtPayBpNm.value = REPLACE(lgF1,Chr(11),"")
		Else
			frm1.txtPayBpCd.value = frm1.txtDealBpCd.value
			frm1.txtPayBpNm.value = frm1.txtDealBpNm.value
		End If
		
		Call CommonQueryRs("A.PARTNER_BP_CD,B.BP_NM", "B_BIZ_PARTNER_FTN A,B_BIZ_PARTNER B", "A.PARTNER_BP_CD = B.BP_CD AND A.PARTNER_FTN  = 'SBI' and DEFAULT_FLAG = " & FilterVar("Y", "''", "S") & " AND A.BP_CD = " & FilterVar(frm1.txtDealBpCd.value, "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) <> "" Then 
			frm1.txtReportBpCd.value = REPLACE(lgF0,Chr(11),"")
			frm1.txtReportBpNm.value = REPLACE(lgF1,Chr(11),"")
		Else
			frm1.txtReportBpCd.value = frm1.txtDealBpCd.value
			frm1.txtReportBpNm.value = frm1.txtDealBpNm.value
		End If
		
	End if
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
	Dim indx

	On Error Resume Next
	Err.Clear 		
	
	ggoSpread.Source = gActiveSpdSheet
	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet()
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadLock()
			Call SetSpread2ColorAr()						
		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
'			Call SetSpread2Lock()			
			Call SetSpread2ColorAr()  
	End Select
	
	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If	
End Sub

'===================================== PrevspdDataRestore()  ========================================
' Name : PrevspdDataRestore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdDataRestore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData.MaxRows
        frm1.vspdData.Row    = indx
        frm1.vspdData.Col    = 0
		
		If frm1.vspdData.Text <> "" Then
			Select Case frm1.vspdData.Text			
				Case ggoSpread.InsertFlag					
					frm1.vspdData.Col = C_ItemSeq					
					Call DeleteHsheet(frm1.vspdData.Text)					
				Case ggoSpread.UpdateFlag		
					For indx1 = 0 To frm1.vspdData3.MaxRows					
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case frm1.vspdData3.Text 
							Case ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1					
								If UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)										
									Call FncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtArNo.Value)
								End If
						End Select
					Next
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtArNo.Value)
			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'===================================== PrevspdDataRestore()  ========================================
' Name : PrevspdData2Restore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdData2Restore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 to frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData	
					        ggoSpread.EditUndo							
						End If
					Next
				Case ggoSpread.UpdateFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.txtArNo.Value) 
						End If
					Next
				Case ggoSpread.DeleteFlag

			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

'========================================================================================================
' Name : fncRestoreDbQuery2																				
' Desc : This function is data query and display												
'========================================================================================================
Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal

	On Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	With frm1
		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text

	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =    " C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END , " & strItemSeq & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
		    
		strFrom = " A_CTRL_ITEM A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_OPEN_AR_DTL C (NOLOCK), A_OPEN_AR_ITEM D (NOLOCK) "
  
		strWhere =     " D.AR_NO =  " & FilterVar(UCase(.txtARNo.value), "''", "S") & "  "
		strWhere = strWhere & " AND D.ITEM_SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.AR_NO  =  C.AR_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
				
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)
					
					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData (strItemSeq)
		End If

		Call LayerShowHide(0)
		Call RestoreToolBar()
	End With

	If Err.number = 0 Then
		fncRestoreDbQuery2 = True
	End If
End Function






'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************


'=======================================================================================================
'   Event Name : vspdData_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData_onfocus()
    If lgIntFlgMode <> parent.OPMD_UMODE Then           
        Call SetToolbar(MENU_MUL_CRT)                                     '버튼 툴바 제어 
    Else        
        Call SetToolbar(MENU_MUL_UPD)                                     '버튼 툴바 제어 
    End If    
End Sub

'=======================================================================================================
'   Event Name : 
'   Event Desc :
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
			.vspdData.Row = NewRow
			.vspdData.Col = C_ItemSeq
			.hItemSeq.value = .vspdData.Text
			ggoSpread.Source = .vspdData2        				
			ggoSpread.ClearSpreadData						
        End With
		frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub       
		End if
        lgCurrRow = NewRow

        Call DbQuery2(lgCurrRow)
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"							'Split 상태코드 
 
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 then
	    Exit Sub
	End if

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

	frm1.vspddata.Row = frm1.vspddata.ActiveRow 
	frm1.vspdData.Col = C_AcctCd
 
	If Len(frm1.vspdData.Text) > 0 Then

	Else
		ggoSpread.Source = frm1.vspdData2        				
		ggoSpread.ClearSpreadData						
	End if 
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
'   Event Name :vspddata_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata_DblClick(ByVal Col,ByVal Row)
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
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

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim tmpDrCrFg
	Dim ii
	Dim iChkAcctForVat
 
	 '---------- Coding part -------------------------------------------------------------
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
 
	With frm1
		.vspddata.Row = Row
		Select Case Col
			Case C_VatNm
		   '부가세 타입 선택시 acct_cd가 부가세 관련인지 check하여 선택하거나 하지 못하게 한다.      
			If lgBlnGetAcctForVat = True Then    
				frm1.vspdData.Col = C_AcctCd
				iChkAcctForVat = False
				For ii = 0 To Ubound(lgArrAcctForVat,1)
					If Trim(frm1.vspdData.Text) = Trim(lgArrAcctForVat(ii)) Then
						iChkAcctForVat = True       
						Exit For
					End If
				Next
				If iChkAcctForVat = False  Then
					frm1.vspdData.Col = C_VatNm
					frm1.vspdData.Text = ""
				End If
			End If 
			.vspddata.Col = Col            
			intIndex = .vspddata.Value
			.vspddata.Col = C_VatType    
			.vspddata.Value = intIndex  
			Call InputCtrlVal(Row)'
		End Select
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		      
		If Row > 0 And Col = C_AcctPB Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenPopup(.Text, 2)
		End If    
		       
		If Row > 0 And Col = C_DeptPB Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenDept(.Text, 1) 
		End If    
		       
		If Row > 0 And Col = C_VatPB Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenPopUp(.Text, 12)
		End If    
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
	Dim iLen
	Dim sPreCtrlVal
	Dim IntRetCD
 	Dim intIndex
 	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0    

    Select Case Col
		Case C_AcctCD      
			If  frm1.vspdData.Text = ggoSpread.InsertFlag Then
 				frm1.vspdData.Col = C_ItemSeq
				frm1.hItemSeq.value = frm1.vspdData.Text
				frm1.vspdData.Col = C_AcctCd
 				If Len(frm1.vspdData.Text) > 0 Then
					frm1.vspdData.Row = Row
					frm1.vspdData.Col = C_itemSeq    
					Call DeleteHsheet(frm1.vspdData.Text)
				    Call DbQuery3 (Row)
				    Call InputCtrlVal(Row)      
				    Call SetSpread2ColorAr()
				End If    
			End If  
		Case C_DeptCd
			frm1.vspdData.Col = C_DeptCd	
			Call deptCd_underChange(frm1.vspdData.Text)	
		Case C_NetAmt
			frm1.vspdData.Col = C_NetLocAmt	
			frm1.vspdData.Text = ""   			
			Call DoSum()
		Case C_NetLocAmt
			Call DoSum()
		Case C_VatNm, C_VatType
'				frm1.vspddata.Col = Col		       			
'			    intIndex = frm1.vspddata.Value
'				frm1.vspddata.Col = C_VatType
'				frm1.vspddata.Value = intIndex		
			    Call InputCtrlVal(Row)'  			
    End Select      
End Sub

'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData_EditChange(ByVal Col , ByVal Row )
                
End Sub

'======================================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspddata_KeyPress(KeyAscii )
     
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************



'=======================================================================================================
'   Event Name : txtArDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtArDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtArDt.Action = 7    
		Call SetFocusToDocument("M")
		Frm1.txtArDt.Focus 
        Call txtArDt_onBlur()    
    End If
End Sub

'=======================================================================================================
'   Event Name : txTblDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txTblDt_DblClick(Button)
    If Button = 1 Then
        frm1.txTblDt.Action = 7   
        Call SetFocusToDocument("M")
		Frm1.txTblDt.Focus 
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7    
        Call SetFocusToDocument("M")
		Frm1.txtDueDt.Focus 
	End If
End Sub

'=======================================================================================================
'   Event Name : txtInvDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtInvDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInvDt.Action = 7    
        Call SetFocusToDocument("M")
		Frm1.txtInvDt.Focus 
    End If
End Sub


'=======================================================================================================
'   Event Name : txtCashAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtCashAmt_Change()
	lgBlnFlgChgValue = True
	If lgQueryOk <> True Then
		frm1.txtCashLocAmt.text = "0"
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtCashLocAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtCashLocAmt_Change()
 lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPrRcptAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtPrRcptAmt_Change()
	lgBlnFlgChgValue = True
 
	If Len(frm1.txtPrRcptAmt.Text) = 0 then
		Call ggoOper.SetReqAttr(frm1.txtPrPaymNo, "D")
	Else
		If UNICDbl(frm1.txtPrRcptAmt.Text) <> 0 then 
			Call ggoOper.SetReqAttr(frm1.txtPrPaymNo, "N")
		Else
			Call ggoOper.SetReqAttr(frm1.txtPrPaymNo, "D")
		End If
	End If
	
	If lgQueryOk <> True Then
		frm1.txtPrRcptLocAmt.text = "0"
	End If	
End Sub

'=======================================================================================================
'   Event Name : txtPrRcptLocAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtPrRcptLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtVatAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtVatAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtVatLocAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtVatLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPayDur_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub txtPayDur_Change()
	lgBlnFlgChgValue = True
End Sub


'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.8 HTML Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************
'==========================================================================================
'   Event Name : DeptCd_underChange(Byval strCode)
'   Event Desc : 
'==========================================================================================
Sub DeptCd_underChange(Byval strCode)
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 

    If Trim(frm1.txtArDt.Text = "") Then    
		Exit sub
    End If
    
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtArDt.Text, gDateFormat,""), "''", "S") & "))"			

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  

		frm1.vspdData.Col = C_deptcd			
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.text = ""
		frm1.vspdData.Col = C_deptnm		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = ""
	End If 
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtDeptCd_OnBlur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtArDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtArDt.Text, gDateFormat,""), "''", "S") & "))"			
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 To jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
		End If
	End if		
		'----------------------------------------------------------------------------------------
	
End Sub

'==========================================================================================
'   Event Name : txtArDt_onBlur
'   Event Desc : 
'==========================================================================================
Sub txtArDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
   If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
	
				If LTrim(RTrim(.txtDeptCd.value)) <> "" And Trim(.txtArDt.Text <> "") Then
					'----------------------------------------------------------------------------------------
						strSelect	=			 " Distinct org_change_id "    		
						strFrom		=			 " b_acct_dept(NOLOCK) "		
						strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
						strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
						strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
						strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtArDt.Text, gDateFormat,""), "''", "S") & "))"			
	
					IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
					If IntRetCD = False  Or Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
						.txtDeptCd.value = ""
						.txtDeptNm.value = ""
						.hOrgChangeId.value = ""
						.txtDeptCd.focus
					End If
				End If
			End With
		'----------------------------------------------------------------------------------------
		End If
	End IF
	
	Call XchLocRate()
End Sub

'======================================================================================================
'송장번호 입력시 송장일자 입력필수 
'======================================================================================================
Sub txtInvNo_OnBlur()
	If Trim(frm1.txtInvNo.value) = "" Then
		Call ggoOper.SetReqAttr(frm1.txtInvDt, "D")
	Else
		Call ggoOper.SetReqAttr(frm1.txtInvDt, "N") 'N:Required, Q:Protected, D:Default
	End If
End Sub

'======================================================================================================
'선하증권번호 입력시 선하증권일자 입력필수 
'======================================================================================================
Sub txtBlNo_OnBlur()
	If Trim(frm1.txtBlNo.value) = "" Then
		Call ggoOper.SetReqAttr(frm1.txtBlDt, "D")
	Else
		Call ggoOper.SetReqAttr(frm1.txtBlDt, "N") 'N:Required, Q:Protected, D:Default
	End If
End Sub

Sub txtPrPaymNo_OnBlur()
	If frm1.txtPrPaymNo.value = "" Then     
		Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "N")  
		Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")  
	End If
End Sub
Sub txtDueDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtInvDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txTblDt_Change()
	lgBlnFlgChgValue = True
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>
<!--'======================================================================================================
'            6. Tag부 
' 기능: Tag부분 설정 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL=NO>
<FORM NAME=frm1 TARGET=MyBizASP METHOD=POST>
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS=CLSMTABP>
						<TABLE ID=MyTab CELLSPACING=0 CELLPADDING=0 onClick="ClickTab1()">
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG HEIGHT=23 SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH=9></TD>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=RIGHT><IMG HEIGHT=23 SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH=10></TD>
							</TR>
						</TABLE>
					</TD>
					<TD CLASS=CLSMTABP>
						<TABLE ID=MyTab CELLSPACING=0 CELLPADDING=0 onClick="ClickTab2()">
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/tab_up_bg.gif"><IMG HEIGHT=23 SRC="../../../CShared/image/table/tab_up_left.gif" WIDTH=9></TD>
								<TD BACKGROUND="../../../CShared/image/table/tab_up_bg.gif" ALIGN=CENTER CLASS=CLSMTABP><FONT COLOR=WHITE>상세채권정보</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/tab_up_bg.gif" ALIGN=RIGHT><IMG HEIGHT=23 SRC="../../../CShared/image/table/tab_up_right.gif" WIDTH=10></TD>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS=Tab11>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS=CLSFLD>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>채권번호</TD>
									<TD CLASS=TD656 NOWRAP><INPUT NAME=txtArNo ALT="채권번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU"><IMG ALIGN=TOP NAME=btnCalType SRC="../../../CShared/image/btnPopup.gif"  TYPE=BUTTON onClick="vbscript:CALL OpenPopUp(frm1.txtArNo.Value, 0)"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD WIDTH="100%">
						
						
						<DIV ID=TabDiv STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL=NO>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>주문처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtDealBpCd ALT="주문처" MAXLENGTH=10 SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG ALIGN=TOP NAME=btnCalType onClick="vbscript:CALL OpenBp(frm1.txtDealBpCd.Value, 3)" SRC="../../../CShared/image/btnPopup.gif"  TYPE=BUTTON>
										<INPUT NAME="txtDealBpNm" ALT="주문처" SIZE="20" tag="24" ></TD>
									<TD CLASS=TD5 NOWRAP>세금계산서발행처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtReportBpCd ALT="세금계산서발행처" MAXLENGTH=10 SIZE=10 STYLE="TEXT-ALIGN: Left" tag="21NXXU"><IMG ALIGN=TOP NAME=btnCalType onClick="vbscript:CALL OpenBp(frm1.txtReportBpCd.Value, 9)" SRC="../../../CShared/image/btnPopup.gif"  TYPE=BUTTON>
										<INPUT  NAME="txtReportbpnm"  ALT="세금계산서발행처" SIZE="18" tag = "24" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>수금처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtPayBpCd ALT="수금처" MAXLENGTH=10 SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG ALIGN=TOP NAME=btnCalType onClick="vbscript:CALL OpenBp(frm1.txtPayBpCd.Value, 4)" src="../../../CShared/image/btnPopup.gif"  TYPE=BUTTON>
										<INPUT  NAME="txtpaybpnm"  ALT="수금처" SIZE="20" tag = "24" ></TD>
									<TD CLASS=TD5 NOWRAP>송장번호|송장일</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvNo" ALT="송장번호" MAXLENGTH="50" SIZE=18 STYLE="TEXT-ALIGN: left" tag="2X">&nbsp;
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ALIGN="TOP" NAME="txtInvDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="송장일" id=OBJECT3></OBJECT>');</SCRIPT>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>채권일자</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME=txtArDt CLASS=FPDTYYYYMMDD tag="22X1" Title=FPDATETIME ALT="채권일자" ID=OBJECT1></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>선하증권번호|선하증권일</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txTblNo" ALT="선하증권번호" MAXLENGTH="35" SIZE=18 tag="2X" >&nbsp;
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ALIGN="TOP" NAME="txTblDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="선하증권일" id=OBJECT4></OBJECT>');</SCRIPT>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>채권만기일자</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME=txtDueDt CLASS=FPDTYYYYMMDD tag="22X1" Title=FPDATETIME ALT="채권만기일자" ID=OBJECT2></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" nowrap>결제기간</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table CELLPADDING=0 CELLSPACING=0>
											<TR>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="결제기간" NAME=txtPayDur CLASSID=<%=gCLSIDFPDS%> ID=fpDoubleSingle5 STYLE="HEIGHT: 20px; WIDTH: 80px" tag="21X70" Title=FPDOUBLESINGLE> <PARAM NAME=MaxValue VALUE="30000"> <PARAM NAME=MinValue VALUE="-30000"></OBJECT>');</SCRIPT></TD>
												<TD NOWRAP>&nbsp;일</TD>
											</TR>
										</Table>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부서</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtDeptCd ALT="부서" MAXLENGTH=10 SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG ALIGN=TOP NAME=btnCalType onClick="vbscript:CALL OpenDept(frm1.txtDeptCd.Value, 0)" SRC="../../../CShared/image/btnPopup.gif"  TYPE=BUTTON>
										<INPUT NAME="txtDeptNm" ALT="부서" SIZE="20" tag ="24" ></TD>
									<TD CLASS=TD5 nowrap>결제방법</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME=txtPayMethCd ALT="결제방법" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnPayMethod ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtPayMethCd.value, 10)">
										<INPUT TYPE=TEXT NAME=txtPayMethNm ALT="결제방법" SIZE="18" style="HEIGHT: 20px; " tag="24X" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계정코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtAcctCd ALT="계정코드" MAXLENGTH=20 SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG ALIGN=TOP NAME=btnCalType onClick="vbscript:CALL OpenPopUp(frm1.txtAcctCd.value,1)" src="../../../CShared/image/btnPopup.gif"  TYPE=BUTTON>
										<INPUT NAME=txtAcctnm ALT="계정코드명" MAXLENGTH=20  tag  ="24"></TD>
									<TD CLASS=TD5 NOWRAP>입금유형</TD>
									<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME=txtPayTypeCd ALT="입금유형" STYLE="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" STYLE="HEIGHT: 21px; WIDTH: 16px" NAME=btnPayMethod ALIGN=TOP TYPE=BUTTON onClick="vbscript:Call OpenPopup(frm1.txtPayTypeCd.value, 11)">
										<INPUT TYPE=TEXT NAME=txtPayTypeNm ALT="입금유형" SIZE="18" style="HEIGHT: 20px; " tag="24X" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>거래통화</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtDocCur ALT="거래통화" MAXLENGTH=3 SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="22XXXU"><IMG ALIGN=TOP NAME=btnCalType onClick="vbscript:CALL OpenPopUp(frm1.txtDocCur.Value,8)" src="../../../CShared/image/btnPopup.gif"  TYPE=BUTTON></TD>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtXchRate CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="환율" tag="24X5Z" ID=OBJECT5></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>선수금번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtPrPaymNo ALT="선수금번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="21XXXU" ><IMG ALIGN=TOP NAME=btnCalType onClick="vbscript:CALL OpenRefPreRcptNo()" src="../../../CShared/image/btnPopup.gif"  TYPE=BUTTON></TD>
									<TD CLASS=TD5 NOWRAP>대금결제참조</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtPaymTerms ALT="대금결제참조" MAXLENGTH=120 SIZE=30 STYLE="TEXT-ALIGN: left" tag ="21"></TD>									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>선수금매출액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtPrRcptAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="선수금매출액" tag="21X2Z" id=OBJECT7></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>선수금매출액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtPrRcptLocAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="선수금매출액(자국통화)" tag="21X2Z" id=OBJECT13></OBJECT>');</SCRIPT></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>현금매출액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtCashAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="현금매출액" tag="21X2Z" id=OBJECT5></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>현금매출액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtCashLocAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="현금매출액(자국통화)" tag="21X2Z" id=OBJECT14></OBJECT>');</SCRIPT></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>외상매출액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtArAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="외상매출액" tag="24X2" id=OBJECT9></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>외상매출액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtArLocAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="외상매출액(자국통화)" tag="24X2" id=OBJECT15></OBJECT>');</SCRIPT></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>채권잔액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtBalAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="반제금액" tag="24X2" id=OBJECT10></OBJECT>');</SCRIPT></TD>

									<TD CLASS=TD5 NOWRAP>채권잔액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtBalLocAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="반제금액(자국통화)" tag="24X2" id=OBJECT16></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>매출총액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtArTotAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="매출총액" tag="24X2" id=OBJECT8></OBJECT>');</SCRIPT></TD>

									<TD CLASS=TD5 NOWRAP>매출총액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtArTotLocAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="매출총액(자국통화)" tag="24X2" id=OBJECT17></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부가세금액</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtVatAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="부가세금액" tag="24X2"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>부가세금액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtVatLocAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="부가세금액(자국통화)" tag="24X2"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>순매출액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtNetAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="순매출액" tag="24X2"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>순매출액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtNetLocAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="순매출액(자국통화)" tag="24X2"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTempGLNo" ALT="결의전표번호" SIZE=19 MAXLENGTH=18  STYLE="TEXT-ALIGN: Left" tag="24XXXU"></TD>
									<TD CLASS=TD5 NOWRAP>전표번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtGlNo ALT="회계전표번호" SIZE=19 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>비고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtDesc ALT="비고" MAXLENGTH=128 SIZE=60 tag="2X"></TD>
									<TD CLASS=TD5 NOWRAP>프로젝트</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="프로젝트" MAXLENGTH=25 SIZE=25 tag="2X"></TD>
								</TR>
							</TABLE>
						</DIV>
						
						
						<DIV ID=TabDiv SCROLL=NO>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR HEIGHT="60%">
									<TD WIDTH="100%" COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" NAME=vspdData WIDTH="100%" tag="2" TITLE=SPREAD id=OBJECT11> <PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD COLSPAN=4>
										<TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD class=TD5 NOWRAP>순매출액</TD>
												<TD class=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtTotNetAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="순매출액" tag="24X2" ID=OBJECT22></OBJECT>');</SCRIPT></TD>
												<TD class=TD5 NOWRAP>순매출액(자국)</TD>
												<TD class=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtTotNetLocAmt CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="순매출액(자국)" tag="24X2" ID=OBJECT22></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR HEIGHT="40%">
									<TD WIDTH="100%" COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" NAME=vspdData2 width="100%" tag="2" TITLE=SPREAD id=OBJECT12> <PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0></OBJECT>');</SCRIPT></TD>

								</TR>
							</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME=MyBizASP WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TYPE=HIDDEN CLASSID=<%=gCLSIDFPSPD%> HEIGHT=100 NAME=vspdData3 WIDTH="100%" tag="2" ID=OBJECT6 TABINDEX="-1"> <PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0></OBJECT>');</SCRIPT>

<TEXTAREA	CLASS=HIDDEN NAME=txtSpread			tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA	CLASS=HIDDEN NAME=txtSpread3		tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT		TYPE=HIDDEN  NAME=txtMode			tag="24" TABINDEX="-1">
<INPUT		TYPE=HIDDEN  NAME=txtUpdtUserId		tag="24" TABINDEX="-1">
<INPUT		TYPE=HIDDEN  NAME=txtInsrtUserId	tag="24" TABINDEX="-1">
<INPUT		TYPE=HIDDEN  NAME=txtMaxRows		tag="24" TABINDEX="-1">
<INPUT		TYPE=HIDDEN  NAME=txtFlgMode		tag="24" TABINDEX="-1">
<INPUT		TYPE=HIDDEN  NAME=htxtArNo			tag="24" TABINDEX="-1">
<INPUT		TYPE=HIDDEN  NAME=hItemSeq			tag="24" TABINDEX="-1">
<INPUT		TYPE=HIDDEN  NAME=hAcctCd			tag="24" TABINDEX="-1">
<INPUT		TYPE=HIDDEN  NAME=txtMaxRows3		tag="24" TABINDEX="-1">
<INPUT		TYPE=hidden	 NAME="hOrgChangeId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID=MousePT NAME=MousePT>
<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
