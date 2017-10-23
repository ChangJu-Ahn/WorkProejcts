
<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Account Receivable
'*  3. Program ID           : a4101ma.asp
'*  4. Program Name         : 채무등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : ap0011
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
'* 13. History              :7
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">           </SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																				'☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
' .Constant는 반드시 대문자 표기.
' .변수 표준에 따름. prefix로 g를 사용함.
' .Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'@PGM_ID
Const BIZ_PGM_QRY_ID  = "a4101mb1.asp"										'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a4101mb2.asp"										'☆: Save 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID  = "a4101mb3.asp"										'☆: Delete 비지니스 로직 ASP명 

Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"						'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1																'☜: Tab의 위치 
Const TAB2 = 2

'@Grid_Column
Dim C_ItemSeq 																'☆: Spread Sheet 의 Columns 인덱스 
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

'@Grid_Row


Const MENU_NEW  = "1110100000011111"
Const MENU_SIN_CRT = "1110100000011111"       
Const MENU_MUL_CRT = "1110111100111111"
Const MENU_SIN_UPD = "1111100000011111"
Const MENU_MUL_UPD = "1111111100111111"

Dim  lgStrPrevKey1
Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3

'Dim  intItemCnt     
Dim  IsOpenPop																'Popup
Dim  gSelframeFlg

Dim  lgCurrRow
Dim  lgReportPopUp
Dim  lgArrAcctForVat
Dim  lgBlnGetAcctForVat
Dim  lgFormLoad
Dim  lgQueryOk			' Queryok여부 (loc_amt =0 check)
Dim  lgstartfnc

<%
Dim dtToday
dtToday = GetSvrDate
%>

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


'======================================================================================================
'            2. Function부 
'
' 내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
' 공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'               2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'=======================================================================================================




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
End Sub

'======================================================================================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE										'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False												'Indicates that no value changed
    lgIntGrpCount = 0														'initializes Group View Size
    lgStrPrevKey = 0														'initializes Previous Key
    lgStrPrevKeyDtl = 0														'initializes Previous Key
    lgLngCurRows = 0														'initializes Deleted Rows Count
    
   	lgstartfnc = False
	lgFormLoad = True
	lgQueryOk  = False    
End Sub

'======================================================================================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtApDt.text =  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDueDt.text =  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDocCur.value = parent.gCurrency
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	frm1.txtXchRate.text = 1

	Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")
	Call ggoOper.SetReqAttr(frm1.txtPayBpCd,   "N") 
	lgBlnFlgChgValue = False											'Indicates that no value changed 
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
'   Function Name : InitData()
'   Function Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub  InitData()

End Sub

'======================================================================================================
' Function Name : InitSpreadSheet()
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet()
    Call initSpreadPosVariables()
        
    With frm1.vspdData
    
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadInit "V20021127",,parent.gAllowDragDropSpread 

		.Redraw = False		
		.MaxCols = C_ItemDesc + 1										'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols													'공통콘트롤 사용 Hidden Column
		.ColHidden = True
		.MaxRows = 0		    
  
		Call AppendNumberPlace("6","3","0")

		Call GetSpreadColumnPos("A")
				
		ggoSpread.SSSetFloat  C_ItemSeq  , "NO"            ,  6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"
		ggoSpread.SSSetEdit   C_AcctCd   , "계정코드"      , 20, ,,20,2
		ggoSpread.SSSetButton C_AcctPB
		ggoSpread.SSSetEdit   C_AcctNm   , "계정코드명"    , 20,3
		ggoSpread.SSSetEdit   C_DeptCd   , "부서"          , 17, ,,10,2
		ggoSpread.SSSetButton C_DeptPB
		ggoSpread.SSSetEdit   C_DeptNm   , "부서명"        , 20,3    
		ggoSpread.SSSetEdit   C_VatType  , "부가세"        , 10,3,,,2        
		ggoSpread.SSSetButton C_VatPB
		ggoSpread.SSSetEdit   C_VatNm    , "부가세유형명"  , 12,3        
		ggoSpread.SSSetFloat  C_NetAmt   , "순매입액"      , 15, "A"  , ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloat  C_NetLocAmt, "순매입액(자국)", 15, parent.ggAmtOfMoneyNo  , ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec		
		ggoSpread.SSSetEdit   C_ItemDesc , "비고"          , 50, ,,128
				
		Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPB)
		Call ggoSpread.MakePairsColumn(C_DeptCd,C_DeptPB)		
		Call ggoSpread.MakePairsColumn(C_VatType,C_VatPB)
		
		.Redraw = True 
    End With

    Call SetSpreadLock() 
End Sub

'======================================================================================================
' Function Name : SetSpreadLock()
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock()
    Dim objSpread

    With frm1
		ggoSpread.Source = .vspdData
		Set objSpread = .vspdData

		objSpread.Redraw = False
		    
		ggoSpread.SpreadLock C_ItemSeq, -1, C_ItemSeq, -1
		ggoSpread.SpreadLock C_AcctCd, -1, C_AcctCd, -1
		ggoSpread.SpreadLock C_AcctPB, -1, C_AcctPB, -1
		ggoSpread.SpreadLock C_AcctNm, -1, C_AcctNm, -1
		ggoSpread.SpreadLock C_DeptNm, -1, C_DeptNm, -1                            
		ggoSpread.SpreadLock C_VatType, -1, C_VatNm, -1                            
		ggoSpread.SSSetRequired  C_DeptCd, -1, -1 
		ggoSpread.SSSetRequired  C_NetAmt, -1, -1

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
		ggoSpread.SSSetRequired C_AcctCd, pvStartRow, pvEndRow				
 		ggoSpread.SSSetProtected C_AcctNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_DeptCd, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_DeptNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_NetAmt, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_VatNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VatType  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VatPB  , pvStartRow, pvEndRow
		   	
		.Col = 1
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True
		.Redraw = True
    End With
End Sub
'======================================================================================================
' Function Name : SetSpread2Color
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpread2ColorAp()
	Dim i

    With frm1
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False  
		For i = 1 to .vspddata2.maxrows
			ggoSpread.SSSetProtected C_DtlSeq   , i, i
			ggoSpread.SSSetProtected C_CtrlCd   , i, i
			ggoSpread.SSSetProtected C_CtrlNm   , i, i
			ggoSpread.SSSetProtected C_CtrlValNm, i, i
			.vspddata2.row = i
			.vspddata2.col = C_DrFg

			If (.vspddata2.text = "Y") Or (.vspddata2.text = "D") Or (.vspddata2.text = "DC") then
				ggoSpread.SSSetRequired C_CtrlVal, i, i
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
	End Select
End Sub

'======================================================================================================
' Name : OpenPopupGL()
' Description : 회계전표POP-UP
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8) 
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
	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
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
	Dim IntRetCD
	
	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 
	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName,  Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

 '========================================================================================================
' Name : Open???()
' Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefPrePaymNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4107RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4107RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPayBpCd.value														'검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtpaybpnm.value   
	arrParam(2) = frm1.txtDocCur.value     
	arrParam(3) = frm1.txtApDt.text			
	arrParam(4) = ""
	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
 
	If arrRet(0) = "" Then  
		Exit Function
	Else  
		Call SetRefPrePaymNo(arrRet)
	End If
End Function

'========================================================================================================
' Name : SetRefPrePaymNo()
' Description : OpenAp Popup에서 Return되는 값 setting
'========================================================================================================
Function  SetRefPrePaymNo(Byval arrRet)
	lgBlnFlgChgValue = True
	
	With frm1
		.txtPrPaymNo.Value    = arrRet(0)													'C_PpNo = 1   
		.txtPayBpCd.Value     = arrRet(3)													'C_DeptCd = 6 
		.txtpaybpnm.Value     = arrRet(4)													'C_DeptNm = 7 
		.txtDocCur.value      = arrRet(8)													'C_DocCur = 9  
		.txtPrPaymAmt.Text    = arrRet(11)
		.txtPrPaymLocAmt.Text = arrRet(12)

		If UCase(Trim(frm1.txtDocCur.value)) <> parent.gCurrency Then
			frm1.txtXchRate.Text = 0
		Else
			frm1.txtXchRate.Text = 1
		End If	
				
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
		
		If .txtPayBpCd.value <> "" Then     
			Call ggoOper.SetReqAttr(.txtPayBpCd,   "Q")  
		Else   
			Call ggoOper.SetReqAttr(.txtPayBpCd,   "N")  
		End If
	 
		If frm1.txtPayBpCd.value <> "" Then     
			Call ggoOper.SetReqAttr(.txtDocCur,   "Q")  
		Else   
			Call ggoOper.SetReqAttr(.txtDocCur,   "N")  
		End If      
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

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtApDt.Text
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
		Call EscDept(iwhere)
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function EscDept(Byval iWhere)
	With frm1
		Select Case iWhere
		     Case "0"
				.txtDeptCd.focus
			Case "2"
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
				.txtDeptCD.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtApDt.text = arrRet(3)
				call txtDeptCd_Onblur()  
				.txtDeptCd.focus

			Case "2"
			    .vspdData.Col = C_DeptCd
			    .vspdData.Text = arrRet(0)
			    .vspdData.Col = C_DeptNm
			    .vspdData.Text = arrRet(1)     
			    Call vspdData_Change(.vspdData.Col, .vspdData.Row)   
				Call SetActiveCell(.vspdData,C_deptcd,.vspdData.ActiveRow ,"M","X","X")	
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
   	arrParam(1) = ""									' 채권과 연계(거래처 유무)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "S"									'B :매출 S: 매입 T: 전체 
	Select Case iWhere
		Case 3
			arrParam(5) = "SUP"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
		Case 9
			arrParam(5) = "INV"		
		Case 4
			arrParam(5) = "PAYTO"		
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
	Dim iCalledAspName
	Dim IntRetCD
	

	If IsOpenPop = True Then Exit Function 
 
	Select Case iWhere
		Case 0

		Case 1
			arrParam(0) = "계정코드팝업"												' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"						' TABLE 명칭 
			arrParam(2) = Trim(strCode)														' Code Condition
			arrParam(3) = ""																' Name Condition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD " & _ 
			    "and C.trans_type = " & FilterVar("AP003", "''", "S") & "  and C.jnl_cd = " & FilterVar("AP", "''", "S") & " "							' Where Condition
			arrParam(5) = "계정코드"													' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"														' Field명(0)
			arrField(1) = "A.Acct_NM"														' Field명(1)
			arrField(2) = "B.GP_CD"															' Field명(2)
			arrField(3) = "B.GP_NM"															' Field명(3)
		 
			arrHeader(0) = "계정코드"													' Header명(0)
			arrHeader(1) = "계정코드명"													' Header명(1)
			arrHeader(2) = "그룹코드"													' Header명(2)
			arrHeader(3) = "그룹명"														' Header명(3) 
		Case 2
			arrParam(0) = "계정코드팝업"												' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B"											' TABLE 명칭 
			arrParam(2) = Trim(strCode)														' Code Condition
			arrParam(3) = ""																' Name Cindition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & " "								' Where Condition
			arrParam(5) = "계정코드"													' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"														' Field명(0)
			arrField(1) = "A.Acct_NM"														' Field명(1)
			arrField(2) = "B.GP_CD"															' Field명(2)
			arrField(3) = "B.GP_NM"															' Field명(3)
		 
			arrHeader(0) = "계정코드"													' Header명(0)
			arrHeader(1) = "계정코드명"													' Header명(1)
			arrHeader(2) = "그룹코드"													' Header명(2)
			arrHeader(3) = "그룹명"   
		Case 3
			arrParam(0) = "공급처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("C", "''", "S") & " "										' Where Condition
			arrParam(5) = "공급처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    
			arrHeader(0) = "공급처"							' Header명(0)
			arrHeader(1) = "공급처명"						' Header명(1)
		Case 4
			If UCase(frm1.txtPayBpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function			
			arrParam(0) = "지급처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("C", "''", "S") & " "										' Where Condition
			arrParam(5) = "지급처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    
			arrHeader(0) = "지급처"							' Header명(0)
			arrHeader(1) = "지급처명"						' Header명(1)
		Case 5       
			arrParam(0) = "사업장팝업"													' 팝업 명칭 
			arrParam(1) = "B_Biz_AREA"														' TABLE 명칭 
			arrParam(2) = strCode															' Code Condition
			arrParam(3) = ""																' Name Cindition
			arrParam(4) = ""																' Where Condition
			arrParam(5) = "사업장"   
 
			arrField(0) = "Biz_AREA_CD"														' Field명(0)
			arrField(1) = "Biz_AREA_NM"														' Field명(1)    
			 
			arrHeader(0) = "사업장"														' Header명(0)
			arrHeader(1) = "사업장명"													' Header명(1)
		Case 8
			If UCase(frm1.txtDocCur.className) = UCase(parent.UCN_PROTECTED) Then Exit Function 
			arrParam(0) = "거래통화팝업"												' 팝업 명칭 
			arrParam(1) = "b_currency"														' TABLE 명칭 
			arrParam(2) = strCode															' Code Condition
			arrParam(3) = ""																' Name Condition
			arrParam(4) = ""																' Where Condition
			arrParam(5) = "거래통화"    
 
			arrField(0) = "CURRENCY"														' Field명(0)
			arrField(1) = "CURRENCY_DESC"													' Field명(1)
			 
			arrHeader(0) = "거래통화"													' Header명(0)
			arrHeader(1) = "거래통화명"													' Header명(1)    

		Case 9
			arrParam(0) = "세금계산서발행처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("C", "''", "S") & " "									' Where Condition
			arrParam(5) = "세금계산서발행처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    
			arrHeader(0) = "세금계산서발행처"							' Header명(0)
			arrHeader(1) = "세금계산서발행처명"						' Header명(1)		
		Case 10
			If  UCase(frm1.txtPayMethCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		 
			arrHeader(0) = "결제방법"													' Header명(0)
			arrHeader(1) = "결제방법명"													' Header명(1)
			arrHeader(2) = "Reference"
			 
			arrField(0) = "B_Minor.MINOR_CD"												' Field명(0)
			arrField(1) = "B_Minor.MINOR_NM"												' Field명(1)
			arrField(2) = "b_configuration.REFERENCE"
			 
			arrParam(0) = "결제방법"													' 팝업 명칭 
			arrParam(1) = "B_Minor,b_configuration"											' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPayMethCd.Value)										' Code Condition
			arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9004", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
			              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"  
			arrParam(5) = "결제방법"													' TextBox 명칭 
		Case 11
			If UCase(frm1.txtPayTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
 
			arrParam(0) = "지급유형"													' 팝업 명칭 
		 
			If Trim(frm1.txtPayMethCd.Value) = "" then
				arrParam(1) = "B_MINOR,B_CONFIGURATION "
				arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
				  & "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "	' Where Condition     
			Else   
				arrParam(1) = "B_Minor,b_configuration"										' TABLE 명칭 
				arrParam(4) = "b_configuration.minor_cd =  " & FilterVar(frm1.txtPayMethCd.Value , "''", "S") & " And b_configuration.Major_Cd=" & FilterVar("B9004", "''", "S") & "  and " & _
				  "b_minor.minor_cd=*b_configuration.REFERENCE and b_configuration.SEQ_NO>" & FilterVar("1", "''", "S") & "  And " & _
				  "b_minor.Major_Cd=" & FilterVar("A1006", "''", "S") & "  and substring(b_configuration.REFERENCE,2,1) <> " & FilterVar("R", "''", "S") & "  "
			End If 

			arrParam(2) = Trim(frm1.txtPayTypeCd.value)										' Code Condition
			arrParam(3) = ""																' Name Condition
			arrParam(5) = "지급유형"													' TextBox 명칭 
	 
			arrField(0) = "B_MINOR.MINOR_CD"												' Field명(0)
			arrField(1) = "B_MINOR.MINOR_NM"												' Field명(1)
			  
			arrHeader(0) = "지급유형"													' Header명(0)
			arrHeader(1) = "지급유형명"													' Header명(1)  
		Case 12
			arrHeader(0) = "부가세유형"													' Header명(0)
			arrHeader(1) = "부가세명"													' Header명(1)
			arrHeader(2) = "부가세Rate"
			 
			arrField(0) = "B_Minor.MINOR_CD"												' Field명(0)
			arrField(1) = "B_Minor.MINOR_NM"												' Field명(1)
			arrField(2) = "b_configuration.REFERENCE"
			 
			arrParam(0) = "부가세유형"													' 팝업 명칭 
			arrParam(1) = "B_Minor,b_configuration"											' TABLE 명칭 
			arrParam(2) = Trim(StrCode)														' Code Condition
		 
			arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9001", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
			              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"  
			arrParam(5) = "부가세유형"													' TextBox 명칭 
	End Select    
	

	
	IsOpenPop = True
	 
	IF iwhere = 0 Then    
		iCalledAspName = AskPRAspName("A4101RA1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4101RA1", "X")
			IsOpenPop = False
			Exit Function
		End If
 
		Dim arrParam_1(8)

		' 권한관리 추가 
		arrParam_1(5) = lgAuthBizAreaCd
		arrParam_1(6) = lgInternalCd
		arrParam_1(7) = lgSubInternalCd
		arrParam_1(8) = lgAuthUsrID	
		
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam_1), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;") 
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")   
	End if
	 
	IsOpenPop = False
 
	If arrRet(0) = "" Then     
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtApNo.focus
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
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End if 
End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtApNo.value = arrRet(0)
				.txtApNo.focus
			Case 1 
				.txtAcctCd.value = arrRet(0)
				.txtAcctNm.value = arrRet(1)
				.txtAcctCd.focus
			Case 2
				.vspdData.Col = C_AcctCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_AcctNm
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow )					' 변경이 일어났다고 알려줌 
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
			    Call vspdData_Change(.vspdData.Col, .vspdData.Row)   
				Call SetActiveCell(.vspdData,C_VatType,.vspdData.ActiveRow ,"M","X","X")
		End Select    
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End if 
End Function

'======================================================================================================
' 기능: Tab Click
' 설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar(MENU_SIN_CRT)													'⊙: 버튼 툴바 제어 
	Else     
	    Call SetToolbar(MENU_SIN_UPD)													'⊙: 버튼 툴바 제어 
	End If
 
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)																		'~~~ 첫번째 Tab 
	gSelframeFlg = TAB1  
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)																		'~~~ 두번째 Tab 
	gSelframeFlg = TAB2
 
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolBar(MENU_MUL_CRT)
	ELSE     
		Call SetToolBar(MENU_MUL_UPD)
	End If 
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

    Call ggoOper.LockField(Document, "N")											'Lock  Suitable  Field    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call InitSpreadSheet()																	'Setup the Spread sheet
	Call InitCtrlSpread()
	Call InitCtrlHSpread()	    
    Call InitVariables()																	'Initializes local global variables    
    
    Call SetToolbar(MENU_NEW)														'버튼 툴바 제어 
'	Call InitComboBox()
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
	
	frm1.txtApNo.focus
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
    If Not chkField(Document, "1") Then									'This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData		:    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2		:    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then  
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")     
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'Clear Contents  Field
    Call ClickTab1()
    Call InitVariables()												'Initializes local global variables
	ggoSpread.Source = frm1.vspdData		:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2		:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3		:	ggoSpread.ClearSpreadData
    '-----------------------
    'Query function call area
    '-----------------------                  
    Call DbQuery()														'☜: Query db data    
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

    ggoSpread.Source = frm1.vspdData		:    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2		:    var2 = ggoSpread.SSCheckChange

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
    Call InitVariables()															'Initializes local global variables
    
	ggoSpread.Source = frm1.vspdData	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2	:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3	:	ggoSpread.ClearSpreadData
    
    frm1.txtApNo.Value = ""
    frm1.txtApNo.focus
    Call SetDefaultVal()  
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
    If lgIntFlgMode <> parent.OPMD_UMODE Then										'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")             'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete()																	'☜: Delete db data
    
    FncDelete = True                                                        
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncSave
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
    '일자관계 체크 : LC발행일(txtLcDt)<=송장일(txtInvDt)<=선하증권일(txtBlDt)<=채권/채무일(txtApDt)
    '================================================================================================
	If frm1.txtBlDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtBlDt.text,frm1.txtApDt.text,frm1.txtBlDt.Alt,frm1.txtApDt.Alt, _
                      "970025",frm1.txtBlDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			frm1.txtBlDt.focus
			Exit Function
		End If
    End If

    If frm1.txtInvDt.Text <> "" Then
		If frm1.txtBlDt.Text = "" Then
			If CompareDateByFormat(frm1.txtInvDt.text,frm1.txtApDt.text,frm1.txtInvDt.Alt,frm1.txtApDt.Alt, _
                       "970025",frm1.txtInvDt.UserDefinedFormat,parent.gComDateType, true) = False Then
				frm1.txtInvDt.focus
				Exit Function
			End If
		Else
			If CompareDateByFormat(frm1.txtInvDt.text,frm1.txtBlDt.text,frm1.txtInvDt.Alt,frm1.txtBlDt.Alt, _
                       "970025",frm1.txtInvDt.UserDefinedFormat,parent.gComDateType, true) = False Then
				frm1.txtInvDt.focus
				Exit Function
			End If
		End If
    End If
    
    If frm1.txtLcDt.Text <> "" Then
		If frm1.txtInvDt.Text = "" Then
			If frm1.txtBlDt.Text = "" Then
				If CompareDateByFormat(frm1.txtLcDt.text,frm1.txtApDt.text,frm1.txtLcDt.Alt,frm1.txtApDt.Alt, _
                        "970025",frm1.txtLcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
					frm1.txtLcDt.focus
					Exit Function
				End If
			Else
				If CompareDateByFormat(frm1.txtLcDt.text,frm1.txtBlDt.text,frm1.txtLcDt.Alt,frm1.txtBlDt.Alt, _
                        "970025",frm1.txtLcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
					frm1.txtLcDt.focus
					Exit Function
				End If
			End If
		Else
			If CompareDateByFormat(frm1.txtLcDt.text,frm1.txtInvDt.text,frm1.txtLcDt.Alt,frm1.txtInvDt.Alt, _
                       "970025",frm1.txtLcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
				frm1.txtLcDt.focus
				Exit Function
			End If   
		End If
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Call ClickTab2()
		Exit Function
    End If

    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then  
    	Call ClickTab1()                                '⊙: Check contents area
		Exit Function
    End If

	If CheckSpread3 = False then
		IntRetCD = DisplayMsgBox("110420", "X", "X", "X")
        Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																'☜: Save db data
    
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
			tmpDrCrFG = "D"
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

	With frm1
		.vspdData.ReDraw = False
 
		ggoSpread.Source = frm1.vspdData 
		ggoSpread.CopyRow

		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,.txtDocCur.value,C_NetAmt,   "A" ,"I","X","X")
	
		Call SetSpreadColor(frm1.vspdData.ActiveRow,  frm1.vspdData.ActiveRow)
		Call MaxSpreadVal(frm1.vspdData, C_ItemSeq, frm1.vspdData.ActiveRow)
		Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)
  
		   
		.vspdData.ReDraw = True
	End With
	Call Dosum()
	    		
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
				DeleteHSheet(.Text)
			End If
        End if
        
        .Redraw = False 
        ggoSpread.Source = frm1.vspdData 
        ggoSpread.EditUndo
        
        Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,Frm1.txtDocCur.value,C_NetAmt,   "A" ,"I","X","X")
		Call DoSum()
		If frm1.vspdData.MaxRows < 1 Then exit Function
  
		.Redraw = True
		.Row = .ActiveRow
		.Col = 0
  
		If .Row = 0 Then
			Exit Function
		Else 
			If .Text = ggoSpread.InsertFlag Then
				.Col = C_AcctCd
				If Len(Trim(.text)) > 0 Then 
					.Col = C_ItemSeq
					frm1.hItemSeq.value = .Text
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.ClearSpreadData		
					Call DbQuery3(.ActiveRow)
					Call SetSpread2ColorAp()
				End If 
			Else
			    .Col = C_ItemSeq
			    frm1.hItemSeq.value = .Text
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.ClearSpreadData		
				Call DbQuery2(.ActiveRow)
			End if
		End If

    End With
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim lngNum
	Dim imRow
	Dim ii
	Dim iCurRowPos
	
    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error stat	

    If gSelframeFlg <> TAB2 Then
		Call ClickTab2																'sstData.Tab = 1
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
		Next  
		.Col = 2																	' 컬럼의 절대 위치로 이동      
		.Row = 	ii - 1
		.Action = 0
        Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)

        Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,iCurRowPos + 1,iCurRowPos + imRow,frm1.txtDocCur.value,C_NetAmt,"A" ,"I","X","X")
        
        .ReDraw = True
    End With

    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
    ggoSpread.Source = frm1.vspdData2										
	ggoSpread.ClearSpreadData		
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
		.Col = 1 
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
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function  FncNext() 
    On Error Resume Next                                               
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
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
	    		
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
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function  DbDelete() 
    DbDelete = False              
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtApNo=" & Trim(frm1.txtApNo.value)						'☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
    
	Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()																'삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "2")									'Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field    
    Call InitVariables()															'Initializes local global variables
    Call SetDefaultVal()
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    
    frm1.txtApNo.Value = ""
    frm1.txtApNo.focus
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbQuery() 
    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    Dim strVal
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001     
			strVal = strVal & "&txtApNo=" & Trim(.htxtApNo.value)					'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001     
			strVal = strVal & "&txtApNo=" & Trim(.txtApNo.value)					'조회 조건 데이타 
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

	Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동 

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
		End If
	    '-----------------------
	    'Reset variables area
	    '-----------------------  
	    Call ggoOper.LockField(Document, "Q")								'This function lock the suitable field        
	    Call SetToolbar(MENU_SIN_UPD) 
	    Call InitVariables()

		lgQueryOk = True	    
	    lgIntFlgMode = parent.OPMD_UMODE											'Indicates that current mode is Update mode
	       
	    If .vspdData.MaxRows > 0 Then
	        .vspdData.Row = 1
	        .vspdData.Col = C_ItemSeq
	        .hItemSeq.Value = .vspdData.Text 
	        Call DbQuery2(1)
	    End If

		If Trim(frm1.txtBlNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtBlDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtBlDt, "N")						'N:Required, Q:Protected, D:Default
		End If

		If Trim(frm1.txtInvNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtInvDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtInvDt, "N")						'N:Required, Q:Protected, D:Default
		End If

		If Trim(frm1.txtLcNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtLcDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtLcDt, "N")						'N:Required, Q:Protected, D:Default
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
	   
	Call DoSum()

	Call txtDocCur_OnChange()
	call txtDeptCd_Onblur()  

	lgBlnFlgChgValue = False
	lgQueryOk= False	
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

	frm1.txtFlgMode.value = lgIntFlgMode         
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
				.Col = C_ItemSeq								'1
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_AcctCd									'2
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_DeptCd									'3
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_VatType								'4
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_NetAmt									'5
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_NetLocAmt								'6
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_ItemDesc								'7
			    strVal = strVal & Trim(.Text) & parent.gRowSep

				lGrpCnt = lGrpCnt + 1          
			End If
		Next
	End With
	frm1.txtMaxRows.value = lGrpCnt-1													'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value =  strDel & strVal												'Spread Sheet 내용을 저장 

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
			    .Col =  1																'C_Seq 
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
 
 	'권한관리추가 start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'권한관리추가 end
		
    frm1.txtMaxRows3.value = lGrpCnt - 1												'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread3.value =  strDel & strVal											'Spread Sheet 내용을 저장 
 
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)											'저장 비지니스 ASP 를 가동 
        
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk(ByVal ApNo)															'☆: 저장 성공후 실행 로직 
    ggoSpread.SSDeleteFlag 1
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
	   frm1.txtApNo.value = ApNo
	End If   
 
	Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    frm1.txtApNo.focus
    Call ClickTab1()
    Call InitVariables()																'Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    
	Call DBQuery()     
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
	Dim i,Indx1
	Dim arrVal,arrTemp
 
	'Err.Clear

	With frm1
	    .vspdData.row = Row
	    .vspdData.col = 1
	    .hItemSeq.Value = .vspdData.Text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	    
	    If CopyFromData(.hItemSeq.Value) = True Then
		 	Call SetSpread2ColorAp()  
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
		    
		strFrom = " A_CTRL_ITEM A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_OPEN_AP_DTL C (NOLOCK), A_OPEN_AP_ITEM D (NOLOCK) "
	 
		strWhere =     " D.AP_NO = " & FilterVar(UCase(.txtAPNo.value), "''", "S")
		strWhere = strWhere & " AND D.ITEM_SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.AP_NO  =  C.AP_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
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
				frm1.vspddata2.row = lngRows 
				frm1.vspddata2.col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspddata2.col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspddata2.col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspddata2.col = C_ColNm
					strColNm = frm1.vspddata2.text 
					frm1.vspddata2.col = C_MajorCd     
					strMajorCd = frm1.vspddata2.text 
				 
					frm1.vspddata2.col = C_CtrlVal
				 
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text , "''", "S") & " " 
				 
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") & " "
					End If     
				 
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.col = C_CtrlValNm
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
				strVal = strVal & Chr(11) & .vspddata2.text
				strVal = strVal & Chr(11) & Chr(12)         
			Next     

			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal 
		End If   
	 
'		Call CopyFromData(.hItemSeq.value)
		Call SetSpread2ColorAp()  
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
	With frm1
        ggoSpread.Source = .vspdData2
        
        Call SetSpread2ColorAp()
        Call txtDocCur_OnChange()
    End With
End Function

'======================================================================================================
' Function Name : SetSpreadFG
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Sub  SetSpreadFG(ByVal pobjSpread , ByVal pMaxRows )
    Dim lngRows 
    
    For lngRows = 1 To pMaxRows
        pobjSpread.Col = 0
        pobjSpread.Row = lngRows
        pobjSpread.Text = ""
    Next
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
  
	Call AutoInputDetail(strAcctCd, Trim(frm1.vspdData.text), frm1.txtApDt.text, Row)
	For ii = 1 To frm1.vspdData2.MaxRows
		frm1.vspddata2.col = C_CtrlVal
		frm1.vspddata2.row = ii
		  
		If Trim(frm1.vspddata2.text) <> "" Then
			Call CopyToHSheet2(frm1.vspdData.ActiveRow,ii)       
		End if
	Next
End Sub

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

'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'=======================================================================================================
Sub DoSum()
	Dim dblToNetAmt
	Dim dblToNetLocAmt

	With frm1.vspdData
		dblToNetAmt = FncSumSheet1(frm1.vspdData,C_NetAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToNetLocAmt = FncSumSheet1(frm1.vspdData,C_NetLocAmt, 1, .MaxRows, False, -1, -1, "V")
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then                     
			frm1.txtTotNetAmt.text = UNIConvNumPCToCompanyByCurrency(dblToNetAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End If 
		frm1.txtTotNetLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblToNetLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")    
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
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 순매입액 
		ggoOper.FormatFieldByObjectOfCur .txtNetAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 현금매입액 
		ggoOper.FormatFieldByObjectOfCur .txtCashAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 선급금매입액 
		ggoOper.FormatFieldByObjectOfCur .txtPrPaymAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 매입총액 
		ggoOper.FormatFieldByObjectOfCur .txtApTotAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 외상매입액 
		ggoOper.FormatFieldByObjectOfCur .txtApAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec  
		' 채무잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec    
		' 순매입액(탭)
		ggoOper.FormatFieldByObjectOfCur .txtTotNetAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec    
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
' Name : CurFormatNumSprSheet()
' Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	Dim ii 
	With frm1
'		For ii = 1 To .vspdData.MaxRows 
'			Call FixDecimalPlaceByCurrency2(frm1.vspdData,ii,.txtDocCur.value,C_NetAmt,"A" ,"X","X")
'      	Next
      	
'       Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,1,-1,.txtDocCur.value,C_NetAmt,"A" ,"I","X","X")         

	
		ggoSpread.Source = frm1.vspdData
		' 순매입액(탭)
		ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : XchLocRate()
'	Description : 환율이 변경되는 Factor 가 변했을 때 수정되는 Local Amt. Setting
'====================================================================================================
Sub XchLocRate()
	Dim ii

	With frm1
		For ii = 1 To .vspdData.MaxRows 
			.vspdData.Row = ii	
			.vspdData.Col = C_NetLocAmt	
			.vspdData.Text = "0"    		
			ggoSpread.Source = .vspdData
			ggoSpread.UpdateRow ii				
		Next	
		.txtTotNetLocAmt.text = "0"
		.txtPrPaymLocAmt.text = ""
		.txtCashLocAmt.text   = ""
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
			Call SetSpread2ColorAp()
			Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,-1,-1,frm1.txtDocCur.value,C_NetAmt,"A" ,"I","X","X")    
							
		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
'			Call SetSpread2Lock()			
			Call SetSpread2ColorAp()  
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
									Call FncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtApNo.Value)
								End If
						End Select
					Next
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtApNo.Value)
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
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.txtApNo.Value) 
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
		    
		strFrom = " A_CTRL_ITEM A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_OPEN_AP_DTL C (NOLOCK), A_OPEN_AP_ITEM D (NOLOCK) "
	 
		strWhere =     " D.AP_NO = " & FilterVar(UCase(.txtAPNo.value), "''", "S")
		strWhere = strWhere & " AND D.ITEM_SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.AP_NO  =  C.AP_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
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
        Call SetToolbar(MENU_MUL_CRT)											'버튼 툴바 제어 
    Else        
        Call SetToolbar(MENU_MUL_UPD)											'버튼 툴바 제어 
    End If    
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1 
            .vspdData.Row = NewRow
            .vspdData.Col = 1
            
            .vspdData.Col = C_ItemSeq
            .hItemSeq.value = .vspdData.Text
			ggoSpread.Source = frm1.vspdData2
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
    gMouseClickStatus = "SPC"								'Split 상태코드 
 
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 then
	    Exit Sub
	End if

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col								'Ascending Sort
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey					'Descending Sort
			lgSortKey = 1
		End If																
		Exit Sub
	End If		


	frm1.vspddata.row = frm1.vspddata.ActiveRow 
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
				Call InputCtrlVal(Row)
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
	           
	        Call OpenDept(frm1.txtDeptCd.Value, 2)
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
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0    

    Select Case Col
		Case   C_AcctCD      
			If  frm1.vspdData.Text = ggoSpread.InsertFlag Then
				frm1.vspdData.Col = C_ItemSeq
			    frm1.hItemSeq.value = frm1.vspdData.Text
			    frm1.vspdData.Col = C_AcctCd
   
			    If Len(frm1.vspdData.Text) > 0 Then
					frm1.vspdData.Row = Row
					frm1.vspdData.Col = 1     
					DeleteHsheet frm1.vspdData.Text    
			        Call DbQuery3(Row)
					Call InputCtrlVal(Row)      
			        Call SetSpread2ColorAp()
			    End If    
			End If  
		Case C_NetAmt
			frm1.vspdData.Col = C_NetLocAmt
			frm1.vspdData.text = ""	
			Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,Frm1.txtDocCur.value,C_NetAmt,  "A" ,"X","X")
			Call DoSum()
		Case C_NetLocAmt
			Call DoSum()
		Case C_VatNm, C_VatType
			 Call InputCtrlVal(Row)'  			
			
    End Select      
End Sub

'======================================================================================================
'   Event Name :vspdData_EditMode
'   Event Desc :
'=======================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_NetAmt
        
            Call EditModeCheck2(frm1.vspdData, Row, frm1.txtDocCur.value,C_NetAmt, "A" ,"I", Mode, "X", "X")
    End Select
End Sub
'======================================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata_DblClick( ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
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
'   Event Name : txtApDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtApDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtApDt.Action = 7        
        Call SetFocusToDocument("M")
		Frm1.txtApDt.Focus 
    End If
	Call txtApDt_OnBlur()
End Sub

'=======================================================================================================
'   Event Name : txtDealBpCd_onChange()
'   Event Desc :  
'=======================================================================================================
Sub  txtDealBpCd_onChange()

    lgBlnFlgChgValue = True
		
	If lgIntFlgMode <> parent.OPMD_UMODE Then 		
		Call CommonQueryRs("A.PARTNER_BP_CD,B.BP_NM", "B_BIZ_PARTNER_FTN A,B_BIZ_PARTNER B", "A.PARTNER_BP_CD = B.BP_CD AND A.PARTNER_FTN  = 'MPA' and DEFAULT_FLAG = " & FilterVar("Y", "''", "S") & "  and A.BP_CD = " & FilterVar(frm1.txtDealBpCd.value, "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) <> "" Then 
			frm1.txtPayBpCd.value = REPLACE(lgF0,Chr(11),"")
			frm1.txtPayBpNm.value = REPLACE(lgF1,Chr(11),"")
		Else
			frm1.txtPayBpCd.value = frm1.txtDealBpCd.value
			frm1.txtPayBpNm.value = frm1.txtDealBpNm.value
		End If
		
		Call CommonQueryRs("A.PARTNER_BP_CD,B.BP_NM", "B_BIZ_PARTNER_FTN A,B_BIZ_PARTNER B", "A.PARTNER_BP_CD = B.BP_CD AND A.PARTNER_FTN  = 'MBI' and DEFAULT_FLAG = " & FilterVar("Y", "''", "S") & " AND A.BP_CD = " & FilterVar(frm1.txtDealBpCd.value, "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) <> "" Then 
			frm1.txtReportBpCd.value = REPLACE(lgF0,Chr(11),"")
			frm1.txtReportBpNm.value = REPLACE(lgF1,Chr(11),"")
		Else
			frm1.txtReportBpCd.value = frm1.txtDealBpCd.value
			frm1.txtReportBpNm.value = frm1.txtDealBpNm.value
		End If
		
	End if


End Sub
'==========================================================================================
'   Event Name : txtApDt_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtApDt_OnBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
   If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
	
				If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtApDt.Text <> "") Then
					'----------------------------------------------------------------------------------------
						strSelect	=			 " Distinct org_change_id "    		
						strFrom		=			 " b_acct_dept(NOLOCK) "		
						strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
						strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
						strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
						strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtApDt.Text, gDateFormat,""), "''", "S") & "))"			
	
					IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 			
					If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
						.txtDeptCd.value = ""
						.txtDeptNm.value = ""
						.hOrgChangeId.value = ""
						.txtDeptCd.focus
					End if

				End If
			End With
		'----------------------------------------------------------------------------------------
		End If
	End IF
  
	If lgQueryOk <> true then
		frm1.txtNetLocAmt.text = "0"
	End if
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_Onblur
'   Event Desc : 
'==========================================================================================

Sub txtDeptCd_Onblur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtApDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
	If Trim(frm1.txtDeptCd.value) <> "" Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtApDt.Text, gDateFormat,""), "''", "S") & "))"			
		
	
		
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
	End If
		'----------------------------------------------------------------------------------------

End Sub

'=======================================================================================================
'   Event Name : txtApDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtApDt_Change() 
    lgBlnFlgChgValue = True
    
	If lgQueryOk <> True Then    
		Call XchLocRate()
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
'   Event Name : txTblDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txTblDt_Change()
    lgBlnFlgChgValue = True
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
'   Event Name : txtDueDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtDueDt_Change()
    lgBlnFlgChgValue = True
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
'   Event Name : txtInvDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtInvDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtLcDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtLcDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtLcDt.Action = 7                        
        Call SetFocusToDocument("M")
		Frm1.txtLcDt.Focus     
    End If
End Sub

'=======================================================================================================
'   Event Name : txtLcDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtLcDt_Change()
    lgBlnFlgChgValue = True
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
'   Event Name : txtPrPaymAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtPrPaymAmt_Change()
	lgBlnFlgChgValue = True
 
	If Len(frm1.txtPrPaymAmt.Text) = 0 Then
		Call ggoOper.SetReqAttr(frm1.txtPrPaymNo, "D")
	Else
		If UNICDbl(frm1.txtPrPaymAmt.Text) <> 0 then
			Call ggoOper.SetReqAttr(frm1.txtPrPaymNo, "N")
		Else
			Call ggoOper.SetReqAttr(frm1.txtPrPaymNo, "D")
		End If
	End If
	
	If lgQueryOk<> True Then
		frm1.txtPrPaymLocAmt.text="0"
	End If	
End Sub

'=======================================================================================================
'   Event Name : txtPrPaymLocAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtPrPaymLocAmt_Change()
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



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.8 HTML Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************




'======================================================================================================
'txtLcNo 입력시 txtLcDt 입력필수 
'======================================================================================================
Sub txtLcNo_OnBlur()
	If Trim(frm1.txtLcNo.value) = "" Then
		Call ggoOper.SetReqAttr(frm1.txtLcDt, "D")
	Else
		Call ggoOper.SetReqAttr(frm1.txtLcDt, "N") 'N:Required, Q:Protected, D:Default
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


'=======================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!--'======================================================================================================
'            6. Tag부 
' 기능: Tag부분 설정 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="no">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
		<TABLE <%=LR_SPACE_TYPE_00%>>
			<TR>
				<TD <%=HEIGHT_TYPE_00%>></TD>
			</TR>
			<TR HEIGHT=23>
				<TD WIDTH="100%">
					<TABLE <%=LR_SPACE_TYPE_10%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
									</TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>상세채무정보</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
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
				<TD WIDTH="100%" CLASS="Tab11">  
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH="100%">
								<FIELDSET CLASS="CLSFLD">
									<TABLE <%=LR_SPACE_TYPE_40%>>
										<TR>
											<TD CLASS="TD5" NOWRAP>채무번호</TD>
											<TD CLASS="TD656" NOWRAP><INPUT NAME="txtApNo" ALT="채무번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript:CALL OpenPopUp(frm1.txtApNo.Value, 0)"></TD>        
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
								
								
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>공급처</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDealBpCd" ALT="공급처" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtDealBpCd.Value, 3)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
																 <INPUT NAME="txtDealBpNm" ALT="공급처" SIZE="20" tag = "24" ></TD>
											<TD CLASS=TD5 NOWRAP>세금계산서발행처</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReportBpCd" ALT="세금계산서발행처" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtReportBpCd.Value, 9)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
																 <INPUT  NAME="txtReportbpnm"  ALT="세금계산서발행처" SIZE="18" tag = "24" ></TD>								
										</TR>					 
										<TR>
											<TD CLASS=TD5 NOWRAP>지급처</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayBpCd" ALT="지급처" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtPayBpCd.Value, 4)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
																 <INPUT  NAME="txtpaybpnm"  ALT="지급처" SIZE="20" tag = "24" ></TD>
											<TD CLASS=TD5 NOWRAP>송장번호|송장일</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvNo" ALT="송장번호" MAXLENGTH="50" SIZE=18 STYLE="TEXT-ALIGN:  left" tag="2XXXXU">&nbsp;
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ALIGN="TOP" NAME="txtInvDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="송장일" id=OBJECT3></OBJECT>');</SCRIPT>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>채무일자</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtApDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="채무일자" id=OBJECT1></OBJECT>');</SCRIPT></TD>																					 
											<TD CLASS=TD5 NOWRAP>선하증권번호|선하증권일</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txTblNo" ALT="선하증권번호" MAXLENGTH="35" SIZE=18 tag="2XXXXU" >&nbsp;
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ALIGN="TOP" NAME="txTblDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="선하증권일" id=OBJECT4></OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>채무만기일자</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDueDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="채무만기일자" id=OBJECT2></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>LC번호|LC발행일</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txTlcNo" ALT="LC번호" MAXLENGTH="35" SIZE=18 STYLE="TEXT-ALIGN: left" tag ="21XXXU">&nbsp;
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ALIGN="TOP" NAME="txtLcDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="LC발행일" id=OBJECT6></OBJECT>');</SCRIPT>
											</TD>										
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>부서</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)"  src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
																 <INPUT NAME="txtDeptNm" ALT="부서" SIZE="18" tag ="24" ></TD>										
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
										    <TD CLASS=TD5 NOWRAP>계정코드</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="계정코드" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU" ><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtAcctCd.value,1)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
																 <INPUT NAME="txtAcctnm" ALT="계정코드명" SIZE="18"  tag  ="24"></TD>										
											<TD CLASS=TD5 nowrap>결제방법</TD>
											<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME="txtPayMethCd" ALT="결제방법" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtPayMethCd.value, 10)">
																   <INPUT TYPE=TEXT NAME="txtPayMethNm" ALT="결제방법" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>거래통화</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtDocCur.Value,8)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
											&nbsp;환율&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtXchRate" align="top" CLASS=FPDS140 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="환율" tag="24X5Z" id=OBJECT5></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>지급유형</TD>
											<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayTypeCd" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtPayTypeCd.value, 11)">
																   <INPUT TYPE=TEXT NAME="txtPayTypeNm" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>																	   								
										</TR>               
										<TR>
											<TD CLASS=TD5 NOWRAP>선급금번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPrPaymNo" ALT="선급금번호" MAXLENGTH="18" STYLE="TEXT-ALIGN: Left" tag="21XXXU" ><IMG align=top name=btnCalType onclick="vbscript:CALL OpenRefPrePaymNo()" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"></TD>
											<TD CLASS=TD5 NOWRAP>대금결제참조</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaymTerms" ALT="대금결제참조" MAXLENGTH="120" SIZE=30 STYLE="TEXT-ALIGN: left" tag ="21"></TD>
										</TR>               
										<TR>
											<TD CLASS=TD5 NOWRAP>선급금매입액</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtPrPaymAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="선급금매입액" tag="21X2Z" id=OBJECT10></OBJECT>');</SCRIPT></TD>

											<TD CLASS=TD5 NOWRAP>선급금매입액(자국통화)</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtPrPaymLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="선급금매입액(자국통화)" tag="21X2Z" id=OBJECT11></OBJECT>');</SCRIPT></TD>
										</TR>      
										<TR>
											<TD CLASS=TD5 NOWRAP>현금매입액</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtCashAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="현금매입액" tag="21X2Z" id=OBJECT12></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>현금매입액(자국통화)</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtCashLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="현금매입액(자국통화)" tag="21X2Z" id=OBJECT13></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>외상매입액</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtApAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="외상매입액" tag="24X2" id=OBJECT14></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>외상매입액(자국통화)</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtApLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="외상매입액(자국통화)" tag="24X2" id=OBJECT15></OBJECT>');</SCRIPT></TD>
										</TR>      										
										<TR>
											<TD CLASS=TD5 NOWRAP>채무잔액</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtBalAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액" tag="24X2" id=OBJECT16></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>채무잔액(자국통화)</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtBalLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액(자국통화)" tag="24X2" id=OBJECT17></OBJECT>');</SCRIPT></TD>
									        <TD CLASS=TD5 NOWRAP></TD>
									        <TD CLASS=TD6 NOWRAP>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>매입총액</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtApTotAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="매입총액" tag="24X2" id=OBJECT18></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>매입총액(자국통화)</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtApTotLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="매입총액(자국통화)" tag="24X2" id=OBJECT19></OBJECT>');</SCRIPT></TD>
										</TR>      										
										<TR>
											<TD CLASS=TD5 NOWRAP>부가세금액</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtVatAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="부가세금액" tag="24X2" id=OBJECT8></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>부가세금액(자국통화)</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtVatLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="부가세금액(자국통화)" tag="24X2" id=OBJECT9></OBJECT>');</SCRIPT></TD>
										</TR>										
										<TR>
											<TD CLASS=TD5 NOWRAP>순매입액</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNetAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="순매입액" tag="24X2"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>순매입액(자국통화)</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNetLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="순매입액(자국통화)" tag="24X2"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGLNo" SIZE=19 MAXLENGTH=18 tag="24" ALT="전표번호"></TD>
											<TD CLASS=TD5 NOWRAP>전표번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGlNo" ALT="전표번호" SIZE=19  MAXLENGTH="18" STYLE="TEXT-ALIGN: Left" tag="24XXXU" ></TD>
										</TR>											
										<TR>
											<TD CLASS=TD5 NOWRAP>비고</TD>
											<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="80" tag="2X" ></TD>
										</TR>
									</TABLE>
								</DIV>
								
								
								<DIV ID="TabDiv"  SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_20%>>
										<TR HEIGHT="60%"  >
											<TD WIDTH="100%" COLSPAN="4">
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD COLSPAN=4>
												<TABLE <%=LR_SPACE_TYPE_20%>>
													<TR>       
														<TD class=TD5 NOWRAP>순매입액</TD>
														<TD class=TD6 NOWRAP>         
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotNetAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="순매입액" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
														<TD class=TD5 NOWRAP>순매입액(자국)</TD>
														<TD class=TD6 NOWRAP>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotNetLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="순매입액(자국)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
													</TR>
												</TABLE>
											</TD>         
										</TR>
										<TR HEIGHT="40%">
											<TD WIDTH="100%" COLSPAN="4">
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData2 width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
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
				<TD WIDTH=100% HEIGHT=<%=BizSize%>>
					<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
				</TD>
			</TR>
		</TABLE>
			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TYPE=hidden CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 tag="2" name=vspdData3 width="100%" id=OBJECT7 TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
					<TEXTAREA Class=hidden name=txtSpread tag="24" TABINDEX="-1"></TEXTAREA>
					<TEXTAREA Class=hidden name=txtSpread3 tag="24" TABINDEX="-1"></TEXTAREA>
					<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="htxtApNo" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="hItemSeq" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="hAcctCd" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="txtMaxRows3" tag="24" TABINDEX="-1">
					<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">
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


