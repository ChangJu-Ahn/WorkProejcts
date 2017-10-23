
<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : (-)ä��/��ݹ��� 
'*  3. Program ID           : a4116ma1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2001/02/10
'*  8. Modified date(Last)  : 2001/02/10
'*  9. Modifier (First)     : CHANG SUNG HEE
'* 10. Modifier (Last)      : CHANG SUNG HEE
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs"></SCRIPT>
<SCRIPT LANGUAGE= VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const BIZ_PGM_QRY_ID = "a4116mb1.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "a4116mb2.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_DEL_ID =  "a4116mb3.asp"
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

Dim C_ArNo 
Dim C_AcctCd 
Dim C_AcctNm 
Dim C_BizCd 
Dim C_BizNm 
Dim C_ArDt 
Dim C_ArDueDt
Dim C_ArAmt 
Dim C_ArRemAmt 
Dim C_ArClsAmt 
Dim C_ArClsLocAmt 
Dim C_ArClsDesc 


Dim  lgStrPrevKey1
Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3

Dim  IsOpenPop	
Dim  lgRetFlag	                'Popup
Dim	 lgQueryOk					' Queryok���� (loc_amt =0 check)
Dim  gSelframeFlg

Dim  lgCurrRow

<%
Dim dtToday
dtToday = GetSvrDate
%>

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.1 Common Group -1
' Description : This part declares 1st common function group
'=======================================================================================================
'*******************************************************************************************************


'======================================================================================================
' Name : initSpreadPosVariables()
' Description : �׸���(��������) �÷� ���� ���� �ʱ�ȭ 
'=======================================================================================================
Sub initSpreadPosVariables()
	C_ArNo  = 1
	C_AcctCd = 2
	C_AcctNm = 3
	C_BizCd = 4
	C_BizNm = 5
	C_ArDt = 6
	C_ArDueDt = 7
	C_ArAmt = 8
	C_ArRemAmt = 9
	C_ArClsAmt = 10
	C_ArClsLocAmt = 11
	C_ArClsDesc = 12
End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE					'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size
        
    lgStrPrevKey = ""									'initializes Previous Key
    lgStrPrevKey1 = ""
    lgStrPrevKeyDtl = 0									'initializes Previous Key
    lgLngCurRows = 0									'initializes Deleted Rows Count
	frm1.txtPaymAmt.text= "0"
	frm1.txtPaymLocAmt.text= "0"
	frm1.txtTotArAmt.text= "0"
	frm1.txtTotArRemAmt.text= "0"
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtAllcDt.text  = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDocCur.value = parent.gCurrency
	frm1.txtXchRate.text = 1
	
	lgBlnFlgChgValue = False
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
' Function Desc : This method initializes spread Sheet Column property
'=======================================================================================================
Sub  InitSpreadSheet()
    Call initSpreadPosVariables()

    With frm1.vspdData
    
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadInit "V20021103",,parent.gAllowDragDropSpread 

		.Redraw = False
		    
		.MaxCols = C_ArClsDesc + 1   
		.Col = .MaxCols
		.ColHidden = True
		.MaxRows = 0
			
		Call GetSpreadColumnPos("A")	
	
		ggoSpread.SSSetEdit  C_ArNo        , "ä�ǹ�ȣ"      , 18	, 3	'1
		ggoSpread.SSSetEdit  C_AcctCd      , "�����ڵ�"      , 10,,,20,2	'2
		ggoSpread.SSSetEdit  C_AcctNm      , "�����ڵ��"    , 15,,,20,2	'3    
		ggoSpread.SSSetEdit  C_BizCd       , "�����"        , 10, 3	'6
		ggoSpread.SSSetEdit  C_BizNm       , "������"      , 15,,,20,2	'7    
		ggoSpread.SSSetDate  C_ArDt        , "ä������"      , 10, 2, parent.gDateFormat  
		ggoSpread.SSSetDate  C_ArDueDt     , "��������"      , 10, 2, parent.gDateFormat  		
		ggoSpread.SSSetFloat C_ArAmt       , "ä�Ǿ�"        , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_ArRemAmt    , "ä���ܾ�"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_ArClsAmt    , "�����ݾ�"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_ArClsLocAmt , "�����ݾ�(�ڱ�)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec    
    	ggoSpread.SSSetEdit  C_ArClsDesc   , "���"          , 20,,,20	'2		
		
		.Redraw = True     	
    End With
      
    Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SpreadLock C_ArNo,-1, C_ArNo
		ggoSpread.SpreadLock C_AcctCd,-1, C_AcctCd
		ggoSpread.SpreadLock C_AcctNm,-1, C_AcctNm
		ggoSpread.SpreadLock C_BizCd,-1, C_BizCd
		ggoSpread.SpreadLock C_BizNm,-1, C_BizNm
		ggoSpread.SpreadLock C_ArDt,-1, C_ArDt
		ggoSpread.SpreadLock C_ArDueDt,-1, C_ArDueDt		
		ggoSpread.SpreadLock C_ArAmt,-1, C_ArAmt
		ggoSpread.SpreadLock C_ArRemAmt,-1, C_ArRemAmt    
		
		ggoSpread.SSSetRequired  C_ArClsAmt, -1, -1 		
		
		.vspdData.ReDraw = True   
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData
		ggoSpread.source = frm1.vspdData
    
		.ReDraw = False
		ggoSpread.SSSetProtected C_ArNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcctCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcctNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BizCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BizNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ArDt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ArDueDt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ArAmt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ArRemAmt, pvStartRow, pvEndRow
		
		ggoSpread.SSSetRequired C_ArClsAmt, pvStartRow, pvEndRow
		.ReDraw = True   
    End With
End Sub

'=========================================================================================================
'	Name : OpenRefOpenAr()
'	Description : Ref ȭ���� call�Ѵ�. 
'========================================================================================================= 
Function OpenRefOpenAr()
	Dim arrRet
	Dim arrParam(11)	
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4112RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4112RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtBpCd.value				' �˻������� ������� �Ķ���� 
	arrParam(1) = frm1.txtBpNm.value				
	arrParam(2) = frm1.txtDocCur.value			
	arrParam(3) = "M"
	arrParam(4) = frm1.txtAllcDt.text			
	arrParam(5) = frm1.txtAllcDt.alt

	' ���Ѱ��� �߰� 
	arrParam(8) = lgAuthBizAreaCd
	arrParam(9) = lgInternalCd
	arrParam(10) = lgSubInternalCd
	arrParam(11) = lgAuthUsrID	

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0, 0) = "" Then
		Exit Function
	Else		
		Call SetRefOpenAr(arrRet)
	End If
End Function

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

			C_ArNo        = iCurColumnPos(1)
			C_AcctCd      = iCurColumnPos(2)
			C_AcctNm      = iCurColumnPos(3)
			C_BizCd       = iCurColumnPos(4) 
			C_BizNm       = iCurColumnPos(5)
			C_ArDt        = iCurColumnPos(6)
			C_ArDueDt     = iCurColumnPos(7)
			C_ArAmt       = iCurColumnPos(8)
			C_ArRemAmt    = iCurColumnPos(9)
			C_ArClsAmt    = iCurColumnPos(10)
			C_ArClsLocAmt = iCurColumnPos(11)
			C_ArClsDesc   = iCurColumnPos(12)
	End select
End Sub

'=========================================================================================================
'	Name : OpenPopupGL()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'=========================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A5120RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
		
	arrParam(0) = Trim(frm1.txtGlNo.value)	'ȸ����ǥ��ȣ 
	arrParam(1) = ""						'Reference��ȣ 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
End Function

'=========================================================================================================
'	Name : openTempglpopup
'	Description :������ǥ  POP-UP
'=========================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A5130RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5130RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	
	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)		'ȸ����ǥ��ȣ 
	arrParam(1) = ""								'Reference��ȣ 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
End Function

'=========================================================================================================
'	Name : SetRefOpenAr()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'=========================================================================================================
Function SetRefOpenAr(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	
	With frm1
	
		.vspdData.focus		
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False	
	
		TempRow = .vspdData.MaxRows												'��: ��������� MaxRows

		For I = TempRow To TempRow + Ubound(arrRet, 1) 
			sFindFg	= "N"
			For x = 1 to TempRow
				.vspdData.Row = x
				.vspdData.Col = C_ArNo				
				If "" & UCase(Trim(.vspdData.Text)) = "" & UCase(Trim(arrRet(I - TempRow, 0))) Then
					sFindFg	= "Y"
				End If
			Next
			If 	sFindFg	= "N" Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
				.vspdData.Row = I + 1				
				.vspdData.Col = 0

				.vspdData.Text = ggoSpread.InsertFlag
				.vspdData.Col = C_ArNo												
				.vspdData.text = arrRet(I - TempRow, 0)				
				.vspdData.Col = C_AcctCd											
				.vspdData.text = arrRet(I - TempRow, 1)				
				.vspdData.Col = C_AcctNm											
				.vspdData.text = arrRet(I - TempRow, 2)				
				.vspdData.Col = C_BizCd												
				.vspdData.text = arrRet(I - TempRow, 5)				
				.vspdData.Col = C_BizNm												
				.vspdData.text = arrRet(I - TempRow, 6)				
				.vspdData.Col = C_ArDt												
				.vspdData.text = arrRet(I - TempRow, 7)				
				.vspdData.Col = C_ArDueDt 											
				.vspdData.text = arrRet(I - TempRow, 8)				
				.vspdData.Col = C_ArAmt												
				.vspdData.text = arrRet(I - TempRow, 11)				
				.vspdData.Col = C_ArRemAmt 											
				.vspdData.text = arrRet(I - TempRow, 13)		
				.vspdData.Col = C_ArClsAmt 											
				.vspdData.text = arrRet(I - TempRow, 13)	
				.vspdData.Col = C_ArClsDesc
				.vspdData.text = arrRet(I - TempRow, 19)							
			End If
		Next	
		
		frm1.txtDocCur.Value = arrRet(0, 22)				
		frm1.txtbpCd.Value = arrRet(0, 20)				
		frm1.txtbpNm.Value = arrRet(0, 21)				
		
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
		
		ggoSpread.SpreadUnlock   C_ArNo  , TempRow + 1, C_AcctCd, .vspdData.MaxRows				'��: Unlock �÷� 
		ggoSpread.ssSetProtected C_ArNo  , TempRow + 1, .vspdData.MaxRows
		ggoSpread.ssSetProtected C_AcctCd, TempRow + 1, .vspdData.MaxRows		
		Call ggoOper.SetReqAttr(frm1.txtAllcDt,   "Q")
		
		Call SetSpreadColor(TempRow + 1, .vspdData.MaxRows)
		Call txtDocCur_OnChange()
		
		.vspdData.ReDraw = True
    End With
End Function
'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBpCd.className = "protected" Then Exit Function	
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "S"							'B :���� S: ���� T: ��ü 
	arrParam(5) = "PAYTO"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
	
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
	Dim iArrParam(8)
	Dim strCd	
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4116RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4116RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	Select Case iWhere
		Case 0	
			If frm1.txtAllcNo.className = "protected" Then Exit Function			
		Case 1
			If frm1.txtBpCd.className = "protected" Then Exit Function		
			IsOpenPop = True
			arrParam(0) = "�ŷ�ó�˾�"
			arrParam(1) = "B_BIZ_PARTNER"				
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "�ŷ�ó"			
	
			arrField(0) = "BP_CD"	
			arrField(1) = "BP_NM"	
    
			arrHeader(0) = "�ŷ�ó"		
			arrHeader(1) = "�ŷ�ó��"								' Header��(1)			
		Case 3
			If frm1.txtDocCur.className = "protected" Then Exit Function		
			IsOpenPop = True
			arrParam(0) = "�ŷ���ȭ�˾�"							' �˾� ��Ī 
			arrParam(1) = "B_CURRENCY"									' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtDocCur.Value)					' Code Condition
			arrParam(3) = ""											' Name Cindition
			arrParam(4) = ""											' Where Condition
			arrParam(5) = "�ŷ���ȭ"			
	
			arrField(0) = "CURRENCY"									' Field��(0)
			arrField(1) = "CURRENCY_DESC"								' Field��(1)
    
			arrHeader(0) = "�ŷ���ȭ"								' Header��(0)
			arrHeader(1) = "�ŷ���ȭ��"
		Case 4			
			IsOpenPop = True
			arrParam(0) = "�����ڵ��˾�"
			arrParam(1) = "A_Acct"				
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "�����ڵ�"			
	
			arrField(0) = "ACCT_CD"	
			arrField(1) = "ACCT_NM"	
    
			arrHeader(0) = "�����ڵ�"		
			arrHeader(1) = "�����ڵ��"								' Header��(1)				
		Case 5	
			If frm1.txtBankCd.className = "protected" Then Exit Function
			IsOpenPop = True
			arrParam(0) = "�����˾�"
			arrParam(1) = "B_BANK"				
			arrParam(2) = Trim(frm1.txtBankCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "����"			
	
			arrField(0) = "BANK_CD"	
			arrField(1) = "BANK_NM"	
    
			arrHeader(0) = "����"		
			arrHeader(1) = "�����"	
		Case 6
			If frm1.txtBankAcct.className = "protected" Then Exit Function
			IsOpenPop = True
			arrParam(0) = "���¹�ȣ�˾�"
			arrParam(1) = "B_BANK, B_BANK_ACCT"				
			arrParam(2) = Trim(frm1.txtBankAcct.Value)
			arrParam(3) = ""
			
			IF Trim(frm1.txtBankCd.Value) = "" Then
				strCd = "B_BANK.BANK_CD = B_BANK_ACCT.BANK_CD "
			Else
				strCd = "B_BANK.BANK_CD = B_BANK_ACCT.BANK_CD AND  B_BANK_ACCT.BANK_CD =  " & FilterVar(frm1.txtBankCd.Value, "''", "S") & " "	
			End IF		
			
			arrParam(4) = strCd
			arrParam(5) = "���¹�ȣ"			
			
		    arrField(0) = "B_BANK_ACCT.BANK_ACCT_NO"	
		    arrField(1) = "B_BANK.BANK_CD"	
		    arrField(2) = "B_BANK.BANK_NM"	
		    
		    arrHeader(0) = "���¹�ȣ"		
		    arrHeader(1) = "����"	
		    arrHeader(2) = "�����"	
		Case 7
		
			
			DIm strWhere
			
			If frm1.txtCheckCd.className = "protected" Then Exit Function
			IsOpenPop = True
			
			arrParam(0) = "������ȣ�˾�"										' �˾� ��Ī 
			arrParam(1) = "f_note a,b_biz_partner b, b_bank c"						' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtCheckCd.Value)								' Code Condition
			arrParam(3) = ""														' Name Condition
			
			If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
				strWhere = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _ 
								& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "	_ 
								& " AND B_CONFIGURATION.MINOR_CD =  " & FilterVar(UCase(frm1.txtInputType.value), "''", "S") & ""
			ELse
				strWhere = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD and B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
								& " and B_CONFIGURATION.SEQ_NO = 2 and B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " " _
								& " And B_minor.minor_cd Not in ( Select  minor_cd  from b_configuration " _ 
								& " where major_cd=" & FilterVar("a1006", "''", "S") & "  and seq_no=4 and reference=" & FilterVar("NO", "''", "S") & " ) " _ 
								& " AND B_CONFIGURATION.MINOR_CD =  " & FilterVar(UCase(frm1.txtInputType.value), "''", "S") & ""
			End if
			
			If CommonQueryRs( " B_MINOR.MINOR_CD" , "B_CONFIGURATION ,  B_MINOR   " , strWhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
				
				Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
					Case "NP"

						'���޾��� 
						arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D3", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"					' Where Condition
						arrParam(5) = "������ȣ"
							
						arrField(4) = "c.bank_nm"    	    					
				
						arrHeader(0) = "������ȣ"											' Header��(0)' �����ʵ��� �� ��Ī				
						arrHeader(4) = "����"												' Header��(1)								
					Case "CP"  
						'���ұ���ī�� 
						arrParam(0) = "���ұ���ī���ȣ�˾�"										' �˾� ��Ī				
						arrParam(1) = "f_note a,b_biz_partner b, b_bank c, b_card_co d "						' TABLE ��Ī				
						arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("CP", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd *= c.bank_cd and a.card_co_cd*=d.card_co_cd "						' Where Condition				
						arrParam(5) = "���ұ���ī���ȣ"					    					' �����ʵ��� �� ��Ī				
				
						arrField(4) = " d.card_co_nm "    	    
									
						arrHeader(0) = "���ұ���ī���ȣ"											' Header��(0)				
						arrHeader(4) = "ī���"												' Header��(1)								
					Case "NE" ' Header��(1)	
						'�輭���� 
						arrParam(4) = "a.note_sts = " & FilterVar("ED", "''", "S") & "  AND a.note_fg = " & FilterVar("D1", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"					' Where Condition
						arrParam(5) = "������ȣ"					    					' �����ʵ��� �� ��Ī					
				
						arrField(4) = "c.bank_nm"    	    				
				
						arrHeader(0) = "������ȣ"											' Header��(0)				
						arrHeader(4) = "����"												' Header��(1)								
					Case Else
						arrParam(4) = "((a.note_sts = " & FilterVar("ED", "''", "S") & "  AND a.note_fg = " & FilterVar("D1", "''", "S") & " ) or (a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D3", "''", "S") & " )) " 
						arrParam(4) = arrParam(4) & " and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"	
						arrParam(5) = "������ȣ"					    					' �����ʵ��� �� ��Ī				
				
						arrField(4) = "c.bank_nm"    	    				
				
						arrHeader(0) = "������ȣ"											' Header��(0)				
						arrHeader(4) = "����"												' Header��(1)								
				End Select 
			
			ENd if
			arrField(0) = "a.Note_no"												' Field��(0)
			arrField(1) =  "F2" & parent.gColSep & "a.Note_amt"						' Field��(1)
			arrField(2) =  "DD" & parent.gColSep & "a.Issue_dt"						' Field��(2)
			arrField(3) = "b.bp_nm"

	
			arrHeader(1) = "�ݾ�"												' Header��(1)
			arrHeader(2) = "������"												' Header��(1)	    
			arrHeader(3) = "�ŷ�ó"												' Header��(1)
		Case 8 
			If frm1.txtInputType.className = "protected" Then Exit Function    
			
			If frm1.txtDocCur.value <> "" Then
				If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
					arrParam(0) = "��������"														' �˾� ��Ī						
					arrParam(1) = "B_MINOR,B_CONFIGURATION "
					arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
					arrParam(3) = ""																	' Name Cindition
					arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
								& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "	' Where Condition								
					arrParam(5) = "��������"														' TextBox ��Ī 
		
					arrField(0) = "B_MINOR.MINOR_CD"													' Field��(0)
					arrField(1) = "B_MINOR.MINOR_NM"													' Field��(1)
	    
					arrHeader(0) = "��������"														' Header��(0)
					arrHeader(1) = "����������"														' Header��(1)		
				Else
					arrParam(0) = "��������"														' �˾� ��Ī						
					arrParam(1) = "B_MINOR,B_CONFIGURATION "
					arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
					arrParam(3) = ""																	' Name Cindition
					arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD and B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
								& " and B_CONFIGURATION.SEQ_NO = 2 and B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " " _
								& " And B_minor.minor_cd Not in ( Select  minor_cd  from b_configuration " _ 
								& " where major_cd=" & FilterVar("a1006", "''", "S") & "  and seq_no=4 and reference=" & FilterVar("NO", "''", "S") & " ) "			' Where Condition								
					arrParam(5) = "��������"														' TextBox ��Ī 
		
					arrField(0) = "B_MINOR.MINOR_CD"													' Field��(0)
					arrField(1) = "B_MINOR.MINOR_NM"													' Field��(1)
	    
					arrHeader(0) = "��������"														' Header��(0)
					arrHeader(1) = "����������"														' Header��(1)		
				End If
			Else
				arrParam(0) = "��������"															' �˾� ��Ī						
				arrParam(1) = "B_MINOR,B_CONFIGURATION "
				arrParam(2) = Trim(frm1.txtInputType.value)												' Code Condition
				arrParam(3) = ""																		' Name Cindition
				arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
							& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "		' Where Condition								
				arrParam(5) = "��������"															' TextBox ��Ī 
		
				arrField(0) = "B_MINOR.MINOR_CD"														' Field��(0)
				arrField(1) = "B_MINOR.MINOR_NM"														' Field��(1)
	    
				arrHeader(0) = "��������"															' Header��(0)
				arrHeader(1) = "����������"															' Header��(1)									
			End If				
		Case 9	'��ݰ����ڵ� 
			If frm1.txtAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "�����ڵ��˾�"															' �˾� ��Ī 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"									' TABLE ��Ī 
			arrParam(2) = ""																			' Code Condition
			arrParam(3) = ""																			' Name Cindition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
							" and C.trans_type = " & FilterVar("AP005", "''", "S") & "  and C.jnl_cd = " & FilterVar(frm1.txtInputType.Value, "''", "S")	' Where Condition
			arrParam(5) = "�����ڵ�"																' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"																	' Field��(0)
			arrField(1) = "A.Acct_NM"																	' Field��(1)
    		arrField(2) = "B.GP_CD"																		' Field��(2)
			arrField(3) = "B.GP_NM"																		' Field��(3)
			
			arrHeader(0) = "�����ڵ�"																' Header��(0)
			arrHeader(1) = "�����ڵ��"																' Header��(1)
			arrHeader(2) = "�׷��ڵ�"																' Header��(2)
			arrHeader(3) = "�׷��"																	' Header��(3)									
	End Select				
	
	IsOpenPop = True

	' ���Ѱ��� �߰� 
	iArrParam(5) = lgAuthBizAreaCd
	iArrParam(6) = lgInternalCd
	iArrParam(7) = lgSubInternalCd
	iArrParam(8) = lgAuthUsrID
		
	If iwhere = 0 Then	
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, iArrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")				
	Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
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
				.txtAllcNo.focus
			Case 1	
				.txtBpCd.focus
			Case 3
				.txtDocCur.focus
			Case 4

			Case 5
				.txtBankCd.focus
			Case 6
				.txtBankAcct.focus
			Case 7	
				.txtCheckCd.focus
			Case 8
				.txtInputType.focus
			Case 9	
				.txtAcctCd.focus
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
				.txtAllcNo.value = arrRet(0)
				.txtAllcNo.focus
			Case 1	
				.txtBpCd.value = arrRet(0)		
				.txtBpNm.value = arrRet(1)
				.txtBpCd.focus
			Case 3
				.txtDocCur.value = arrRet(0)		
				Call txtDocCur_OnChange()
				.txtDocCur.focus
			Case 4

			Case 5
				.txtBankCd.value = arrRet(0)		
				.txtBankNm.value = arrRet(1)			    		
				.txtBankCd.focus
			Case 6
				.txtBankAcct.value = arrRet(0)		
				.txtBankCd.value = arrRet(1)		
				.txtBankNm.value = arrRet(2)	
				.txtBankAcct.focus
			Case 7	
				.txtCheckCd.value = arrRet(0)
				.txtCheckCd.focus
			Case 8
				.txtInputType.value = arrRet(0)		 	
				.txtInputTypeNm.value = arrRet(1)		 	
				Call txtInputType_OnChange()
				.txtInputType.focus
			Case 9	
				.txtAcctCd.value = arrRet(0)		
				.txtAcctnm.value = arrRet(1)							
				.txtAcctCd.focus
		End Select				
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End if	
End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = "protected" Then Exit Function
			
	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtAllcDt.Text
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = "F"									' �������� ���� Condition  
	
	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function
'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtAllcDt.text = arrRet(3)
				call txtDeptCd_Onblur()  
				.txtDeptCd.focus
	    End Select
	End With
End Function 
'======================================================================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=======================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ ù��° Tab 
	gSelframeFlg = TAB1
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar("1110101100001111")										'��: ��ư ���� ����						 
	Else				 
	    Call SetToolbar("1111101100001111")										'��: ��ư ���� ���� 
	End If               
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ ù��° Tab 
	gSelframeFlg = TAB2
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar("1110101100001111")										'��: ��ư ���� ����						 
	Else				 
	    Call SetToolbar("1111101100001111")										'��: ��ư ���� ���� 
	End If               	
	
	Call SetSumItem()
End Function



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************



'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub  Form_Load()
    Call LoadInfTB19029()																		'Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
                         
    Call ggoOper.LockField(Document, "N")														'Lock  Suitable  Field    
    Call InitSpreadSheet()																		'Setup the Spread sheet
    Call txtInputType_onChange()
    Call InitVariables()																		'Initializes local global variables
    Call SetDefaultVal()
    
    Call SetToolbar("1110101100001111")															'��ư ���� ����	                     
	frm1.txtAllcNo.focus
	
	' ���Ѱ��� �߰� 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' ����� 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ� 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' ���� 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing	
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
    Dim var1
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then															'This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var1 = True  Then		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")	    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")														'Clear Contents  Field
	ggoSpread.Source = frm1.vspddata
	ggoSpread.ClearSpreadData
    Call InitVariables()																		'Initializes local global variables	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																				'��: Query db data
           
    FncQuery = True																
    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
	Dim var1
	    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")													'��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")													'Clear Condition Field
    Call ggoOper.LockField(Document, "N")													'Lock  Suitable  Field
    Call txtInputType_onChange()
    Call InitVariables()																	'Initializes local global variables
    Call SetDefaultVal()    
    Call txtDocCur_OnChange()
    
	ggoSpread.Source = frm1.vspddata
	ggoSpread.ClearSpreadData
    
    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus
	lgBlnFlgChgValue = FALSE
	
    FncNew = True                                                          
    		
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
    If lgIntFlgMode <> parent.OPMD_UMODE Then												'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
		Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")						'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then																'��: Delete db data
       Exit Function																		'��:
    End If					
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    
    FncDelete = True                                                        
    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1
	
    FncSave = False                                                         
    
    Err.Clear                                                               
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False Then										'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'��: Display Message(There is no changed data.)
		Exit Function
		
    End If
    
    If Not chkField(Document, "2") Then														'��: Check required field(Single area)
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then												'��: Check contents area
		Exit Function
    End If
    
    If Not chkAllcDate() Then
		Exit Function
    End If 
    
    If chkInputType= False Then
		Exit Function
    End If          
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																				'��: Save db data
    
    FncSave = True                                                       
    		
	Set gActiveElement = document.activeElement    

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
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")	'��: "Will you destory previous data"
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
    
		.vspdData.ReDraw = True
	End With
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
	DIm i
	With frm1.vspdData
		If .Maxrows < 1 Then Exit Function
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo     
	
		If .MaxRows < 1 Then 
			Call ggoOper.SetReqAttr(frm1.txtAllcDt,   "N")
			Exit Function
		End if					
						

		For i = .MaxRows to 0 Step -1 
			.Row= i
			.Col =0			
			If Trim(frm1.vspddata.text) = ggoSpread.InsertFlag Then 
				Call ggoOper.SetReqAttr(frm1.txtAllcDt,   "Q")
				Exit Function
			End if
						
			Call ggoOper.SetReqAttr(frm1.txtAllcDt,   "N")
		Next

		Call DoSum()
	End With	
	    		
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow() 

End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows 
    
    If frm1.vspdData.Maxrows < 1 Then Exit Function
    
	ggoSpread.Source = frm1.vspdData	
    lDelRows = ggoSpread.DeleteRow
    
	Call DoSum()
		    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next    
    Calll parent.FncPrint()   
    	    		
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
' Function Desc : ȭ�� �Ӽ�, Tab���� 
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
	Call FncExport(parent.C_SINGLEMULTI)
		    		
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
	Dim var1
	
	FncExit = False

	ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    
    If lgBlnFlgChgValue = True or var1 = True Then  '��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
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
    strVal = strVal & "&txtAllcNo=" & Trim(frm1.txtAllcNo.value)			'��: ���� ���� ����Ÿ 

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbDelete = True                                                         

End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()												        '���� ������ ���� ���� 
	Call ggoOper.ClearField(Document, "1")                           '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                           'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                            'Lock  Suitable  Field
    Call txtInputType_onChange()

	ggoSpread.Source = frm1.vspddata
	ggoSpread.ClearSpreadData

    Call InitVariables()                                                      'Initializes local global variables
    Call SetDefaultVal()
    
    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus
	lgBlnFlgChgValue = FALSE
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
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: 
			strVal = strVal & "&txtAllcNo=" & Trim(.htxtAllcNo.value)				'��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: 
			strVal = strVal & "&txtAllcNo=" & Trim(.txtAllcNo.value)				'��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    End With

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

	Call RunMyBizASP(MyBizASP, strVal)										    '��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function  DbQueryOk()
	Dim strTemp
	
    Call SetSpreadLock() 
    '-----------------------
    'Reset variables area
    '-----------------------        
    lgIntFlgMode = parent.OPMD_UMODE
    Call SetToolbar("1111101100001111")									'��: ��ư ���� ���� 
    
	Call DoSum()
	call txtDeptCd_Onblur()  
	
	strTemp = frm1.txtXchRate.text
	Call txtDocCur_OnChange()
	frm1.txtXchRate.text = strTemp
	
    lgBlnFlgChgValue = False	
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

	frm1.txtFlgMode.value = lgIntFlgMode									
	'-----------------------
    'Data manipulate area
    '-----------------------
    ' Data ���� ��Ģ 
    ' 0: Sheet��, 1: Flag , 2: Row��ġ, 3~N: �� ����Ÿ 

    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    ggoSpread.Source = frm1.vspdData
    
	With frm1.vspdData
		For lngRows = 1 To .MaxRows
		    .Row = lngRows
			.Col = 0
				
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else
					strVal = strVal & "C" & parent.gColSep  					'��: C=Create, Row��ġ ���� 
			        .Col = C_ArNo								'1
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_AcctCd
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_ArDt
			        strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gColSep		        
			        strVal = strVal & Trim(frm1.txtDocCur.value) & parent.gColSep
			        .Col = C_ArClsAmt
			        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
			        .Col = C_ArClsLocAmt		            
			        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep	
			        .Col = C_ArClsDesc		            
					strVal = strVal & Trim(.Text) & parent.gRowSep		   		        	                  
			            
			        lGrpCnt = lGrpCnt + 1	
			End Select				
		Next
	End With	
	
	With frm1
		.txtMaxRows.value = lGrpCnt-1										'Spread Sheet�� ����� �ִ밹�� 
		.txtSpread.value =  strDel & strVal									'Spread Sheet ������ ���� 

		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'���� �����Ͻ� ASP �� ���� 
        
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function  DbSaveOk(ByVal AllcNo)													'��: ���� ������ ���� ���� 
    ggoSpread.SSDeleteFlag 1
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtAllcNo.value = AllcNo
	End If	  
	
	lgBlnFlgChgValue = False
	
	Call FncQuery()	
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************


'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================
'=======================================================================================================
' Function Name : chkAllcDate
' Function Desc : 
'========================================================================================================
Function chkAllcDate()
	Dim intI
	
	chkAllcDate = True
	With frm1
		For intI = 1 To .vspdData.Maxrows
			.vspdData.Row = intI
			.vspdData.Col = C_ArDt

			If CompareDateByFormat(.vspdData.Text,.txtAllcDt.Text,"ä������",.txtAllcDt.Alt, _
		    	               "970025",.txtAllcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   .txtAllcDt.focus
			   chkAllcDate = False
			   Exit Function
			End If
		Next
	End With
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

'======================================================================================================
' Function Name : SetSumItem
' Function Desc :
'=======================================================================================================
Function  SetSumItem()
    Dim DblTotClsAmt 
    Dim DblTotClsLocAmt 
    Dim DblTotDcLocAmt 
    Dim DblTotDcAmt 
    Dim lngRows 

	With frm1.vspdData 
		If .MaxRows > 0 Then    
		    For lngRows = 1 To .MaxRows
		        .Row = lngRows
		        .Col = C_ArClsAmt	'6
		        If .Text = "" Then
		            DblTotClsAmt = DblTotClsAmt + 0
		        Else
		            DblTotClsAmt = DblTotClsAmt + CDbl(.Text)
		        End If
		        
		        .Col = C_ArClsLocAmt	'8
		        If .Text = "" Then
		            DblTotClsLocAmt = DblTotClsLocAmt + 0
		        Else
		            DblTotClsLocAmt = DblTotClsLocAmt + CDbl(.Text)
		        End If                      
		    Next 
		End If     
    End With        
        
	frm1.txtPaymAmt.Text = DblTotClsAmt
	frm1.txtPaymLocAmt.Text = DblTotClsLocAmt		 
End Function

'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'=======================================================================================================
Sub DoSum()
	Dim dblToArAmt			'ä�Ǿ�	���� 
	Dim dblToArRemAmt		'ä���ܾ� ���� 
	Dim dblToArClsAmt		'�����ݾ� ���� 

	With frm1.vspdData
		dblToArAmt = FncSumSheet1(frm1.vspdData,C_ArAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToArRemAmt = FncSumSheet1(frm1.vspdData,C_ArRemAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToArClsAmt = (FncSumSheet1(frm1.vspdData,C_ArClsAmt, 1, .MaxRows, False, -1, -1, "V")) * -1
		
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
			frm1.txtTotArAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToArAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotArRemAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToArRemAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtPaymAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToArClsAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End If	
	End With	
End Sub

'=======================================================================================================
'   Event Name : txtInputType_Change()
'   Event Desc :  
'=======================================================================================================
Sub  txtInputType_onChange()
	Dim IntRetCD
    lgBlnFlgChgValue = True
	
	If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(frm1.txtInputType.value , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
			Case "CS" 
				frm1.txtCheckCd.value  = ""
				frm1.txtBankCd.value   = ""
				frm1.txtBankAcct.value = ""
				frm1.txtAcctCd.value   = ""
				frm1.txtAcctNm.value   = ""														
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
			Case "DP"  			' ������ 
				frm1.txtCheckCd.value  = ""
				frm1.txtAcctCd.value   = ""					
				frm1.txtAcctNm.value   = ""																		
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
			Case "NO" 	 	' ���޾���/�������� 
				If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
					frm1.txtBankCd.value   = ""
					frm1.txtBankAcct.value = ""			
					frm1.txtAcctCd.value   = ""					
					frm1.txtAcctNm.value   = ""											
					Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
					Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "N")	
				Else
					IntRetCD = DisplayMsgBox("111524","X","X","X")  
					frm1.txtInputType.value   = ""
					frm1.txtInputTypeNm.value = ""			
					frm1.txtAcctCd.value      = ""
					frm1.txtAcctNm.value      = ""
					Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")						
					Exit Sub
				End If								
			Case Else
				frm1.txtCheckCd.value  = ""
				frm1.txtBankCd.value   = ""
				frm1.txtBankAcct.value = ""		
				frm1.txtAcctCd.value   = ""					
				frm1.txtAcctNm.value   = ""																		
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
		End Select
	End If	
	If 	frm1.txtInputType.value = "" Then
		frm1.txtCheckCd.value  = ""
		frm1.txtBankCd.value   = ""
		frm1.txtBankAcct.value = ""		
		frm1.txtAcctCd.value   = ""					
		frm1.txtAcctNm.value   = ""																		
		Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
		Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
	End If	
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
	Dim iRef

    lgBlnFlgChgValue = True
	
	If CommonQueryRs( "reference" , "b_configuration" , " major_cd=" & FilterVar("a1006", "''", "S") & "  and minor_cd =  " & FilterVar(frm1.txtinputtype.value , "''", "S") & " and seq_no=4 " , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		iRef = split(lgF0,Chr(11))
		If lgQueryOk = False Then
			If UCase(Trim(frm1.txtDocCur.value)) <> UCase(parent.gCurrency) Then 
				If iRef(0) = "NO" Then
					frm1.txtInputType.value = ""
					frm1.txtInputTypeNm.value = ""
					frm1.txtAcctCd.value = ""
					frm1.txtAcctNm.value = ""					
					frm1.txtCheckCd.value = ""
					Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")
				End If
			End If			
		End If
	
	Else
		frm1.txtInputType.value = ""
		Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")
	End If

	Call CurFormatNumericOCX()
	Call CurFormatNumSprSheet()	
	
    If lgQueryOk = False Then
		If UCase(Trim(frm1.txtDocCur.value)) <> UCase(parent.gCurrency) Then 
			frm1.txtXchRate.Text = "0"
		Else
			frm1.txtXchRate.Text = "1"
		End If
	End If	
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' ��ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtPaymAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' ä�Ǿ� 
		ggoOper.FormatFieldByObjectOfCur .txtTotArAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' ä���ܾ� 
		ggoOper.FormatFieldByObjectOfCur .txtTotArRemAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' �����ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtPaymAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' ä�Ǿ� 
		ggoSpread.SSSetFloatByCellOfCur C_ArAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' ä���ܾ� 
		ggoSpread.SSSetFloatByCellOfCur C_ArRemAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' �����ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_ArClsAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
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
' Description : �̵��� �÷��� ������ ���� 
'====================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'===================================== PopRestoreSpreadColumnInf()  ======================================
' Name : PopRestoreSpreadColumnInf()
' Description : �÷��� ���������� ������ 
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





'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0101111111")
	
	Set gActiveSpdSheet = frm1.vspdData        		
    gMouseClickStatus = "SPC"							'Split �����ڵ� 
 	
	If frm1.vspdData.Maxrows = 0 then
	    Exit Sub
	End if

	If Row <= 0 Then
		Exit Sub
	End If		
End Sub

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : �󼼳��� �׸����� (��Ƽ)�÷��� �ʺ� �����ϴ� ��� 
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
'   Event Desc : Spread Split �����ڵ� 
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
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData_EditChange(ByVal Col , ByVal Row )
                
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
	Dim ArAmt
	Dim ArClsAmt
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0             
	
	With frm1.vspdData
		.Row = Row
		.Col = C_ArAmt
		ArAmt = .Text
		
		Select Case Col
			Case C_ArAmt
				Call DoSum()
			Case C_ArRemAmt
				Call DoSum()
			Case C_ArClsAmt
				.Col = C_ArClsAmt
				ArClsAmt = .Text
				If (UNICDbl(ArAmt) > 0 And UNICDbl(ArClsAmt) < 0) Or (UNICDbl(ArAmt) < 0 And UNICDbl(ArClsAmt) > 0) then
					
					.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(.Text) * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				End If
				frm1.vspdData.Col = C_ArClsLocAmt
				frm1.vspdData.text = ""				
				Call DoSum()
			Case C_ArClsLocAmt
				Call DoSum()
		End Select
	End With
End Sub

'======================================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspddata_KeyPress(KeyAscii )
     
End Sub



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************
'*******************************************************************************************************
'==========================================================================================
'   Event Name : txtDeptCd_Onblur
'   Event Desc : 
'==========================================================================================

Sub txtDeptCd_Onblur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtAllcDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	If Trim(frm1.txtDeptCd.value) <> "" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtAllcDt.Text, gDateFormat,""), "''", "S") & "))"			
		
	
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

'==========================================================================================
'   Event Name : txtAllcDt_onBlur
'   Event Desc : 
'==========================================================================================
Sub txtAllcDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
  	lgBlnFlgChgValue = True
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtAllcDt.Text <> "") Then
			'----------------------------------------------------------------------------------------
				strSelect	=			 " Distinct org_change_id "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtAllcDt.Text, gDateFormat,""), "''", "S") & "))"			
	
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
End Sub




'=======================================================================================================
'   Event Name : txtAllcDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtAllcDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAllcDt.Action = 7    
        Call txtAllcDt_onBlur() 
        Call SetFocusToDocument("M")
		Frm1.txtAllcDt.Focus                            
    End If
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>
<!--
 '#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### 
 -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>(-)ä��/��ݹ���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>								
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A>&nbsp;|&nbsp;<A href="vbscript:OpenRefOpenAr()">ä�ǹ߻�����</A></TD>								
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>��ݹ�ȣ</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAllcNo" ALT="��ݹ�ȣ" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript: Call OpenPopup(frm1.txtAllcNo.value,0)"></TD>								
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtAllcDt" CLASS=FPDTYYYYMMDD tag="23" Title="FPDATETIME" ALT="�����" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>����ó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="23NXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="����ó��"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag=23NXXU" ALT="�μ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)"> <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="�μ���"></TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME="txtInputType" ALT="��������" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtInputType.value, 8)">
													   <INPUT TYPE=TEXT NAME="txtInputTypeNm" ALT="��������" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>																	   
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBankCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="21NXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankCd.value,5)"> <INPUT TYPE=TEXT NAME="txtBankNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="�����"></TD>											
								<TD CLASS=TD5 NOWRAP>������ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCheckCd" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="21XXXU" ALT="������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCheckCd.value,7)"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���¹�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT  TYPE=TEXT NAME="txtBankAcct" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" tag="21XXXU" ALT="���¹�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankAcct.value,6)"></TD>																						
								<TD CLASS=TD5 NOWRAP>��ݰ����ڵ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="�����ڵ�" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtAcctCd.value,9)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
												 <INPUT NAME="txtAcctnm" ALT="�����ڵ��" MAXLENGTH="20"  tag  ="24"></TD>
							</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������ǥ��ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="������ǥ��ȣ"> </TD>																						
									<TD CLASS="TD5" NOWRAP>��ǥ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="��ǥ��ȣ"></TD>								
								</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="22NXXU" STYLE="TEXT-ALIGN: left" ALT="�ŷ���ȭ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtDocCur.value,3)"></TD>
								<TD CLASS=TD5 NOWRAP>ȯ��</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtXchRate" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="ȯ��" tag="24X5Z" ></OBJECT>');</SCRIPT></TD>											
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ݾ�</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtPaymAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="��ݾ�" tag="24X2" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP >��ݾ�(�ڱ���ȭ)</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtPaymLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="��ݾ�(�ڱ���ȭ)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtAllcDesc" SIZE=90 MAXLENGTH=128 tag="21XXX" ALT="���"></TD>
							</TR>						
							<TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD  COLSPAN="4">
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>							
								<TD CLASS=TD5 NOWRAP>ä�Ǿ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="ä�Ǿ�" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD class=TD5 STYLE="WIDTH : 0px;"></TD>
								<TD CLASS=TD5 NOWRAP>ä���ܾ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArRemAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="ä���ܾ�" tag="24X2" id=OBJECT3></OBJECT>');</SCRIPT></TD>											
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
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>	
</TABLE>
<TEXTAREA class=hidden name=txtSpread		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread1		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread2		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3		tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows1"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtAllcNo"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden	 NAME="hOrgChangeId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 tag="2" width="100%" TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

