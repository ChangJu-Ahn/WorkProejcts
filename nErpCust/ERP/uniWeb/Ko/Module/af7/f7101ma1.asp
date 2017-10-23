
<%@ LANGUAGE="VBSCRIPT" %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : PRERECEIPT
'*  3. Program ID           : f7101ma1
'*  4. Program Name         : ������ ��� 
'*  5. Program Desc         : ������ ��� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/09/25
'*  8. Modified date(Last)  : 2002/11/18
'*  9. Modifier (First)     : Hee Jung, Kim
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'=======================================================================================================
'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">				</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_QRY_ID	= "f7101mb1.asp"											'�����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID	= "f7101mb2.asp"											'�����Ͻ� ���� ASP�� 
Const BIZ_PGM_DEL_ID	= "f7101mb3.asp"											'�����Ͻ� ���� ASP�� 

Const PreReceiptJnlType = "PR"

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

Dim C_SEQ		
Dim C_RCPT_TYPE	
Dim C_RCPT_TYPE_PB
Dim C_RCPT_TYPE_NM
Dim C_RCPT_ACCT	
Dim C_RCPT_ACCT_PB
Dim C_RCPT_ACCT_NM
Dim C_AMT		
Dim C_LOC_AMT	
Dim C_BANK_CD	
Dim C_BANK_PB	
Dim C_BANK_NM	
Dim C_BANK_ACCT	
Dim C_BANK_ACCT_PB
Dim C_NOTE_NO	
Dim C_NOTE_NO_PB
Dim C_COL_END
Dim C_STTL_DESC

Dim IsOpenPop																		'Popup
Dim	lgFormLoad
Dim	lgQueryOk
Dim lgstartfnc

'2002.01.10 �߰� ���� ;form load .. default time �������ֱ�.
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
     C_SEQ				= 1
	 C_RCPT_TYPE		= 2
	 C_RCPT_TYPE_PB  	= 3
	 C_RCPT_TYPE_NM	    = 4
	 C_RCPT_ACCT		= 5
	 C_RCPT_ACCT_PB 	= 6
	 C_RCPT_ACCT_NM 	= 7
	 C_AMT				= 8
	 C_LOC_AMT			= 9
	 C_BANK_CD			= 10
	 C_BANK_PB			= 11
	 C_BANK_NM			= 12
	 C_BANK_ACCT		= 13
	 C_BANK_ACCT_PB 	= 14
	 C_NOTE_NO			= 15
	 C_NOTE_NO_PB		= 16
	 C_STTL_DESC        = 17
End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE												'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False														'Indicates that no value changed
    lgIntGrpCount = 0																'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgstartfnc=False
	lgFormLoad=True	
    lgStrPrevKey = ""																'initializes Previous Key
    lgQueryOk = false
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	frm1.txtDocCur.value = parent.gCurrency
	frm1.txtPrrcptDt.text = UniConvDateAToB("<%=dtToday%>",parent.gServerDateFormat,parent.gDateFormat)

<% If gIsShowLocal <> "N" Then %>
	frm1.txtXchRate.text	= 1
<% Else %>
	frm1.txtXchRate.Value	= 1
<% End If %>
	frm1.hOrgChangeId.value = parent.gChangeOrgId

	lgBlnFlgChgValue = False
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables() 

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread    
		.ReDraw = False	
        
		.MaxCols = C_STTL_DESC + 1													'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
    	.Col = .MaxCols																'������Ʈ�� ��� Hidden Column
    	.ColHidden = True
		.MaxRows = 0    	
    		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit		C_SEQ,			"����"          , 5,	2,	-1,	3
		ggoSpread.SSSetEdit		C_RCPT_TYPE,	"�Ա�����"      ,10, , ,	2, 2
		ggoSpread.SSSetButton	C_RCPT_TYPE_PB
		ggoSpread.SSSetEdit		C_RCPT_TYPE_NM,	"�Ա�������"    ,15,	,	,	50
		ggoSpread.SSSetEdit		C_RCPT_ACCT,	"�Աݰ����ڵ�"  ,12, , ,	20, 2
		ggoSpread.SSSetButton	C_RCPT_ACCT_PB
		ggoSpread.SSSetEdit		C_RCPT_ACCT_NM,	"�Աݰ����ڵ��",15,	,	,	30		
		ggoSpread.SSSetFloat	    C_AMT,			"�ݾ�"          ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C_LOC_AMT,		"�ݾ�(�ڱ�)"    ,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C_BANK_CD,		"����"          ,10, , ,	10, 2
		ggoSpread.SSSetButton	C_BANK_PB
		ggoSpread.SSSetEdit		C_BANK_NM,		"�����"        ,15, , ,	30
		ggoSpread.SSSetEdit		C_BANK_ACCT,	"���¹�ȣ"      ,15, , ,	30, 2
		ggoSpread.SSSetButton	C_BANK_ACCT_PB
		ggoSpread.SSSetEdit		C_NOTE_NO,		"������ȣ"      ,30, , ,	30, 2
		ggoSpread.SSSetButton	C_NOTE_NO_PB
		ggoSpread.SSSetEdit		C_STTL_DESC,       "���", 20,,,128
		
	    If Trim(UCase(gIsShowLocal)) = "N" Then        
			Call ggoSpread.SSSetColHidden(C_LOC_AMT,C_LOC_AMT,True)
		End If                		
		Call ggoSpread.MakePairsColumn(C_RCPT_TYPE,C_RCPT_TYPE_PB)
        Call ggoSpread.MakePairsColumn(C_RCPT_ACCT,C_RCPT_ACCT_PB)
        Call ggoSpread.MakePairsColumn(C_BANK_CD,C_BANK_PB)
        Call ggoSpread.MakePairsColumn(C_NOTE_NO,C_NOTE_NO_PB)
        Call ggoSpread.MakePairsColumn(C_BANK_ACCT,C_BANK_ACCT_PB)

		.ReDraw = True
	End With
	
	Call SetSpreadLock() 	
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
				
		ggoSpread.SpreadLock	C_SEQ,			-1,	C_SEQ
		ggoSpread.SpreadLock	C_RCPT_TYPE_NM,	-1,	C_RCPT_TYPE_NM
		ggoSpread.SpreadLock	C_RCPT_ACCT_NM,	-1,	C_RCPT_ACCT_NM		
		ggoSpread.SpreadLock	C_BANK_CD,		-1,	C_BANK_CD
		ggoSpread.SpreadLock	C_BANK_PB,		-1,	C_BANK_PB
		ggoSpread.SpreadLock	C_BANK_NM,		-1,	C_BANK_NM
		ggoSpread.SpreadLock	C_BANK_ACCT,	-1,	C_BANK_ACCT
		ggoSpread.SpreadLock	C_BANK_ACCT_PB,	-1,	C_BANK_ACCT_PB
		ggoSpread.SpreadLock	C_NOTE_NO,		-1,	C_NOTE_NO
		ggoSpread.SpreadLock	C_NOTE_NO_PB,	-1,	C_NOTE_NO_PB
		
		ggoSpread.SSSetRequired		C_RCPT_TYPE, -1,-1
		ggoSpread.SSSetRequired		C_RCPT_ACCT, -1,-1		
		ggoSpread.SSSetRequired		C_AMT      , -1,-1
		
		.vspdData.ReDraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
   	With frm1
   		.vspdData.ReDraw = False
		
		ggoSpread.SSSetProtected	    C_SEQ,          pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_RCPT_TYPE,	pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_RCPT_ACCT,	pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_AMT,      	pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected	    C_RCPT_TYPE_NM,	pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected	    C_RCPT_ACCT_NM,	pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected	    C_BANK_NM,  	pvStartRow,	pvEndRow
		
		.vspdData.ReDraw = True
	End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
            C_SEQ				= iCurColumnPos(1)
	        C_RCPT_TYPE		    = iCurColumnPos(2)
	        C_RCPT_TYPE_PB	    = iCurColumnPos(3)
	        C_RCPT_TYPE_NM	    = iCurColumnPos(4)
	        C_RCPT_ACCT		    = iCurColumnPos(5)
	        C_RCPT_ACCT_PB	    = iCurColumnPos(6)
	        C_RCPT_ACCT_NM	    = iCurColumnPos(7)
	        C_AMT				= iCurColumnPos(8)
	        C_LOC_AMT			= iCurColumnPos(9)
	        C_BANK_CD			= iCurColumnPos(10)
	        C_BANK_PB			= iCurColumnPos(11)
	        C_BANK_NM			= iCurColumnPos(12)
	        C_BANK_ACCT		    = iCurColumnPos(13)
	        C_BANK_ACCT_PB	    = iCurColumnPos(14)
	        C_NOTE_NO			= iCurColumnPos(15)
	        C_NOTE_NO_PB		= iCurColumnPos(16)
	        C_STTL_DESC         = iCurColumnPos(17) 
    End Select    
End Sub

'============================================================
'ȸ����ǥ �˾� 
'============================================================
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
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'ȸ����ǥ��ȣ 
	arrParam(1) = ""						'Reference��ȣ 
	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'============================================================
'������ǥ �˾� 
'============================================================
Function OpenPopupTempGL()
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
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'ȸ����ǥ��ȣ 
	arrParam(1) = ""						'Reference��ȣ 
	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
'   Function Name : OpenPopupPR()
'   Function Desc : 
'=======================================================================================================
Function OpenPopupPR()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	
	iCalledAspName = AskPRAspName("f7101ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f7101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.ShowModalDialog(iCalledAspName, Array(Window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		frm1.txtPrrcptNo.focus
		Exit Function
	Else
		frm1.txtPrrcptNo.value = arrRet(0)
		frm1.txtPrrcptNo.focus
	End If	

	
End Function

'=======================================================================================================
'Description : �ΰ������� �˾� 
'=======================================================================================================
Function OpenVatType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
      
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ΰ��������˾�"													' �˾� ��Ī 
	arrParam(1) = "B_MINOR a , a_jnl_acct_assn b "			                			' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtVatType.Value)
	arrParam(3) = ""
	arrParam(4) = "A.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " AND A.MINOR_CD=B.JNL_CD AND B.TRANS_TYPE=" & FilterVar("FR001", "''", "S") & ""	' WHERE ����		
	arrParam(5) = "�ΰ����ڵ�"														' �����ʵ��� �� ��Ī 
	
    arrField(0) = "A.MINOR_CD"															' Field��(0)
    arrField(1) = "A.MINOR_NM"															' Field��(1)
    
    arrHeader(0) = "�ΰ�������"														' Header��(0)
    arrHeader(1) = "�ΰ���������"													' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtVatType.focus
		Exit Function
	Else
		Call SetVatType(arrRet)
	End If	
End Function
'------------------------------------------  OpenPopupDept()  ------------------------------------------------
'	Name : OpenPopupDept()
'	Description : Dept Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'�μ��ڵ� 
	arrParam(1) = frm1.txtPrrcptDt.Text			'��¥(Default:������)
	arrParam(2) = lgUsrIntCd							'�μ�����(lgUsrIntCd)
	arrParam(3) = "F"
	IsOpenPop = True

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus	
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If
	
	
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'	Name : SetAcctCd()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetVatType(byval arrRet)
	frm1.txtVatType.Value    = arrRet(0)		
	frm1.txtVatTypeNm.Value    = arrRet(1)	
	Call txtVatType_OnChange	
	frm1.txtVatType.focus		
	lgBlnFlgChgValue = True
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
				.txtPrrcptDt.text = arrRet(3)
				.txtDeptCd.focus      
				Call txtDeptCd_OnChange()  
        End Select
	End With
End Function  
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBpCd.className = parent.UCN_PROTECTED Then Exit Function

	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopup(iwhere)
		Exit Function
	Else
		Call SetPopup(arrRet,iWhere)
		lgBlnFlgChgValue = True
	End If
End Function
'=======================================================================================================
'	Name : OpenPopup()
'	Description : �����ڵ��˾� 
'=======================================================================================================
Function OpenPopup(strCode, strWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case UCase(strWhere)
		Case "BP"
			If frm1.txtBpCd.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = "�ŷ�ó �˾�"									' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER A" 								' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "�ŷ�ó"										' �����ʵ��� �� ��Ī 

		    arrField(0) = "A.BP_CD"											' Field��(0)
		    arrField(1) = "A.BP_NM"											' Field��(1)
    
		    arrHeader(0) = "�ŷ�ó�ڵ�"									' Header��(0)
			arrHeader(1) = "�ŷ�ó��"									' Header��(1)
		Case "CURR"
			If frm1.txtDocCur.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = "��ȭ �˾�"									' �˾� ��Ī 
			arrParam(1) = "B_CURRENCY A"									' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "��ȭ"										' �����ʵ��� �� ��Ī 

		    arrField(0) = "A.CURRENCY"										' Field��(0)
		    arrField(1) = "A.CURRENCY_DESC"									' Field��(1)
    
		    arrHeader(0) = "��ȭ�ڵ�"									' Header��(0)
			arrHeader(1) = "��ȭ��"										' Header��(1)
		Case "BANK"
			arrParam(0) = "���� �˾�"									' �˾� ��Ī 
			arrParam(1) = "B_BANK A"										' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "����"										' �����ʵ��� �� ��Ī 

		    arrField(0) = "A.BANK_CD"										' Field��(0)
		    arrField(1) = "A.BANK_NM"										' Field��(1)
    
		    arrHeader(0) = "�����ڵ�"									' Header��(0)
			arrHeader(1) = "�����"										' Header��(1)
		Case "BANK_ACCT"
			arrParam(0) = "���¹�ȣ �˾�"								' �˾� ��Ī 
			arrParam(1) = "F_DPST A, B_BANK B"								' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A.BANK_CD=B.BANK_CD"								' Where Condition
			
			frm1.vspdData.Col = C_BANK_CD
			
			If "" & Trim(frm1.vspdData.Text) <> "" Then
				arrParam(4) = arrParam(4) & " AND A.BANK_CD =  " & FilterVar(frm1.vspdData.Text, "''", "S") & " "
			End If
		
			arrParam(5) = "���¹�ȣ"											' �����ʵ��� �� ��Ī 

		    arrField(0) = "A.BANK_ACCT_NO"											' Field��(0)
		    arrField(1) = "A.BANK_CD"												' Field��(1)
		    arrField(2) = "B.BANK_NM"
    
		    arrHeader(0) = "���¹�ȣ"											' Header��(0)
			arrHeader(1) = "�����ڵ�"											' Header��(1)
			arrHeader(2) = "�����"
		Case "RCPT"	'�Ա����� 
			arrParam(0) = "�Ա����� �˾�"
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
						& " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD AND B.SEQ_NO = 3 AND B.REFERENCE = " & FilterVar("PR", "''", "S") & " "
			arrParam(5) = "�Ա�����"
	
			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
			    
			arrHeader(0) = "�Ա������ڵ�"
			arrHeader(1) = "�Ա�������"
		Case "RCPTACCT"	'�Աݰ����ڵ� 
			arrParam(0) = "�����ڵ��˾�"										' �˾� ��Ī 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE ��Ī 
			arrParam(2) = ""														' Code Condition
			arrParam(3) = ""														' Name Condition
			
			frm1.vspdData.Col = C_RCPT_TYPE
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
							" and C.trans_type = " & FilterVar("fr001", "''", "S") & " and C.jnl_cd = " & FilterVar(frm1.vspdData.Text, "''", "S")         ' Where Condition
			arrParam(5) = "�����ڵ�"											' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"												' Field��(0)
			arrField(1) = "A.Acct_NM"												' Field��(1)
    		arrField(2) = "B.GP_CD"													' Field��(2)
			arrField(3) = "B.GP_NM"													' Field��(3)
			
			arrHeader(0) = "�����ڵ�"											' Header��(0)
			arrHeader(1) = "�����ڵ��"											' Header��(1)
			arrHeader(2) = "�׷��ڵ�"											' Header��(2)
			arrHeader(3) = "�׷��"												' Header��(3)						
		Case "PRRCPTTYPE"
			If frm1.txtPrrcptType.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = frm1.txtPrrcptType.Alt									' �˾� ��Ī 
			arrParam(1) = "a_jnl_item a , a_jnl_acct_assn b "	 						' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtPrrcptType.Value)							' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = "a.jnl_type = " & FilterVar(PreReceiptJnlType, "''", "S")	' Where Condition
			arrParam(4) = arrParam(4) & " and a.jnl_cd=b.jnl_cd "
			arrParam(4) = arrParam(4) & " AND B.TRANS_TYPE = " & FilterVar("FR001", "''", "S") & "" 			
			arrParam(5) = frm1.txtPrrcptType.Alt									' �����ʵ��� �� ��Ī 

		    arrField(0) = "A.JNL_CD"												' Field��(0)
		    arrField(1) = "A.JNL_NM"												' Field��(1)
    
		    arrHeader(0) = frm1.txtPrrcptType.Alt									' Header��(0)
			arrHeader(1) = frm1.txtPrrcptTypeNm.Alt									' Header��(1)
		Case "BIZAREA"
			arrParam(0) = "���ݽŰ����� �˾�"									' �˾� ��Ī 
			arrParam(1) = "B_TAX_BIZ_AREA"	 										' TABLE ��Ī 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = ""														' Where Condition
			arrParam(5) = "���ݽŰ������ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "TAX_BIZ_AREA_CD"											' Field��(0)
			arrField(1) = "TAX_BIZ_AREA_NM"											' Field��(0)
    
			arrHeader(0) = "���ݽŰ������ڵ�"									' Header��(0)
			arrHeader(1) = "���ݽŰ������"									' Header��(0)			
		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False

	If arrRet(0) = "" Then
		Call EscPopUp(strWhere)
		Exit Function
	Else		
		Call SetPopUp(arrRet, strWhere)
	End If
End Function	
'=======================================================================================================
'	Description : SetPopUp
'=======================================================================================================
Sub SetPopUp(byVal arrRet, byval strWhere)
	Select Case UCase(strWhere)
		Case "BP"
			frm1.txtBpCd.value = arrRet(0)
			frm1.txtBpNm.value = arrRet(1)
			frm1.txtBpCd.focus
			lgBlnFlgChgValue = True
		Case "CURR"
			frm1.txtDocCur.value = arrRet(0)
			Call txtDocCur_OnChange()
		    Call XchLocRate()
			frm1.txtDocCur.focus
			lgBlnFlgChgValue = True
		Case "BANK"
			With frm1.vspdData
				.Col = C_BANK_CD
				.Text = arrRet(0)
				.Col = C_BANK_NM
				.Text = arrRet(1)
				Call vspdData_Change(.Col, .Row)
				Call SetActiveCell(frm1.vspdData,C_BANK_CD,frm1.vspdData.ActiveRow ,"M","X","X")
			End With
		Case "BANK_ACCT"
			With frm1.vspdData
				.Col = C_BANK_ACCT
				.Text = arrRet(0)
				.Col = C_BANK_CD
				.Text = arrRet(1)
				.Col = C_BANK_NM
				.Text = arrRet(2)
				Call vspdData_Change(.Col, .Row)
				Call SetActiveCell(frm1.vspdData,C_BANK_ACCT,frm1.vspdData.ActiveRow ,"M","X","X")
			End With
		Case "RCPT"
			With frm1.vspdData
				.Col = C_RCPT_TYPE
				.Text = arrRet(0)
				.Col = C_RCPT_TYPE_NM
				.Text = arrRet(1)
				Call vspdData_Change(.Col, .Row)
				Call SetActiveCell(frm1.vspdData,C_RCPT_TYPE,frm1.vspdData.ActiveRow ,"M","X","X")
			End With
		Case "RCPTACCT"
			With frm1.vspdData
				.Col = C_RCPT_ACCT
				.Text = arrRet(0)
				.Col = C_RCPT_ACCT_NM
				.Text = arrRet(1)
				Call vspdData_Change(.Col, .Row)
				Call SetActiveCell(frm1.vspdData,C_RCPT_ACCT,frm1.vspdData.ActiveRow ,"M","X","X")
			End With
		Case "PRRCPTTYPE"
			frm1.txtPrrcptType.value = arrRet(0)
			frm1.txtPrrcptTypeNm.value = arrRet(1)
			frm1.txtPrrcptType.focus
			lgBlnFlgChgValue = True
		Case "BIZAREA"
			frm1.txtBizAreaCD.value = arrRet(0)
			frm1.txtBizAreaNM.value = arrRet(1)
			frm1.txtBizAreaCD.focus
			lgBlnFlgChgValue = True
		Case Else
			Exit Sub
	End Select
End Sub


'=======================================================================================================
'	Description : EscPopUp
'=======================================================================================================
Sub EscPopUp(byval strWhere)
	Select Case UCase(strWhere)
		Case "BP"
			frm1.txtBpCd.focus
		Case "CURR"
			frm1.txtDocCur.focus
		Case "BANK"
				Call SetActiveCell(frm1.vspdData,C_BANK_CD,frm1.vspdData.ActiveRow ,"M","X","X")
		Case "BANK_ACCT"
				Call SetActiveCell(frm1.vspdData,C_BANK_ACCT,frm1.vspdData.ActiveRow ,"M","X","X")
		Case "RCPT"
				Call SetActiveCell(frm1.vspdData,C_RCPT_TYPE,frm1.vspdData.ActiveRow ,"M","X","X")
		Case "RCPTACCT"
				Call SetActiveCell(frm1.vspdData,C_RCPT_ACCT,frm1.vspdData.ActiveRow ,"M","X","X")
		Case "PRRCPTTYPE"
			frm1.txtPrrcptType.focus
		Case "BIZAREA"
			frm1.txtBizAreaCD.focus
		Case Else
			Exit Sub
	End Select
End Sub
'=======================================================================================================
'	Description : ������ȣ �˾� 
'=======================================================================================================
Function OpenPopupNote(strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strNoteFg

	If IsOpenPop = True Then Exit Function	

	With frm1.vspdData
		.Col = C_RCPT_TYPE
		
		Select Case UCase(.Text)
			Case "NP"
				strNoteFg = "D3"
			Case "NR"
				strNoteFg = "D1"
			Case "CR"
				strNoteFg = "CR"
			Case "CR"
				strNoteFg = "CR"
			Case Else
				Exit Function
			End Select
	End With

	if strNoteFg <> "CR" then
		arrParam(0) = "������ȣ �˾�"								' �˾� ��Ī 
	else 
		arrParam(0) = "����ī�� �˾�"								' �˾� ��Ī 
	end if 
	arrParam(1) = "F_NOTE A, B_BIZ_PARTNER B, B_BANK C, B_CARD_CO D"						' TABLE ��Ī 
	if strNoteFg <> "CR" then
		arrParam(2) = Trim(strCode)										' Code Condition
	else
		arrParam(2) = ""										' Code Condition
	end if
	arrParam(3) = ""												' Name Cindition
	arrParam(4) = "A.NOTE_FG= " & FilterVar(strNoteFg, "''", "S") & "  AND A.NOTE_STS=" & FilterVar("BG", "''", "S") & " AND A.BP_CD=B.BP_CD "
	arrParam(4) = arrParam(4) & " AND a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd "

'-- �μ��ڵ� 
			If lgInternalCd <> "" Then
				arrParam(4) = " A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = " A.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			Else
				arrParam(4) = ""
			End If


' ����� 
			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

' �ۼ��� 
			' ���Ѱ��� �߰� 
			If lgAuthUsrID <> "" Then
				arrParam(4) = " A.INSRT_USR_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If
			
	if strNoteFg <> "CR" then
		arrParam(5) = "������ȣ"									' �����ʵ��� �� ��Ī 
	else 
		arrParam(5) = "����ī���ȣ"									' �����ʵ��� �� ��Ī 
	end if

	if strNoteFg <> "CR" then
    arrField(0) = "A.NOTE_NO"										' Field��(0)
    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"' Field��(1)
    arrField(2) = "DD" & parent.gColSep & "ISSUE_DT"' Field��(2)
    arrField(3) = "DD" & parent.gColSep & "A.DUE_DT"	' Field��(3)
    arrField(4) = "A.BP_CD"											' Field��(4)
    arrField(5) = "B.BP_NM"											' Field��(5)
    
    arrHeader(0) = "������ȣ"									' Header��(0)
	arrHeader(1) = "�����ݾ�"									' Header��(1)
	arrHeader(2) = "��������"									' Header��(2)
	arrHeader(3) = "��������"									' Header��(3)
	arrHeader(4) = "�ŷ�ó�ڵ�"									' Header��(4)
	arrHeader(5) = "�ŷ�ó��"									' Header��(5)
    else
    arrField(0) = "A.NOTE_NO"										' Field��(0)
    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"' Field��(1)
    arrField(2) = "DD" & parent.gColSep & "ISSUE_DT"' Field��(2)
    arrField(3) = "B.BP_NM"											' Field��(4)
    arrField(4) = "D.CARD_CO_NM"											' Field��(5)
    
    arrHeader(0) = "����ī���ȣ"									' Header��(0)
	arrHeader(1) = "�ݾ�"									' Header��(1)
	arrHeader(2) = "������"									' Header��(2)
	arrHeader(3) = "�ŷ�ó"									' Header��(4)
	arrHeader(4) = "ī���"									' Header��(5)
    end if
	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		Call SetActiveCell(frm1.vspdData,C_NOTE_NO,frm1.vspdData.ActiveRow ,"M","X","X")
		Exit Function
	End If	
	
	With frm1
		.vspdData.Col	= C_NOTE_NO
		.vspdData.Text	= arrRet(0)
		.vspdData.Col	= C_AMT
		.vspdData.Text	= arrRet(1)
		.vspdData.Col	= C_LOC_AMT
		.vspdData.Text	= arrRet(1)
		
	    Call vspdData_Change(.vspdData.Col, .vspdData.Row)
	    Call SetActiveCell(frm1.vspdData,C_NOTE_NO,frm1.vspdData.ActiveRow ,"M","X","X")
	End With
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
Sub Form_Load()
    Call LoadInfTB19029()                                                     'Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                         
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitVariables()                                                      'Initializes local global variables
	Call FncNew()

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
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'======================================================================================================
Function FncQuery() 
    Dim IntRetCD
    
    FncQuery = False                    
    lgstartfnc = True                                       
    
    Err.Clear																			'Protect system from crashing
	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
      	    Exit Function
    	End If
    End If
	'-----------------------
    'Check contents area
    '----------------------- 
    If Not chkField(Document, "1") Then													'This function check indispensable field
		Exit Function
    End If    
    
    Call InitVariables()																'Initializes local global variables
    
    frm1.vspdData.MaxRows = 0
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then													'This function check indispensable field
		Exit Function
    End If
	'-----------------------
    'Query function call area
    '----------------------- 
    frm1.hCommand.value = "LOOKUP"
    Call DbQuery()																		'Query db data
       
    FncQuery = True		
    lgstartfnc = False
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False                                                          
	'-----------------------
	'Check previous data area
	'-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "1")										'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
	Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
	
	Call InitVariables()																	'Initializes local global variables
    Call InitSpreadSheet()
	Call SetDefaultVal()
	Call SetToolbar("1110110100101111")
	
	Call txtDocCur_OnChange()
    frm1.txtPrrcptNo.focus 
	
	lgBlnFlgChgValue = False
	FncNew = True                                                           
	lgFormLoad = True																	' tempgldt read
    lgQueryOk = False
    lgstartfnc = False    
    
    Set gActiveElement = document.activeElement
End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete() 
    Dim IntRetCD
	FncDelete = False
		
	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")					'�����Ͻðڽ��ϱ�?  
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	'-----------------------
	'Precheck area
	'-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then											'Check if there is retrived data
        intRetCD = DisplayMsgBox("900002","x","x","x")                                
    	Exit Function
    End If
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete()																		'��: Delete db data
    
    FncDelete = True
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	Dim IntRetCD 
	
	FncSave = False
	
	Err.Clear                                                               

    ggoSpread.Source = frm1.vspdData												'��: Preset spreadsheet pointer 
    
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then		'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")							'��: Display Message(There is no changed data.)
        Exit Function
    End If

    If Not chkField(Document, "2") Then													'��: Check required field(Single area)
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData												'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then									'��: Check required field(Multi area)
		Exit Function
    End If
	'-----------------------
	'Save function call area
	'-----------------------
	Call DbSave()																			'��: Save db data
	
	FncSave = True
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()
   	frm1.vspdData.ReDraw = False
    	
    If frm1.vspdData.MaxRows < 1 then Exit Function
    	
	ggoSpread.Source = frm1.vspdData	
	ggoSpread.CopyRow
	
	MaxSpreadVal frm1.vspdData, C_Seq , frm1.vspdData.ActiveRow
	
	Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)

	frm1.vspdData.Col = C_RCPT_TYPE
	frm1.vspdData.Text = ""

	frm1.vspdData.Col = C_RCPT_TYPE_NM
	frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 
    if frm1.vspdData.MaxRows < 1 then Exit Function

	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow(Byval pvRowcnt)
	Dim imRow
    Dim ii
    Dim iCurRowPos
	
	On Error Resume Next                                                          '��: If process fails
    Err.Clear   
	
    FncInsertRow = False    
    
    If IsNumeric(Trim(pvRowcnt)) Then 
		imRow  = Cint(pvRowcnt)
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
			Call MaxSpreadVal(frm1.vspdData, C_Seq, ii)
		Next  
		.Col = 2																	' �÷��� ���� ��ġ�� �̵�      
		.Row = 	ii - 1
		.Action = 0

        Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)
        .ReDraw = True
	End With        

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call FncPrint()                                              
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev()
	Dim IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                     '�ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
      	    Exit Function
    	End If
    End If    
	
    Call InitVariables()                                                      'Initializes local global variables
    Call InitSpreadSheet()
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
		Exit Function
    End If

	frm1.hCommand.value = "PREV"
	Call DbQuery()
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
	Dim IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then										'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")									'�ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
      	    Exit Function
    	End If
    End If
    
    Call InitVariables																'Initializes local global variables
    Call InitSpreadSheet
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then												'This function check indispensable field
       Exit Function
    End If

	frm1.hCommand.value = "NEXT"
	Call DbQuery()
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call FncExport(parent.C_SINGLEMULTI)										
	    		
	Set gActiveElement = document.activeElement    

End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'=======================================================================================================
Function FncFind() 
    Call FncFind(parent.C_SINGLEMULTI , True)                               
	    		
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

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
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
'=======================================================================================================
Function DbDelete() 
    Dim strVal
    
    DbDelete = False																'��: Processing is NG 
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPrrcptNo=" & Trim(frm1.txtPrrcptNo.value)				'��: ���� ���� ����Ÿ 

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
    
	Call RunMyBizASP(MyBizASP, strVal)												'��: �����Ͻ� ASP �� ���� 
    
    DbDelete = True																	'��: Processing is NG
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()																'���� ������ ���� ���� 
	Call ggoOper.ClearField(Document, "1")									'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
	Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field
	
	Call txtDocCur_OnChange()

	Call InitVariables()																'Initializes local global variables
    Call InitSpreadSheet()
	Call SetDefaultVal()
	Call SetToolbar("1110110100101111")

    frm1.txtPrrcptNo.focus 
	Set gActiveElement = document.activeElement
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
	Dim strVal

	DbQuery = False                                                         
	
	Call LayerShowHide(1)
	
	With frm1
       	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: 
       	strVal = strVal & "&txtPrrcptNo=" & Trim(.txtPrrcptNo.value)				'��ȸ ���� ����Ÿ 
       	strVal = strVal & "&txtCommand=" & Trim(.hCommand.value)
       	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End With

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
	
	Call RunMyBizASP(MyBizASP, strVal)												'�����Ͻ� ASP �� ���� 
	
	DbQuery = True                                                          
	lgQueryOk = True 	
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================================
Function DbQueryOk()																'��ȸ ������ ������� 
	Dim strTemp, varData

  	lgQueryOk=True

	If frm1.vspdData.MaxRows > 0 Then 
		Call SetSpreadLock()

		frm1.vspdData.Row = 1
		frm1.vspdData.Col = C_RCPT_TYPE
		varData = frm1.vspdData.text
		Call subVspdSettingChange(C_RCPT_TYPE,1,frm1.vspdData.Maxrows)
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
	
	Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field	
	Call SetToolbar("1111111111111111")									'��ư ���� ���� 
<% If gIsShowLocal <> "N" Then %>	
	strTemp = frm1.txtXchRate.Text
<% Else %>
	strTemp = frm1.txtXchRate.Value
<% End If %>
	
<% If gIsShowLocal <> "N" Then %>			
	frm1.txtXchRate.Text = strTemp
<% Else %>
	frm1.txtXchRate.Value = strTemp
<% End If %>	
	If Trim(frm1.txtVatType.Value) <>"" Then
		Call CommonQueryRs (" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("B9001", "''", "S") & " And Minor_cd =  " & FilterVar(frm1.txtVatType.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		frm1.txtVatTypeNm.value=replace(lgF0,chr(11),"")
	End If
	
	Call txtDeptCd_OnChange()  
	Call txtDocCur_OnChange()
	Call CheckNextPrev()
	
	lgBlnFlgChgValue = False
	lgQueryOk=false	
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	Dim IntRows 
	Dim IntCols 
	Dim lGrpcnt 
	Dim strVal
	Dim strDel
	
	DbSave = False                                                          
	
	On Error Resume Next                                                   
	Err.Clear 
		
	Call LayerShowHide(1)
	
	strVal = ""
	strDel = ""
	
	With frm1
		.txtMode.value = parent.UID_M0002											'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode											'��: �ű��Է�/���� ���� 
	End With
	'-----------------------
	'Data manipulate area
	'-----------------------
	' Data ���� ��Ģ 
	' 0: Flag , 1: Row��ġ, 2~N: �� ����Ÿ 
	lGrpCnt = 1
	
	With frm1.vspdData
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.InsertFlag	'Create
					strVal = strVal & "C" & parent.gColSep & IntRows & parent.gColSep
				Case ggoSpread.UpdateFlag	'Update
					strVal = strVal & "U" & parent.gColSep & IntRows & parent.gColSep
				Case ggoSpread.DeleteFlag	'Delete
					strDel = strDel & "D" & parent.gColSep & IntRows & parent.gColSep
			End Select
			
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.Col = C_SEQ
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_RCPT_TYPE
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_RCPT_ACCT
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_AMT
					strVal = strVal & UNIConvNum(.Text,0) & parent.gColSep
					.Col = C_LOC_AMT
					strVal = strVal & UNIConvNum(.Text,0) & parent.gColSep
					.Col = C_BANK_CD
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_BANK_ACCT
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_NOTE_NO
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_STTL_DESC
					strVal = strVal & Trim(.Text) & parent.gRowSep				    '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 
					
					lGrpCnt = lGrpCnt + 1

				Case ggoSpread.DeleteFlag
					.Col = C_SEQ
					strDel = strDel & Trim(.Text) & parent.gRowSep				    '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 
					
					lGrpcnt = lGrpcnt + 1             
			End Select
		Next
	End With

	frm1.txtMaxRows.value = lGrpCnt-1												'��: Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread.value = strDel & strVal											'��: Spread Sheet ������ ���� 

	'���Ѱ����߰� start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'���Ѱ����߰� end
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: ���� �����Ͻ� ASP �� ���� 
	
	DbSave = True                                                           
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function DbSaveOk()																	'��: ���� ������ ���� ���� 
   	lgBlnFlgChgValue = False	
	frm1.vspdData.MaxRows = 0

	Call FncQuery()
End Function





'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************





'==========================================================================================
'   Event Name : subVspdSettingChange
'   Event Desc : 
'==========================================================================================
Sub subVspdSettingChange(ByVal Col , ByVal Row,  ByVal Row2)	
	dim intIndex
	dim strval
	Dim lRow
	

	For lRow = Row To Row2
		frm1.vspddata.col = C_RCPT_TYPE
		frm1.vspddata.Row = lRow
		strval = frm1.vspdData.Text
		
		IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
			Select Case UCase(lgF0)
				Case "DP" & Chr(11)           '�������� ��� ' 						
					ggoSpread.SSSetRequired	C_BANK_ACCT,		 lRow, lRow			
					ggoSpread.SpreadUnLock   C_BANK_ACCT,         lRow, C_BANK_ACCT
					ggoSpread.SpreadUnLock   C_BANK_ACCT_PB,      lRow, C_BANK_ACCT_PB
					ggoSpread.SSSetEdit	    C_BANK_ACCT, "�������ڵ�", 25, 0, lRow, 30    
					ggoSpread.SSSetRequired	C_BANK_ACCT,      lRow, lRow	
					ggoSpread.SpreadLock     C_NOTE_NO,		 lRow, C_NOTE_NO,lRow   '������ȣ protect
					ggoSpread.SSSetProtected C_NOTE_NO,       lRow, lRow						
					ggoSpread.SpreadLock     C_NOTE_NO_PB,  lRow, C_NOTE_NO_PB,lRow          
				Case "NO" & Chr(11) 						
					ggoSpread.SpreadUnLock   C_NOTE_NO,        lRow, C_NOTE_NO,       lRow
					ggoSpread.SpreadUnLock   C_NOTE_NO_PB,   lRow, C_NOTE_NO_PB,  lRow
					ggoSpread.SpreadLock     C_BANK_ACCT,      lRow, C_BANK_ACCT,     lRow   
					ggoSpread.SpreadLock     C_BANK_ACCT_PB, lRow, C_BANK_ACCT_PB,lRow
					ggoSpread.SSSetProtected C_BANK_ACCT,      lRow, lRow								
					ggoSpread.SSSetEdit      C_NOTE_NO, "������ȣ", 30, 0, lRow, 30	
					ggoSpread.SSSetRequired  C_NOTE_NO,        lRow, lRow
				Case Else 
					ggoSpread.SpreadLock     C_BANK_ACCT,      lRow, C_BANK_ACCT,     lRow   			
					ggoSpread.SpreadLock     C_BANK_ACCT_PB, lRow, C_BANK_ACCT_PB,lRow
					ggoSpread.SSSetProtected C_BANK_ACCT,      lRow, lRow							
					ggoSpread.SpreadLock     C_NOTE_NO,        lRow, C_NOTE_NO,     lRow
					ggoSpread.SpreadLock     C_NOTE_NO_PB,   lRow, C_NOTE_NO_PB,lRow		
					ggoSpread.SSSetProtected C_NOTE_NO,        lRow, lRow													
			End Select
		End if
	Next	
End Sub	

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' �����ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtPrrcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' �����ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' û��ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtSttlAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' �ܾ� 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' �ΰ����ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec	
		' ȯ�� 
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtDocCur.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' �ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_AMT,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

Sub CheckNextPrev() 
	Dim IntRetCD

	Select Case Trim(frm1.txtAfterLookUp.value)
		Case "D"
		Case "900012"
			IntRetCD = DisplayMsgBox("900012","X","X","X") 
		Case "900011"				
			IntRetCD = DisplayMsgBox("900011","X","X","X") 
	End Select
End Sub


'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************




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


'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim ARow, ACol
	
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	With frm1.vspdData
		ARow = .ActiveRow
		ACol = .ActiveCol
		
		If (Col = C_RCPT_TYPE) Or (Col = C_RCPT_TYPE_NM) Then
			.Col = C_RCPT_TYPE
			If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND MINOR_CD =  " & FilterVar(.Text , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
				Select Case UCase(lgF0)
					Case "DP" & Chr(11)
						.Col = C_NOTE_NO
						If (.Text <> "") Then .Text = ""
						ggoSpread.SSSetRequired		C_BANK_CD,		Row,	Row
						ggoSpread.SSSetRequired		C_BANK_ACCT,	Row,	Row
						ggoSpread.SSSetProtected	    C_NOTE_NO,		Row,	Row
						ggoSpread.SpreadUnLock		C_BANK_PB,		Row,	C_BANK_PB,	Row
						ggoSpread.SpreadUnLock		C_BANK_ACCT_PB,	Row,	C_BANK_ACCT_PB,	Row
						ggoSpread.SSSetProtected	    C_NOTE_NO_PB,	Row,	Row
					Case "NO" & Chr(11)
						.Col = C_BANK_CD
						If (.Text <> "") Then .Text = ""
						.Col = C_BANK_ACCT
						If (.Text <> "") Then .Text = ""
						ggoSpread.SSSetProtected	C_BANK_CD,		    Row,	Row
						ggoSpread.SSSetprotected	C_BANK_ACCT,	    Row,	Row
						ggoSpread.SpreadUnLock		C_NOTE_NO,		Row,	Row
						ggoSpread.SSSetRequired		C_NOTE_NO,		Row,	Row
						ggoSpread.SSSetProtected	C_BANK_PB,		    Row,	Row
						ggoSpread.SSSetProtected	C_BANK_ACCT_PB,	    Row,	Row
						ggoSpread.SpreadUnLock		C_NOTE_NO_PB,	Row,	C_NOTE_NO_PB,	Row
					Case Else
						.Col = C_BANK_CD
						If (.Text <> "") Then .Text = ""
						.Col = C_BANK_ACCT
						If (.Text <> "") Then .Text = ""
						.Col = C_NOTE_NO
						If (.Text <> "") Then .Text = ""
						ggoSpread.SSSetProtected	C_BANK_CD,		Row,	Row
						ggoSpread.SSSetprotected	C_BANK_ACCT,	Row,	Row
						ggoSpread.SSSetProtected	C_NOTE_NO,		Row,	Row
						ggoSpread.SSSetProtected	C_BANK_PB,		Row,	Row
						ggoSpread.SSSetProtected	C_BANK_ACCT_PB,	Row,	Row
						ggoSpread.SSSetProtected	C_NOTE_NO_PB,	Row,	Row
				End Select
			Else
				.Col = C_BANK_CD
				If (.Text <> "") Then .Text = ""
				.Col = C_BANK_ACCT
				If (.Text <> "") Then .Text = ""
				.Col = C_NOTE_NO
				If (.Text <> "") Then .Text = ""
				ggoSpread.SSSetProtected	C_BANK_CD,		Row,	Row
				ggoSpread.SSSetprotected	C_BANK_ACCT,	Row,	Row
				ggoSpread.SSSetProtected	C_NOTE_NO,		Row,	Row
				ggoSpread.SSSetProtected	C_BANK_PB,		Row,	Row
				ggoSpread.SSSetProtected	C_BANK_ACCT_PB,	Row,	Row
				ggoSpread.SSSetProtected	C_NOTE_NO_PB,	Row,	Row
			End If
			
			frm1.vspdData.Col  = C_RCPT_ACCT
			frm1.vspdData.Text = ""
			frm1.vspdData.Col  = C_RCPT_ACCT_Nm
			frm1.vspdData.Text = ""		
		End If
		
		.Col = ACol
		Select Case Col
			Case C_AMT
				.col=C_LOC_AMT
				.text=""
		End Select
	End With
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
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"	'Split �����ڵ� 
	
	Set gActiveSpdSheet = frm1.vspdData
	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If
    
    Call SetPopupMenuItemInf("1101111111")
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
'======================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strTemp
    Dim intPos1
    Dim bankCode
	Dim intRetCd
	Dim strData
	
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
				Case C_RCPT_TYPE_PB
					.Col = C_RCPT_TYPE
					.Row = Row
					Call OpenPopup(.Text, "RCPT")
				Case C_RCPT_ACCT_PB
					.Col = C_RCPT_ACCT
					.Row = Row
					Call OpenPopup(.Text, "RCPTACCT")
				Case C_BANK_PB
					.Col = C_BANK_CD
					.Row = Row
					Call OpenPopup(.Text, "BANK")
				Case C_BANK_ACCT_PB
					.Col = C_BANK_ACCT
					.Row = Row
					Call OpenPopup(.Text, "BANK_ACCT")
				Case C_NOTE_NO_PB
					.Col = C_NOTE_NO
					.Row = Row
					Call OpenPopupNote(.Text)
				Case Else
			End Select
		End If
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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





'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************





'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtPrrcptDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPrrcptDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtPrrcptDt.Focus
		        
    End If
End Sub

'=======================================================================================================
'   Event Name :txtIssuedDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDt.Action = 7
       	Call SetFocusToDocument("M")
		Frm1.txtIssuedDt.Focus
	
    End If
End Sub

'==========================================================================================
'   Event Name : txtPrrcptDt_Change
'   Event Desc : 
'==========================================================================================
Sub txtPrrcptDt_Change()
    Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtPrrcptDt.Text <> "") Then
	
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtPrrcptDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtDeptCd.value = ""
					frm1.txtDeptNm.value = ""
					frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
		End If
	End If
	
    lgBlnFlgChgValue = True
	Call XchLocRate()
End Sub

'==========================================================================================
'   Event Name : txtXchRate_Change
'   Event Desc : 
'==========================================================================================
Sub txtXchRate_Change()
    lgBlnFlgChgValue = True
	
	if lgQueryOk <> TRUE then 
		Dim ii

		With frm1
			For ii = 1 To .vspdData.MaxRows 
				.vspdData.Row = ii	
				.vspdData.Col = C_LOC_AMT	
				.vspdData.Text = "" 
				 ggoSpread.Source = .vspdData
				 ggoSpread.UpdateRow ii	
			Next	
			.txtVAtLocAmt.text="0"

		End With
	End if
End Sub 

'==========================================================================================
'   Event Name : txtVAtLocAmt_Change
'   Event Desc : 
'==========================================================================================
Sub txtVAtLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtVatAmt_Change()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub  txtVatAmt_Change()
	lgBlnFlgChgValue = True

	If UCase(Trim(frm1.txtDocCur.value)) <> UCase(parent.gCurrency) Then
		frm1.txtVatLocAmt.Text = "0"
	End If
	
	If UNIConvNum(frm1.txtVatAmt.Text,0) <> 0 Or Trim(frm1.txtVatType.value) <> "" Then
		Call ggoOper.SetReqAttr(frm1.txtVatType, "N")
		Call ggoOper.SetReqAttr(frm1.txtVatAmt, "N")
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCD, "N")				
	Else
		Call ggoOper.SetReqAttr(frm1.txtVatType, "D")
		Call ggoOper.SetReqAttr(frm1.txtVatAmt, "D")
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCD, "D")				
	End If
End Sub

'==========================================================================================
'   Event Name : txtVatType_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtVatType_OnChange()
    lgBlnFlgChgValue = True
 
    If Trim(frm1.txtVatType.value) <>"" Or UNIConvNum(frm1.txtVatAmt.Text,0) <> 0  Then
		Call ggoOper.SetReqAttr(frm1.txtVatType, "N")
		Call ggoOper.SetReqAttr(frm1.txtVatAmt, "N")
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCD, "N")		
	Else
		Call ggoOper.SetReqAttr(frm1.txtVatType, "D")
		Call ggoOper.SetReqAttr(frm1.txtVatAmt, "D")
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCD, "D")		
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtBizAreaCD_OnChange()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub  txtBizAreaCD_OnChange()
	lgBlnFlgChgValue = True
	
	If UNIConvNum(frm1.txtVatAmt.Text,0) <> 0 Or Trim(frm1.txtVatType.value) <> ""  Then
		Call ggoOper.SetReqAttr(frm1.txtVatType, "N")
		Call ggoOper.SetReqAttr(frm1.txtVatAmt, "N")
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCD, "N")				
	Else
		Call ggoOper.SetReqAttr(frm1.txtVatType, "D")
		Call ggoOper.SetReqAttr(frm1.txtVatAmt, "D")
		Call ggoOper.SetReqAttr(frm1.txtBizAreaCD, "D")				
	End If
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    
    If lgQueryOk <> True Then
	<% If gIsShowLocal <> "N" Then %>    
			frm1.txtXchRate.Text = "0" 
	<% Else %>
			frm1.txtXchRate.Value = "0" 
	<% End If %>    
	End If	
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If	   
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_nChange
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtPrrcptDt.Text = "") Then    
		Exit sub
    End If

	'----------------------------------------------------------------------------------------
	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtPrrcptDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
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

    lgBlnFlgChgValue = True

End Sub

'==========================================================================================
'   Event Name : txtIssuedDt_Change
'   Event Desc : 
'=========================================================================================
Sub txtIssuedDt_Change()
    lgBlnFlgChgValue = True
End Sub

'===================================== XchLocRate()  ======================================
'	Name : XchLocRate()
'	Description : ��ȭ�� ����ɰ�� ��ȭ�� ���� �ڱ��ݾ� 
'====================================================================================================
Sub XchLocRate()
	Dim ii

	With frm1
		For ii = 1 To .vspdData.MaxRows 
			.vspdData.Row = ii	
			.vspdData.Col = C_LOC_AMT	
			.vspdData.Text = ""    	
			 ggoSpread.Source = .vspdData
			 ggoSpread.UpdateRow ii	
		Next	
		.txtVAtLocAmt.text="0"
		If UCase(Trim(frm1.txtDocCur.Value)) <> UCase(Trim(parent.gCurrency)) Then
			.txtXchRate.Text = "0" 
		Else			
			.txtXchRate.Text = "1" 		
		End If					
	End With
End Sub

Sub chkLimitFg_onchange()
	If frm1.chkLimitFg.checked = True Then
		frm1.txtLimitFg.value = "Y"
	Else
		frm1.txtLimitFg.value = "N"	
	End If
	lgBlnFlgChgValue = True	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>

<!--'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�����ݹ�ȣ</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtPrrcptNo" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="�����ݹ�ȣ" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrrcptNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopupPR"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>����������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPrrcptType" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="����������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrrcptType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup('','PRRCPTTYPE')">&nbsp;<INPUT TYPE=TEXT NAME="txtPrrcptTypeNm" SIZE=25 tag="24" ALT="������������"></TD>
								<TD CLASS="TD5" NOWRAP><LABEL FOR=chkConfFg>���Ű���</LABEL></TD>
								<TD CLASS="TD6" NOWRAP><INPUT type="checkbox" CLASS="STYLE CHECK"  NAME=chkLimitFg ID=chkLimitFg tag="1" onclick=chkLimitFg_onchange()></TD>
							</TR>						
							<TR>
								<TD CLASS="TD5" NOWRAP>�Ա�����</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtPrrcptDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�Ա�����" tag="22X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�ŷ�ó�ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 'BP')">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24" ALT="�ŷ�ó��"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�μ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�μ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenpopupDept(frm1.txtDeptCd.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 tag="24" ALT="ȸ��μ���"></TD>
								<TD CLASS="TD5" NOWRAP>������ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" SIZE=30 MAXLENGTH=30 tag="24XXXU" ALT="������ȣ" ></TD>
							</TR>
<%	If gIsShowLocal <> "N" Then	%>														
							<TR>
								<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" TYPE="Text" SIZE=10 MAXLENGTH=3 tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.value, 'CURR')"></TD>
								<TD CLASS="TD5" NOWRAP>ȯ��</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 80px" title=FPDOUBLESINGLE ALT="ȯ��" tag="21X5Z" id=fpDoubleSingle1></OBJECT>');</SCRIPT></TD>
							</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtXchRate" TABINDEX="-1">
<%	End If %>														
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ݾ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtPrrcptAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�" tag="24X2" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
<%	If gIsShowLocal <> "N" Then	%>	                            
								<TD CLASS="TD5" NOWRAP>�����ݾ�(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtPrrcptLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle3></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtPrrcptLocAmt" TABINDEX="-1">
<%	End If %>							
								<TD CLASS="TD5" NOWRAP>�����ݾ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtClsAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�" tag="24X2" id=fpDoubleSingle4></OBJECT>');</SCRIPT></TD>
<%	If gIsShowLocal <> "N" Then	%>	                            
								<TD CLASS="TD5" NOWRAP>�����ݾ�(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtClsLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle5></OBJECT>');</SCRIPT></TD>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtClsLocAmt" TABINDEX="-1">
<%	End If %>								                            
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>û��ݾ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSttlAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="û��ݾ�" tag="24X2" id=fpDoubleSingle6></OBJECT>');</SCRIPT></TD>
<%	If gIsShowLocal <> "N" Then	%>	                            
								<TD CLASS="TD5" NOWRAP>û��ݾ�(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSttlLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="û��ݾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle7></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtSttlLocAmt" TABINDEX="-1">
<%	End If %>								                            							
								<TD CLASS="TD5" NOWRAP>�ܾ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�ܾ�" tag="24X2" id=fpDoubleSingle8></OBJECT>');</SCRIPT></TD>
<%	If gIsShowLocal <> "N" Then	%>	                            	                            
								<TD CLASS="TD5" NOWRAP>�ܾ�(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�ܾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle9></OBJECT>');</SCRIPT></TD>
<!--	                            <INPUT TYPE=HIDDEN NAME="txtVatType" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtVatAmt" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtVAtLocAmt" TABINDEX="-1">-->
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ΰ�������</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="�ΰ�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType()">&nbsp;<INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="24" ALT="�ΰ�������"></TD>
								<TD CLASS="TD5" NOWRAP>������Ʈ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProjectNo"  SIZE=14 MAXLENGTH=25 TAG="21xxxU" ALT="������Ʈ"></TD>	                     
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>�ΰ����ݾ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtVatAmt style="HEIGHT: 20px; Right: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="�ΰ����ݾ�" tag="21X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>�ΰ����ݾ�(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtVAtLocAmt style="HEIGHT: 20px; Right: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="�ΰ����ݾ�(�ڱ�)" tag="21X2Z" id=fpDoubleSingle3></OBJECT>');</SCRIPT></TD>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtBalLocAmt" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtVatType" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtVatAmt" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtVAtLocAmt" TABINDEX="-1">
<%	End If %>		
								
							</TR>	
							<TR>
								<TD CLASS="TD5" NOWRAP>���ݽŰ�����</TD>
								<TD CLASS="TD6" NOWRAP ><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 ALT="���ݽŰ�����" tag="21XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup('','BizArea')">
														<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=25 MAXLENGTH=50  ALT="���ݽŰ�����" tag="24" ></TD>
								<TD CLASS="TD5" NOWRAP>��꼭����</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssuedDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��꼭����" tag="11X1"></OBJECT>');</SCRIPT>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>������ǥ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="������ǥ��ȣ"></TD>
								<TD CLASS="TD5" NOWRAP>ȸ����ǥ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="ȸ����ǥ��ȣ"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtPrrcptDesc" SIZE=90 MAXLENGTH=128 tag="2X" ALT="���"></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> tag="2" HEIGHT="100%" name=vspdData width="100%" id=fpSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"        tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     tag="24" TABINDEX="-1">
<INPUT TYPE=TEXT NAME="hDocumentNo1"     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCommand"       tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtLimitFg"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtAfterLookUp" tag="24" TABINDEX="-1">
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

