
<%@ LANGUAGE="VBSCRIPT" %>

<!--===================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : RECEIPT
'*  3. Program ID		    : f5101ma1
'*  4. Program Name         : �������� ��� 
'*  5. Program Desc         : �������� ��� ���� ���� ��ȸ 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/10/11
'*  8. Modified date(Last)  : 2002/03/25
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Soo Min, Oh
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************
'==========================================  1.1.1 Style Sheet  ==========================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '��: indicates that All variables must be declared in advance
                                                          '��: indicates that All variables must be declared in advance 

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
<%
StrSvrDate = GetSvrDate
%>

 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Const BIZ_PGM_ID  = "f5101mb1.asp"										'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 = "f5101mb2.asp"										'��: �����Ͻ� ���� ASP�� 
Const JUMP_PGM_ID_NOTE_CHG = "f5107ma1"									'���������� 

Dim C_GL_DT	
Dim C_SEQ	
Dim C_DR_CR_FG
Dim C_DR_CR_FG_NM
Dim C_ITEM_AMT
Dim C_ACCT_CD
Dim C_ACCT_NM
Dim C_ITEM_DESC
Dim C_GL_NO	
Dim C_TEMP_GL_NO
Dim C_COL_END

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
 '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE   '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '��: Indicates that no value changed
    lgIntGrpCount = 0           '��: Initializes Group View Size

	lgStrPrevKey = ""
	lgLngCurRows = 0                            'initializes Deleted Rows Count
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'��: ����� ���� �ʱ�ȭ 

    lgSortKey = 1
	lgPageNo  = ""

	frm1.hOrgChangeId.value = parent.gChangeOrgId		
    lgBlnFlgChgValue = False
End Sub


sub initSpreadPosVariables()

	C_GL_DT		= 1
	C_SEQ			= 2
	C_DR_CR_FG	= 3
	C_DR_CR_FG_NM	= 4
	C_ITEM_AMT	= 5
	C_ACCT_CD		= 6
	C_ACCT_NM		= 7
	C_ITEM_DESC	= 8
	C_GL_NO		= 9
	C_TEMP_GL_NO	= 10
	C_COL_END		= 11

end sub
'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

 '******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

 if frm1.cboNoteFg.length > 0 then
       frm1.cboNoteFg.selectedindex = 0
    end if

	frm1.txtIssueDt.Text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtDueDt.Text   = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
	
	frm1.hOrgChangeId.value = Parent.gChangeOrgId
	frm1.txtNoteNoQry.focus 

End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()

    call initSpreadPosVariables()
    
    Dim sList
    
    With frm1
    
		.vspdData.MaxCols = C_COL_END
		.vspdData.Col = .vspdData.MaxCols	:	.vspdData.ColHidden = True				'��: ������Ʈ�� ��� Hidden Column
		.vspdData.MaxRows = 0
		ggoSpread.Source = .vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
        Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit C_SEQ, "����", 8, , , 3
		ggoSpread.SSSetDate C_GL_DT, "����", 12, , Parent.gDateFormat
		ggoSpread.SSSetCombo C_DR_CR_FG, "���뱸��", 12
		ggoSpread.SSSetCombo C_DR_CR_FG_NM, "���뱸��", 12
		ggoSpread.SSSetFloat C_ITEM_AMT, "�ݾ�", 17, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit C_ACCT_CD, "�����ڵ�", 15, , , 20
		ggoSpread.SSSetEdit C_ACCT_NM, "������", 25, , , 30
		ggoSpread.SSSetEdit C_ITEM_DESC, "����", 35, , , 128
		ggoSpread.SSSetEdit C_GL_NO, "��ǥ��ȣ", 15, , , 18
		ggoSpread.SSSetEdit C_TEMP_GL_NO, "������ǥ��ȣ", 15, , , 18
 
        Call ggoSpread.SSSetColHidden(C_SEQ,C_SEQ,True)
        Call ggoSpread.SSSetColHidden(C_DR_CR_FG,C_DR_CR_FG,True)
        Call ggoSpread.SSSetColHidden(C_ACCT_CD,C_ACCT_CD,True)
        Call SetSpreadLock                                              '�ٲ�κ� 

    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		.ReDraw = False
		 ggoSpread.SpreadLockWithOddEvenRowColor()    
		.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1

		.vspdData.ReDraw = False

		.vspdData.ReDraw = True
    
    End With

End Sub



Function InitCombo()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	                   'Select                 From        Where                Return value list  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteFg ,lgF0  ,lgF1  ,Chr(11))
    
    
    '��������            'Select                 From        Where                Return value list  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteSts ,lgF0  ,lgF1  ,Chr(11))
    
    '�������           'Select                 From        Where                Return value list  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1005", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    Call SetCombo2(frm1.cboPlace ,lgF0  ,lgF1  ,Chr(11))

	
    '�ڼ�Ÿ������       'Select                 From        Where                Return value list  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1009", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboRcptFg ,lgF0  ,lgF1  ,Chr(11))

    Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("A1012", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DR_CR_FG
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DR_CR_FG_NM    

End Function


Function InitCombobox()

	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("A1012", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DR_CR_FG
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DR_CR_FG_NM    

End Function
 '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenNoteInfo()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	iCalledAspName = AskPRAspName("f5101ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f5101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		frm1.txtNoteNoQry.focus
		Exit Function
	Else
		frm1.txtNoteNoQry.value  = arrRet(0)
		frm1.txtNoteNoQry.focus
	End If	

End Function

 '------------------------------------------  OpenPopUp()  ---------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function  OpenPopUpNoteNo()
	Dim strNoteFg
	Dim IntRetCd
	
	strNoteFg = frm1.cboNoteFg.Value
	
	If strNoteFg = "" Then
	    IntRetCD = DisplayMsgBox("141327","x","x","x")	'���������� ���� �Է��Ͻʽÿ�.
		Exit Function		      	
	ElseIf strNoteFg = "D3" Then 
		Call OpenPopUp(frm1.txtNoteNo.Value, 1) '���޾��� 
	Else  
	    IntRetCD = DisplayMsgBox("141220","x","x","x")	'������ȣ�� ���� �Է����ֽʽÿ�.
		Exit Function		
    End If	
End Function  

'==================================================================================
'	Name : OpenPopUp()
'	Description : �����˾� ���� 
'==================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = parent.gChangeOrgId
	
	Select Case iWhere
		Case 1		' ������ȣ 
			If frm1.txtNoteNo.className = Parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = "���޾��� �˾�"				' �˾� ��Ī 
			arrParam(1) = "F_NOTE_NO A, B_BANK B"					' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.STS = " & FilterVar("NP", "''", "S") & "  AND A.BANK_CD = B.BANK_CD"							' Where Condition
			arrParam(5) = "���޾�����ȣ"					' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Note_NO"						' Field��(0)
			arrField(1) = "B.BANK_CD"						' Field��(1)
			arrField(2) = "B.BANK_NM"						' Field��(2)
    
			arrHeader(0) = "���޾�����ȣ"					' Header��(0)
			arrHeader(1) = "��������"						' Header��(1)
			arrHeader(2) = "���������"						' Header��(2)
			
		Case 5		' �������� 
			arrParam(0) = "���� �˾�"	' �˾� ��Ī 
			arrParam(1) = "B_BANK"			 			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�����ڵ�"					' �����ʵ��� �� ��Ī 

			arrField(0) = "BANK_CD"						' Field��(0)
			arrField(1) = "BANK_NM"					' Field��(1)
    
			arrHeader(0) = "�����ڵ�"					' Header��(0)
			arrHeader(1) = "�����"						' Header��(1)

	End Select
  
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'�μ��ڵ� 
	arrParam(1) = frm1.txtIssueDt.Text			'��¥(Default:������)
	arrParam(2) = "1"							'�μ�����(lgUsrIntCd)
	arrParam(3) = "F"							'�μ�����(lgUsrIntCd)
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCD.focus
		Exit Function
	End If
	
	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	frm1.txtIssueDt.text = arrRet(3)
	Call txtDeptCD_Change()
	frm1.txtDeptCD.focus
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
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
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function



Function OpenPopuptempGL()

	Dim arrRet
	Dim arrParam(8)	
    Dim iCalledAspName
	
	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_TEMP_GL_NO
			arrParam(0) = Trim(.Text)	'������ǥ��ȣ 
			arrParam(1) = ""			'Reference��ȣ 
		End If
	End With	
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

Function OpenPopupGL()


	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_Gl_No
			arrParam(0) = Trim(.Text)	'��ǥ��ȣ 
			arrParam(1) = ""			'Reference��ȣ 
		End If
	End With
	
	IsOpenPop = True
   
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function


 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1		' ������ȣ 
				.txtNoteNo.focus
			Case 2		' �μ� 
				.txtDeptCD.focus
			Case 3		' �ڽ�Ʈ��Ÿ 
				.txtCostCD.focus
			Case 4		' �ŷ�ó 
				.txtBpCd.focus
			Case 5		' �������� 
				.txtBankCd.focus
		End Select

	End With
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1		' ������ȣ 
				.txtNoteNo.value = arrRet(0)
				.txtBankCd.value = arrRet(1)
				.txtBankNM.value = arrRet(2)
				.txtNoteNo.focus
				lgBlnFlgChgValue = True
			Case 4		' �ŷ�ó 
				.txtBpCd.value = arrRet(0)
				.txtBpNM.value = arrRet(1)
				.txtBpCd.focus
				lgBlnFlgChgValue = True
			Case 5		' �������� 
				.txtBankCd.value = arrRet(0)
				.txtBankNM.value = arrRet(1)
				.txtBankCd.focus
				lgBlnFlgChgValue = True
		End Select

	End With
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
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
					
			call MainQuery()
	
		Case JUMP_PGM_ID_NOTE_CHG	'���������� 
			strTemp = frm1.txtNoteNoQry.value 			
			Call WriteCookie("NOTE_NO", strTemp)

		Case Else
			Exit Function
	End Select
End Function	

'========================================================================================================
'	Desc : ȭ���̵� 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ����Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call CookiePage(strPgmId)
    Call PgmJump(strPgmId)
End Function

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
 '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029							'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.ClearField(Document, "1")      '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")		'��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")		'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet                                                        'Setup the Spread sheet
	Call InitCombo
 
    Call SetDefaultVal
 
	Call cboNoteFg_OnChange
    Call InitVariables							'��: Initializes local global variables
    
	Call SetToolbar("1110100000001111")
    Call CookiePage("FORM_LOAD")

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


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
            C_GL_DT	= iCurColumnPos(1)
            C_SEQ	= iCurColumnPos(2)
            C_DR_CR_FG = iCurColumnPos(3)
            C_DR_CR_FG_NM = iCurColumnPos(4)
            C_ITEM_AMT = iCurColumnPos(5)
            C_ACCT_CD = iCurColumnPos(6)
            C_ACCT_NM = iCurColumnPos(7)
            C_ITEM_DESC = iCurColumnPos(8)
            C_GL_NO	 = iCurColumnPos(9)
            C_TEMP_GL_NO = iCurColumnPos(10)
            C_COL_END = iCurColumnPos(11)
    End Select    
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	For intRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = intRow
		frm1.vspdData.Col = C_DR_CR_FG
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_DR_CR_FG_NM
		frm1.vspdData.value = intindex
	Next
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtIssueDt.Focus             
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt_Change()

    Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2


	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtIssueDt.Text <> "") Then
	
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtIssueDt.Text, gDateFormat,""), "''", "S") & "))"

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
End Sub

'=======================================================================================================
'   Event Name :txtDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtDueDt.Focus          
    End If
End Sub

Sub txtDeptCD_Change()

    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii
	If Trim(frm1.txtDeptCd.value) = "" and Trim(frm1.txtIssueDt.Text = "") Then		Exit Sub

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtIssueDt.Text, gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
					
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
		'----------------------------------------------------------------------------------------

     lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDueDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCashRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtNoteAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtSttlAmt_Change()
    lgBlnFlgChgValue = True
End Sub


 '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************
Sub cboNoteFg_OnChange()							'create ���� event (field not clear)
	with frm1
	
		If .cboNoteFg.value = "D3" then         
			Call ElementVisible(.btnNoteNo, 1)
		Else
			Call ElementVisible(.btnNoteNo, 0)
		End if		
			
		Select Case frm1.cboNoteFg.value
			Case "D1"	'�������� 
				.txtCashRate.Text = ""
				Call ggoOper.SetReqAttr(.txtCashRate, "N")	'N:Required, Q:Protected, D:Default				
			Case "D3"	'���޾��� 
				.txtCashRate.Text = ""
				Call ggoOper.SetReqAttr(.txtCashRate, "Q")	'N:Required, Q:Protected, D:Default				
			Case Else
				.txtCashRate.Text = ""
				Call ggoOper.SetReqAttr(.txtCashRate, "Q")	'N:Required, Q:Protected, D:Default
		End Select
	
	End with

	lgBlnFlgChgValue = True
End Sub

Sub cboNoteFg_OnChange1()							'dbqueryok ���� event (field not clear)
	with frm1
		Select Case frm1.cboNoteFg.value
			Case "D1"	'��������				
				Call ggoOper.SetReqAttr(.txtCashRate, "N")	'N:Required, Q:Protected, D:Default				
			Case "D3"	'���޾���			
				Call ggoOper.SetReqAttr(.txtCashRate, "Q")	'N:Required, Q:Protected, D:Default				
			Case Else				
				Call ggoOper.SetReqAttr(.txtCashRate, "Q")	'N:Required, Q:Protected, D:Default
		End Select
	End with
End Sub

Sub cboPlace_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboRcptFg_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtNoteNo_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtDeptCD_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj
	If Trim(frm1.txtDeptCd.value = "") Then		Exit sub
	If Trim(frm1.txtIssueDt.value = "") Then		Exit sub
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtIssueDt.Text, gDateFormat,""), "''", "S") & "))"			

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
		'----------------------------------------------------------------------------------------
	lgBlnFlgChgValue = True
End Sub

Sub txtPublisher_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtBpCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtBankCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtNoteDesc_OnChange()
	lgBlnFlgChgValue = True
End Sub



Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    
    Set gActiveSpdSheet = frm1.vspdData
    
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
	End If
'    Call SetPopupMenuItemInf("0000111111")
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
     Dim iColumnName
    
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgPageNo <> "" Then                         
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
    
End Sub


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing


    '-----------------------
    'Check previous data area
    '----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x") '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")			'��: Clear Contents  Field
    Call SetDefaultVal
    Call InitVariables								'��: Initializes local global variables

	frm1.vspdData.MaxRows = 0
	
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then		'��: This function check indispensable field
       Exit Function
    End If
    
    Call ggoOper.LockField(Document, "N")		'��: This function lock the suitable field

  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery									'��: Query db data
       
    FncQuery = True									'��: Processing is OK
    
    Set gActiveElement = document.activeElement

End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False      '��: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
	If lgBlnFlgChgValue = True Then
	    IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") '�� �ٲ�κ� 
	     If IntRetCD = vbNo Then
	         Exit Function
	     End If
	End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")	'��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")  '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")   '��: Lock  Suitable  Field
    Call SetDefaultVal
    Call cboNoteFg_OnChange
    Call InitVariables						'��: Initializes local global variables
    
	frm1.vspdData.MaxRows = 0

    Call SetToolbar("1110100000000011")										'��: ��ư ���� ���� 

    FncNew = True							'��: Processing is OK

	frm1.txtNoteNoQry.focus 
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False														'��: Processing is NG
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        intRetCD = DisplayMsgBox("900002","x","x","x")  '�� �ٲ�κ� 
        'Call MsgBox("��ȸ���Ŀ� ������ �� �ֽ��ϴ�.", vbInformation)
        Exit Function
    End If
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")  '�� �ٲ�κ� 
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete															'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001","x","x","x")  '�� �ٲ�κ� 
		Exit Function
	End If
    
  '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
    
  '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
    Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
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

Function FncInsertRow() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)												'��: ȭ�� ���� 
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                     '��:ȭ�� ����, Tab ���� 
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
    Call InitCombobox()    
    Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
    Err.Clear                                                               '��: Protect system from crashing
    
    DbDelete = False														'��: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNo.value)		'��: ���� ���� ����Ÿ 
    strVal = strVal & "&cboNoteFg=" & Trim(frm1.cboNoteFg.value)		'��: ���� ���� ����Ÿ 

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True                                                         '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call FncNew()
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
Dim strVal
    
    Err.Clear                                                               '��: Protect system from crashing
    
    DbQuery = False                                                         '��: Processing is NG
    
	Call LayerShowHide(1)
	
	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode	=" & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
			strVal = strVal & "&txtNoteNoQry=" & Trim(.hNoteNo.value)		'��: ��ȸ ���� ����Ÿ 
		Else		
			strVal = BIZ_PGM_ID & "?txtMode	=" & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
			strVal = strVal & "&txtNoteNoQry=" & Trim(.txtNoteNoQry.value)		'��: ��ȸ ���� ����Ÿ 
		End If
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgPageNo	=" & lgPageNo         
			strVal = strVal & "&txtMaxRows	=" & .vspdData.MaxRows
	End With

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()							'��: ��ȸ ������ ������� 

 	Call ggoOper.LockField(Document, "Q")		'��: This function lock the suitable field
 	Call cboNoteFg_OnChange1
	Call InitData
	Call SetToolbar("1111100000011111")
	
	lgIntFlgMode = Parent.OPMD_UMODE					'��: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	
End Function



'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 
Dim strVal

    Err.Clear																'��: Protect system from crashing
	DbSave = False															'��: Processing is NG
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value = Parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode

		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk(Byval ptxtNoteNo)			'��: ���� ������ ���� ���� 

    Select Case lgIntFlgMode
		Case Parent.OPMD_CMODE
			frm1.txtNoteNoQry.value = ptxtNoteNo
    End Select

    Call InitVariables
    call MainQuery()

End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopuptempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>������ȣ</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtNoteNoQry" NAME="txtNoteNoQry" SIZE=30 MAXLENGTH=30 tag="12XXXU"ALT="������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteQry" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenNoteInfo"></TD>
								<TR>		
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
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg" NAME="cboNoteFg" ALT="��������" STYLE="WIDTH: 100px" tag="23X"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>������ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtNoteNo" NAME="txtNoteNo" SIZE=30 MAXLENGTH=30  tag="23XXXU" ALT="������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteNo" ALIGN=top TYPE="BUTTON" tag="23X" ONCLICK="vbscript:Call OpenPopUpNoteNo()"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptCD" NAME="txtDeptCD" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="�μ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON"  ONCLICK="vbscript:Call OpenPopUpDept(frm1.txtDeptCD.Value, 2)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 MAXLENGTH=40 STYLE="TEXT-ALIGN: left" tag="24X" ALT="�μ�"></TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteSts" NAME="cboNoteSts" ALT="��������" STYLE="WIDTH: 100px" tag="24X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=10 MAXLENGTH=10   tag="22XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value, 4)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM" NAME="txtBpNM" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="24X" ALT="�ŷ�ó"> </TD>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10   tag="22XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 5)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="����"> </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssueDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="22X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtDueDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="22X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpXRate name=txtCashRate CLASS=FPDS140 title=FPDOUBLESINGLE ALT="������" tag="22X5Z"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtNoteAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�" tag="22X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtSttlAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboPlace" NAME="cboPlace" ALT="�������" STYLE="WIDTH: 132px" tag="2XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>�ڼ�Ÿ������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboRcptFg" NAME="cboRcptFg" ALT="�ڼ�Ÿ������" STYLE="WIDTH: 132px" tag="2XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtPublisher" NAME="txtPublisher" SIZE=15 MAXLENGTH=20 tag="2XX" ALT="������"></TD>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtNoteDesc" NAME="txtNoteDesc" SIZE=40 MAXLENGTH=128  tag="2XX" ALT="���"></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_NOTE_CHG)">����������</A>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT= <%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="2" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteNo" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="horgchangeid" tag="2" TABINDEX="-1">
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

