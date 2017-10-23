
<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1 %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Account Receivable
'*  3. Program ID           : a3112ma.asp
'*  4. Program Name         : ����ä�ǵ�� 
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2003/01/07
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'            1. �� �� �� 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc ����   
' ���: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=vbscript>

Option Explicit                 '��: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
' .Constant�� �ݵ�� �빮�� ǥ��.
' .���� ǥ�ؿ� ����. prefix�� g�� �����.
' .Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

'@PGM_ID
Const BIZ_PGM_QRY_ID = "a3112mb1.asp"							'��: Head Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "a3112mb2.asp"							'��: Save �����Ͻ� ���� ASP�� 
Const BIZ_PGM_DEL_ID = "a3112mb3.asp"


Const TAB1 = 1													'��: Tab�� ��ġ 
Const TAB2 = 2

Dim  IsOpenPop													'Popup
Dim	 lgFormLoad
Dim	 lgQueryOk													' Queryok���� (loc_amt =0 check)
Dim  lgstartfnc

Dim dtToday
dtToday = "<%=GetSvrDate%>"

'======================================================================================================
'            2. Function�� 
'
' ���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
' �������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'               2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'=======================================================================================================

'======================================================================================================
' Name : InitVariables()
' Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE						'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False								'Indicates that no value changed
	lgstartfnc=False
	lgFormLoad=True    
    lgQueryOk= False    
	lgstartfnc=False
	lgFormLoad=True
	lgQueryOk= False

End Sub
'======================================================================================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtArDt.text  =  UniConvDateAToB(dtToday, parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDueDt.text  =  UniConvDateAToB(dtToday, parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtGlDt.text =  UniConvDateAToB(dtToday, parent.gServerDateFormat,gDateFormat)
'	frm1.cboArType.value = "NT" 
	frm1.txtDocCur.value = parent.gCurrency
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	frm1.txtDeptCd.value = parent.gDepart
	
	If gIsShowLocal <> "N" Then
		frm1.txtXchRate.text = "1"
	Else
		frm1.txtXchRate.value = "1"
	End if

	lgBlnFlgChgValue = False								'Indicates that no value changed 
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
'   Function Name : OpenPopUpgl()
'   Function Desc : 
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1) 
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
 
	arrParam(0) = Trim(frm1.txtGlNo.value)											'ȸ����ǥ��ȣ 
	arrParam(1) = ""																'Reference��ȣ 

	IsOpenPop = True
	  
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
End Function
 '------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtArDt.Text
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = "F"									' �������� ���� Condition  
	
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
				.txtDeptCD.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtArDt.text = arrRet(3)
				call txtDeptCd_OnBlur()  
				.txtDeptCd.focus
        End Select
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

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "B"							'B :���� S: ���� T: ��ü 
	Select Case iWhere
		Case 3
			arrParam(5) = "SOL"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
		Case 9
			arrParam(5) = "INV"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
		Case 4
			arrParam(5) = "PAYER"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
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

	If IsOpenPop = True Then Exit Function 
 
	Select Case iWhere
		Case 0
			arrParam(0) = "ä�������˾�"										' �˾� ��Ī 
			arrParam(1) = "A_OPEN_AR A,B_ACCT_DEPT B,B_BIZ_PARTNER C"				' TABLE ��Ī 
			arrParam(2) = ""														' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = "A.DEPT_CD = B.DEPT_Cd AND A.DEAL_BP_CD = C.BP_CD AND A.AR_TYPE = " & FilterVar("NR", "''", "S") & "  "         ' Where Condition
			arrParam(5) = "ä�ǹ�ȣ"   
 
			arrField(0) = "A.Ar_NO"													' Field��(0)
			arrField(1) = "CONVERT(VARCHAR(40),A.Ar_DT)"							' Field��(1)
			arrField(2) = "B.DEPT_NM"												' Field��(2)
			arrField(3) = "A.DOC_CUR"												' Field��(3) 
			arrField(4) = "C.BP_FULL_NM"											' Field��(4) 
			arrField(5) = "CONVERT(VARCHAR(15),A.Ar_AMT)"							' Field��(5)
			arrField(6) = "CONVERT(VARCHAR(15),A.VAT_AMT)"							' Field��(6)
			 
			arrHeader(0) = "ä�ǹ�ȣ"											' Header��(0)
			arrHeader(1) = "ä����"												' Header��(1)
			arrHeader(2) = "�μ���"												' Header��(2)
			arrHeader(3) = "�ŷ���ȭ"											' Header��(3)
			arrHeader(4) = "�ŷ�ó��"											' Header��(4)
			arrHeader(5) = "ä�Ǳݾ�"											' Header��(5)
			arrHeader(6) = "�ΰ����ݾ�"											' Header��(6)
		Case 1
			arrParam(0) = "�����ڵ��˾�"										' �˾� ��Ī 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE ��Ī 
			arrParam(2) = Trim(strCode)												' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD " & _ 
			    "and C.trans_type = " & FilterVar("ar005", "''", "S") & "  and C.jnl_cd = " & FilterVar("AR", "''", "S") & " "					' Where Condition
			arrParam(5) = "�����ڵ�"											' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"												' Field��(0)
			arrField(1) = "A.Acct_NM"												' Field��(1)
			arrField(2) = "B.GP_CD"													' Field��(2)
			arrField(3) = "B.GP_NM"													' Field��(3)
		 
			arrHeader(0) = "�����ڵ�"											' Header��(0)
			arrHeader(1) = "�����ڵ��"											' Header��(1)
			arrHeader(2) = "�׷��ڵ�"											' Header��(2)
			arrHeader(3) = "�׷��"												' Header��(3)
		Case 3
			arrParam(0) = "�ֹ�ó�˾�"						' �˾� ��Ī 
			arrParam(1) = "b_biz_partner"						' TABLE ��Ī 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "�ֹ�ó"			
	
			arrField(0) = "BP_CD"								' Field��(0)
			arrField(1) = "BP_NM"								' Field��(1)
    
    
			arrHeader(0) = "�ֹ�ó"							' Header��(0)
			arrHeader(1) = "�ֹ�ó��"						' Header��(1)
		Case 4
			If UCase(frm1.txtPayBpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
			arrParam(0) = "����ó�˾�"						' �˾� ��Ī 
			arrParam(1) = "b_biz_partner"						' TABLE ��Ī 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "����ó"			
	
			arrField(0) = "BP_CD"								' Field��(0)
			arrField(1) = "BP_NM"								' Field��(1)
    
    		arrHeader(0) = "����ó"							' Header��(0)
			arrHeader(1) = "����ó��"						' Header��(1)		Case 5       
			
		Case 5
			arrParam(0) = "������˾�"											' �˾� ��Ī 
			arrParam(1) = "B_Biz_AREA"												' TABLE ��Ī 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = ""														' Where Condition
			arrParam(5) = "�����"		
 
			arrField(0) = "Biz_AREA_CD"												' Field��(0)
			arrField(1) = "Biz_AREA_NM"												' Field��(1)    
			 
			arrHeader(0) = "�����"												' Header��(0)
			arrHeader(1) = "������"											' Header��(1)
		Case 8
			arrParam(0) = "�ŷ���ȭ�˾�"										' �˾� ��Ī 
			arrParam(1) = "b_currency"												' TABLE ��Ī 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = ""														' Where Condition
			arrParam(5) = "�ŷ���ȭ"    
 
			arrField(0) = "CURRENCY"												' Field��(0)
			arrField(1) = "CURRENCY_DESC"											' Field��(1)
			 
			arrHeader(0) = "�ŷ���ȭ"											' Header��(0)
			arrHeader(1) = "�ŷ���ȭ��"											' Header��(1)    
		Case 9
			arrParam(0) = "���ݰ�꼭����ó�˾�"						' �˾� ��Ī 
			arrParam(1) = "b_biz_partner"						' TABLE ��Ī 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "���ݰ�꼭����ó"			
	
			arrField(0) = "BP_CD"								' Field��(0)
			arrField(1) = "BP_NM"								' Field��(1)
    
    
			arrHeader(0) = "���ݰ�꼭����ó"							' Header��(0)
			arrHeader(1) = "���ݰ�꼭����ó��"						' Header��(1)
		


		Case 10
			If  UCase(frm1.txtPayMethCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
 
			arrHeader(0) = "�������"											' Header��(0)
			arrHeader(1) = "���������"											' Header��(1)
			arrHeader(2) = "Reference"
			 
			arrField(0) = "B_Minor.MINOR_CD"										' Field��(0)
			arrField(1) = "B_Minor.MINOR_NM"										' Field��(1)
			arrField(2) = "b_configuration.REFERENCE"
			 
			arrParam(0) = "�������"											' �˾� ��Ī 
			arrParam(1) = "B_Minor,b_configuration"									' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtPayMethCd.Value)								' Code Condition
		 
			arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9004", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
			              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"  
			arrParam(5) = "�������"											' TextBox ��Ī 
		Case 11
			if Trim(frm1.txtPayMethCd.Value) = "" then
				Call DisplayMsgBox("205152","X" , "�������","X")
				Exit Function
			End if

			If UCase(frm1.txtPayTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
 
			arrParam(0) = "�Ա�����"											' �˾� ��Ī 
			arrParam(1) = "B_MINOR,B_CONFIGURATION," _
				& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & " "_
				& "And MINOR_CD= " & FilterVar(frm1.txtPayMethCd.value, "''", "S") & " And SEQ_NO>=2)C" ' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtPayTypeCd.value)								' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
			      & "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & " ," & FilterVar("R", "''", "S") & " )"					' Where Condition
			   
			arrParam(5) = "�Ա�����"											' TextBox ��Ī 
	 
			arrField(0) = "B_MINOR.MINOR_CD"										' Field��(0)
			arrField(1) = "B_MINOR.MINOR_NM"										' Field��(1)
			  
			arrHeader(0) = "�Ա�����"											' Header��(0)
			arrHeader(1) = "�Ա�������"											' Header��(1)  
	End Select    
 
	IsOpenPop = True
	 
	If iwhere = 0 Then  
		iCalledAspName = AskPRAspName("a3112ra1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3112ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
	   
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
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
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtArNo.focus
			Case 1 
				.txtAcctCd.focus
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
		End Select    
	End With
 
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If 
End Function

'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================

'======================================================================================================
'            3. Event�� 
' ���: Event �Լ��� ���� ó�� 
' ����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'=======================================================================================================

'======================================================================================================
' Name : Form_Load()
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub  Form_Load()
    Call LoadInfTB19029()															'Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
										parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
                         
    Call ggoOper.LockField(Document, "N")											'Lock  Suitable  Field    
    Call InitVariables()															'Initializes local global variables    

    Call SetToolbar("1110100000001111")												'��ư ���� ���� 
	Call SetDefaultVal()
 
	frm1.txtArNo.focus
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'=======================================================================================================
'   Event Name : txtArDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtArDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtArDt.Action = 7   
        Call  txtArDt_OnBlur()    
        Call SetFocusToDocument("M")
		Frm1.txtArDt.Focus
		
    End If
End Sub
'==========================================================================================
'   Event Name : txtArDt_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtArDt_OnBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
   If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
	
				If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtArDt.Text <> "") Then
					'----------------------------------------------------------------------------------------
						strSelect	=			 " Distinct org_change_id "    		
						strFrom		=			 " b_acct_dept(NOLOCK) "		
						strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
						strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
						strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
						strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtArDt.Text, gDateFormat,""), "''", "S") & "))"			
	
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


'=======================================================================================================
'   Event Name : txtDealBpCd_onChange()
'   Event Desc :  
'=======================================================================================================
Sub  txtDealBpCd_onChange()

    lgBlnFlgChgValue = True
	If lgIntFlgMode <> parent.OPMD_UMODE Then 		
		frm1.txtPayBpCd.value = frm1.txtDealBpCd.value
		frm1.txtPayBpNm.value = frm1.txtDealBpNm.value
		frm1.txtReportBpCd.value = frm1.txtDealBpCd.value
		frm1.txtReportBpNm.value = frm1.txtDealBpNm.value
	End if

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
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
	End if	
		'----------------------------------------------------------------------------------------
End Sub

'=======================================================================================================
'   Event Name : txtArDt_Change()
'   Event Desc : 
'=======================================================================================================
Sub  txtArDt_Change() 
    lgBlnFlgChgValue = True

    If lgQueryOk <> True Then
		frm1.txtNetLocAmt.text = "0"
	End if
End Sub
'=======================================================================================================
'   Event Name : txTblDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txTblDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7     
        Call SetFocusToDocument("M")
		Frm1.txtDueDt.Focus           
    End If
End Sub
'=======================================================================================================
'   Event Name : txtGlDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGlDt.Action = 7   
		Call SetFocusToDocument("M")
		Frm1.txtGlDt.Focus                       
    End If
End Sub

'=======================================================================================================
'   Event Name : txtGlDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtGlDt_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
'   Event Name : txtDueDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtDueDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInvDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtInvDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtCashAmt_Change()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub  txtCashAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtCashLocAmt_Change()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub  txtCashLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPrRcptAmt_Change()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub  txtPrRcptAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPrRcptLocAmt_Change()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub  txtPrRcptLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtNetAmt_Change()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub  txtNetAmt_Change()
	lgBlnFlgChgValue = True

	If lgQueryOk <> True Then
		frm1.txtNetLocAmt.text = "0"
	End If	
End Sub

'=======================================================================================================
'   Event Name : txtNetLocAmt_Change()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub  txtNetLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPayDur_Change()
'   Event Desc : Single�� �����ʵ尡 �ٲ������ check�Ѵ�.
'=======================================================================================================
Sub txtPayDur_Change()
	lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'�����ȣ �Է½� �������� �Է��ʼ� 
'======================================================================================================
Sub txtInvNo_OnBlur()
	If Trim(frm1.txtInvNo.value) = "" Then
		Call ggoOper.SetReqAttr(frm1.txtInvDt, "D")
	Else
		Call ggoOper.SetReqAttr(frm1.txtInvDt, "N") 'N:Required, Q:Protected, D:Default
	End If
End Sub

'======================================================================================================
'�������ǹ�ȣ �Է½� ������������ �Է��ʼ� 
'======================================================================================================
Sub txtBlNo_OnBlur()
	If Trim(frm1.txtBlNo.value) = "" Then
		Call ggoOper.SetReqAttr(frm1.txtBlDt, "D")
	Else
		Call ggoOper.SetReqAttr(frm1.txtBlDt, "N") 'N:Required, Q:Protected, D:Default
	End If
End Sub

Sub txTGlDt_Change()
	lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'            4. Common Function�� 
' ���: Common Function
' ����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================

'======================================================================================================
'            5. Interface�� 
' ���: Interface
' ����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'=======================================================================================================

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    lgstartfnc = True
    
    Err.Clear                                                               
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then										'This function check indispensable field
       Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then  
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")     
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    Call InitVariables()													'Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------                  
    Call DbQuery()															'��: Query db data    
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
     
    FncNew = False  
    lgstartfnc = True                                                         
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")												'Clear Condition Field
    Call ggoOper.LockField(Document, "N")												'Lock  Suitable  Field    
    Call InitVariables()																'Initializes local global variables
    call SetDefaultVal()
    
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
    If lgIntFlgMode <> parent.OPMD_UMODE Then											'Check if there is retrived data
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
    Call DbDelete               '��: Delete db data
    
    FncDelete = True                                                        
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
 
    FncSave = False                                                         
    
    Err.Clear                                                               
    

    If lgBlnFlgChgValue = False Then										'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")				'��: Display Message(There is no changed data.)
		Exit Function
    End If
    
    If Len(frm1.txtNetAmt.Text) = 0 then
		Call DisplayMsgBox("970021","X",frm1.txtNetAmt.alt,"X")  
		Exit Function
    ElseIf UNICDbl(frm1.txtNetAmt.Text) = 0 then
		Call DisplayMsgBox("141704","X",frm1.txtNetAmt.alt,"X")  
		Exit Function
    End if
    
    If Not chkField(Document, "2") Then										'��: Check required field(Single area)
		Exit Function
    End If
    '================================================================================================
    '���ڰ��� üũ : LC������(txtLcDt)<=������(txtInvDt)<=����������(txtBlDt)<=ä��/ä����(txtArDt)
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
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																'��: Save db data
    
    FncSave = True                                                       
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
 
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
    
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
Function FncSplitColumn()

End Function

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
 
	FncExit = False

	If lgBlnFlgChgValue = True Then														'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")					'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function  DbDelete() 
    DbDelete = False              
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtArNo=" & Trim(frm1.txtArNo.value)    '��: ���� ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)          '��: �����Ͻ� ASP �� ���� 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()																'���� ������ ���� ���� 
	Call ggoOper.ClearField(Document, "2")											'Clear Condition Field
    Call ggoOper.LockField(Document, "N")											'Lock  Suitable  Field    
    Call InitVariables()															'Initializes local global variables
    Call SetDefaultVal()
    
    frm1.txtArNo.Value = ""
    frm1.txtArNo.focus
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
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'��: 
			strVal = strVal & "&txtArNo=" & Trim(.htxtArNo.value)					'��ȸ ���� ����Ÿ 
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'��: 
			strVal = strVal & "&txtArNo=" & Trim(.txtArNo.value)					'��ȸ ���� ����Ÿ 
		End If
    End With

	Call RunMyBizASP(MyBizASP, strVal)												'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function  DbQueryOk()
	With frm1
		'-----------------------
		'Reset variables area
		'-----------------------  
		Call ggoOper.LockField(Document, "Q")										'This function lock the suitable field        
		Call SetToolbar("1111100000001111") 
		call InitVariables()
		
		lgQueryOk= True
				
		lgIntFlgMode = parent.OPMD_UMODE											'Indicates that current mode is Update mode
	 
		Call txtDocCur_OnChange()        
		Call txtDeptCd_OnBlur()
		If Trim(frm1.txtInvNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtInvDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtInvDt, "N")								'N:Required, Q:Protected, D:Default
		End If
		If Trim(frm1.txtBlNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtBlDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtBlDt, "N")								'N:Required, Q:Protected, D:Default
		End If
	 
		lgBlnFlgChgValue = False
		lgQueryOk= False
	End With 
End Function


'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    DbSave = False                                                          
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   
    Err.Clear 

	frm1.txtFlgMode.value = lgIntFlgMode         
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data ���� ��Ģ 
    ' 0: Sheet��, 1: Flag , 2: Row��ġ, 3~N: �� ����Ÿ 

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'���� �����Ͻ� ASP �� ���� 
        
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function  DbSaveOk(ByVal ArNo)														'��: ���� ������ ���� ���� 
    If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtArNo.value = ArNo
	End If   
 
	Call ggoOper.ClearField(Document, "2")											'Clear Contents  Field
	Call InitVariables()															'Initializes local global variables
	Call DBquery()     
End Function

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then                     
		Call CurFormatNumericOCX()
	End If    

	If lgQueryOk <> True Then
		If UCase(Trim(frm1.txtDocCur.Value)) <> UCase(Trim(parent.gCurrency)) Then
			frm1.txtXchRate.Text = "0" 
		Else			
			frm1.txtXchRate.Text = "1" 		
		End If			
		frm1.txtNetLocAmt.text = "0"
	End If   	
End Sub

'===================================== CurFormatNumericOCX()  =======================================
' Name : CurFormatNumericOCX()
' Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' �ܻ����� 
		ggoOper.FormatFieldByObjectOfCur .txtNetAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' ä���ܾ� 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' ȯ�� 
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtDocCur.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>
<!--'======================================================================================================
'            6. Tag�� 
' ���: Tag�κ� ���� 
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
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>����ä�ǵ��</font></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
								</TR>
							</TABLE>
						</TD>
						<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
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
									<TD CLASS="TD5" NOWRAP>ä�ǹ�ȣ</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtArNo" ALT="ä�ǹ�ȣ" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript:CALL OpenPopUp(frm1.txtArNo.Value, 0)"></TD>        
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
									<TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDealBpCd" ALT="�ֹ�ó" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtDealBpCd.Value,3)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									<INPUT NAME="txtDealBpNm" ALT="�ֹ�ó" SIZE="20" tag = "24" ></TD>
									<TD CLASS=TD5 NOWRAP>�����ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvNo" ALT="�����ȣ" MAXLENGTH="50" SIZE=20 STYLE="TEXT-ALIGN:  left" tag="2XXXXU"></TD>       </TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayBpCd" ALT="����ó" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtPayBpCd.Value,4)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									<INPUT  NAME="txtpaybpnm"  ALT="����ó" SIZE="20" tag = "24" ></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/a3112ma1_OBJECT3_txtInvDt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���ݰ�꼭����ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReportBpCd" ALT="���ݰ�꼭����ó" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtReportBpCd.Value,9)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									<INPUT  NAME="txtReportbpnm"  ALT="���ݰ�꼭����ó" SIZE="20" tag = "24" ></TD>        
									<TD CLASS=TD5 NOWRAP>�������ǹ�ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txTblNo" ALT="�������ǹ�ȣ" MAXLENGTH="35" SIZE=20 tag="2XXXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�μ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="�μ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									<INPUT NAME="txtDeptNm" ALT="�μ�" SIZE="20" tag ="24" ></TD>
									<TD CLASS=TD5 NOWRAP>����������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/a3112ma1_OBJECT4_txTblDt.js'></script></TD>               
								</TR>
								<TR><TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="�����ڵ�" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU" ><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtAcctCd.value,1)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
										<INPUT NAME="txtAcctnm" ALT="�����ڵ��" MAXLENGTH="20"  tag  ="24"></TD>          
									<TD CLASS="TD5" nowrap>�����Ⱓ</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP>
													<script language =javascript src='./js/a3112ma1_fpDoubleSingle5_txtPayDur.js'></script>
												</TD>
												<TD NOWRAP>
													&nbsp;��
												</TD>
											</TR>
										</Table>
									</TD>
								</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ä������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_OBJECT1_txtArDt.js'></script></TD>        
								<TD CLASS="TD5" nowrap>�������</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayMethCd" ALT="�������" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtPayMethCd.value, 10)">
									<INPUT TYPE=TEXT NAME="txtPayMethNm" ALT="�������" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>

					       </TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ä�Ǹ�������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_OBJECT2_txtDueDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>�Ա�����</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayTypeCd" ALT="�Ա�����" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtPayTypeCd.value, 11)">
									<INPUT TYPE=TEXT NAME="txtPayTypeNm" ALT="�Ա�����" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
					       </TR>
					       <TR>
<% If gIsShowLocal <> "N" Then %>
								<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="22XXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtDocCur.Value,8)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
								&nbsp;&nbsp;ȯ��<script language =javascript src='./js/a3112ma1_OBJECT5_txtXchRate.js'></script></TD>                
<% ELSE %>
									<INPUT TYPE=HIDDEN NAME="txtDocCur"   TABINDEX="-1">
									<INPUT TYPE=HIDDEN NAME="txtXchRate"  TABINDEX="-1">									
<% End If %>         
						        <TD CLASS=TD5 NOWRAP>��ݰ�������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaymTerms" ALT="��ݰ�������" MAXLENGTH="120" SIZE=30 STYLE="TEXT-ALIGN: left" tag ="21"></TD>        
							</TR>               
							<TR>
								<TD CLASS=TD5 NOWRAP>��ǥ����</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a3112ma1_OBJECT1_txtGlDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>��ǥ��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGlNo" ALT="��ǥ��ȣ" SIZE="19" MAXLENGTH="18" STYLE="TEXT-ALIGN: Left" tag="24XXXU" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ܻ�����</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_I456811636_txtNetAmt.js'></script></TD>
<% If gIsShowLocal <> "N" Then %>         
							    <TD CLASS=TD5 NOWRAP>�ܻ�����(�ڱ���ȭ)</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_I471998035_txtNetLocAmt.js'></script></TD>
							</TR>
							<TR>
<% ELSE %>
								<INPUT TYPE=HIDDEN NAME="txtNetLocAmt"   TABINDEX="-1">
<% End If %>       
								<TD CLASS=TD5 NOWRAP>ä���ܾ�</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_I623698359_txtBalAmt.js'></script></TD>
<% If gIsShowLocal <> "N" Then %>        
								<TD CLASS=TD5 NOWRAP>ä���ܾ�(�ڱ���ȭ)</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_I870568823_txtBalLocAmt.js'></script></TD>
<% ELSE %>
										<INPUT TYPE=HIDDEN NAME="txtBalLocAmt"   TABINDEX="-1">
<% End If %>       
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDesc" ALT="���" MAXLENGTH="128" SIZE="60" tag="2X" ></TD>
							    <TD CLASS=TD5 NOWRAP>������Ʈ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="������Ʈ" MAXLENGTH=25 SIZE=25 tag="2X"></TD>
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
		<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TabIndex="-1"></IFRAME>
	</TD>
 </TR>
</TABLE>
	<INPUT TYPE=hidden NAME="txtMode" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="htxtArNo" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="hAcctCd" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
</BODY>
</HTML>

