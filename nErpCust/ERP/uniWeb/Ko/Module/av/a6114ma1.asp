<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Change
'*  3. Program ID           : A6114MA1
'*  4. Program Name         : ��꼭���� 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/11/21
'*  8. Modified date(Last)  : 2003/10/22
'*  9. Modifier (First)     : ��ȣ�� 
'* 10. Modifier (Last)      : ����� 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
#########################################################################################################
												1. �� �� �� 
##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit    												'��: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================
'@PGM_ID														'�����Ͻ� ���� ASP�� 
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"			'ȯ������ �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID	=	"a6114mb1.asp" 	
Const JUMP_PGM_ID_GL_ENTRY = "a5104ma1"							'ȸ����ǥ��� 
Const JUMP_PGM_ID_TEMP_GL_ENTRY = "a5101ma1"					'������ǥ��� 

'@Grid_Column
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const	ToolBar	=	"1100100000001111"

Dim IsOpenPop						                        'Popup
Dim gSelframeFlg                                            'Current Tab Page

<%
Dim dtToday
dtToday = GetSvrDate
%>

'======================================================================================================
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'=======================================================================================================

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE									'Indicates that current mode is Create mode
    lgIntGrpCount = 0													'initializes Group View Size
    lgStrPrevKey = ""													'initializes Previous Key	
	lgBlnFlgChgValue = False											'Indicates that no value changed
	Err.Clear
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	Dim svrDate
	Dim strYear, strMonth, strDay
    svrDate					 = "<%=GetSvrDate%>"
	Call ExtractDateFrom(svrDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)	
	frm1.txtIssuedDt.text	=  UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	frm1.txtDocCur.value	= parent.gCurrency
	frm1.txtXchRate.text	= 1

	lgBlnFlgChgValue = False
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'======================================================================================================
'   Function Name : OpenVatNoInfo()
'   Function Desc : 
'=======================================================================================================
Function OpenVatNoInfo(Byval strCode, Byval Cond)
	Dim iCalledAspName
	Dim arrRet
		
	If IsOpenPop = True Then Exit Function	

	iCalledAspName = AskPRAspName("a6114ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a6114ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	     
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtVatNo.focus
		Exit Function
	Else
		Call SetVatNoInfo(arrRet,Cond)	
	End If	
End Function

'======================================================================================================
'   Function Name : SetChgNoInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetVatNoInfo(Byval arrRet, Byval Cond)
	Select Case Cond
		Case "VatNo"
			frm1.txtVatNo.focus
			frm1.txtVatNo.Value	= arrRet(0)
	End Select	
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

	arrParam(0) = strCode								' Code Condition
   	arrParam(1) = ""									' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""									' FrDt
	arrParam(3) = ""									' ToDt
	arrParam(4) = "T"									' B :���� S: ���� T: ��ü 
	arrParam(5) = ""									' SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetBpCd(arrRet)
		lgBlnFlgChgValue = True
	End If	
End Function

'=======================================================================================================
'	Name : OpenBpCd()
'	Description : Bp Cd PopUp
'=======================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ŷ�ó �˾�"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "�ŷ�ó�ڵ�"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "�ŷ�ó�ڵ�"		
    arrHeader(1) = "�ŷ�ó��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetBpCd(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetBpCd()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetBpCd(byval arrRet)
	frm1.txtBpCd.focus
	frm1.txtBpCd.Value    = arrRet(0)		
	frm1.txtBpNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'	Name : OpenReportBizArea()
'	Description : Bp Cd PopUp
'=======================================================================================================
Function OpenReportBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ݽŰ����� �˾�"	                ' �˾� ��Ī 
	arrParam(1) = "B_TAX_BIZ_AREA"			        	' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtReportBizArea.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "���ݽŰ������ڵ�"			        '�����ʵ��� �� ��Ī 
	
    arrField(0) = "TAX_BIZ_AREA_CD"	                           ' Field��(0)
    arrField(1) = "TAX_BIZ_AREA_NM"	                           ' Field��(1)
    
    arrHeader(0) = "���ݽŰ������ڵ�"		               ' Header��(0)
    arrHeader(1) = "���ݽŰ������"		               ' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtReportBizArea.focus	
		Exit Function
	Else
		Call SetReportBizArea(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetReportBizArea()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetReportBizArea(byval arrRet)
	frm1.txtReportBizArea.focus	
	frm1.txtReportBizArea.Value    = arrRet(0)		
	frm1.txtReportBizAreaNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'	Name : OpenBizArea()
'	Description : Bp Cd PopUp
'=======================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "�Ű����� �˾�"	                ' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"			        	' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBizArea.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "�Ű������ڵ�"			        '�����ʵ��� �� ��Ī 
	
	arrField(0) = "BIZ_AREA_CD"	                           ' Field��(0)
    arrField(1) = "BIZ_AREA_NM"	                           ' Field��(1)
    
    arrHeader(0) = "�Ű������ڵ�"		               ' Header��(0)
    arrHeader(1) = "�Ű������"		               ' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizArea.focus
		Exit Function
	Else
		Call SetBizArea(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetBizArea()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetBizArea(byval arrRet)
	frm1.txtBizArea.focus
	frm1.txtBizArea.Value    = arrRet(0)		
	frm1.txtBizAreaNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'	Name : OpenAcctCd()
'	Description : Bp Cd PopUp
'=======================================================================================================
Function OpenAcctCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����ڵ��˾�"	                ' �˾� ��Ī 
	arrParam(1) = "A_ACCT"			        	' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "�����ڵ�"			        '�����ʵ��� �� ��Ī 
	
    arrField(0) = "ACCT_CD"	                           ' Field��(0)
    arrField(1) = "ACCT_NM"	                           ' Field��(1)
    
    arrHeader(0) = "�����ڵ�"		               ' Header��(0)
    arrHeader(1) = "�����ڵ��"		               ' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetAcctCd(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetAcctCd()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetAcctCd(byval arrRet)
	frm1.txtAcctCd.focus
	frm1.txtAcctCd.Value    = arrRet(0)		
	frm1.txtAcctNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'	Name : OpenVatType()
'	Description : Bp Cd PopUp
'=======================================================================================================
Function OpenVatType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
      
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ΰ��������˾�"	                ' �˾� ��Ī 
	arrParam(1) = "B_MINOR"			                	' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtVatType.Value)
	arrParam(3) = ""
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9001", "''", "S") & " "			
	arrParam(5) = "�ΰ����ڵ�"			        '�����ʵ��� �� ��Ī 
	
    arrField(0) = "MINOR_CD"	                           ' Field��(0)
    arrField(1) = "MINOR_NM"	                           ' Field��(1)
    
    arrHeader(0) = "�ΰ�������"		               ' Header��(0)
    arrHeader(1) = "�ΰ���������"		               ' Header��(1)
    
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

'=======================================================================================================
'	Name : SetAcctCd()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetVatType(byval arrRet)
	frm1.txtVatType.focus
	frm1.txtVatType.Value   = arrRet(0)		
	frm1.txtVatTypeNm.Value = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'	Name : OpenCurrency()
'	Description : Currency PopUp
'=======================================================================================================
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ŷ���ȭ �˾�"	
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "�ŷ���ȭ"
	
    arrField(0) = "CURRENCY"	
    arrField(1) = "CURRENCY_DESC"	
    
    arrHeader(0) = "�ŷ���ȭ"		
    arrHeader(1) = "�ŷ���ȭ��"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDocCur.focus
		Exit Function
	Else
		Call SetCurrency(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetCurrency()
'	Description : Currency Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetCurrency(byval arrRet)
	frm1.txtDocCur.focus
	frm1.txtDocCur.value    = arrRet(0)	
	Call CurFormatNumericOCX()	
	lgBlnFlgChgValue = True
End Function

'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================
'=======================================================================================================
'Description : ������ǥ �������� �˾� 
'=======================================================================================================
Function OpenPopupTempGL()
	Dim iCalledAspName
	Dim arrRet, RetFlag
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'������ǥ��ȣ 
	arrParam(1) = Trim(frm1.txtRefNo.value)		'Reference��ȣ 

	If Trim(frm1.txtTempGlNo.value) = "" Then
		RetFlag = DisplayMsgBox("970000","X" , frm1.txtTempGlNo.Alt, "X") 	
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'=======================================================================================================
'Description : ȸ����ǥ �������� �˾� 
'=======================================================================================================
Function OpenPopupGL()
	Dim iCalledAspName
	Dim arrRet,RetFlag
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'ȸ����ǥ��ȣ 
	arrParam(1) = Trim(frm1.txtRefNo.value)		'Reference��ȣ 

	If Trim(frm1.txtGlNo.value) = "" Then
		RetFlag = DisplayMsgBox("970000","X" , frm1.txtGlNo.Alt, "X") 	
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim arrData
	
	'arrData = InitCombo("A1003", "frm1.cboIoFg")	'����/���ⱸ�� 
	'arrData = InitCombo("A1007", "frm1.cboConfFg")		'���ο���		
	 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1003", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboIoFg ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox2()
	Dim arrData
	
	'arrData = InitCombo("A1003", "frm1.cboIoFg")	'����/���ⱸ�� 
	'arrData = InitCombo("A1007", "frm1.cboConfFg")		'���ο���		
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("A1007", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
End Sub

'========================================================================================================
'	Desc : Cookie Setting
'========================================================================================================
Sub CookiePage(ByVal Kubun)
'	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp	
	
	Select Case Kubun		
		Case "FORM_LOAD"
			strTemp = ReadCookie("VAT_NO")
			Call WriteCookie("VAT_NO", "")
			
			If strTemp = "" then Exit Sub
						
			frm1.txtVatNo.value = strTemp
					
			Call ggoOper.SetReqAttr(frm1.txtVatNo,   "Q")	
					
			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("VAT_NO", "")
				Exit Sub 
			End If
			
			Call FncQuery()
		Case JUMP_PGM_ID_GL_ENTRY
			If frm1.txtGLNo.value = "" Then
				Call WriteCookie("GL_NO", "")
				Exit Sub 
			End If
			Call WriteCookie("GL_NO", frm1.txtGLNo.value)
			
			strtemp = ReadCookie("GL_NO")
		Case JUMP_PGM_ID_TEMP_GL_ENTRY	
			If Not (frm1.txtGLNo.value = "" AND frm1.txtTempGLNo.value <> "")  Then
				Call WriteCookie("TEMP_GL_NO", "")
				Exit Sub 
			End If
			
			Call WriteCookie("TEMP_GL_NO", frm1.txtTempGLNo.value)	
		Case Else
			Exit Sub
	End Select
End Sub

'========================================================================================================
'	Desc : ȭ���̵� 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	Dim strTemp

	'-----------------------
	'Check previous data area
	'------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Select Case strPgmId		
		Case JUMP_PGM_ID_GL_ENTRY
			If frm1.txtGLNo.value = "" Then	
				IntRetCD = DisplayMsgBox("113100", "X","X","X")			
				Exit Function			
			End If	
		Case JUMP_PGM_ID_TEMP_GL_ENTRY	
			If  frm1.txtTempGLNo.value = ""  Then
				IntRetCD = DisplayMsgBox("114100", "X","X","X")			
				Exit Function			
			End If				
	End Select
	
    Call CookiePage(strPgmId)
    
    Call PgmJump(strPgmId)
End Function

'======================================================================================================
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()
    Call GetGlobalVar
    Call LoadInfTB19029                                                     'Load table , B_numeric_format
        
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    
    frm1.txtNetAmt.AllowNull =false
    frm1.txtNetLocAmt.AllowNull =false
    frm1.txtVatAmt.AllowNull =false
    frm1.txtVatLocAmt.AllowNull =false        
    Call InitVariables                                                      'Initializes local global variables    
    Call SetDefaultVal
	Call InitComboBox
	Call InitComboBox2  
	Call FncNew()
	Call CookiePage("FORM_LOAD")
    Call SetToolbar(ToolBar)										'��ư ���� ���� 
    frm1.txtVatNo.focus 
End Sub

'======================================================================================================
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================

'======================================================================================================
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'=======================================================================================================

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear																			'Protect system from crashing
	'-----------------------
    'Check previous data area
    '-----------------------     
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
      	    Exit Function
    	End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")												'Clear Contents  Field
    Call InitVariables																	'Initializes local global variables    
    'Call InitComboBox

	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then													'This function check indispensable field
		Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '----------------------- 
    If  DbQuery	= False Then
		Exit Function
	End If																				'Query db data
       
    FncQuery = True															
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	Call CurFormatNumericOCX()
End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete() 
    
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim lDelRows, intRows
	FncSave = False
			
	Err.Clear                                                               
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then				'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")	'��ȸ�� ���� �Ͻʽÿ�.
        Exit Function
    End If
    
    If lgBlnFlgChgValue = False  Then  '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")            '��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then               '��: Check required field(Single area)
       Exit Function
    End If
	
	'-----------------------
	'Save function call area
	'-----------------------
	IF DbSave = False Then
		Exit Function
	END IF
							                                                '��: Save db data	 
	FncSave = True
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()

End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel()

End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow()

End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow()

End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    

End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next

End Function

'=======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call Parent.FncExport(parent.C_SINGLEMULTI)										
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , True)                               
End Function

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	
	If lgBlnFlgChgValue = True  Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete() 
   
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()																'���� ������ ���� ���� 
	
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
        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'��: 
        strVal = strVal     & "&txtVatNo=" & Trim(.txtVatNo.value)					'��ȸ ���� ����Ÿ 
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)												'�����Ͻ� ASP �� ���� 
	
	DbQuery = True                                                          
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================================
Function DbQueryOk()																'��ȸ ������ ������� 
    lgIntFlgMode = parent.OPMD_UMODE    
	Call SetToolbar(ToolBar)														'��ư ���� ����	
	Call CurFormatNumericOCX()
	Call ggoOper.LockField(Document, "Q")

	lgBlnFlgChgValue = False
    Set gActiveElement = document.ActiveElement   
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	Dim IntRows 
	Dim IntCols 
	
	DbSave = False                                                          
	
	On Error Resume Next                                                   
	
	Call LayerShowHide(1)
	
	'Call SetSumItem	
	
	With frm1
		.txtMode.value = parent.UID_M0002											'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode											'��: �ű��Է�/���� ���� 
	End With
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	' Data ���� ��Ģ 
	' 0: Flag , 1: Row��ġ, 2~N: �� ����Ÿ	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)												'��: ���� �����Ͻ� ASP �� ����	
	
	DbSave = True                                                           
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
   	lgBlnFlgChgValue = false	

    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    Call InitVariables                                                      'Initializes local global variables    
    'Call InitComboBox
	Call DbQuery	
End Function

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
Function cboIoFg_onChange		
	lgBlnFlgChgValue = True		
End Function

Function cboConfFg_onChange
	lgBlnFlgChgValue = True	
End Function

Function txtReportBizArea_onblur()
	If frm1.txtReportBizArea.value = "" Then
		frm1.txtReportBizAreaNm.value = ""
	End If
End Function

Function txtBizArea_onblur()
	If frm1.txtBizArea.value = "" Then
		frm1.txtBizAreaNm.value = ""		
	End If	
End Function

Function txtAcctCd_onblur()
	If frm1.txtAcctCd.value = "" Then
		frm1.txtAcctNm.value = ""
	End If
End Function

Function txtBpCd_onblur()
	If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = ""
	End If
End Function

Function cboMadeVatFg_onChange		
	lgBlnFlgChgValue = True		
End Function

'=======================================================================================================
'   Event Name : txtIssuedDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtIssuedDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtIssuedDt_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtIssuedDt_Change()
    lgBlnFlgChgValue = True
End Sub

Function txtXchRate_Change
	lgBlnFlgChgValue = True
End Function

Function txtVatAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtVatLocAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtNetAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtNetLocAmt_Change
	lgBlnFlgChgValue = True
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'===================================================================================================='@@
Sub CurFormatNumericOCX()
	With frm1
		'�����ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtNetAmt , .txtDocCur.Value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'�뺯�ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt , .txtDocCur.Value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'�ΰ����� 
		ggoOper.FormatFieldByObjectOfCur .txtVatRate, .txtDocCur.Value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'ȯ�� 
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtDocCur.Value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
	End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'======================================================================================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��꼭����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>					
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
									<TD CLASS="TD5" NOWRAP>��꼭��ȣ</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtVatNo" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="��꼭��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript: Call OpenVatNoInfo(frm1.txtVatNo.value,'VatNo')"></TD>
								</TR>
							</TABLE>        
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% >
						<TABLE <%=LR_SPACE_TYPE_60%>>	
									<TR>
										<TD CLASS="TD5" NOWRAP>��꼭��ȣ</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatNo1" SIZE=20 MAXLENGTH=18 tag="14XXXU" ALT="��꼭��ȣ"> <!-- <IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenVatNoInfo(frm1.txtVatNo1.value,'VatNo1')"> --> </TD> 
										<TD CLASS="TD5" NOWRAP>������</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a6114ma1_fpDateTime1_txtIssuedDt.js'></script></TD>
									</TR>					
									<TR>
									    <TD CLASS="TD5" NOWRAP>�Ű������ڵ�</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtReportBizArea" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�Ű������ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReportBizArea" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenReportBizArea()">&nbsp;<INPUT TYPE=TEXT NAME="txtReportBizAreaNm" SIZE=20 tag="24" ALT="�Ű������ڵ�"></TD>
								        <TD CLASS="TD5" NOWRAP>�߻�������ڵ�</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizArea" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�߻�������ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 tag="24" ALT="�߻�������ڵ�"></TD>
									</TR>
									<TR>
									    <TD CLASS="TD5" NOWRAP>�ŷ�ó�ڵ�</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�ŷ�ó�ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24" ALT="�ŷ�ó�ڵ�"></TD>																		
								        <TD CLASS="TD5" NOWRAP>����ڵ�Ϲ�ȣ</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOwnRgst" SIZE=35 MAXLENGTH=128 tag="14X" ALT="����ڵ�Ϲ�ȣ"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenAcctCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=20 tag="24" ALT="�����ڵ�"></TD>																		
								        <TD CLASS="TD5" NOWRAP>������ȣ</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" SIZE=18 MAXLENGTH=30 tag="24X" ALT="������ȣ">
								    </TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>����/���ⱸ��</TD>
							            <TD CLASS="TD6" NOWRAP><SELECT NAME="cboIoFg" ALT="����/���ⱸ��" tag="21" STYLE="WIDTH: 100px"  ><OPTION VALUE=""></OPTION></SELECT></TD>							            
										<TD CLASS="TD5" NOWRAP></TD>
										<TD CLASS="TD6" NOWRAP></TD>
								   	</TR>
								   	
									<TR>
										<TD CLASS="TD5" NOWRAP>�ΰ�������</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�ΰ�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType()">&nbsp;<INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="24" ALT="�ΰ�������"></TD>
										<TD CLASS="TD5" NOWRAP>�ΰ�����</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/a6114ma1_fpDoubleSingle9_txtVatRate.js'></script>&nbsp;%
									    </TD>
<!--										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatRate" SIZE=10 MAXLENGTH=10 tag="21" ALT="�ΰ�����">&nbsp;%-->
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
										<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" TYPE="Text" MAXLENGTH=3 SIZE=20 tag="14XXXU" ></TD><!--<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCurrency()"></TD>-->
										<TD CLASS=TD5 NOWRAP>ȯ��</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/a6114ma1_fpDoubleSingle9_txtXchRate.js'></script>
									    </TD>
									</TR>																		
									<TR>
										<TD CLASS="TD5" NOWRAP>���ް���</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a6114ma1_fpDoubleSingle8_txtNetAmt.js'></script></TD>
										<TD CLASS="TD5" NOWRAP>���ް���(�ڱ�)</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a6114ma1_fpDoubleSingle9_txtNetLocAmt.js'></script></TD>
							        </TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>����</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a6114ma1_fpDoubleSingle8_txtVatAmt.js'></script></TD>
										<TD CLASS="TD5" NOWRAP>����(�ڱ�)</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a6114ma1_fpDoubleSingle9_txtVatLocAmt.js'></script></TD>
							        </TR>
									<TR>							
										<TD CLASS="TD5" NOWRAP>�Ű���</TD>
										<TD CLASS="TD6" NOWRAP><SELECT NAME="cboMadeVatFg" ALT="�Ű���" tag="22" STYLE="WIDTH: 100px"  ><OPTION VALUE="Y">Yes</OPTION><OPTION  Selected VALUE="N">No</OPTION></SELECT></TD>
										<TD CLASS="TD5" NOWRAP>���λ���</TD>
										<TD CLASS="TD6" NOWRAP><SELECT NAME="cboConfFg" ALT="���λ���" STYLE="WIDTH: 100px" tag="24" ><OPTION VALUE=""></OPTION></SELECT></TD>
									</TR>
									<TR>							
										<TD CLASS="TD5" NOWRAP>������ǥ��ȣ</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTempGLNo" SIZE=20 MAXLENGTH=18  tag="24" ALT="������ǥ��ȣ"></TD>
										<TD CLASS="TD5" NOWRAP>ȸ����ǥ��ȣ</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGLNo" MAXLENGTH=18 SIZE=20 tag="24" ALT="��ǥ��ȣ" ></TD>
									</TR>
									<TR>							
										<TD CLASS="TD5" NOWRAP>ä����ȣ</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAPNo" SIZE=20 MAXLENGTH=18  tag="24" ALT="ä����ȣ"></TD>
										<TD CLASS="TD5" NOWRAP>ä�ǹ�ȣ</TD>
										<TD CLASS="TD6" NOWRAP><INPUT NAME="txtARNo" ALT="ä�ǹ�ȣ" TYPE="Text" MAXLENGTH=18 SIZE=25 tag="24" ></TD>
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
						<A ONCLICK="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_GL_ENTRY)">ȸ����ǥ���</a>&nbsp;|
						<A ONCLICK="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_TEMP_GL_ENTRY)">������ǥ���</a>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=10>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"         tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"   tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"  tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"	  tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"	  tag="24">
<INPUT TYPE=HIDDEN NAME="htxtVatNo"	  tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
