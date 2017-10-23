<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : ����BOM���� 
'*  2. Function Name        : 
'*  3. Program ID           :  p1711oa1.asp
'*  4. Program Name         :  ����BOM ���� ��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2005/02/12
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Cho Yong Chill
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  <!-- '��: �ش� ��ġ�� ���� �޶���, ��� ��� -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1711ob1.asp"

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim IsOpenPop
Dim lgExpType

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False     
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtBomNo.value = "E"
	frm1.txtBaseDt.text = StartDate
		
End Sub
'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseDt.Focus
    End If
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 

    If parent.gPlant <> "" and CheckPlant(parent.gPlant) = True Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  	 
	End If    
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel, UnloadMode)
    
End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

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
Function FncQuery()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function
Function FncSave()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function
Function FncNew()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function
Function FncDelete()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function
Function FncInsertRow()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function
Function FncDeleteRow()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function
Function FncCopy()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function
Function FncCancel()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function



'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================

Function BtnPrint() 
    Dim strVal
    
    '----------------------------------------------
    '- Call Query ASP
    '----------------------------------------------
	Call BtnDisabled(1)	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
	If frm1.txtPlantCd.value = "" or CheckPlant(frm1.txtPlantCd.value) = False Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		Call BtnDisabled(0)
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
		  
	If Not chkField(Document, "X") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
	   Exit Function
	End If
		  
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
    
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtBaseDt=" & UNIConvDate(frm1.txtBaseDt.text)
	
    '------------ �������� Setting ----------------
      '----------------------------------------------
	   
	If frm1.rdoExplodType1.checked = True And frm1.rdoPrintType1.checked  = True Then	  '�������̰�  �ܴܰ� 
		strVal = strval & "&rdoPrintType=" & "1"
		lgExpType = "1"										  
	ElseIf frm1.rdoExplodType1.checked = True And frm1.rdoPrintType2.checked = True Then  '�������̰�  �ܴܰ� 
		strVal = strval & "&rdoPrintType=" & "3"										  
		lgExpType = "3"
	ElseIf frm1.rdoExplodType2.checked = True And frm1.rdoPrintType1.checked = True Then  '�������̰�  �ٴܰ� 
		strVal = strval & "&rdoPrintType=" & "2"										  
		lgExpType = "2"
	Else																				  '�������̰�  �ٴܰ� 
		strVal = strval & "&rdoPrintType=" & "4"										  
		lgExpType = "4"
	End If

    '----------------------------------------------
      '-------------- End Of �������� ---------------
   
    strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBOMNo.value)
    strVal = strVal & "&BtnType=" & "1"
    
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
	
End Function


'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview() 
    Dim strVal
    
    '----------------------------------------------
    '- Call Query ASP
    '----------------------------------------------
    Call BtnDisabled(1)	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
	If frm1.txtPlantCd.value = "" or CheckPlant(frm1.txtPlantCd.value) = False Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		Call BtnDisabled(0)
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

			  
	If Not chkField(Document, "X") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
	   Exit Function
	End If

    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
    
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtBaseDt=" & UNIConvDate(frm1.txtBaseDt.text)
	
    '------------ �������� Setting ----------------
      '----------------------------------------------
	   
	If frm1.rdoExplodType1.checked = True And frm1.rdoPrintType1.checked  = True Then	  '�������̰� �ܴܰ� 
		strVal = strval & "&rdoPrintType=" & "1"
		lgExpType = "1"										  
	ElseIf frm1.rdoExplodType1.checked = True And frm1.rdoPrintType2.checked = True Then  '�������̰� �ܴܰ� 
		strVal = strval & "&rdoPrintType=" & "3"										  
		lgExpType = "3"
	ElseIf frm1.rdoExplodType2.checked = True And frm1.rdoPrintType1.checked = True Then  '�������̰� �ٴܰ� 
		strVal = strval & "&rdoPrintType=" & "2"										  
		lgExpType = "2"
	Else																				  '�������̰� �ٴܰ� 
		strVal = strval & "&rdoPrintType=" & "4"										  
		lgExpType = "4"
	End If

    '----------------------------------------------
      '-------------- End Of �������� ---------------
   
    strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBOMNo.value)
	strVal = strVal & "&BtnType=" & "0"
	
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
	    		
End Function

'========================================================================================
' Function Name : PrevExecOk()
' Function Desc : BOM Temp ���̺� ������ ������ �����ϸ� EasyBase�� Open�Ѵ�.
'========================================================================================


Function PrevExecOk()
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim sRepName
	
	Dim strUrl, strEbrFile
	Dim arrParam, arrField, arrHeader
		
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UCase(Trim(frm1.txtItemCd.value))
	
	If frm1.rdoPrintType2.checked = True Then
		If Trim(frm1.txtBOMNo.value) = "" Then
			var3 = "0"
		Else
			var3 = Trim(frm1.txtBOMNo.value)
		End If
		
		sRepName = "P1711OA2"
	Else	
		var3 = Trim(frm1.txtBOMNo.value)
		sRepName = "p1711oa1"
	End If
	
	var4 = frm1.txtSpId.value 
	var5 = lgExpType
	var6 = UNIConvDate(frm1.txtBaseDt.text)
	
	strEbrFile = AskEBDocumentName(sRepName, "EBR")
	
	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|item_cd|" & var2 
	strUrl = strUrl & "|bom_no|" & var3 
	strUrl = strUrl & "|user_id|" & var4 
	strUrl = strUrl & "|exp_type|" & var5
	strUrl = strUrl & "|base_dt|" & var6
	
	Call FncEBRPrevIew(strEbrFile , strUrl)
	
	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement

End Function


'========================================================================================
' Function Name : PrintExecOk()
' Function Desc : BOM Temp ���̺� ������ ������ �����ϸ� EasyBase�� Open�Ѵ�.
'========================================================================================


Function PrintExecOk()

	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim sRepName
	
	Dim strUrl, strEbrFile
		
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UCase(Trim(frm1.txtItemCd.value))
	
	If frm1.rdoPrintType2.checked = True Then
		If Trim(frm1.txtBOMNo.value) = "" Then
			var3 = "0"
		Else
			var3 = Trim(frm1.txtBOMNo.value)
		End If
		
		sRepName = "P1711OA2"
	Else	
		var3 = Trim(frm1.txtBOMNo.value)
		sRepName = "p1711oa1"
	End If
	
	var4 = frm1.txtSpId.value 
	var5 = lgExpType
	var6 = UNIConvDate(frm1.txtBaseDt.text)

	strEbrFile = AskEBDocumentName(sRepName, "EBR")

	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|item_cd|" & var2 
	strUrl = strUrl & "|bom_no|" & var3 
	strUrl = strUrl & "|user_id|" & var4 
	strUrl = strUrl & "|exp_type|" & var5
	strUrl = strUrl & "|base_dt|" & var6
	
'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	call FncEBRprint(EBAction, strEbrFile, strUrl)
'----------------------------------------------------------------
	Call BtnDisabled(0)
	
	frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement

End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

Function OpenPlantCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"												' �˾� ��Ī 
	arrParam(1) = "B_PLANT A, P_PLANT_CONFIGURATION B"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y'"		' Where Condition
	arrParam(5) = "����"													' TextBox ��Ī 
	
    arrField(0) = "A.PLANT_CD"												' Field��(0)
    arrField(1) = "A.PLANT_NM"												' Field��(1)
    
    arrHeader(0) = "����"												' Header��(0)
    arrHeader(1) = "�����"												' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" or CheckPlant(frm1.txtPlantCd.value) = False Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "14"							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"	
    arrField(2) = 3								' Field��(1) : "ITEM_ACCT"
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
	
End Function
'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBOMNo.className) = "PROTECTED" Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012","X", "ǰ��","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True

	arrParam(0) = "BOM�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_MINOR"							' TABLE ��Ī 
	
	arrParam(2) = Trim(frm1.txtBomNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	
	arrParam(5) = "BOM Type"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "BOM Type"					' Header��(0)
    arrHeader(1) = "BOM Ư��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBomNo.focus
	
End Function


Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)  
End Function

Function SetItemCd(ByVal arrRet)
	frm1.txtItemCd.value = arrRet(0) 
	frm1.txtItemNm.value = arrRet(1)
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup���� return�� �� 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)

	frm1.txtBomNo.Value    = arrRet(0)		

End Function

'========================================================================================
' Function Name : CheckPlant
' Function Desc : ����Configuration�� ����������� ������ �Ǿ����� Check
'========================================================================================
Function CheckPlant(ByVal sPlantCd)	
														
    Err.Clear																

    CheckPlant = False
    
	Dim arrVal, strWhere, strFrom

	If Trim(sPlantCd) <> "" Then
	
		strFrom = "B_PLANT A, P_PLANT_CONFIGURATION B"
		strWhere = 				" A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y' AND"
		strWhere = strWhere & 	" A.PLANT_CD = " & FilterVar(sPlantCd, "''", "S")

		If Not CommonQueryRs("A.PLANT_NM", strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
    		Exit Function
		End If
	End If

	CheckPlant = True
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE CLASS="BasicTB" CELLSPACING=0 >
	    		<TR>
	    		    <TD HEIGHT=10 WIDTH=100%>
						<!--<FIELDSET CLASS="CLSFLD">-->
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="X2XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=40 tag="X4" ALT="�����"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=30 MAXLENGTH=18 tag="X2XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 MAXLENGTH=40 tag="X4XXXU" ALT="ǰ���"></TD></TD>
								</TR>							
								<TR>
								    <TD CLASS="TD5" NOWRAP>BOM Type</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBOMNo" SIZE=5 MAXLENGTH=3 tag="X4XXXU" ALT="BOM Type"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBOMNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBomNo()"></TD>
								</TR>					
								<!--<INPUT TYPE="hidden" NAME="txtBOMNo">-->
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p1711oa1_fpDateTime1_txtBaseDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						<!--</FIELDSET>-->
					</TD>
				</TR>
	    		<TR>
					<TD HEIGHT=10 WIDTH=100%>
					    <!--<FIELDSET CLASS="CLSFLD">-->
					        <TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType1" CLASS="RADIO" CHECKED><LABEL FOR="rdoPrintType1">������</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType2" CLASS="RADIO" ><LABEL FOR="rdoPrintType2">������</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ܰ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoExplodType" ID="rdoExplodType1" CLASS="RADIO"><LABEL FOR="rdoExplodType1">�ܴܰ�</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoExplodType" ID="rdoExplodType2" CLASS="RADIO" CHECKED><LABEL FOR="rdoExplodType2">�ٴܰ�</LABEL></TD>
								</TR>
					        </TABLE>
					    <!--</FIELDSET>-->
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtSpId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <INPUT TYPE="hidden" NAME="uname">
    <INPUT TYPE="hidden" NAME="dbname">
    <INPUT TYPE="hidden" NAME="filename">
    <INPUT TYPE="hidden" NAME="condvar">
	<INPUT TYPE="hidden" NAME="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>
