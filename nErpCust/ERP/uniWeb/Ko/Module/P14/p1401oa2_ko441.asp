<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production 
'*  2. Function Name        : 
'*  3. Program ID           :  p1401oa2.asp
'*  4. Program Name         :  ���������� ��� 
'*  5. Program Desc         :
'*  6. Comproxy List        :            
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/04/12
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Hong Chang Ho
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1401ob2.asp"

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim IsOpenPop

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

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtBaseQty.Text = 1
	frm1.txtBaseDt.text = StartDate
	frm1.txtBomNo.value = "1"
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        	frm1.txtPlantCd.value = lgPLCd
	End If
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

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
    Call AppendNumberPlace("6","4","0")      
	Call AppendNumberRange("6","1","1000")                                    '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,FALSE,,ggStrMinPart,ggStrMaxPart)
    
    Call GetValue_ko441()
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
	Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
    
    If parent.gPlant <> "" Then
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

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
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
   Dim strVal, strUrl
   Dim arrParam, arrField, arrHeader
	
   '----------------------------------------------
   '- Call Query ASP
   '----------------------------------------------
   
	Call BtnDisabled(1)	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
		  
	If Not chkField(Document, "X") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
	   Exit Function
	End If
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 

    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBOMNo.value)
    strVal = strVal & "&BtnType=" & "1"
    strVal = strVal & "&txtBaseDt=" & UNIConvDate(frm1.txtBaseDt.Text)		
    strVal = strVal & "&txtBaseQty=" & UNIConvNum(frm1.txtBaseQty.Text, 0)		

    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
    
End Function


'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview() 
    Dim strVal, strUrl
    Dim arrParam, arrField, arrHeader
	
    '----------------------------------------------
    '- Call Query ASP
    '----------------------------------------------
	Call BtnDisabled(1)	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
		  
	If Not chkField(Document, "X") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
	   Exit Function
	End If
	
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 

    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBOMNo.value)
    strVal = strVal & "&BtnType=" & "0"
	strVal = strVal & "&txtBaseDt=" & UNIConvDate(frm1.txtBaseDt.Text)		
    strVal = strVal & "&txtBaseQty=" & UNIConvNum(frm1.txtBaseQty.Text, 0)
    
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
	
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
	Dim var7
	Dim var8
		
	Dim strUrl, strEbrFile
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UCase(Trim(frm1.txtItemCd.value))
	var3 = UCase(Trim(frm1.txtBOMNo.value))
	var4 = StartDate
	
	If frm1.rdoPrintOrder1.checked  = True Then	  
		var5 = "P_BOM_FOR_EXPLOSION.CHILD_ITEM_CD"										 
	Else 
		var5 = "B_ITEM.ITEM_NM"	
	End If
	
	var6 = Trim(frm1.txtSpId.value )
	var7 = UNIConvDate(frm1.txtBaseDt.text)
	var8 = UNIConvNum(frm1.txtBaseQty.Text,0)

	strEbrFile = AskEBDocumentName("P1401OA2", "EBR")

	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|item_cd|" & var2 
	strUrl = strUrl & "|bom_no|" & var3 
	strUrl = strUrl & "|user_id|" & var6 
	strUrl = strUrl & "|current_date|" & var4 
	strUrl = strUrl & "|sort_list|" & var5
	strUrl = strUrl & "|base_dt|" & var7 
	strUrl = strUrl & "|base_qty|" & var8  
	
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
	Dim var7
	Dim var8
			
	Dim strUrl, strEbrFile
		
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UCase(Trim(frm1.txtItemCd.value))
	var3 = UCase(Trim(frm1.txtBOMNo.value))
	var4 = StartDate
	
	If frm1.rdoPrintOrder1.checked  = True Then	  
		var5 = "P_BOM_FOR_EXPLOSION.CHILD_ITEM_CD"										 
	Else 
		var5 = "B_ITEM.ITEM_NM"	
	End If
	
	var6 = Trim(frm1.txtSpId.value )
	var7 = UNIConvDate(frm1.txtBaseDt.text)
	var8 = frm1.txtBaseQty.Text
	
	strEbrFile = AskEBDocumentName("P1401OA2", "EBR")

	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|item_cd|" & var2 
	strUrl = strUrl & "|bom_no|" & var3 
	strUrl = strUrl & "|user_id|" & var6 
	strUrl = strUrl & "|current_date|" & var4 
	strUrl = strUrl & "|sort_list|" & var5 
	strUrl = strUrl & "|base_dt|" & var7 
	strUrl = strUrl & "|base_qty|" & var8  
	
	call FncEBRPrevIew(strEbrFile, strUrl)

	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
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

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"			' �˾� ��Ī 
	arrParam(1) = "B_PLANT"			' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""	' Name Cindition
	arrParam(4) = ""				' Where Condition
	arrParam(5) = "����"			' TextBox ��Ī 
	
    	arrField(0) = "PLANT_CD"			' Field��(0)
    	arrField(1) = "PLANT_NM"			' Field��(1)
    
    	arrHeader(0) = "����"			' Header��(0)
    	arrHeader(1) = "�����"			' Header��(1)
    
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
'	Description : Item PopUps
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
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

	If IsOpenPop = True Then
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�������������</font></td>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="X2XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=40 tag="X4" ALT="�����"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>��ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=30 MAXLENGTH=18 tag="X2XXXU" ALT="��ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="ǰ���">&nbsp;</TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>BOM Type</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBOMNo" SIZE=5 MAXLENGTH=3 tag="X2XXXU" ALT="BOM Type"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBOMNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBomNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p1401oa2_fpDateTime1_txtBaseDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���ؼ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p1401oa2_I431221002_txtBaseQty.js'></script>								
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
	    		<TR>
					<TD HEIGHT=10 WIDTH=100%>
					    <FIELDSET CLASS="CLSFLD">
					        <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS=TD5 NOWRAP>��¼���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPrintOrder" ID="rdoPrintOrder1" CLASS="RADIO" tag="XX" CHECKED><LABEL FOR="rdoPrintOrder1">ǰ���ڵ��</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoPrintOrder" ID="rdoPrintOrder2" CLASS="RADIO" tag="XX" ><LABEL FOR="rdoPrintOrder2">ǰ����</LABEL></TD>
								</TR>
							</TABLE>
					    </FIELDSET>
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
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>
