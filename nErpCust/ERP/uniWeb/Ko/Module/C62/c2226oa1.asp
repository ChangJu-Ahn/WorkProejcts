<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--'**********************************************************************************************
'*  1. Module Name          : Costing 
'*  2. Function Name        : 
'*  3. Program ID           : c2226oa1.asp
'*  4. Program Name         : ǥ�ؿ��� ���� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/06/21
'*  8. Modified date(Last)  : 2002/06/21
'*  9. Modifier (First)     : Cho Ig Sung
'* 10. Modifier (Last)      : Cho Ig Sung
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'=========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit		'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
 Const BIZ_PGM_QRY_ID = "c2226ob1.asp"

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size
Dim IsOpenPop
Dim lgExpType


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False     
End Sub


'========================================================================================================= 
Sub SetDefaultVal()
End Sub


'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "PA") %>
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

'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "X",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
    
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement  	 
    
End Sub


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

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview() 
		
	If Not chkField(Document, "X") Then
		Exit Function
	End If

    If Trim(frm1.txtFrItemCd.value) > Trim(frm1.txtToItemCd.value) Then
		Call DisplayMsgBox("970025", "X", frm1.txtFrItemCd.alt, frm1.txtToItemCd.alt)
		frm1.txtFrItemCd.focus
		Exit Function
	End If

	Call BtnDisabled(1)	

	If frm1.OptFlag1.checked = True Then
		Call BatchExe(0)
	Else
		Call PrevExecOk()
	End If
End Function


'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================

Function BtnPrint() 
		
	If Not chkField(Document, "X") Then
		Exit Function
	End If
 
    If Trim(frm1.txtFrItemCd.value) > Trim(frm1.txtToItemCd.value) Then
		Call DisplayMsgBox("970025", "X", frm1.txtFrItemCd.alt, frm1.txtToItemCd.alt)
		frm1.txtFrItemCd.focus
		Exit Function
	End If
    
	Call BtnDisabled(1)	
			  
	If frm1.OptFlag1.checked = True Then
		Call BatchExe(1)
	Else
		Call PrintExecOk()
	End If
	
End Function


Function BatchExe(ByVal strBtnType)
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001
    
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
    strVal = strVal & "&txtFrItemCd=" & Trim(frm1.txtFrItemCd.value)
	strVal = strVal & "&txtToItemCd=" & Trim(frm1.txtToItemCd.value)
	strVal = strVal & "&BtnType=" & strBtnType
	
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

End Function



'========================================================================================
' Function Name : PrevExecOk()
' Function Desc : BOM Temp ���̺� ������ ������ �����ϸ� EasyBase�� Open�Ѵ�.
'========================================================================================
Function PrevExecOk()
	Dim var1, var2, var3, var4, StrEbrFile, strUrl
		
	var1	= UCase(Trim(frm1.txtPlantCd.value))
	var2	= UCase(Trim(frm1.txtFrItemCd.value))
	var3	= UCase(Trim(frm1.txtToItemCd.value))
	var4	= frm1.txtSpId.value
	
	strUrl = "PlantCd|" & var1
	strUrl = strUrl & "|FrItemCd|" & var2
	strUrl = strUrl & "|ToItemCd|" & var3

	If frm1.OptFlag1.checked = True Then
		strUrl = strUrl & "|SpId|" & var4

		StrEbrFile = "c2220oa1"
	Else
		StrEbrFile = "c2220oa2"
	End If
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
	
	call FncEBRPrevIew(ObjName , strUrl)
	
	Call BtnDisabled(0)

End Function


'========================================================================================
' Function Name : PrintExecOk()
' Function Desc : BOM Temp ���̺� ������ ������ �����ϸ� EasyBase�� Open�Ѵ�.
'========================================================================================


Function PrintExecOk()
	Dim var1, var2, var3, var4, StrEbrFile, strUrl
		
	var1	= UCase(Trim(frm1.txtPlantCd.value))
	var2	= UCase(Trim(frm1.txtFrItemCd.value))
	var3	= UCase(Trim(frm1.txtToItemCd.value))
	var4	= frm1.txtSpId.value
	
	strUrl = "PlantCd|" & var1
	strUrl = strUrl & "|FrItemCd|" & var2
	strUrl = strUrl & "|ToItemCd|" & var3

	If frm1.OptFlag1.checked = True Then
		strUrl = strUrl & "|SpId|" & var4

		StrEbrFile = "c2220oa1"
	Else
		StrEbrFile = "c2220oa2"
	End If
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
	
	call FncEBRprint(EBAction, ObjName, strUrl)
	
	Call BtnDisabled(0)
End Function


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

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
		
	arrField(0) = "PLANT_CD"				' Field��(0)
	arrField(1) = "PLANT_NM"				' Field��(1)
	    
	arrHeader(0) = "����"				' Header��(0)
	arrHeader(1) = "�����"				' Header��(1)
	    
	arrRet = window.showModalDialog("../../comasp/ADOcommonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlantCd(arrRet)
	End If	
	
End Function
'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenIremCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.Value)       ' Plant Code

	If iWhere = "From" Then
		arrParam(1) = Trim(frm1.txtFrItemCd.Value)	' Item Code
	Else
		arrParam(1) = Trim(frm1.txtToItemCd.Value)	' Item Code
	End If

	arrParam(2) = "12"						        ' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"	
'    arrField(2) = 3								' Field��(1) : "ITEM_ACCT"
    
	arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent,arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	  If iWhere = "From" Then
	     frm1.txtFrItemCd.focus
	  ElseIf iWhere = "To" Then
		 frm1.txtToItemCd.focus
	  End If				    
		Exit Function
	Else
		Call SetItemCd(arrRet, iWhere)
	End If	
End Function

Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.focus
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)  
End Function

Function SetItemCd(ByVal arrRet, ByVal iWhere)
	If iWhere = "From" Then
		frm1.txtFrItemCd.focus
		frm1.txtFrItemCd.value = arrRet(0) 
		frm1.txtFrItemNm.value = arrRet(1)
	Else
		frm1.txtToItemCd.focus
		frm1.txtToItemCd.value = arrRet(0) 
		frm1.txtToItemNm.value = arrRet(1)
	End If
End Function


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǥ�ؿ��� ����</font></td>
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��±���</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="OptFlag" CHECKED ID="OptFlag1" VALUE="Y" tag="25"><LABEL FOR="OptFlag1">����</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="OptFlag" ID="OptFlag2" VALUE="N" tag="25"><LABEL FOR="OptFlag2">����</LABEL></SPAN></TD>
							</TR>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="X2XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=40 tag="X4" ALT="�����"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>ǰ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFrItemCd" SIZE=20 MAXLENGTH=18 tag="X2XXXU" ALT="����ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd('From')">&nbsp;<INPUT TYPE=TEXT NAME="txtFrItemNm" SIZE=30 MAXLENGTH=40 tag="X4XXXU" ALT="ǰ���"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>��</TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=20 MAXLENGTH=18 tag="X2XXXU" ALT="����ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd('To')">&nbsp;<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=30 MAXLENGTH=40 tag="X4XXXU" ALT="ǰ���"></TD>
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtSpId" tag="24" TABINDEX = "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX = "-1">
    <input type="hidden" name="dbname" TABINDEX = "-1">
    <input type="hidden" name="filename" TABINDEX = "-1">
    <input type="hidden" name="condvar" TABINDEX = "-1">
	<input type="hidden" name="date" TABINDEX = "-1">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>
