<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production 
'*  2. Function Name        : 
'*  3. Program ID           :  p1504oa1.asp
'*  4. Program Name         :  �ڿ��� Shift���ܵ�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/12/13
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Ryu Sung Won
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
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************


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
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim IsOpenPop
'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

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
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P","NOCOOKIE","QA") %>
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
    Call InitVariables                                                      '��: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtResourceCd1.focus
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

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)	                                  '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
 Function FncPrint() 
    Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint()
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	
	Dim strUrl

	Call BtnDisabled(1)	
	
	If frm1.txtResourceCd1.value = "" Then
		frm1.txtResourceNm1.value = ""
	End If
	
	If frm1.txtResourceCd2.value = "" Then
		frm1.txtResourceNm2.value = ""
	End If
	
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
		Call DisplayMsgBox("971012","X", "����","X")
		Call BtnDisabled(0)	
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	
	If frm1.txtResourceCd1.value = "" Then
		var2 = "0"
	Else
		var2 = UCase(Trim(frm1.txtResourceCd1.value))
	End If
	
	if frm1.txtResourceCd2.value = "" then 
		var3 = "zzzzzzz"
	else
		var3 = UCase(Trim(frm1.txtResourceCd2.value))	
	End If
	
	If frm1.rdoSortBy1.checked  = True Then	  
		var4 = "P_SHIFT_EXCEPTION.RESOURCE_CD"										 
	Else 
		var4 = "P_RESOURCE.DESCRIPTION"	
	End If

	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|resource_cd1|" & var2 
	strUrl = strUrl & "|resource_cd2|" & var3 
	strUrl = strUrl & "|sort_by|" & var4 
	
'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	Dim ObjName
	ObjName = AskEBDocumentName("p1504oa1","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	
    
    frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement
    
End Function


'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview() 
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	
	Dim strUrl
	
	Call BtnDisabled(1)	
	
	If frm1.txtResourceCd1.value = "" Then
		frm1.txtResourceNm1.value = ""
	End If
	
	If frm1.txtResourceCd2.value = "" Then
		frm1.txtResourceNm2.value = ""
	End If
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
		Call DisplayMsgBox("971012","X", "����","X")
		Call BtnDisabled(0)	
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	
	If frm1.txtResourceCd1.value = "" Then
		var2 = "0"
	Else
		var2 = UCase(Trim(frm1.txtResourceCd1.value))
	End If
	
	if frm1.txtResourceCd2.value = "" then 
		var3 = "zzzzzzz"	
	else
		var3 = UCase(Trim(frm1.txtResourceCd2.value))
	End If
	
	If frm1.rdoSortBy1.checked  = True Then	  
		var4 = "P_SHIFT_EXCEPTION.RESOURCE_CD"										 
	Else 
		var4 = "P_RESOURCE.DESCRIPTION"	
	End If
	
	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|resource_cd1|" & var2 
	strUrl = strUrl & "|resource_cd2|" & var3 
	strUrl = strUrl & "|sort_by|" & var4 
	
	Dim ObjName
	ObjName = AskEBDocumentName("p1504oa1","ebr")
	Call FncEBRPreview(ObjName, strUrl)
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement

End Function

'========================================================================================
' Function Name : PrevExecOk()
' Function Desc : BOM Temp ���̺� ������ ������ �����ϸ� EasyBase�� Open�Ѵ�.
'========================================================================================


Function PrevExecOk()

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

	arrParam(0) = "�����˾�"			' �˾� ��Ī 
	arrParam(1) = "B_PLANT"			' TABLE ��Ī 
	arrParam(2) = UCase(Trim(frm1.txtPlantCd.Value))	' Code Condition
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

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource1 PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource(ByVal strCode, ByVal iPos)

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)


	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "�ڿ��˾�"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & " "			
	arrParam(5) = "�ڿ�"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ�"		
    arrHeader(1) = "�ڿ���"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet, iPos)
	End If
	
	Call SetFocusToDocument("M")
	If iPos = 0 Then
		frm1.txtResourceCd1.focus
	ElseIf iPos = 1 Then
		frm1.txtResourceCd2.focus
	End If
		
End Function

'------------------------------------------  SetPlantCd()  --------------------------------------------------
'	Name : SetPlantCd()
'	Description : Plant  Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)  
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(ByVal arrRet, ByVal iPos)
	If iPos = 0 Then
		frm1.txtResourceCd1.Value    = arrRet(0)		
		frm1.txtResourceNm1.Value    = arrRet(1)		
	ElseIf iPos = 1 Then
		frm1.txtResourceCd2.Value    = arrRet(0)		
		frm1.txtResourceNm2.Value    = arrRet(1)		
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

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
		<TD HEIGHT=5 colspan="2">&nbsp;</TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ڿ���Shift�������</font></td>
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
									<TD CLASS=TD5 NOWRAP>�ڿ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd1" SIZE=15 MAXLENGTH=10 tag="X1XXXU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource frm1.txtResourceCd1.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm1" SIZE=30 tag="X4" ALT="�ڿ���">&nbsp;~&nbsp;</TD>								
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd2" SIZE=15 MAXLENGTH=10 tag="X1XXXU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource frm1.txtResourceCd2.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm2" SIZE=30 tag="X4" ALT="�ڿ���">&nbsp;</TD>								
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoSortBy" ID="rdoSortBy1" CLASS="RADIO" tag="XX" CHECKED><LABEL FOR="rdoSortBy1">�ڿ��ڵ�</LABEL>&nbsp;&nbsp;&nbsp;
													       <INPUT TYPE="RADIO" NAME="rdoSortBy" ID="rdoSortBy2" CLASS="RADIO" tag="XX" ><LABEL FOR="rdoSortBy2">�ڿ���</LABEL></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TabIndex="-1">
    <input type="hidden" name="dbname" TabIndex="-1">
    <input type="hidden" name="filename" TabIndex="-1">
    <input type="hidden" name="condvar" TabIndex="-1">
	<input type="hidden" name="date" TabIndex="-1">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>
