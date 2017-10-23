<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production 
'*  2. Function Name        : 
'*  3. Program ID           :  p1203oa1.asp
'*  4. Program Name         :  ����� ���� ��� 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 2002/12/18
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
<Script LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim IsOpenPop

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
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "P","NOCOOKIE","OA") %>		
End Sub

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
		frm1.txtItemCd1.focus
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
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLE)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
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
	Dim var5
	Dim var6
	
	Dim strUrl, strEbrFile

	Call BtnDisabled(1)	
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
		Call DisplayMsgBox("971012","X", "����","X")
		Call BtnDisabled(0)	
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
		
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	
	if frm1.txtItemCd1.value = "" then 
		var2 = "0"	
	else
		var2 = Trim(frm1.txtItemCd1.value)
	End If
	
	if frm1.txtItemCd2.value = "" then 
		var3 = "zzzzzzzzzzzzzzzzzz"	
	else
		var3 = Trim(frm1.txtItemCd2.value)
	End If
	
	if frm1.txtRoutingNo1.value = "" then 
		var4 = "0"	
	else
		var4 = Trim(frm1.txtRoutingNo1.value)
	End If
	
	if frm1.txtRoutingNo2.value = "" then
		var5 = "zzzzzzz"
	else	
		var5 = Trim(frm1.txtRoutingNo2.value)
	end if
	
	If frm1.rdoPrintType1.checked  = True Then	  
		var6 = "P1203OA1"
	ElseIf frm1.rdoPrintType2.checked = True Then
		var6 = "P1203OA2"
	ElseIf frm1.rdoPrintType3.checked = True Then 
		var6 = "P1203OA3"
	Else 
		var6 = "P1203OA4"	
	End If	
	
	strEbrFile = AskEBDocumentName(var6, "EBR")

	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|item_cd1|" & var2 
	strUrl = strUrl & "|item_cd2|" & var3 	
	strUrl = strUrl & "|rout_no1|" & var4 
	strUrl = strUrl & "|rout_no2|" & var5 
	
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
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview() 
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
		
	Dim strUrl, strEbrFile
	Dim arrParam, arrField, arrHeader
	
	Call BtnDisabled(1)	
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
		Call DisplayMsgBox("971012","X", "����","X")
		Call BtnDisabled(0)	
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
		
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	
	if frm1.txtItemCd1.value = "" then 
		var2 = "0"	
	else
		var2 = Trim(frm1.txtItemCd1.value)
	End If
	
	if frm1.txtItemCd2.value = "" then 
		var3 = "zzzzzzzzzzzzzzzzzz"	
	else
		var3 = Trim(frm1.txtItemCd2.value)
	End If
	
	if frm1.txtRoutingNo1.value = "" then 
		var4 = "0"	
	else
		var4 = Trim(frm1.txtRoutingNo1.value)
	End If
	
	if frm1.txtRoutingNo2.value = "" then
		var5 = "zzzzzzz"
	else	
		var5 = Trim(frm1.txtRoutingNo2.value)
	end if
	
	If frm1.rdoPrintType1.checked  = True Then	  
		var6 = "P1203OA1"
	ElseIf frm1.rdoPrintType2.checked = True Then
		var6 = "P1203OA2"
	ElseIf frm1.rdoPrintType3.checked = True Then 
		var6 = "P1203OA3"
	Else 
		var6 = "P1203OA4"	
	End If	
	
	strEbrFile = AskEBDocumentName(var6, "EBR")

	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|item_cd1|" & var2 
	strUrl = strUrl & "|item_cd2|" & var3 	
	strUrl = strUrl & "|rout_no1|" & var4 
	strUrl = strUrl & "|rout_no2|" & var5 
	
	call FncEBRPrevIew(strEbrFile, strUrl)
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
'	Name : OpenIremCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal strCode, ByVal iPos)
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
	arrParam(1) = strCode	' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"	
    arrField(2) = 3								' Field��(1) : "ITEM_ACCT"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet, iPos)
	End If	
	
	Call SetFocusToDocument("M")
	If iPos = 1 Then
		frm1.txtItemCd1.focus
	Elseif iPos = 2 Then
		frm1.txtItemCd2.focus
	End If
	
End Function

'===========================================================================
' Function Name : OpenRoutingNo
' Function Desc : OpenRoutingNo Reference Popup
'===========================================================================
Function OpenRoutingNo(ByVal strCode, ByVal iPos)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtRoutingNo1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True
		
	arrParam(0) = "����� �˾�"					' �˾� ��Ī 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE ��Ī 
	arrParam(2) = strCode		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S") ' & "And ITEM_CD =" & "'" & Trim(frm1.txtItemCd.value)	& "'"							' Where Condition
	arrParam(5) = "�����"			
	
    arrField(0) = "ROUT_NO"							' Field��(0)
    arrField(1) = "DESCRIPTION"						' Field��(1)
    arrField(2) = "MAJOR_FLG"						' Field��(1)
    
    arrHeader(0) = "�����"						' Header��(0)
    arrHeader(1) = "����ø�"					' Header��(1)
    arrHeader(2) = "�ֶ���ÿ���"				' Header��(1)
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
	If arrRet(0) <> "" Then
		Call SetRoutingNo(arrRet, iPos)
	End If
	
	Call SetFocusToDocument("M")
	If iPos = 1 Then
		frm1.txtRoutingNo1.focus
	Elseif iPos = 2 Then
		frm1.txtRoutingNo2.focus
	End If
	
End Function

Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)  
End Function

Function SetItemCd(ByVal arrRet, ByVal iPos)
	If iPos = 1 Then
		frm1.txtItemCd1.value = arrRet(0) 	
	Elseif iPos = 2 Then
		frm1.txtItemCd2.value = arrRet(0) 	
	End If
End Function

'------------------------------------------  SetRoutingNo()  --------------------------------------------------
'	Name : SetRoutingNo()
'	Description : RoutingNo Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRoutingNo(Byval arrRet, ByVal iPos)
	If iPos = 1 Then
		frm1.txtRoutingNo1.value = arrRet(0)	
	ElseIf iPos = 2 Then
		frm1.txtRoutingNo2.value = arrRet(0)	
	End If
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������������</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="X2XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="�����"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=30 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd1.value, 1">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=30 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd2.value, 2"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutingNo1" SIZE=15 MAXLENGTH=7 tag="X1XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutingNo1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutingNo frm1.txtRoutingNo1.value, 1">&nbsp;~&nbsp;
										<INPUT TYPE=TEXT NAME="txtRoutingNo2" SIZE=15 MAXLENGTH=7 tag="X1XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutingNo2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutingNo frm1.txtRoutingNo2.value, 2"></TD>
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
									<TD CLASS=TD5 NOWRAP>��¹��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType1" CLASS="RADIO" tag="XX" CHECKED><LABEL FOR="rdoPrintType1">������������</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType2" CLASS="RADIO" tag="XX" ><LABEL FOR="rdoPrintType2">�����/�����������</LABEL><TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType3" CLASS="RADIO" tag="XX" ><LABEL FOR="rdoPrintType3">�����/����/��ǰ�������������</LABEL><TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType4" CLASS="RADIO" tag="XX" ><LABEL FOR="rdoPrintType4">�����/����/�ڿ������������</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
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
