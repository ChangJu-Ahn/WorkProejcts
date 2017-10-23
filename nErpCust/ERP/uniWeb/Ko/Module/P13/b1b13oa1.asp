<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b13oa1.asp
'*  4. Program Name         : ��üǰ����� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : 
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
	Else
		frm1.txtPlantCd.focus 	 
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
		Exit Function
	End If

	var1 = UCase(Trim(frm1.txtPlantCd.value))
		
	if frm1.txtItemCd1.value = "" then 
		frm1.txtItemNm1.value = ""
		var2 = "0"	
	else
		var2 = UCase(Trim(frm1.txtItemCd1.value))
	End If


	if frm1.txtItemCd2.value = "" then 
		frm1.txtItemNm2.value = ""
		var3 = "zzzzzzzzzzzzzzzzzz"	
	else
		var3 = UCase(Trim(frm1.txtItemCd2.value))
	End If
	
	
	If frm1.rdoValidFlg2.checked  = True Then	  
		var4 = "Y"									
		var5 = "Y"											 
	ElseIf frm1.rdoValidFlg3.checked  = True Then
		var4 = "N"	
		var5 = "N"								
	Else
		var4 = "Y"	
		var5 = "N"											 
	End If
	
	
	If frm1.rdoSortBy1.checked  = True Then	  
		var6 = "B_ALTERNATIVE_ITEM.ITEM_CD ASC"										 
	Else 
		var6 = "B_ITEM.ITEM_NM ASC"	
	End If
	
	strEbrFile = AskEBDocumentName("B1B13OA1", "EBR")

	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|fr_item_cd|" &  var2
	strUrl = strUrl & "|to_item_cd|" &  var3 
	strUrl = strUrl & "|valid_flg|" &  var4 
	strUrl = strUrl & "|valid_flg1|" &  var5 
	strUrl = strUrl & "|sort_list|" &  var6  
	
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
	
	Call BtnDisabled(1)	
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
		Call DisplayMsgBox("971012","X", "����","X")
		Call BtnDisabled(0)	
		frm1.txtPlantCd.focus 
		Exit Function
	End If

	var1 = UCase(Trim(frm1.txtPlantCd.value))
		
	if frm1.txtItemCd1.value = "" then 
		frm1.txtItemNm1.value = ""
		var2 = "0"	
	else
		var2 = UCase(Trim(frm1.txtItemCd1.value))
	End If


	if frm1.txtItemCd2.value = "" then 
		frm1.txtItemNm2.value = ""
		var3 = "zzzzzzzzzzzzzzzzzz"	
	else
		var3 = UCase(Trim(frm1.txtItemCd2.value))
	End If
	
	
	If frm1.rdoValidFlg2.checked  = True Then	  
		var4 = "Y"									
		var5 = "Y"											 
	ElseIf frm1.rdoValidFlg3.checked  = True Then
		var4 = "N"	
		var5 = "N"								
	Else
		var4 = "Y"	
		var5 = "N"											 
	End If
	
	
	If frm1.rdoSortBy1.checked  = True Then	  
		var6 = "B_ALTERNATIVE_ITEM.ITEM_CD ASC"										 
	Else 
		var6 = "B_ITEM.ITEM_NM ASC"	
	End If
	
	strEbrFile = AskEBDocumentName("B1B13OA1", "EBR")

	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|fr_item_cd|" &  var2
	strUrl = strUrl & "|to_item_cd|" &  var3 
	strUrl = strUrl & "|valid_flg|" &  var4 
	strUrl = strUrl & "|valid_flg1|" &  var5 
	strUrl = strUrl & "|sort_list|" &  var6  

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
'	Name : OpenItemCd()
'	Description : Item1 PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal strCode, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
		
	If IsOpenPop = True Then Exit Function
		
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.Value)						   ' Plant Code
	arrParam(1) = strCode	' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"	
    arrField(2) = 3								' Field��(1) : "ITEM_ACCT"
    
	iCalledAspName = AskPRAspName("B1B11PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA1", "X")
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
	If iPos = 0 Then
		frm1.txtItemCd1.focus
	Else
		frm1.txtItemCd2.focus
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

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : ItemCd1 Popup���� return�� �� 
'---------------------------------------------------------------------------------------------------------
 

Function SetItemCd(ByVal arrRet, ByVal iPos)	
	If iPos = 0 Then
		frm1.txtItemCd1.value = arrRet(0) 
		frm1.txtItemNm1.value = arrRet(1)
	ElseIf iPos = 1 Then
		frm1.txtItemCd2.value = arrRet(0) 
		frm1.txtItemNm2.value = arrRet(1)
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��üǰ���</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>		
								<TR>	
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="X2XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=40 tag="X4" ALT="�����"></TD>
								</TR>						
								<TR>
									<TD CLASS="TD5" NOWRAP>����ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=30 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ���ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd1.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=30 MAXLENGTH=40 tag="X4" ALT="ǰ���ڵ�">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=30 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ���ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd2.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=30 MAXLENGTH=40 tag="X4" ALT="ǰ���ڵ�">&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>��ȿǰ�񱸺�</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" CHECKED ID="rdoValidFlg1" tag="X1"><LABEL FOR="rdoValidFlg1">��ü</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg2" tag="X1"><LABEL FOR="rdoValidFlg2">��</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg3" tag="X1"><LABEL FOR="rdoValidFlg3">�ƴϿ�</LABEL>																	
									</TD>
								</TR>
																					
								<TR>
									<TD CLASS="TD5" NOWRAP>��¼���</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSortBy" CHECKED ID="rdoSortBy1" tag="X1"><LABEL FOR="rdoSortBy1">����ǰ���ڵ��</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSortBy" ID="rdoSortBy2" tag="X1"><LABEL FOR="rdoSortBy2">����ǰ����</LABEL>																	
									</TD>
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
				</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
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
