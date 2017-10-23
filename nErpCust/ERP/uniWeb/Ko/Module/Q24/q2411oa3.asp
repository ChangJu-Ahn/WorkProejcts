<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2111OA3
'*  4. Program Name         : ������� 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop          
Dim strYr
Dim strMonth
Dim strDay
Call ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gServerDateType, strYr, strMonth, strDay)

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
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
	frm1.cboInspClassCd.value		= "S"
	Call ggoOper.FormatDate(frm1.txtStartDt, Parent.gDateFormat, 3)
	frm1.txtStartDt.Text = strYr
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","OA") %>
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspClassCd , lgF0, lgF1, Chr(11))
End Sub

'------------------------------------------  OpenPlant() -------------------------------------------------
'	Name : OpenPlant()
'	Description :Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false 
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			

    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	

    arrHeader(0) = "�����ڵ�"		
    arrHeader(1) = "�����"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam,arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	End If	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false 
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'���������� �ʿ��մϴ� 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "ǰ��"																	' �˾� ��Ī 
	arrParam(1) = "B_Item_By_Plant,B_Item"												' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemCd.Value)													' Code Condition
	arrParam(3) = ""													' Name Condition
	arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd"
	arrParam(4) = arrParam(4) & "  And B_Item_By_Plant.Plant_Cd =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & "" 			' Where Condition
	arrParam(5) = "ǰ��"																	' TextBox ��Ī 
	
	arrField(0) = "B_Item_By_Plant.Item_Cd"					' Field��(0)
	arrField(1) = "B_Item.Item_NM"				' Field��(1)
	arrField(2) = "B_Item.SPEC"					' Field��(2)
	
	arrHeader(0) = "ǰ���ڵ�"						' Header��(0)
	arrHeader(1) = "ǰ���"					' Header��(1)
	arrHeader(2) = "�԰�"						' Header��(2)
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If Trim(arrRet(0)) <> "" Then
		frm1.txtItemCd.Value = Trim(arrRet(0))
		frm1.txtItemNm.Value = Trim(arrRet(1))
	End If
	frm1.txtItemCd.Focus
	Set gActiveElement = document.activeElement
	OpenItem = true	
End Function

 '------------------------------------------  OpenBp()  -------------------------------------------------
'	Name : OpenBp()
'	Description : Supplier PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp()
	OpenBp = false
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ŷ�ó"						' �˾� ��Ī 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBpCd.Value)		' Code Condition
	arrParam(3) = ""				' Name Cindition
	arrParam(4) = "( BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")  )" 	' Where Condition
	arrParam(5) = "�ŷ�ó"						' TextBox ��Ī 
	
	arrField(0) = "BP_CD"							' Field��(0)
	arrField(1) = "BP_NM"							' Field��(1)
	
	arrHeader(0) = "�ŷ�ó�ڵ�"					' Header��(0)
	arrHeader(1) = "�ŷ�ó��"					' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
	End If
	frm1.txtBpCd.Focus
	Set gActiveElement = document.activeElement	
	OpenBp = TRUE
End Function

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	
	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables                                                      '��: Initializes local global variables
	
	 '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemCd.focus 
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

'=======================================================================================================
'   Event Name : txtStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStartDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtStartDt.Action = 7   
		Call SetFocusToDocument("M")
        frm1.txtStartDt.Focus 
	End If 
End Sub

Function txtStartDt_KeyPress(KeyAscii)
	txtStartDt_KeyPress = false
	
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
	
	txtStartDt_KeyPress = true
End Function

 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery()
	FncQuery = true
	
	If FncBtnPreview = False Then Exit function
	
	FncQuery = True                                                   '�̸����� Call
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	FncFind = false 
	
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
    
    FncFind = true
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
	
	Call Parent.FncPrint()
	
	FncPrint = True
End Function

'========================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function FncBtnPrint() 
	dim var1, var2, var3, var4, var5, var6, var7, var8
	dim condvar

	Dim strEbrFile
	Dim objName
	Dim strUrl
	
	FncBtnPrint = false
	
	If Not chkField(Document, "1") Then	Exit Function
	
	If Plant_Item_Check = False Then Exit Function		

	var1 = Trim(frm1.txtPlantCd.value)
	
	var2 = Trim(frm1.txtStartDt.Text)
	
	var4 = Trim(frm1.cboInspClassCd.value)
	
	var6 = Trim(frm1.txtBpCd.value)

	If var6  = "" Then
		var6 = "%"
	End If
	
	var7 = Trim(frm1.txtItemCd.value)
	If var7  = "" Then
		var7 = "%"
	End If
	
' ������ ��ȸŸ�� ����(ǰ��,�ŷ�ó��,ǰ��/�ŷ�ó��) 
	If frm1.RadioOutputType.rdoCase2.Checked Then
		strEbrFile = "Q2411OA32"
	ElseIf frm1.RadioOutputType.rdoCase3.Checked Then
		strEbrFile = "Q2411OA33"
	ElseIf frm1.RadioOutputType.rdoCase4.Checked Then
		strEbrFile = "Q2411OA34"
	End if
	
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	strUrl = strUrl & "Yr|" & var2 
	strUrl = strUrl & "|InspClassCd|" & var4 
	strUrl = strUrl & "|PlantCd|" & var1
	strUrl = strUrl & "|BpCd|" & var6
	strUrl = strUrl & "|ItemCd|" & var7
		
	call FncEBRprint(EBAction, objName, strUrl)

	FncBtnPrint = True			
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview() 
	dim var1, var2, var3, var4, var5, var6, var7, var8
	dim condvar

	Dim strEbrFile
	Dim objName
	Dim strUrl
	
	FncBtnPreview = false
	
	If Not chkField(Document, "1") Then	Exit Function
	
	If Plant_Item_Check = False Then Exit Function		
	
	var1 = Trim(frm1.txtPlantCd.value)
	
	var2 = Trim(frm1.txtStartDt.Text)
	
	var4 = Trim(frm1.cboInspClassCd.value)
	
	var6 = Trim(frm1.txtBpCd.value)

	If var6  = "" Then
		var6 = "%"
	End If
	
	var7 = Trim(frm1.txtItemCd.value)
	If var7  = "" Then
		var7 = "%"
	End If
	
' ������ ��ȸŸ�� ����(ǰ��,�ŷ�ó��,ǰ��/�ŷ�ó��) 
	If frm1.RadioOutputType.rdoCase2.Checked Then
		strEbrFile = "Q2411OA32"
	ElseIf frm1.RadioOutputType.rdoCase3.Checked Then
		strEbrFile = "Q2411OA33"
	ElseIf frm1.RadioOutputType.rdoCase4.Checked Then
		strEbrFile = "Q2411OA34"
	End if
	
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	strUrl = strUrl & "Yr|" & var2 
	strUrl = strUrl & "|InspClassCd|" & var4 
	strUrl = strUrl & "|PlantCd|" & var1
	strUrl = strUrl & "|BpCd|" & var6
	strUrl = strUrl & "|ItemCd|" & var7
	
	call FncEBRPreview(objName, strUrl)

	FncBtnPreview = true		
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function


'========================================================================================
' Function Name : Plant_Item_Check
'========================================================================================
Function Plant_Item_Check()
	
	Plant_Item_Check = False

	With frm1
 
		'-----------------------
		'Check Plant CODE  
		'-----------------------
		If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
  
			Call DisplayMsgBox("125000","X","X","X")
			.txtPlantNm.Value = ""
			.txtPlantCd.Focus
			Set gActiveElement = document.activeElement
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		.txtPlantNm.Value = lgF0(0)

		If .txtItemCd.value <> "" Then
			If  CommonQueryRs(" B.ITEM_NM "," B_ITEM_BY_PLANT A, B_ITEM B ", " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S") & " AND A.ITEM_CD = " & FilterVar(.txtItemCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

				If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(.txtItemCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
					Call DisplayMsgBox("122600","X","X","X")
					.txtItemNm.Value = ""
					.txtItemCd.Focus
					Set gActiveElement = document.activeElement
					Exit function
				Else
					lgF0 = Split(lgF0, Chr(11))
					.txtItemNm.Value = lgF0(0)
					Call DisplayMsgBox("122700","X","X","X")
					.txtItemCd.Focus
					Set gActiveElement = document.activeElement
					Exit function
				End If
			End If
			lgF0 = Split(lgF0, Chr(11))
			.txtItemNm.Value = lgF0(0)
		Else
			.txtItemNm.Value = ""
		End if 
	 
		If .txtBpCd.value <> "" Then
			If 	CommonQueryRs(" BP_NM, BP_TYPE, usage_flag, in_out_flag "," B_Biz_Partner ", " BP_CD = " & FilterVar(.txtBpCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
						
				Call DisplayMsgBox("229927","X","X","X")
				.txtBpNm.Value = ""
				.txtBpCd.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			lgF1 = Split(lgF1, Chr(11))
			lgF2 = Split(lgF2, Chr(11))
			lgF3 = Split(lgF3, Chr(11))
			.txtBpNm.Value = lgF0(0)

			If Trim(lgF2(0)) <> "Y" Then
				Call DisplayMsgBox("179021","X","X","X")
				.txtBpCd.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
			If Trim(lgF1(0)) <> "C" and Trim(lgF1(0)) <> "CS" Then
				Call DisplayMsgBox("179020","X","X","X")
				.txtBpCd.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
			If Trim(lgF3(0)) <> "O" Then
				Call DisplayMsgBox("17C003","X","X","X")
				.txtBpCd.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
		End If
	End With       

 Plant_Item_Check = True

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
		<TD HEIGHT=5 colspan="2">&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100% colspan="2">
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ϰ˻翬��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% CLASS="Tab11" HEIGHT=* colspan="2">
			<TABLE CLASS="BasicTB" CELLSPACING=0 STYLE="HEIGHT: 100%">	
	    		<TR>
					<TD WIDTH=100%>
						<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT: 100%">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0 STYLE="HEIGHT: 100%">
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" TAG="12XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnPlantCd ONCLICK=vbscript:OpenPlant() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm" TAG="14X"></TD>
                                    				</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/q2411oa3_fpDateTime1_txtStartDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�˻�з�</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="�˻�з�" STYLE="WIDTH: 150px" tag="14"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="ǰ��" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnItemCd align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenItem()">
												<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=20 ALT="�ŷ�ó" tag="11XXXU"><IMG align=top height=20 name=btnBpCd onclick="vbscript:OpenBp()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
												<INPUT TYPE=TEXT NAME="txtBpnm" SIZE=20 MAXLENGTH=20 tag="14" >
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X" Checked><LABEL FOR="rdoCase2">ǰ��</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase3" TAG="1X"><LABEL FOR="rdoCase3">�ŷ�ó��</LABEL>              
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase4" TAG="1X"><LABEL FOR="rdoCase4">ǰ��/�ŷ�ó��</LABEL>              
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
	    <TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
				    	<BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;              
		                <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>                
		            </TD>                  
				</TR>
			</TABLE>
		</TD>  
	</TR>              
	<TR>              
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm "  tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>              
		</TD>              
	</TR>              
</TABLE>              
</FORM>              
<DIV ID="MousePT" NAME="MousePT">              
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>              
</DIV>              
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST"> 
    <input type="hidden" name="uname" tabindex=-1>
    <input type="hidden" name="dbname" tabindex=-1>
    <input type="hidden" name="filename" tabindex=-1>
    <input type="hidden" name="condvar" tabindex=-1>
	<input type="hidden" name="date" tabindex=-1>
</FORM>                
</BODY>              
</HTML>
