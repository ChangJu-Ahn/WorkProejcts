<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : ���ݾ׻����۾� 
'*  3. Program ID           : s4151ba1_ko441
'*  4. Program Name         : ���ݾ׻����۾� 
'*  5. Program Desc         :  
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2008/08/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*                           
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "s4151bb1_ko441.asp"
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim EndDate

' �ý��� ��¥ 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'==========================================
Sub InitVariables()
End Sub

'=========================================
Sub SetDefaultVal()
    frm1.txtFr_dt.focus
	frm1.txtFr_dt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtTo_dt.Text = EndDate
    frm1.txtconBp_cd.focus 
	
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub


'=========================================
Sub Form_Load()

    Call LoadInfTB19029														'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables														
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'��: ��ư ���� ���� 
End Sub

'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub


'===========================================================================
Function OpenConSItemDC(Byval strCode, Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 Select Case iWhere
 Case 5
  arrParam(1) = "B_BIZ_PARTNER"      <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode) <%' Code Condition%>
  arrParam(3) = ""                                    <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"             <%' Where Condition%>
  arrParam(5) = "��"          <%' TextBox ��Ī %>
 
  arrField(0) = "BP_CD"           <%' Field��(0)%>
  arrField(1) = "BP_NM"           <%' Field��(1)%>
    
  arrHeader(0) = "��"          <%' Header��(0)%>
  arrHeader(1) = "����"            <%' Header��(1)%>
  frm1.txtconBp_cd.focus 
 Case 0
  arrParam(1) = "b_item"                           <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode)  <%' Code Condition%>
  arrParam(3) = ""                           <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "ǰ��"       <%' TextBox ��Ī %>
 
  arrField(0) = "item_cd"        <%' Field��(0)%>
  arrField(1) = "item_nm"        <%' Field��(1)%>
  arrField(2) = "spec"        <%' Field��(1)%>
    
  arrHeader(0) = "ǰ��"       <%' Header��(0)%>
  arrHeader(1) = "ǰ���"       <%' Header��(1)%> 
  arrHeader(2) = "�԰�"       <%' Header��(1)%>  
 Case 1
  arrParam(1) = "B_USER_DEFINED_MINOR"                           <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode)  <%' Code Condition%>
  arrParam(3) = ""                           <%' Name Cindition%>
  arrParam(4) = "UD_MAJOR_CD='ZZ002'"         <%' Where Condition%>
  arrParam(5) = "���TYPE"       <%' TextBox ��Ī %>
 
  arrField(0) = "UD_MINOR_CD"        <%' Field��(0)%>
  arrField(1) = "UD_MINOR_NM"        <%' Field��(1)%>
    
  arrHeader(0) = "���TYPE"       <%' Header��(0)%>
  arrHeader(1) = "���TYPE��"       <%' Header��(1)%> 
 Case 4					'���� 
	arrParam(1) = "B_PLANT"								
	arrParam(2) = Trim(strCode)				
	arrParam(4) = ""									
	arrParam(5) = "����"							

	arrField(0) = "PLANT_CD"							
	arrField(1) = "PLANT_NM"							

	arrHeader(0) = "����"							
	arrHeader(1) = "�����"							
	
	frm1.txtPlantCode.focus

 End Select
    
    arrParam(3) = "" 
 arrParam(0) = arrParam(5)        <%' �˾� ��Ī %>

	Select Case iWhere
		Case 0
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
		Case Else
           arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

 IsOpenPop = False

 If arrRet(0) = "" Then
	 Exit Function
 Else
	 With frm1
	  Select Case iWhere
	  Case 0
	   .txtconItem_cd.value = arrRet(0) 
	   .txtconItem_nm.value = arrRet(1)   
	  Case 1
	   .txtconOutType.value = arrRet(0) 
	   .txtconOutTypeNm.value = arrRet(1)   
	  Case 4
		.txtPlantCode.value = arrRet(0) 
		.txtPlantName.value = arrRet(1)   
	  Case 5
	   .txtconBp_cd.value = arrRet(0) 
	   .txtconBp_Nm.value = arrRet(1)    
	  End Select
	 End With
 End If 
 
End Function


Function txtconBp_cd_OnChange()
    txtconBp_cd_OnChange = true
    
    If  frm1.txtconBp_cd.value = "" Then
        frm1.txtconBp_nm.value = ""
        frm1.txtconBp_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" BP_NM "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(frm1.txtconBp_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","�ŷ�ó","x")

            frm1.txtconBp_nm.value = ""
	        frm1.txtconBp_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtconBp_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function


Function txtPlantCode_OnChange()
    If  frm1.txtPlantCode.value <> "" Then
        if   CommonQueryRs(" plant_nm "," B_PLANT "," plant_cd =  " & FilterVar(frm1.txtPlantCode.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtPlantName.value = ""
            Call  DisplayMsgBox("970000", "x","�����ڵ�","x")
	        frm1.txtPlantCode.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtPlantName.value = Replace(lgF0, Chr(11), "")
	    End If
	else 
		 frm1.txtPlantName.value=""
    End If

End Function

Function txtconItem_cd_OnChange()
   txtconItem_cd_OnChange = true
    
    If  frm1.txtconItem_cd.value = "" Then
        frm1.txtconItem_nm.value = ""
        frm1.txtconItem_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" ITEM_NM "," B_ITEM "," ITEM_CD = " & FilterVar(frm1.txtconItem_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","ǰ���ڵ�","x")

            frm1.txtconItem_nm.value = ""
	        frm1.txtconItem_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtconItem_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function

'=========================================
Sub txtFr_dt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFr_dt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtFr_dt.Focus
    End If
End Sub

'=========================================
Sub txtTo_dt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTo_dt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtTo_dt.Focus
    End If
End Sub

'=====================================================
 Function FncPrint() 
	Call parent.FncPrint()
End Function

'=====================================================
 Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function


'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect()
	Dim  strVal
	Dim  IntRetCD

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0) 
		Exit Function
	End If

	If ValidDateCheck(frm1.txtFr_dt, frm1.txtTo_dt) = False Then Exit Function

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If LayerShowHide(1) = False Then
	     Call BtnDisabled(0)
	     Exit Function
	End If
	ExeReflect = False                                                          '��: Processing is NG
    
	On Error Resume Next                                                   '��: Protect system from crashing
	Call BtnDisabled(1)  

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	strVal = strVal & "&txtconBp_cd=" & frm1.txtconBp_cd.value
	strVal = strVal & "&txtPlantCode=" & frm1.txtPlantCode.value
	strVal = strVal & "&txtconItem_cd=" & frm1.txtconItem_cd.value
    strVal = strVal & "&txtFr_dt=" & frm1.txtFr_dt.text
    strVal = strVal & "&txtTo_dt=" & frm1.txtTo_dt.text

    if  frm1.txtVol_flag1.checked = true then
	    strVal = strVal & "&txtVol_flag=Y"  
	else
	    strVal = strVal & "&txtVol_flag=N" 
    end if
    if  frm1.txtQty_flag1.checked = true then
	    strVal = strVal & "&txtQty_flag=Y"
	else
	    strVal = strVal & "&txtQty_flag=N"
    end if

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

	ExeReflect = True                                                           '��: Processing is NG
	Call BtnDisabled(0)
End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function ExeReflectOk()				            '��: ���� ������ ���� ���� 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	window.status = "�۾� �Ϸ�"

End Function

Function ExeReflectNo()				            '��: ó���� ����Ÿ�� �����ϴ�.
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("800161","X","X","X")
	window.status = "�۾� �Ϸ�"

End Function

'=====================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ݾ׻����۾�</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>> 
							<TR> 
                                 <TD CLASS="TD5" NOWRAP>����</TD>
                                 <TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="����" TYPE="Text" MAXLENGTH=10 SiZE=12  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconBp_cd.value,5">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
                                <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFr_dt" CLASS=FPDTYYYYMMDD tag="12X1X" ALT="��������" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;       
                                                       <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtTo_dt" CLASS=FPDTYYYYMMDD tag="12X1X" ALT="���������" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>       
							</TR>
                            <TR>
		                     <TD CLASS="TD5" NOWRAP>��&nbsp;&nbsp;��</TD>
		                     <TD CLASS="TD6"><INPUT NAME="txtPlantCode" TYPE="Text" ALT="����" MAXLENGTH=10 SiZE=12 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtPlantCode.value,4">&nbsp;<INPUT NAME="txtPlantName" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
							</TR>
                            <TR>
                             <TD CLASS="TD5" NOWRAP>ǰ&nbsp;&nbsp;��</TD>
                             <TD CLASS="TD6"><INPUT NAME="txtconItem_cd" ALT="ǰ��" TYPE="Text" MAXLENGTH=18 SiZE=12  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconItem_cd.value,0">&nbsp;<INPUT NAME="txtconItem_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
                            </TR>
						    <TR>
								<TD CLASS=TD5 NOWRAP>Volume��������</TD>
				                <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtVol_flag" TAG="11" VALUE="Y" CHECKED ID="txtVol_flag1"><LABEL FOR="txtVol_flag1">Y</LABEL>&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtVol_flag" TAG="11" VALUE="N" ID="txtVol_flag2"><LABEL FOR="txtVol_flag2">N</LABEL></TD>
						    </TR>
						    <TR>
								<TD CLASS=TD5 NOWRAP>�����йݿ�����</TD>
				                <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtQty_flag" TAG="11" VALUE="Y" CHECKED ID="txtQty_flag1"><LABEL FOR="txtQty_flag1">Y</LABEL>&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtQty_flag" TAG="11" VALUE="N" ID="txtQty_flag2"><LABEL FOR="txtQty_flag2">N</LABEL></TD>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>����</BUTTON>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>
