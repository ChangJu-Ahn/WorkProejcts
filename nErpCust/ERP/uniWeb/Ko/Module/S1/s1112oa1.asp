<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1112OA1
'*  4. Program Name         : ����ǰ��ܰ���� 
'*  5. Program Desc         : ����ǰ��ܰ���� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/01/15
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 2002/12/16 Include ������� ���ر� 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim IsOpenPop          

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
	
End Sub

'========================================================================================================= 
Sub SetDefaultVal()

	frm1.txtBp_Cd.focus 
	frm1.txtValid_From_Dt.Text = EndDate

End Sub

'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","S","NOCOOKIE","OA") %>
	<% Call LoadBNumericFormatA("Q","S","NOCOOKIE","OA") %>
End Sub

'--------------------------------------------------------------------------------------------------------- 
Function OpenConPop1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��"						<%' �˾� ��Ī %>
	arrParam(1) = "B_BIZ_PARTNER"		            <%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtBP_Cd.value)		    <%' Code Condition%>
	arrParam(3) = ""                            	<%' Name Cindition%>
	arrParam(4) = "BP_TYPE IN(" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"	        <%' Where Condition%>
	arrParam(5) = "��"						<%' TextBox ��Ī %>
	
	arrField(0) = "BP_CD"					        <%' Field��(0)%>
	arrField(1) = "BP_NM"					        <%' Field��(1)%>
    
	arrHeader(0) = "��"						<%' Header��(0)%>
	arrHeader(1) = "����"						<%' Header��(1)%>

	frm1.txtBP_Cd.focus 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop1(arrRet)
	End If

End Function

'--------------------------------------------------------------------------------------------------------- 
Function OpenConPop2()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��"						<%' �˾� ��Ī %>
	arrParam(1) = "b_item"		                    <%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtItem_Cd.value)		<%' Code Condition%>
	arrParam(3) = ""		                        <%' Name Cindition%>
	arrParam(4) = ""	                            <%' Where Condition%>
	arrParam(5) = "ǰ��"						<%' TextBox ��Ī %>
	
	arrField(0) = "item_cd"					        <%' Field��(0)%>
	arrField(1) = "item_nm"					        <%' Field��(1)%>
    
	arrHeader(0) = "ǰ��"						<%' Header��(0)%>
	arrHeader(1) = "ǰ���"						<%' Header��(1)%>

	frm1.txtItem_Cd.focus 
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			
	IsOpenPop = False
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop2(arrRet)
	End If

End Function

'--------------------------------------------------------------------------------------------------------- 
Function SetConPop1(Byval arrRet)
	With frm1	
		.txtBP_Cd.Value		= arrRet(0)
		.txtBP_Nm.Value		= arrRet(1)
	End With
End Function

'--------------------------------------------------------------------------------------------------------- 
Function SetConPop2(Byval arrRet)
	With frm1	
		.txtItem_Cd.Value		= arrRet(0)
		.txtItem_Nm.Value		= arrRet(1)
	End With
End Function

'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables														'��: Initializes local global variables
    <% '----------  Coding part  -------------------------------------------------------------%>
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'��: ��ư ���� ���� 

End Sub

'========================================================================================================= 
Sub txtValid_From_Dt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValid_From_Dt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtValid_From_Dt.Focus
    End If
End Sub

'========================================================================================
 Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
 Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function

'========================================================================================
Function FncQuery() 
	FncQuery = true   
End Function

'========================================================================================
Function BtnPrint() 
	Dim strUrl
	
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If

    <%'--��������� �����ϴ� �κ� ���� %>
	dim var1, var2, var3
	
	var1 = UniConvDateToYYYYMMDD(frm1.txtValid_From_Dt.Text,parent.gDateFormat,parent.gServerDateType)

	
	If UCase(frm1.txtBP_Cd.value) = "" Then
		var2 = "%"	
	Else
		var2 = FilterVar(Trim(UCase(frm1.txtBP_Cd.value)), "" ,  "SNM") 
	End If


	If UCase(frm1.txtItem_Cd.value) = "" Then
		var3 = "%"	
	Else
		var3 = FilterVar(Trim(UCase(frm1.txtItem_Cd.value)), "" ,  "SNM")
	End If

	strUrl = strUrl & "Valid_From_Dt|" & var1
	strUrl = strUrl & "|BP_cd|" & var2
	strUrl = strUrl & "|Item_cd|" & var3
			
'----------------------------------------------------------------
' Print �Լ����� ȣ�� 
'----------------------------------------------------------------
OBjName = AskEBDocumentName("s1112oa1","ebr")    
Call FncEBRprint(EBAction, OBjName, strUrl)
'----------------------------------------------------------------
		
End Function

'========================================================================================
Function BtnPreview() 
    
	
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If

	
	Dim var1, var2, var3
	Dim strUrl
	Dim arrParam, arrField, arrHeader
		
<%
'	Ư�����ڸ� �Ѱ��ٶ��� �ƽ�Ű �ڵ尪���� ��ȯ�� ���־�� �Ѵٴ±��� 
'	"%" ---> %
'	""  ---> %32 �� �ٲپ� �ּž� �մϴ�.
'	�ƽ�Ű�ڵ� 25�� %�̰� 32�� space�Դϴ�.
'	SQL 7.0������ ""�� " "�� ���� �ν��ϴ����� 
%>

		
	 	var1 = UniConvDateToYYYYMMDD(frm1.txtValid_From_Dt.Text,parent.gDateFormat,parent.gServerDateType)
	
	If UCase(frm1.txtBP_Cd.value) = "" Then
		var2 = "%"	
	Else
		var2 = FilterVar(Trim(UCase(frm1.txtBP_Cd.value)), "" ,  "SNM") 
	End If


	If UCase(frm1.txtItem_Cd.value) = "" Then
		var3 = "%"	
	Else
		var3 = FilterVar(Trim(UCase(frm1.txtItem_Cd.value)), "" ,  "SNM")
	End If


	strUrl = strUrl & "Valid_From_Dt|" & var1
	strUrl = strUrl & "|BP_cd|" & var2
	strUrl = strUrl & "|Item_cd|" & var3
	

	OBjName = AskEBDocumentName("S1112oa1","ebr")    
	Call FncEBRPreview(OBjName, strUrl)		
		
End Function

'========================================================================================
Function FncExit()
	
	FncExit = True
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ǰ��ܰ�</font></td>
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
									<TD CLASS="TD5" NOWRAP>���������</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s1112oa1_fpDateTime1_txtValid_From_Dt.js'></script>
												</TD>
												
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBP_Cd" ALT="��" TYPE="Text" MAXLENGTH="10" SIZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop1">&nbsp;<INPUT NAME="txtBP_Nm" TYPE="Text" MAXLENGTH="50" SIZE=30  tag="14"></TD>								
							    </TR>
							    <TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem_Cd" ALT="ǰ��" TYPE="Text" MAXLENGTH="18" SIZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnITEM_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop2">&nbsp;<INPUT NAME="txtItem_Nm" TYPE="Text" MAXLENGTH="40"  SIZE=30  tag="14"></TD>								
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
					<TD>
					    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
					    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON>
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
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX="-1">
    <input type="hidden" name="dbname" TABINDEX="-1">
    <input type="hidden" name="filename" TABINDEX="-1">
    <input type="hidden" name="condvar" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>
