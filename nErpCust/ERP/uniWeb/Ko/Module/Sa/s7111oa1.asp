<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : ���� Nego ���� 
'*  3. Program ID           : S7111OA1
'*  4. Program Name         : ���� Nego ���� 
'*  5. Program Desc         : ���� Nego ���� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/07/18
'*  8. Modified date(Last)  : 2000/07/18
'*  9. Modifier (First)     : Cho Sung Hyun
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 2002/12/17 Include ������� ���ر� 
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
	frm1.ConApplicant.focus 
	frm1.NegoFromDt.text = StartDate
	frm1.NegoToDt.text = EndDate
End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'========================================================================================================= 
Function OpenConPop()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "������"						<%' �˾� ��Ī %>
	arrParam(1) = "B_BIZ_PARTNER"					<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.ConApplicant.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "BP_TYPE IN (" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"         	<%' Where Condition%>
	arrParam(5) = "������"						<%' TextBox ��Ī %>
			
	arrField(0) = "BP_CD"							<%' Field��(0)%>
	arrField(1) = "BP_NM"							<%' Field��(1)%>
		    
	arrHeader(0) = "������"						<%' Header��(0)%>
	arrHeader(1) = "�����ڸ�"					<%' Header��(1)%>

	frm1.ConApplicant.focus
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop(arrRet)
	End If

End Function

'========================================================================================================= 
Function SetConPop(Byval arrRet)
	With frm1	
		.ConApplicant.Value		= arrRet(0)
		.ConApplicantNm.Value		= arrRet(1)
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
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================= 
Sub NegoFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.NegoFromDt.Action = 7
		Call SetFocusToDocument("P")	
		Frm1.NegoFromDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

Sub NegoToDt_DblClick(Button)
    If Button = 1 Then
        frm1.NegoToDt.Action = 7
		Call SetFocusToDocument("P")	
		Frm1.NegoToDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Function FncPrint() 
	Call parent.FncPrint()
End Function
'========================================================================================================= 
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function
'========================================================================================================= 
Function FncQuery() 
    FncQuery = true
End Function

'========================================================================================================= 
Function BtnPrint() 
	Dim strUrl
	
    '** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� ũ�ų� ���ƾ� �Ҷ� **
	If ValidDateCheck(frm1.NegoFromDt, frm1.NegoToDt) = False Then Exit Function

    
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If

    <%'--��������� �����ϴ� �κ� ���� %>
	dim var1, var2 ,var3
	
	
	If UCase(frm1.ConApplicant.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.ConApplicant.value)), "" ,  "SNM")
	End If

 	
		var2 = UniConvDateToYYYYMMDD(frm1.NegoFromDt.text,parent.gDateFormat,parent.gServerDateType)
		var3 = UniConvDateToYYYYMMDD(frm1.NegoToDt.text,parent.gDateFormat,parent.gServerDateType)

	
	<%'--��������� �����ϴ� �κ� ���� - �� %>
	
'    On Error Resume Next                                                    '��: Protect system from crashing
    

    <%'--��������� �����ϴ� �κ� ���� %>
	strUrl = strUrl & "ConApplicant|" & var1 & "|NegoFromDt|" & var2 & "|NegoToDt|" & var3 



		'----------------------------------------------------------------
		' Print �Լ����� ȣ�� 
		'----------------------------------------------------------------
			OBjName = AskEBDocumentName("s7111oa1","ebr")    
			Call FncEBRprint(EBAction, OBjName, strUrl)
		'----------------------------------------------------------------
End Function

'========================================================================================================= 
Function BtnPreview()  
    

	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� ũ�ų� ���ƾ� �Ҷ� **
	If ValidDateCheck(frm1.NegoFromDt, frm1.NegoToDt) = False Then Exit Function

    
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

	If UCase(frm1.ConApplicant.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.ConApplicant.value)), "" ,  "SNM")
	End If

	 	var2 = UniConvDateToYYYYMMDD(frm1.NegoFromDt.text,parent.gDateFormat,parent.gServerDateType)
		var3 = UniConvDateToYYYYMMDD(frm1.NegoToDt.text,parent.gDateFormat,parent.gServerDateType)
	
		
		
		strUrl = strUrl & "ConApplicant|" & var1 & "|NegoFromDt|" & var2 & "|NegoToDt|" & var3 
	
		OBjName = AskEBDocumentName("s7111oa1","ebr")    
		Call FncEBRPreview(OBjName, strUrl)		
		
End Function
'========================================================================================================= 
Function FncExit()
	FncExit = True
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���� Nego ����</font></td>
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
									<TD CLASS=TD5>������</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="ConApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPop" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPop">&nbsp;<INPUT TYPE=TEXT NAME="ConApplicantNm" SIZE=25 TAG="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>NEGO �Ⱓ</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s7111oa1_fpDateTime1_NegoFromDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s7111oa1_fpDateTime2_NegoToDt.js'></script>
									</TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> SRC= "../../blank.htm" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX ="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX ="-1">
    <input type="hidden" name="dbname" TABINDEX ="-1">
    <input type="hidden" name="filename" TABINDEX ="-1">
    <input type="hidden" name="condvar" TABINDEX ="-1">
	<input type="hidden" name="date" TABINDEX ="-1">
</FORM>
</BODY>
</HTML>
