<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : �����Ȳ ��� 
'*  3. Program ID           : s4111oa1
'*  4. Program Name         : �����Ȳ ��� 
'*  5. Program Desc         : �����Ȳ ��� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/18
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : �չ��� 
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'*                            2003/06/11 ǥ�عݿ� 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Dim IsOpenPop          
Dim EndDate

' �ý��� ��¥ 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'==========================================
Sub InitVariables()
End Sub

'=========================================
Sub SetDefaultVal()
    frm1.txtReqdlvyFromDt.focus
	frm1.txtReqdlvyFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtReqdlvyToDt.Text = EndDate
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'==========================================
Function OpenConPop1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"					        
	arrParam(1) = "b_minor A, I_MOVETYPE_CONFIGURATION B"	
	arrParam(2) = Trim(frm1.txtDN_TYPE.value)		        
	'arrParam(3) = Trim(frm1.txtDn_TypeNm.value)		    
	arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & " "	
	arrParam(5) = "��������"					         

	arrField(0) = "A.MINOR_CD"							     
	arrField(1) = "A.MINOR_NM"							      

	arrHeader(0) = "��������"					         
	arrHeader(1) = "�������¸�"					          

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtDN_TYPE.focus 

	If arrRet(0) <> "" Then
		frm1.txtDN_TYPE.Value	= arrRet(0)
		frm1.txtDN_TYPE_NM.Value= arrRet(1)
	End If

End Function

'==========================================
Function OpenConPop2()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��ǰó"							
	arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"		
	arrParam(2) = Trim(frm1.txtShipToParty.value)		
	arrParam(3) = Trim(frm1.txtShipToPartyNm.value)		
	arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SSH", "''", "S") & " " _
					& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE IN (" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"
	arrParam(5) = "��ǰó"							
		
	arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"			
	arrField(1) = "PARTNER.BP_NM"						
	arrField(2) = "PARTNER_FTN.BP_CD"					
	arrField(3) = "PARTNER_FTN.PARTNER_FTN"				
	arrField(4) = "PARTNER_FTN.USAGE_FLAG"				
	    
	arrHeader(0) = "��ǰó"							
	arrHeader(1) = "��ǰó��"						
	arrHeader(2) = "�ŷ�ó�ڵ�"						
	arrHeader(3) = "�ŷ�óŸ��"						
	arrHeader(4) = "��뿩��"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtShipToParty.focus 

	If arrRet(0) <> "" Then
		frm1.txtShipToParty.Value	= arrRet(0)
		frm1.txtShipToPartyNm.Value	= arrRet(1)
	End If

End Function

'=========================================
Sub Form_Load()

    Call LoadInfTB19029														'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables
    Call GetValue_ko441()														
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'��: ��ư ���� ���� 
End Sub

'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================
Sub txtReqdlvyFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqdlvyFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtReqdlvyFromDt.Focus
    End If
End Sub

'=========================================
Sub txtReqdlvyToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqdlvyToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtReqdlvyToDt.Focus
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

'========================================
Function BtnPreview_OnClick()
	Call BtnPrint("N")
End Function

'========================================
Function BtnPrint_OnClick()
	Call BtnPrint("Y")
End Function

'=====================================================
Function BtnPrint(ByVal pvStrPrint) 

Dim vargBizArea,vargPlant,vargSalesGrp,vargSalesOrg 

   If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(frm1.txtReqdlvyFromDt, frm1.txtReqdlvyToDt) = False Then Exit Function

	Dim iStrUrl
    
    ' ������� 
	If Trim(frm1.txtDN_TYPE.value) = "" Then
		iStrUrl = "DN_TYPE|%"
	Else
		iStrUrl = "DN_TYPE|" & Replace(UCase(Trim(frm1.txtDN_TYPE.value)), "'" ,  "''")
	End If
    
	If Trim(frm1.txtShipToParty.value) = "" Then
		iStrUrl = iStrUrl & "|ShipToParty|%"
	Else
		iStrUrl = iStrUrl & "|ShipToParty|" & Replace(UCase(Trim(frm1.txtShipToParty.value)), "'" ,  "''")
	End If
	If lgBACd <> "" Then
		vargBizArea = " AND S_DN_HDR.BIZ_AREA =  " & FilterVar(Trim(UCase(lgBACd)), "" ,  "S")
	Else
		vargBizArea = ""
	End If
	If lgPLCd <> "" Then
		vargPlant = " AND S_DN_DTL.PLANT_CD = " & FilterVar(Trim(UCase(lgPLCd)), "" ,  "S")
	Else
		vargPlant = ""
	End If
	If lgSGCd <> "" Then
		vargSalesGrp = " AND S_DN_HDR.SALES_GRP = " & FilterVar(Trim(UCase(lgSGCd)), "" ,  "S")
	Else
		vargSalesGrp = ""
	End If
	If lgSOCd <> "" Then
		vargSalesOrg = " AND S_DN_HDR.SALES_ORG = " & FilterVar(Trim(UCase(lgSOCd)), "" ,  "S")
	Else
		vargSalesOrg = ""
	End If

	iStrUrl = iStrUrl & "|ReqdlvyFromDt|" & UniConvDateToYYYYMMDD(frm1.txtReqdlvyFromDt.Text,parent.gDateFormat,parent.gServerDateType)
	iStrUrl = iStrUrl & "|ReqdlvyToDt|" & UniConvDateToYYYYMMDD(frm1.txtReqdlvyToDt.Text,parent.gDateFormat,parent.gServerDateType)
	istrUrl = istrUrl & "|gBizArea|" & vargBizArea 
	istrUrl = istrUrl & "|gPlant|" & vargPlant 
	istrUrl = istrUrl & "|gSalesGrp|" & vargSalesGrp
	istrUrl = istrUrl & "|gSalesOrg|" & vargSalesOrg 
	
	OBjName = AskEBDocumentName("s4111oa1_ko441","ebr")    

	If pvStrPrint = "N" Then
		' �̸����� 
		Call FncEBRPreview(ObjName, iStrUrl)
	Else
		' ��� 
		Call FncEBRprint(EBAction, ObjName, iStrUrl)
	End If
		
End Function
'=====================================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����Ȳ</font></td>
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
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s4111oa1_fpDateTime1_txtReqdlvyFromDt.js'></script>
											</TD>
											<TD>
												&nbsp;~&nbsp;
											</TD>
											<TD>
												<script language =javascript src='./js/s4111oa1_fpDateTime2_txtReqdlvyToDt.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDN_TYPE" ALT="��������" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDN_TYPE" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop1()">&nbsp;<INPUT NAME="txtDN_TYPE_NM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ǰó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShipToParty" ALT="��ǰó" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShipToParty" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop2()">&nbsp;<INPUT NAME="txtShipToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
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
					<TD valign=top>
						<BUTTON NAME="BtnPreview" CLASS="CLSSBTN" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="BtnPrint" CLASS="CLSSBTN" Flag=1>�μ�</BUTTON>
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
