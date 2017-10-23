<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5101ma1
'*  4. Program Name         : ������ǥ���� 
'*  5. Program Desc         : ������ǥ�� ���Ͽ� ���� �Ǵ� ��������ϴ� ��� 
'*  6. Component List       : PAGG015.dll
'*  7. Modified date(First) : 2000/10/14
'*  8. Modified date(Last)  : 2003/10/31
'*  9. Modifier (First)     : Chang Goo,Kang
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="a5103ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBSCRIPT">

Option Explicit  

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate

iDBSYSDate = "<%=GetSvrDate%>"

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029      

	With frm1
		Call FormatDATEField(.txtFromReqDt)
		Call LockObjectField(.txtFromReqDt,"R")
		Call FormatDATEField(.txtToReqDt)
		Call LockObjectField(.txtToReqDt,"R")

'		Call LockHTMLField(.cboConfFg,"R")
		Call LockHTMLField(.txtDeptNm,"P")
		Call LockHTMLField(.txtGlInputTypeNm,"P")        

		Call FormatDATEField(.GIDate)    
		Call LockObjectField(.GIDate,"O")		


		Call InitSpreadSheet                                                    '��: Setup the Spread sheet
		Call InitComboBox
		Call SetDefaultVal
		Call InitVariables                                                      '��: Initializes local global variables
		Call SetToolbar("110000000000111")										'��: ��ư ���� ����    	

		.txtDeptCd.focus
		.btnConf.disabled  = True
		.btnUnCon.disabled = True
	End With		
	
	' ���Ѱ��� �߰� 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc)
	
	' ����� 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ� 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' ���� 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text

	Set xmlDoc = Nothing
	
End Sub

'========================================================================================================= 
Sub InitComboBox()

	Dim iData

	'iData = CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1007", "''", "S") & "  order by minor_cd ")	
	'Response.Write " Call SetCombo3(frm1.cboConfFg,""" & iData & """)" & vbCrLf
	Call CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1007", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConfFg, lgF0, lgF1, Chr(11))

End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### %>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
						<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopuptempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>  WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5"NOWRAP>��������</TD>
									<TD CLASS="TD6"NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtFromReqDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��������" class=required></OBJECT>');</SCRIPT> ~
 										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtToReqDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��������" class=required></OBJECT>');</SCRIPT>										
 									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14" class=protected readonly=true tabindex="-1">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>�μ��ڵ�</TD>
									<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDeptCd" SIZE=10  MAXLENGTH=10  tag="11XXXU" ALT="�μ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
														  <INPUT TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 tag="14X" class=protected readonly=true tabindex="-1"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd1.Value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14" class=protected readonly=true tabindex="-1"></TD>
								</TR>
									<TD CLASS="TD5"NOWRAP>���λ���</TD>
									<TD CLASS="TD6"NOWRAP><SELECT NAME="cboConfFg" tag="12" STYLE="WIDTH:82px:" Alt="���λ���" class=required></SELECT>
									<TD CLASS="TD5"NOWRAP>��ǥ�Է°��</TD>
									<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtGlInputType" SIZE=10  MAXLENGTH=10 tag="11XXXU" ALT="��ǥ�Է°��" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtGlInputType.Value, 1)">
										 <INPUT TYPE=TEXT ID="txtGlInputTypeNm" NAME="txtGlInputTypeNm" SIZE=20 tag="14X" ALT="��ǥ�Է°�θ�" class=protected readonly=true tabindex=-1>
									</TD>
								</TR>
								<TR>
										        
									<TD CLASS="TD5"NOWRAP>��ǥ����</TD>
									<TD CLASS="TD6"NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=GIDate name=GIDate CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="��ǥ����"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>���ǹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtTempGlNoFr" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="���۰��ǹ�ȣ" STYLE="TEXT-ALIGN:left" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNoFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtTempGlNoFr.Value,2)">&nbsp;~&nbsp;
														   <INPUT TYPE="Text" NAME="txtTempGlNoTo" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="������ǹ�ȣ" STYLE="TEXT-ALIGN:left" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNoTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtTempGlNoTo.Value,3)">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��ǥ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtGlNoFr" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="������ǥ��ȣ" STYLE="TEXT-ALIGN:left" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNoFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtGlNoFr.Value,4)">&nbsp;~&nbsp;
														   <INPUT TYPE="Text" NAME="txtGlNoTo" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="������ǥ��ȣ" STYLE="TEXT-ALIGN:left" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNoTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtGlNoTo.Value,5)">
									</TD>
									<TD CLASS="TD5" NOWRAP>������ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtRefNo" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="������ȣ" STYLE="TEXT-ALIGN:left" STYLE="text-transform:uppercase"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
				     <TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="60%">
								<TD  WIDTH="100%" colspan=4>
								    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" colspan="4">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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

	
<TR HEIGHT="20">
	<TD WIDTH="100%" >
  		<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
			<TD WIDTH=10>&nbsp</TD>
			<TD><BUTTON NAME="btnConf" CLASS="CLSSBTN" OnClick="VBScript:Call fnBttnConf()">�ϰ�����</BUTTON>&nbsp<BUTTON NAME="btnUnCon" CLASS="CLSSBTN" OnClick="VBScript:Call fnBttnUnConf()">�ϰ����</BUTTON></TD>
			<TD WIDTH=10>&nbsp</TD>
			</TR>
  		</TABLE> 
	</TD>
</TR>
<TR>
    <TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
    <!--<TD WIDTH="100%" HEIGHT=30%><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=100 TABINDEX="-1"></IFRAME></TD>-->
</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDeptCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBizCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hcboConfFg"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtWorkFg"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="hFromReqDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="��������" id=fpDateTime1 TABINDEX="-1"></OBJECT>');</SCRIPT>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="hToReqDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="��������" id=fpDateTime2 TABINDEX="-1"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=280 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP"   METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname"       TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname"      TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename"    TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar"     TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date">	
</Form>
</BODY>
</HTML>
