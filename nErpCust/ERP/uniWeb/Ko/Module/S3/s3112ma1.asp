<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ְ��� 
'*  3. Program ID           : S3112MA1
'*  4. Program Name         : ���ֳ������ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/11/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho in kuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="s3112ma1.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															

<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

</SCRIPT>


<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
 <TR >
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ֳ���</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right><A href="vbscript:OpenBOMRef">BOM����</A>&nbsp;|&nbsp;<A href="vbscript:OpenStockDtlRef">�����Ȳ����</A>&nbsp;|&nbsp;<A href="vbscript:OpenSoDtlRef">���ֳ�������</A>&nbsp;|&nbsp;<A href="vbscript:OpenStyleRef">Ŭ��������</A>&nbsp;<A ID="txtOpenPrjRef" STYLE="DISPLAY: none" href="vbscript:OpenPrjRef">|&nbsp;������Ʈ��������</A></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR HEIGHT=*>
  <TD WIDTH=100% CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>>
    <TR>
     <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD HEIGHT=20 WIDTH=100%>
      <FIELDSET CLASS="CLSFLD">
       <TABLE <%=LR_SPACE_TYPE_40%>>
        <TR>
         <TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConSoNo" ALT="���ֹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSoDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSoDtl()"></TD>
         <TD CLASS="TDT"></TD>
         <TD CLASS="TD6"></TD>
        </TR>
       </TABLE>
      </FIELDSET>
     </TD>
    </TR>
    <TR>
     <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD CLASS="TD5" NOWRAP>�ֹ�ó</TD>
        <TD CLASS="TD6"><INPUT NAME="txtSoldToParty" ALT="�ֹ�ó" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="24XXXU" class = protected readonly = true>&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true></TD>
        <TD CLASS="TD5" NOWRAP>���ֹ���ȣ</TD>
        <TD CLASS="TD6"><INPUT NAME="txtCustPoNo" ALT="���ֹ���ȣ" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24XXXU" class = protected readonly = true></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>���ּ��ݾ�</TD>
        <TD CLASS="TD6">
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtNetAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
           </TD>
           <TD>
            &nbsp;<INPUT NAME="txtCurrency" ALT="" TYPE="Text" MAXLENGTH=3 SiZE=4 tag="24XXXU" class = protected readonly = true>
           </TD>
          </TR>
         </TABLE>
        </TD>
        <TD CLASS="TD5" NOWRAP>�ΰ�������</TD>
        <TD CLASS="TD6"><INPUT NAME="txtVatIncFlag" ALT="�ΰ�������" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="24XXXU" class = protected readonly = true>&nbsp;<INPUT NAME="txtVatIncFlagNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>����</TD>
        <TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="����" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true></TD>
        <TD CLASS="TD5" NOWRAP></TD>
        <TD CLASS="TD6" NOWRAP></TD>
       </TR>
       <TR>
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
        </TD>
       </TR>
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR HEIGHT=20>
  <TD WIDTH=100%>
   <TABLE <%=LR_SPACE_TYPE_30%>>
       <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD>
      <BUTTON NAME="btnConfirm" CLASS="CLSSBTN">Ȯ��ó��</BUTTON>&nbsp;
      <BUTTON NAME="btnATPCheck" CLASS="CLSSBTN">ATP Check</BUTTON>&nbsp;
      <BUTTON NAME="btnCTPCheck" CLASS="CLSSBTN">CTP Check</BUTTON>&nbsp;
      <BUTTON NAME="btnDNCheck" CLASS="CLSSBTN">���Ͽ�ûó��</BUTTON>&nbsp;
      <BUTTON NAME="btnAvlStkRef" CLASS="CLSSBTN">���������Ȳ</BUTTON>
     </TD>
     <TD WIDTH=* Align=right><A HREF = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_SOHDR_ID)">���ֵ��</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_SOSCHE_ID)">ȸ�䳳����ȸ</A></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%>  FRAMEBORDER=0 SCROLLING=no  noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtConfirm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="RdoConfirm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConfirmFlg" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtShipToParty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSoDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSoType" tag="24" TABINDEX="-1">  
<INPUT TYPE=HIDDEN NAME="txtHNetAmt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVATAmt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN  NAME="txtHVATType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HReqDlvyDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HPriceFlag" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HExportFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HCiFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HRetItemFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHPreSONo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSoNo" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVatRate" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHVATIncFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVATIncFlagNm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHMaintNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPayTermsCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHDealType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtCtpCDFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHTrackingNORule" tag="14" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHDnReqFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="RdoDnReq" tag="24" TABINDEX="-1">

</FORM>
 <DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  