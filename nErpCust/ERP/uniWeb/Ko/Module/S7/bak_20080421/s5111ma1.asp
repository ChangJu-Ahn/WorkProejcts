<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ����
'*  2. Function Name        : ����ä�ǰ���
'*  3. Program ID           : S5111MA1
'*  4. Program Name         : ����ä�ǵ��
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G111.cSBillHdrSvr,PS3G102.cLookupSoHdrSvr,PB5CS41.cLookupBizPartnerSvr
'*							  PS4G119.cSLkLcHdrSvr,PB5CS41.cLookupBizPartnerSvr	
'*							  PS7G115.cSPostOpenArSvr
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : 3rd ȭ�� Layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� Layout
'*                            -2001/12/18 : Date ǥ������
'*                            -2002/11/15 : UI ǥ������
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="S5111ma1_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
Dim EndDate

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2KCM.inc" -->
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
     <TD CLASS="CLSSTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>T_����ä���Ϲ�</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD CLASS="CLSSTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
       <TR>
        <td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>T_ä�Ǳݾ�����</font></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD CLASS="CLSSTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
       <TR>
        <td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>T_��������</font></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right><A href="vbscript:OpenSORef">��������</A>&nbsp;|&nbsp;<A href="vbscript:OpenLCRef">L/C����</A>&nbsp;|&nbsp;<A href="vbscript:OpenDNRef">��������</A></TD>
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
         <TD CLASS="TD5" NOWRAP>����ä�ǹ�ȣ</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConBillNo" ALT="����ä�ǹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConBillNo()"></TD>
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
     <TD WIDTH=100% VALIGN=TOP>

      <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD CLASS="TD5" NOWRAP>����ä�ǹ�ȣ</TD>
        <TD CLASS="TD6"><INPUT NAME="txtBillNo" ALT="����ä�ǹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="25XXXU" STYLE="text-transform:uppercase"></TD>
        <TD CLASS=TD5 NOWRAP>����ä������</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillTypeCd" ALT="����ä������" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="23XXXU" STYLE="text-transform:uppercase" class=required ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 0">&nbsp;<INPUT NAME="txtBillTypeNm" TYPE="Text" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
        <TD CLASS=TD5 NOWRAP>����ä����</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtBillDt" CLASS="FPDTYYYYMMDD" tag="22X1" ALT="����ä����" Title="FPDATETIME"></OBJECT>');</SCRIPT>
           </TD>
          </TR>
         </TABLE>
        </TD>
        <TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoNo" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;
         <LABEL FOR="chkSoNo">���ֹ�ȣ����</LABEL><INPUT TYPE=CHECKBOX NAME="chkSoNo" tag="25X" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid">
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>����ó</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillToPartyCd" ALT="����ó" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 8">&nbsp;<INPUT NAME="txtBillToPartyNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldtoPartyCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" ALT="�ֹ�ó" class = protected readonly = true TABINDEX="-1" >&nbsp;<INPUT NAME="txtSoldtoPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>����ó</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayerCd" ALT="����ó" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 1">&nbsp;<INPUT NAME="txtPayerNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP></TD>
        <TD CLASS=TD6 NOWRAP></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>���ݿ����׷�</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtToBizAreaCd" ALT="���ݿ����׷�" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 2">&nbsp;<INPUT NAME="txtToBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>�����׷�</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>ȯ��</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchgRate" CLASS="FPDS100" ALT="ȯ��" tag="22X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
        </TD>
        <TD CLASS=TD5 NOWRAP>ȭ�����</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur1" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT�������</TD>
        <TD CLASS=TD6 NOWRAP>
         <INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVATCalcType" TAG="21" VALUE="1" CHECKED ID="rdoVATCalcType1">
         <LABEL FOR="rdoVATCalcType1">����</LABEL>
         <INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVATCalcType" TAG="21" VALUE="2" ID="rdoVATCalcType2">
         <LABEL FOR="rdoVATCalcType2">����</LABEL>
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT���Ա���</TD>
        <TD CLASS=TD6 NOWRAP>
         <INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVatIncFlag" TAG="21" VALUE="1" CHECKED ID="rdoVatIncFlag1">
         <LABEL FOR="rdoVatIncFlag1">����</LABEL>
         <INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVatIncFlag" TAG="21" VALUE="2" ID="rdoVatIncFlag2">
         <LABEL FOR="rdoVatIncFlag2">����</LABEL>
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT����</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <INPUT NAME="txtVatType" TYPE="Text" MAXLENGTH="5" SIZE=10 ALT="VAT����" tag="23XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 10">&nbsp;
           </TD>
           <TD>
            <INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="25" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1">
           </TD>
          </TR>
         </TABLE>
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT��</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtVatRate" CLASS="FPDS100" tag="24X5ZU"></OBJECT>');</SCRIPT>&nbsp;<LABEL><b>%</b></LABEL>
           </TD>
           
          </TR>
         </TABLE>
        </TD>
       </TR>         
       <TR>
        <TD CLASS=TD5 NOWRAP>�������</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermsCd" Alt="�������" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 5">&nbsp;<INPUT NAME="txtPayTermsNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"  class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>�����Ⱓ</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle8 NAME="txtPayDur" CLASS="FPDS100" Alt="�����Ⱓ" tag="21X6Z" Title="FPDOUBLESINGLE"> </OBJECT>');</SCRIPT>&nbsp;<LABEL>��.</LABEL>
        </TD>
       </TR>       
       <TR>
        <TD CLASS=TD5 NOWRAP>�Ա�����</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTypeCd" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="�Ա�����" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 4">&nbsp;<INPUT NAME="txtPayTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"  class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>���ݸ�����</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtPlanIncomeDt" CLASS="FPDTYYYYMMDD" tag="21X1" ALT="���ݸ�����" Title="FPDATETIME"></OBJECT>');</SCRIPT>
           </TD>
          </TR>
         </TABLE>
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>���ݽŰ�����</LABEL></TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaCd" ALT="���ݽŰ�����" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 3">&nbsp;<INPUT NAME="txtTaxBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP></TD>
        <TD CLASS=TD6 NOWRAP></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>��ݰ�����������</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaytermsTxt" ALT="��ݰ�������" TYPE="Text" MAXLENGTH="120" SIZE=39 tag="21"></TD>
        <TD CLASS=TD5 NOWRAP>���</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemark" ALT="���" TYPE="Text" MAXLENGTH="120" SIZE=39 tag="21"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>Ȯ������</TD>
        <TD CLASS=TD6 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoPostFlag" id="rdoPostFlagY" value="Y" tag = "24">
          <label for="rdoPostFlagY">Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoPostFlag" id="rdoPostFlagN" value="N" tag = "24" checked>
          <label for="rdoPostFlagN">��Ȯ��</label>
        </TD>
        <TD CLASS=TD5 NOWRAP>��ǥ��ȣ</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctNo" ALT="��ǥ��ȣ" TYPE="Text" MAXLENGTH="18" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD HEIGHT=20 WIDTH=100% CLASS=TD6 COLSPAN=4>
         <FIELDSET ID="filTaxNo" CLASS="CLSFLD" TITLE="���ݰ�꼭�ڵ�����">
         <LEGEND ALIGN=LEFT><INPUT TYPE=CHECKBOX NAME="chkTaxNo" tag="21" Class="Check"><LABEL FOR="chkTaxNo">���ݰ�꼭�ڵ����࿩��</LABEL></LEGEND>
          <TABLE <%=LR_SPACE_TYPE_40%>>
           <TR>
            <TD CLASS=TD5 NOWRAP><LABEL ID="lblTaxBillNo">���ݰ�꼭��ȣ</LABEL></TD>
            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBillNo" ALT="���ݰ�꼭��ȣ" TYPE="Text" MAXLENGTH="30" SIZE=30 tag="24XXXU" class = protected readonly = true TABINDEX="-1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillTaxNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTaxNo()"></TD>
            <TD CLASS=TD6 NOWRAP></TD>
            <TD CLASS=TD6 NOWRAP></TD>
           </TR>
          </TABLE>
         </FIELDSET>
        </TD>
       </TR>
      </TABLE>
      </DIV>

      <DIV ID="TabDiv" SCROLL=no>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD CLASS=TD5 NOWRAP>����ä�Ǳݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
          <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtBillAmt" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
           </TD>
           <TD>
            &nbsp;<INPUT NAME="txtDocCur2" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU"  class = protected readonly = true TABINDEX="-1">
           </TD>
          </TR>
          </TABLE>
        </TD>
        <TD CLASS=TD5 NOWRAP>����ä���ڱ��ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtBillAmtLoc" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
           </TD>
           <TD>
            &nbsp;<INPUT NAME="txtLocCur" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU"  class = protected readonly = true TABINDEX="-1">
           </TD>
          </TR>
         </TABLE>
        </TD>
       </TR>  
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT�ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtVatAmt" ALT="VAT�ݾ�" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT�ڱ��ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtVatLocAmt" ALT="VAT�ڱ��ݾ�" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
       </TR>        
       <TR>
        <TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtDepositAmt" Alt="�����ݾ�" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
        <TD CLASS=TD5 NOWRAP>�������ڱ���</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtDepositAmtLoc" Alt="�������ڱ���" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>        
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>�Ѹ���ä�Ǳݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtTotBillAmt" Alt="�Ѹ���ä�Ǳݾ�" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
        <TD CLASS=TD5 NOWRAP>�Ѹ���ä���ڱ��ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtTotBillAmtLoc" Alt="�Ѹ���ä���ڱ��ݾ�" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>        
       </TR>   
       <TR>
        <TD CLASS=TD5 NOWRAP>�Ѽ��ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtIncomeAmt" Alt="�Ѽ��ݾ�" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
        <TD CLASS=TD5 NOWRAP>�Ѽ����ڱ���</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtIncomeLocAmt" Alt="�Ѽ����ڱ���" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>        
       </TR>   
       <% Call SubFillRemBodyTD5656(10) %>
      </TABLE>
      </DIV>

      <DIV ID="TabDiv" SCROLL=no>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD CLASS=TD5 NOWRAP>�����</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBeneficiaryCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" Alt="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 6">&nbsp;<INPUT NAME="txtBeneficiaryNm" TYPE="Text" MAXLENGTH="20" SIZE=23 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>�絵��</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtApplicantCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" Alt="�絵��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 7">&nbsp;<INPUT NAME="txtApplicantNm" TYPE="Text" MAXLENGTH="20" SIZE=23 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>L/C������ȣ</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>Amend����</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCAmendSeq" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>L/C��ȣ</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE="Text" MAXLENGTH="35" SIZE=35 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>���FOB�ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtAcceptFobAmt" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
       </TR>
       <% Call SubFillRemBodyTD5656(12) %>
      </TABLE>
      </DIV>
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
     <TD><BUTTON NAME="btnPostFlag" CLASS="CLSMBTN">Ȯ��</BUTTON>&nbsp;
      <BUTTON NAME="btnGLView" CLASS="CLSMBTN">��ǥ��ȸ</BUTTON>&nbsp;
      <BUTTON NAME="btnPreRcptView" CLASS="CLSMBTN">��������Ȳ</BUTTON></TD>
     <TD WIDTH=* Align=Right><a href = "VBSCRIPT:JumpChgCheck(BIZ_BillDtl_JUMP_ID)">����ä�ǳ������</a>&nbsp;|&nbsp;<a href = "vbscript:JumpChgCheck(BIZ_BillCollect_JUMP_ID)">����ä�Ǽ��ݳ������</a></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtBillCommand" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtHBillNo" tag="24" TABINDEX ="-1">

<INPUT TYPE=HIDDEN NAME="txtRefFlag" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtChkSoNo" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtchkTaxNo" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtSupplyFlag" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtSts" tag="24" TABINDEX ="-1">

<INPUT TYPE=HIDDEN NAME="txtVatCalcType" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtVatIncFlag" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtRetItemFlag" tag="24" TABINDEX ="-1">

<INPUT TYPE=HIDDEN NAME="txtGLNo" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtTempGLNo" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtBatchNo" tag="24"TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtCreditRotDay" tag="24" TABINDEX ="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX ="-1"></iframe>
  </DIV>
</BODY>
</HTML>
