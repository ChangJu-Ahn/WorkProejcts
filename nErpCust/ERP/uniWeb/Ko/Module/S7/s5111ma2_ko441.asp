<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5111MA2
'*  4. Program Name         : ���ܸ���ä�ǵ�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G111.cSBillHdrSvr,PS3G102.cLookupSoHdrSvr,PB5CS41.cLookupBizPartnerSvr
'*							  PS4G119.cSLkLcHdrSvr,PB5CS41.cLookupBizPartnerSvr	
'*							  PS7G115.cSPostOpenArSvr
'*  7. Modified date(First) : 2002/11/15
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : Ahn Tae Hee
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : 3rd ȭ�� Layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� Layout
'*                            -2001/12/18 : Date ǥ������ 
'*                            -2001/12/26 : VAT �������� �߰� 
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
<SCRIPT LANGUAGE="VBScript"   SRC="S5111ma2_ko441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
' User-defind Variables
'========================================
Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
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
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
       <TR>
        <td background="../../../CShared/../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ܸ���ä��</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD CLASS="CLSMTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
       <TR>
        <td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ä�Ǳݾ�����</font></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right><A href="vbscript:OpenBillRef">��������ä������</A></TD>
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
         <TD CLASS=TD5 NOWRAP>����ä�ǹ�ȣ</TD>
         <TD CLASS=TD6 NOWRAP><INPUT NAME="txtConBillNo" ALT="����ä�ǹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConBillNo()"></TD>
         <TD CLASS=TDT NOWRAP></TD>
         <TD CLASS=TD6 NOWRAP></TD>
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
        <TD CLASS=TD5 NOWRAP>����ä�ǹ�ȣ</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillNo" ALT="����ä�ǹ�ȣ" TYPE="Text" MAXLENGTH="18" SIZE=30 tag="25XXXU" STYLE="text-transform:uppercase"></TD>
        <TD CLASS=TD5 NOWRAP>��������ä�ǹ�ȣ</TD>
        <TD CLASS=TD6 NOWRAP>
         <INPUT TYPE=TEXT NAME="txtRefBillNo" SIZE=20  MAXLENGTH=18 TAG="24XXXU" ALT="��������ä�ǹ�ȣ" class = protected readonly = true TABINDEX="-1">
         <LABEL FOR="chkRefBillNoFlg">����ä�ǹ�ȣ����</LABEL>
         <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="25X" VALUE="Y" NAME="chkRefBillNoFlg" ID="chkRefBillNoFlg">
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>����ä����</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDateTime1_txtBillDt.js'></script>
        </TD>
        <TD CLASS=TD5 NOWRAP>����ä������</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillTypeCd" ALT="����ä������" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="23XXXU" STYLE="text-transform:uppercase" class=required ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 9">&nbsp;<INPUT NAME="txtBillTypeNm" TYPE="Text" SIZE=24.5 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldtoPartyCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU" ALT="�ֹ�ó" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 11">&nbsp;<INPUT NAME="txtSoldtoPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>      
        <TD CLASS=TD5 NOWRAP>����ó</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillToPartyCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU" ALT="����ó" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 6" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtBillToPartyNm" TYPE="Text" SIZE=24.5 tag="24" class = protected readonly = true TABINDEX="-1"></TD>        
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>����ó</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayerCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU" ALT="����ó" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 1">&nbsp;<INPUT NAME="txtPayerNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP></TD>
        <TD CLASS=TD6 NOWRAP></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>�����׷�</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd" ALT="�����׷�" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="23XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 8">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>���ݿ����׷�</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtToBizAreaCd" ALT="���ݿ����׷�" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="22XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 2">&nbsp;<INPUT NAME="txtToBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=24.5 tag="24" class = protected readonly = true TABINDEX="-1"></TD>     
       </TR>             
       <TR>
        <TD CLASS=TD5 NOWRAP>ȭ��</TD>
        <TD CLASS=TD6 NOWRAP>
         <INPUT NAME="txtDocCur1" ALT="ȭ��" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="23XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 10"></TD>
        </TD>
        <TD CLASS=TD5 NOWRAP>ȯ��</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle3_txtXchgRate.js'></script>        
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT�������</TD>
        <TD CLASS=TD6 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoVatCalcType" id="rdoVatCalcType1" value="1" tag = "21" checked>
          <label ID="lblVatCalcType1" for="rdoVatCalcType1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoVatCalcType" id="rdoVatCalcType2" value="2" tag = "21" >
          <label ID="lblVatCalcType2" for="rdoVatCalcType2">����</label>
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT���Ա���</TD>
        <TD CLASS=TD6 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoVatIncFlag" id="rdoVatIncFlag1" value="1" tag = "21" checked>
          <label ID="lblVatIncFlag1" for="rdoVatIncFlag1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoVatIncFlag" id="rdoVatIncFlag2" value="2" tag = "21">
          <label ID="lblVatIncFlag2" for="rdoVatIncFlag2">����</label>
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT����</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <INPUT NAME="txtVatType" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="23XXXU" ALT="VAT����" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 7">&nbsp;
           </TD>
           <TD>
            <INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="25" SIZE=25 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;&nbsp;&nbsp;
           </TD>

          </TR>
         </TABLE>
        </TD>
           <TD CLASS=TD5 NOWRAP>VAT��</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>       
           <TD>
            <script language =javascript src='./js/s5111ma2_fpDoubleSingle5_txtVatRate.js'></script>&nbsp;<LABEL><b>%</b></LABEL>
           </TD>
          </TR>
         </TABLE>
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>�������</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermsCd" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="23XXXU" ALT="�������" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 5">&nbsp;<INPUT NAME="txtPayTermsNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>�����Ⱓ</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle5_txtPayDur.js'></script>&nbsp;<LABEL>��</LABEL>
        </TD>   
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>�Ա�����</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTypeCd" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="�Ա�����" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 4">&nbsp;<INPUT NAME="txtPayTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>        
        <TD CLASS=TD5 NOWRAP>���ݸ�����</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDateTime2_txtPlanIncomeDt.js'></script>
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaCd" ALT="���ݽŰ�����" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 3">&nbsp;<INPUT NAME="txtTaxBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       <TD CLASS=TD5 NOWRAP>B/L����</TD>
        <TD CLASS=TD6 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoBlFlag" id="rdoBlFlagY" value="Y" tag = "21">
          <label for="rdoPostFlagY">��</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoBlFlag" id="rdoBlFlagN" value="N" tag = "21" checked>
          <label for="rdoPostFlagN">�ƴϿ�</label></TD>
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
          <label for="rdoPostFlagN">��Ȯ��</label></TD>
        <TD CLASS="TD5">��ǥ��ȣ</TD>
        <TD CLASS="TD6"><INPUT NAME="txtAcctNo" ALT="��ǥ��ȣ" TYPE="Text" MAXLENGTH="18" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
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
            <script language =javascript src='./js/s5111ma2_fpDoubleSingle1_txtBillAmt.js'></script></TD>
           </TD>
           <TD>
            &nbsp;<INPUT NAME="txtDocCur2" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">
           </TD>
          </TR>
          </TABLE>
        </TD>
        <TD CLASS=TD5 NOWRAP>����ä���ڱ��ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <script language =javascript src='./js/s5111ma2_fpDoubleSingle5_txtBillAmtLoc.js'></script>
           </TD>
           <TD>
            &nbsp;<INPUT NAME="txtLocCur" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">
           </TD>
          </TR>
         </TABLE>
        </TD>
       </TR>  
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT�ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle2_txtVatAmt.js'></script>       
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT�ڱ��ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle6_txtVatLocAmt.js'></script>       
        </TD>
       </TR>        
       <TR>
        <TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle4_txtDepositAmt.js'></script>       
        </TD>
        <TD CLASS=TD5 NOWRAP>�������ڱ���</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle7_txtDepositAmtLoc.js'></script>       
        </TD>        
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>�Ѹ���ä�Ǳݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle4_txtTotBillAmt.js'></script>       
        </TD>
        <TD CLASS=TD5 NOWRAP>�Ѹ���ä���ڱ��ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle7_txtTotBillAmtLoc.js'></script>       
        </TD>        
       </TR>   
       <TR>
        <TD CLASS=TD5 NOWRAP>�Ѽ��ݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle4_txtIncomeAmt.js'></script>       
        </TD>
        <TD CLASS=TD5 NOWRAP>�Ѽ����ڱ���</TD>
        <TD CLASS=TD6 NOWRAP>
         <script language =javascript src='./js/s5111ma2_fpDoubleSingle7_txtIncomeLocAmt.js'></script>       
        </TD>        
       </TR>   
       <% Call SubFillRemBodyTD5656(10) %>
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
     <TD WIDTH=* Align=Right><a href = "VBSCRIPT:JumpChgCheck(BIZ_BillDtl_JUMP_ID)">���ܸ���ä�ǳ������</a>&nbsp;|&nbsp;<a href = "vbscript:JumpChgCheck(BIZ_BillCollect_JUMP_ID)">����ä�Ǽ��ݳ������</a></TD>
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

<INPUT TYPE=HIDDEN NAME="txtHBillNo" tag="14" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtHRefBillNo" tag="14" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtHExceptFlg" tag="14" TABINDEX ="-1">

<INPUT TYPE=HIDDEN NAME="txtHExportFlag" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtSts" tag="24" TABINDEX ="-1">

<INPUT TYPE=HIDDEN NAME="txtchkTaxNo" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtVatCalcType" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtVatIncFlag" tag="24" TABINDEX ="-1">

<INPUT TYPE=HIDDEN NAME="txtHRefFlag" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtGLNo" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtTempGLNo" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtBatchNo" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtCreditRotDay" tag="24" TABINDEX ="-1">
<INPUT TYPE=HIDDEN NAME="txtRefBillNoFlg" tag="24" TABINDEX ="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX ="-1"></iframe>
  </DIV>
</BODY>
</HTML>
