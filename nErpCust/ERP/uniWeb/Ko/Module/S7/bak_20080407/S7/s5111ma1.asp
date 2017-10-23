<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 영업
'*  2. Function Name        : 매출채권관리
'*  3. Program ID           : S5111MA1
'*  4. Program Name         : 매출채권등록
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G111.cSBillHdrSvr,PS3G102.cLookupSoHdrSvr,PB5CS41.cLookupBizPartnerSvr
'*							  PS4G119.cSLkLcHdrSvr,PB5CS41.cLookupBizPartnerSvr	
'*							  PS7G115.cSPostOpenArSvr
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : 3rd 화면 Layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 Layout
'*                            -2001/12/18 : Date 표준적용
'*                            -2002/11/15 : UI 표준적용
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

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>T_매출채권일반</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD CLASS="CLSSTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
       <TR>
        <td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>T_채권금액정보</font></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD CLASS="CLSSTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
       <TR>
        <td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSSTAB"><font color=white>T_무역정보</font></td>
        <td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right><A href="vbscript:OpenSORef">수주참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenLCRef">L/C참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenDNRef">출하참조</A></TD>
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
         <TD CLASS="TD5" NOWRAP>매출채권번호</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConBillNo" ALT="매출채권번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConBillNo()"></TD>
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
        <TD CLASS="TD5" NOWRAP>매출채권번호</TD>
        <TD CLASS="TD6"><INPUT NAME="txtBillNo" ALT="매출채권번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="25XXXU" STYLE="text-transform:uppercase"></TD>
        <TD CLASS=TD5 NOWRAP>매출채권형태</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillTypeCd" ALT="매출채권형태" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="23XXXU" STYLE="text-transform:uppercase" class=required ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 0">&nbsp;<INPUT NAME="txtBillTypeNm" TYPE="Text" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
        <TD CLASS=TD5 NOWRAP>매출채권일</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtBillDt" CLASS="FPDTYYYYMMDD" tag="22X1" ALT="매출채권일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
           </TD>
          </TR>
         </TABLE>
        </TD>
        <TD CLASS=TD5 NOWRAP>수주번호</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoNo" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;
         <LABEL FOR="chkSoNo">수주번호지정</LABEL><INPUT TYPE=CHECKBOX NAME="chkSoNo" tag="25X" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid">
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>발행처</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillToPartyCd" ALT="발행처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 8">&nbsp;<INPUT NAME="txtBillToPartyNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>주문처</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldtoPartyCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" ALT="주문처" class = protected readonly = true TABINDEX="-1" >&nbsp;<INPUT NAME="txtSoldtoPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>수금처</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayerCd" ALT="수금처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 1">&nbsp;<INPUT NAME="txtPayerNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP></TD>
        <TD CLASS=TD6 NOWRAP></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>수금영업그룹</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtToBizAreaCd" ALT="수금영업그룹" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 2">&nbsp;<INPUT NAME="txtToBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>영업그룹</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>환율</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchgRate" CLASS="FPDS100" ALT="환율" tag="22X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
        </TD>
        <TD CLASS=TD5 NOWRAP>화폐단위</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur1" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT적용기준</TD>
        <TD CLASS=TD6 NOWRAP>
         <INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVATCalcType" TAG="21" VALUE="1" CHECKED ID="rdoVATCalcType1">
         <LABEL FOR="rdoVATCalcType1">개별</LABEL>
         <INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVATCalcType" TAG="21" VALUE="2" ID="rdoVATCalcType2">
         <LABEL FOR="rdoVATCalcType2">통합</LABEL>
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT포함구분</TD>
        <TD CLASS=TD6 NOWRAP>
         <INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVatIncFlag" TAG="21" VALUE="1" CHECKED ID="rdoVatIncFlag1">
         <LABEL FOR="rdoVatIncFlag1">별도</LABEL>
         <INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVatIncFlag" TAG="21" VALUE="2" ID="rdoVatIncFlag2">
         <LABEL FOR="rdoVatIncFlag2">포함</LABEL>
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT유형</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <INPUT NAME="txtVatType" TYPE="Text" MAXLENGTH="5" SIZE=10 ALT="VAT유형" tag="23XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 10">&nbsp;
           </TD>
           <TD>
            <INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="25" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1">
           </TD>
          </TR>
         </TABLE>
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT율</TD>
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
        <TD CLASS=TD5 NOWRAP>결제방법</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermsCd" Alt="결제방법" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 5">&nbsp;<INPUT NAME="txtPayTermsNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"  class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>결제기간</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle8 NAME="txtPayDur" CLASS="FPDS100" Alt="결제기간" tag="21X6Z" Title="FPDOUBLESINGLE"> </OBJECT>');</SCRIPT>&nbsp;<LABEL>일.</LABEL>
        </TD>
       </TR>       
       <TR>
        <TD CLASS=TD5 NOWRAP>입금유형</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTypeCd" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="입금유형" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 4">&nbsp;<INPUT NAME="txtPayTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"  class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>수금만기일</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtPlanIncomeDt" CLASS="FPDTYYYYMMDD" tag="21X1" ALT="수금만기일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
           </TD>
          </TR>
         </TABLE>
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>세금신고사업장</LABEL></TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaCd" ALT="세금신고사업장" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 3">&nbsp;<INPUT NAME="txtTaxBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP></TD>
        <TD CLASS=TD6 NOWRAP></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>대금결제참조사항</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaytermsTxt" ALT="대금결제참조" TYPE="Text" MAXLENGTH="120" SIZE=39 tag="21"></TD>
        <TD CLASS=TD5 NOWRAP>비고</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemark" ALT="비고" TYPE="Text" MAXLENGTH="120" SIZE=39 tag="21"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>확정여부</TD>
        <TD CLASS=TD6 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoPostFlag" id="rdoPostFlagY" value="Y" tag = "24">
          <label for="rdoPostFlagY">확정</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoPostFlag" id="rdoPostFlagN" value="N" tag = "24" checked>
          <label for="rdoPostFlagN">미확정</label>
        </TD>
        <TD CLASS=TD5 NOWRAP>전표번호</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctNo" ALT="전표번호" TYPE="Text" MAXLENGTH="18" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD HEIGHT=20 WIDTH=100% CLASS=TD6 COLSPAN=4>
         <FIELDSET ID="filTaxNo" CLASS="CLSFLD" TITLE="세금계산서자동발행">
         <LEGEND ALIGN=LEFT><INPUT TYPE=CHECKBOX NAME="chkTaxNo" tag="21" Class="Check"><LABEL FOR="chkTaxNo">세금계산서자동발행여부</LABEL></LEGEND>
          <TABLE <%=LR_SPACE_TYPE_40%>>
           <TR>
            <TD CLASS=TD5 NOWRAP><LABEL ID="lblTaxBillNo">세금계산서번호</LABEL></TD>
            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBillNo" ALT="세금계산서번호" TYPE="Text" MAXLENGTH="30" SIZE=30 tag="24XXXU" class = protected readonly = true TABINDEX="-1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillTaxNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTaxNo()"></TD>
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
        <TD CLASS=TD5 NOWRAP>매출채권금액</TD>
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
        <TD CLASS=TD5 NOWRAP>매출채권자국금액</TD>
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
        <TD CLASS=TD5 NOWRAP>VAT금액</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtVatAmt" ALT="VAT금액" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT자국금액</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtVatLocAmt" ALT="VAT자국금액" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
       </TR>        
       <TR>
        <TD CLASS=TD5 NOWRAP>적립금액</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtDepositAmt" Alt="적립금액" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
        <TD CLASS=TD5 NOWRAP>적립금자국액</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtDepositAmtLoc" Alt="적립금자국액" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>        
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>총매출채권금액</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtTotBillAmt" Alt="총매출채권금액" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
        <TD CLASS=TD5 NOWRAP>총매출채권자국금액</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtTotBillAmtLoc" Alt="총매출채권자국금액" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>        
       </TR>   
       <TR>
        <TD CLASS=TD5 NOWRAP>총수금액</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtIncomeAmt" Alt="총수금액" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
        <TD CLASS=TD5 NOWRAP>총수금자국액</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtIncomeLocAmt" Alt="총수금자국액" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>        
       </TR>   
       <% Call SubFillRemBodyTD5656(10) %>
      </TABLE>
      </DIV>

      <DIV ID="TabDiv" SCROLL=no>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD CLASS=TD5 NOWRAP>양수자</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBeneficiaryCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" Alt="양수자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 6">&nbsp;<INPUT NAME="txtBeneficiaryNm" TYPE="Text" MAXLENGTH="20" SIZE=23 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>양도자</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtApplicantCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" Alt="양도자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr 7">&nbsp;<INPUT NAME="txtApplicantNm" TYPE="Text" MAXLENGTH="20" SIZE=23 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>L/C관리번호</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>Amend차수</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCAmendSeq" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>L/C번호</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE="Text" MAXLENGTH="35" SIZE=35 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>통관FOB금액</TD>
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
     <TD><BUTTON NAME="btnPostFlag" CLASS="CLSMBTN">확정</BUTTON>&nbsp;
      <BUTTON NAME="btnGLView" CLASS="CLSMBTN">전표조회</BUTTON>&nbsp;
      <BUTTON NAME="btnPreRcptView" CLASS="CLSMBTN">선수금현황</BUTTON></TD>
     <TD WIDTH=* Align=Right><a href = "VBSCRIPT:JumpChgCheck(BIZ_BillDtl_JUMP_ID)">매출채권내역등록</a>&nbsp;|&nbsp;<a href = "vbscript:JumpChgCheck(BIZ_BillCollect_JUMP_ID)">매출채권수금내역등록</a></TD>
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
