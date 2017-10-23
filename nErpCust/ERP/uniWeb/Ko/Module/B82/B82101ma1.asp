<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
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

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="B82101ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../B81/B81COMM.vbs"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit                                       

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->    
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim BaseDt,BaseDtTo

BaseDt = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
BaseDtTo = UniConvDateAToB("2999-12-31", parent.gServerDateFormat, parent.gDateFormat)
  
'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
    <% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
    <% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>        
End Sub

'======================================================================================================
'	Name : FetchWebSvrIp()
'	Description : 
'=======================================================================================================
Function FetchWebSvrIp()	

	Dim gHttpWebSvrIPURL
	
	gHttpWebSvrIPURL =  "http://<%= request.servervariables("server_name") %>"	
	FetchWebSvrIp = Split(gHttpWebSvrIPURL, "/")(2)
	
End Function

</SCRIPT>

<SCRIPT language=javascript>

//======================================================================================================
//	Name : ViewFile()
//	Description : 
//=======================================================================================================
function ViewFile(sMode, sRet){
	
	var strWebSvrIp;
	
	document.FR_ATTWIZ.SetLanguage('K');	
	document.FR_ATTWIZ.SetModUpload();
	document.FR_ATTWIZ.SetServerAutoDelete(1);
	document.FR_ATTWIZ.SetFileUIMode(1);
	document.FR_ATTWIZ.SetServerOption(0,0);	
    document.FR_ATTWIZ.SetServerInfo(FetchWebSvrIp(), '7775');
	document.FR_ATTWIZ.ViewFile(sMode, sRet);
}	

</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>

<BODY SCROLL="No" TABINDEX="-1" >
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%> >
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
                                <TD BACKGROUND"../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" WIDTH="10" HEIGHT="23"></td>
                                <TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" CLASS="CLSMTAB" ALIGN="center"><FONT COLOR=white><%=Request("strASPMnuMnuNm")%></font></td>
                                <TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right><A href="vbscript:OpenBasicItem()">����ǰ���ڵ�����</A>&nbsp;</A>
                                            <A href="vbscript:OpenReReqRef()">���Ƿڳ���</A>&nbsp;</A></TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD WIDTH=100% CLASS="Tab11">
            <TABLE <%=LR_SPACE_TYPE_20%>>
                <TR>
                    <TD WIDTH=100% HEIGHT=* VALIGN="TOP">
                    
                        <TABLE <%=LR_SPACE_TYPE_20%>>
                            <TR>
                                <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
                            </TR>
                            <TR>
                                <TD HEIGHT=20 WIDTH=100%>
                                    <FIELDSET CLASS="CLSFLD">
                                        <TABLE <%=LR_SPACE_TYPE_40%>>
                                            <TR>
                                                <TD CLASS=TD5 NOWRAP>�Ƿڹ�ȣ</TD>
                                                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtarReqNo" ALT="�Ƿڹ�ȣ" TYPE="Text" SiZE=18 MAXLENGTH=18   tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenReqNo()"></TD>
                                                <TD CLASS=TD5 NOWRAP>ǰ���ڵ�</TD>
                                                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtarItemCd" ALT="ǰ���ڵ�" TYPE="Text" SiZE=15 MAXLENGTH=18   tag="24XXXU"> 
                                                                     <INPUT NAME="txtarItemNm" ALT="ǰ���"   TYPE="Text" SiZE=25   tag="24XXXU"></TD>

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
                                    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=Yes>
                                    <!--<FIELDSET CLASS="CLSFLD">-->
                                    <TABLE <%=LR_SPACE_TYPE_60%>>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>�Ƿڹ�ȣ</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtReqNo" ALT="�Ƿڹ�ȣ" TYPE="Text" SiZE=23 MAXLENGTH=18   tag="24XXXU">    <INPUT NAME="txtStatus" ALT="Status" TYPE="Text" SiZE=10 MAXLENGTH=18 style="font-weight:bold;text-align:center;" tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>�Ϸ��û����</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82101ma1_fpDateTime1_txtEndReqDt.js'></script></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>ǰ�����</TD>
                                            <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct"  CLASS=cboNormal TAG="22" ALT="ǰ�����"><OPTION VALUE=""></OPTION></SELECT></TD>
                                            <TD CLASS=TD5 NOWRAP>ǰ�񱸺�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemKind"   ALT="ǰ�񱸺�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('1')">
                                                                 <INPUT NAME="txtItemKindNm" ALT="ǰ�񱸺и�" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>��з�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl1" ALT="��з�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('2')">
                                                                 <INPUT NAME="txtItemLvl1Nm" ALT="��з���" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>�ߺз�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl2" ALT="�ߺз�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('3')">
                                                                 <INPUT NAME="txtItemLvl2Nm" ALT="�ߺз���" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                                            
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>�Һз�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl3" ALT="�Һз�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('4')">
                                                                 <INPUT NAME="txtItemLvl3Nm" ALT="�Һз���" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>Serial No</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSerialNo" ALT="SerialNo" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU"></TD>
                                        </TR>
                                        <TR>
                                           <TD CLASS=TD5 NOWRAP>�Ļ�����</TD>  
                                           <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoDerive" ID="rdoDerive1" Value="Y" CLASS="RADIO" tag="2"><LABEL FOR="rdoDerive1">��</LABEL>
                                                                <INPUT TYPE="RADIO" NAME="rdoDerive" ID="rdoDerive2" Value="N" CLASS="RADIO" tag="2" CHECKED><LABEL FOR="rdoDerive2">�ƴϿ�</LABEL>
                                                                <INPUT TYPE= HIDDEN NAME="hrdoDerive"  SIZE= 10 MAXLENGTH=10  TAG="24" ALT="�Ļ�����"></TD>                     
                                            <TD CLASS=TD5 NOWRAP>����ǰ���ڵ�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBasicItem"   ALT="����ǰ���ڵ�" TYPE="Text" SiZE=15 MAXLENGTH=18   tag="24XXXX">
                                                                 <INPUT NAME="txtBasicItemNm" ALT="�����ڵ��"   TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                        </TR>
                                        <TR>                         
                                            <TD CLASS=TD5 NOWRAP>�̽��ο�</TD>
                                            <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemVer"  CLASS=cboNormal TAG="2" ALT="�̽��ο�"><OPTION VALUE=""></OPTION></SELECT></TD>
                                            <TD CLASS=TD5 NOWRAP>ǰ���</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm" ALT="ǰ���"   TYPE="Text" SiZE=40 MAXLENGTH=40  tag="22XXXU">
                                                                 <INPUT TYPE= HIDDEN NAME="htxtItemCd" ALT="ǰ��"   TYPE="Text" SiZE=20 MAXLENGTH=18  tag="24"></TD>
                                        </TR>    
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>ǰ�����ĸ�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm2" ALT="����ǰ���" TYPE="Text" SiZE=40  maxlength=40 tag="21XXXX"></TD>
                                            <TD CLASS=TD5 NOWRAP>�԰�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpec" ALT="�԰�" TYPE="Text" SiZE=37 tag="22XXXU" maxlength=40><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript:OpenCategory()"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>�󼼱԰�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpec2" ALT="�󼼱԰�" TYPE="Text" SiZE=40 maxlength=40  tag="21XXXX"></TD>
                                            <TD CLASS=TD5 NOWRAP>������</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemUnit" ALT="������" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('8')"></TD>
                                        </TR>
                                         <TR>
                                            <TD CLASS=TD5 NOWRAP>���ޱ���</TD>
                                            <TD CLASS=TD6 NOWRAP><SELECT NAME="cboPurType"  CLASS=cboNormal TAG="21" ALT="���ޱ���"><OPTION VALUE=""></OPTION></SELECT></TD>
                                            <TD CLASS=TD5 NOWRAP>���ձ��ű���</TD>  
                                            <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoUnifyPurFlg" ID="rdoUnifyPurFlg1" Value="Y" CLASS="RADIO" tag="21X"><LABEL FOR="rdoUnifyPurFlg1">��</LABEL>
                                                                 <INPUT TYPE="RADIO" NAME="rdoUnifyPurFlg" ID="rdoUnifyPurFlg2" Value="N" CLASS="RADIO" tag="21X" CHECKED><LABEL FOR="rdoUnifyPurFlg2">�ƴϿ�</LABEL>
                                                                 <INPUT TYPE= HIDDEN NAME="hrdoUnifyPurFlg"  SIZE= 10 MAXLENGTH=10  TAG="24" ALT="���ձ��ű���"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>����ó</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPurVendor" ALT="����ó" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('9')">
                                                                 <INPUT NAME="txtPurVendorNm" ALT="����ó��" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>���ű׷�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPurGroup" ALT="���ű׷�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('10')">
                                                                 <INPUT NAME="txtPurGroupNm" ALT="���ű׷��" TYPE="Text" SiZE=20   tag="24XXXU"></TD>                                                            
                                        </TR>                                        
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>�Ƿ���</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtreq_user" ALT="�Ƿ���" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('12')">
                                                                 <INPUT NAME="txtreq_user_Nm" ALT="�Ƿ��ڸ�" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>�Ƿ�����</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82101ma1_fpDateTime1_txtReqDt.js'></script></TD>                                                            
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>Net�߷�</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82101ma1_OBJECT1_txtNetWeight.js'></script>&nbsp;
                                                                 <INPUT TYPE=TEXT NAME="txtNetWeightUnit" SIZE=5 MAXLENGTH=3 tag="21XXXU" ALT="Net�߷�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWeightUnit" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup('13')"></TD>
                                            <TD CLASS=TD5 NOWRAP>Gross�߷�</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82101ma1_OBJECT2_txtGrossWeight.js'></script>&nbsp;
                                                                 <INPUT TYPE=TEXT NAME="txtGrossWeightUnit" SIZE=5 MAXLENGTH=3 tag="21XXXU" ALT="Gross�߷�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrossWeightUnit" align = top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup('14')"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>CBM(����)</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82101ma1_OBJECT3_txtCBM.js'></script></TD>
                                            <TD CLASS=TD5 NOWRAP>CBM����</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCBMInfo" SIZE=40 MAXLENGTH=50 tag="21" ALT="CBM����"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>HS�ڵ�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHSCd" SIZE=13 MAXLENGTH=15 tag="21XXXU" ALT="HS�ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnHsCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup('15')">
                                                                 <INPUT NAME="txtHSNm" ALT="HS�ڵ��" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82101ma1_OBJECT4_txtValidFromDt.js'></script> &nbsp;~&nbsp;
                                                                 <script language =javascript src='./js/b82101ma1_OBJECT5_txtValidToDt.js'></script></TD>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>�����ȣ</TD>
                                            <TD CLASS=TD6 ColSpan=3><INPUT NAME="txtDocNo"  ALT="�����ȣ" TYPE="Text" SiZE=40 MAXLENGTH=20   tag="21XXXX">
                                                                    <INPUT NAME="txtFileNm" ALT="���ϸ�" TYPE="Text"   SiZE=45 MAXLENGTH=100  tag="24XXXX">
                                                                    <INPUT style="FONT-SIZE: 9pt; WIDTH: 100px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 19px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbScript:OpenDocFile()' type=button value="�������ϰ���" id=button1 name=button1 tag="22"></TD>
										</TR>                                       
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>�Ƿڻ���</TD>
                                            <TD CLASS=TD6 ColSpan=3><TEXTAREA  NAME="txtReqReason" tag="22xxx" rows = 2 cols=90  ALT="�Ƿڻ���"></TEXTAREA>
                                                                    <INPUT TYPE=HIDDEN NAME="htxtReqReason"  SIZE= 50 MAXLENGTH=200  TAG="24" ALT="�Ƿڻ���"></TD>
                                        </TR>
                                        <TR>
                                              <TD CLASS=TD5 NOWRAP>���</TD>
                                            <TD CLASS=TD6 ColSpan=3><INPUT  NAME="txtRemark" tag="21xxx"  SiZE=107 MAXLENGTH=200   ALT="���"></TEXTAREA>
                                                                    <INPUT TYPE=HIDDEN NAME="htxtRemark"  SIZE= 50 MAXLENGTH=100  TAG="24" ALT="���"></TD>
                                        </TR>                                                
                                    </TABLE>
                                    </DIV>
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
    <TR HEIGHT=12>
        <TD <%=HEIGHT_TYPE_03%> WIDTH=100%>
            <TABLE <%=LR_SPACE_TYPE_20%>>
                    <TR>
                        <TD>
                            <BUTTON NAME="btnRun" CLASS="CLSMBTN" ONCLICK="vbscript:RunReReq()" >���Ƿ�</BUTTON>                            
                        </TD>
                    </TR>
            </TABLE>
        </TD>
    </TR>
    <TR >
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtReReq"      TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtStatus"     TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtInternalCd" TAG="24">

<INPUT TYPE=HIDDEN NAME="htxtFilePath"   TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtIdFile"     TAG="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
    <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
    <IFRAME NAME="FR_ATTWIZ" SRC="../../Notice/FrAttwiz.html" MARGINWIDTH=0 MARGINHEIGHT=0 WIDTH=0 HEIGHT=0 ></IFRAME>
</DIV>
</BODY>
</HTML>

