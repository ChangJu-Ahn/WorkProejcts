<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          :
*  2. Function Name        :
*  3. Program ID           :
*  4. Program Name         :
*  5. Program Desc         :
*  6. Comproxy List        :
*  7. Modified date(First) :
*  8. Modified date(Last)  :
*  9. Modifier (First)     :
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="B82108ma1.vbs"></SCRIPT>
<Script Language="VBScript">

Option Explicit                                       

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim BaseDt 

BaseDt = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	
'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="No" TABINDEX="-1">
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" CLASS="CLSMTAB" ALIGN="center"><FONT COLOR=white>ǰ��/�԰ݺ����Ƿڽ���</font></td>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</A></TD>
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
						            <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=Yes>
						            <!--<FIELDSET CLASS="CLSFLD">-->
									<TABLE <%=LR_SPACE_TYPE_60%>>
	    				        	    <TR>
              							    <TD CLASS=TD5 NOWRAP>�Ƿڹ�ȣ</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReqNo" ALT="�Ƿڹ�ȣ" TYPE="Text" SiZE=21 MAXLENGTH=18   tag="24XXXU">&nbsp;&nbsp;<INPUT NAME="txtStatus" ALT="�������" TYPE="Text" SiZE=11 MAXLENGTH=18 style="font-weight:bold;text-align:center;"  tag="24XXXU"></TD>
				            				<TD CLASS=TD5 NOWRAP>ǰ���ڵ�</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemCd" ALT="ǰ���ڵ�" TYPE="Text" SiZE=15 MAXLENGTH=18   tag="24XXXU">
				            				                     <INPUT NAME="txtItemNm" ALT="ǰ���"   TYPE="Text" SiZE=25 tag="24XXXU"></TD>
				            			</TR>
				            			<TR>
				            				<TD CLASS=TD5 NOWRAP>ǰ�����ĸ�</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm2" ALT="ǰ�����ĸ�" TYPE="Text" SiZE=38   tag="24XXXX"></TD>
				            				<TD CLASS=TD5 NOWRAP>�԰�</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpec"    ALT="�԰�" TYPE="Text" SiZE=38 tag="24XXXU"></TD>
              							</TR>	
				            			<TR>
				            			    <TD CLASS=TD5 NOWRAP>�󼼱԰�</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpec2"   ALT="�󼼱԰�" TYPE="Text" SiZE=38   tag="24XXXX"></TD>
				            				<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemAcct"   ALT="ǰ�����" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemAcctNm" ALT="ǰ�������" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
				            			</TR>
				            			<TR>                                            
                                            <TD CLASS=TD5 NOWRAP>ǰ�񱸺�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemKind"   ALT="ǰ�񱸺�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemKindNm" ALT="ǰ�񱸺и�" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>��з�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl1" ALT="��з�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemLvl1Nm" ALT="��з���" TYPE="Text" SiZE=25   tag="24XXXU"></TD>                     
                                        </TR>
                                        <TR>                                            
                                            <TD CLASS=TD5 NOWRAP>�ߺз�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl2" ALT="�ߺз�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemLvl2Nm" ALT="�ߺз���" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>�Һз�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl3" ALT="�Һз�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemLvl3Nm" ALT="�Һз���" TYPE="Text" SiZE=25   tag="24XXXU"></TD>                
                                        </TR>
                                        <TR>                                            
                                            <TD CLASS=TD5 NOWRAP>Serial No</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSerialNo" ALT="SerialNo" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                            <TD CLASS=TD5 NOWRAP>�̽��ο�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemVer" ALT="�̽��ο�" TYPE="Text" SiZE=10 MAXLENGTH=18   tag="24XXXX"></TD>
                                        </TR>
                                        <TR>
                                           <TD CLASS=TD5 NOWRAP>�Ļ�����</TD>  
                                           <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoDerive" ID="rdoDerive1" Value="Y" CLASS="RADIO" tag="24"><LABEL FOR="rdoDerive1">��</LABEL>
                                                                <INPUT TYPE="RADIO" NAME="rdoDerive" ID="rdoDerive2" Value="N" CLASS="RADIO" tag="24" CHECKED><LABEL FOR="rdoDerive2">�ƴϿ�</LABEL>
                                                                <INPUT TYPE= HIDDEN NAME="hrdoDerive"  SIZE= 10 MAXLENGTH=10  TAG="24" ALT="�Ļ�����"></TD>                     
                                            <TD CLASS=TD5 NOWRAP>����ǰ���ڵ�</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBasicItem"   ALT="����ǰ���ڵ�" TYPE="Text" SiZE=15 MAXLENGTH=18   tag="24XXXX">
                                                                 <INPUT NAME="txtBasicItemNm" ALT="�����ڵ��"   TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                        </TR>				            							            			
              							<TR>
              							    <TD CLASS="TD5" NOWRAP>��������</TD>
				            				<TD CLASS="TD6"><SELECT NAME="txtRgrade"  CLASS=cboNormal TAG="24" ALT="��������"><OPTION VALUE=""></OPTION></SELECT></TD>
				            				<TD CLASS="TD5" NOWRAP>�������</TD>
				            				<TD CLASS="TD6"><SELECT NAME="txtTgrade"  CLASS=cboNormal TAG="24" ALT="�������"><OPTION VALUE=""></OPTION></SELECT></TD>
				            			</TR>
				            			<TR>
              							    <TD CLASS="TD5" NOWRAP>���Ű���</TD>
				            				<TD CLASS="TD6"><SELECT NAME="txtPgrade"  CLASS=cboNormal TAG="24" ALT="���Ű���"><OPTION VALUE=""></OPTION></SELECT></TD>
				            				<TD CLASS="TD5" NOWRAP>ǰ������</TD>
				            				<TD CLASS="TD6"><SELECT NAME="txtQgrade"  CLASS=cboNormal TAG="24" ALT="ǰ������"><OPTION VALUE=""></OPTION></SELECT></TD>
				            			</TR>				            			     						    		
	        						</TABLE>
							        </DIV>
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
              							    <TD CLASS="TD5" NOWRAP>�Ϸ�����</TD>
				            				<TD CLASS="TD6"><script language =javascript src='./js/b82108ma1_fpDateTime1_txtEndDt.js'></script></TD>
	                   						<TD CLASS="TD5" NOWRAP>�̰�����</TD>
				            				<TD CLASS="TD6"><script language =javascript src='./js/b82108ma1_fpDateTime1_txtTransDt.js'></script></TD>
              							</TR>									    
	        						    <TR>
				            			    <TD CLASS=TD5 NOWRAP>�Ƿ���</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReqId" ALT="�Ƿ���" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
				            				                     <INPUT NAME="txtReqIdNm" ALT="�Ƿ��ڸ�" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
				            			    <TD CLASS=TD5 NOWRAP>�Ƿ�����</TD>
				            				<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82108ma1_fpDateTime1_txtReqDt.js'></script></TD>				            				                
	        						    </TR>
				            			<TR>
                                            <TD CLASS=TD5 NOWRAP>����ǰ���</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtNewItemNm"  ALT="ǰ���"     TYPE="Text" SiZE=40   tag="24XXXU"></TD>
	                   						<TD CLASS=TD5 NOWRAP>����ǰ�����ĸ�</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNewItemNm2" ALT="ǰ�����ĸ�" TYPE="Text" SiZE=40   tag="24XXXX"></TD>
				            			</TR>	
				            			<TR>	
				            				<TD CLASS=TD5 NOWRAP>����԰�</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNewSpec"    ALT="�԰�" TYPE="Text" SiZE=40 tag="24XXXU">
	        						        <TD CLASS=TD5 NOWRAP>����󼼱԰�</TD>
				            				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNewSpec2"   ALT="�󼼱԰�" TYPE="Text" SiZE=40   tag="24XXXX"></TD>
				            			</TR>       						    
	        						    <TR>
              							    <TD CLASS=TD5 NOWRAP>�Ƿڻ���</TD>
				            				<TD CLASS=TD6 ColSpan=3><TEXTAREA  NAME="txtReqReason" tag="24xxx" rows = 3 cols=90  ALT="�Ƿڻ���"></TEXTAREA>
																	<INPUT TYPE=HIDDEN NAME="htxtReqReason"  SIZE= 50 MAXLENGTH=200  TAG="24" ALT="�Ƿڻ���"></TD>
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
							<BUTTON NAME="Btn1" CLASS="CLSMBTN" ONCLICK="vbscript:BtnR()" >��������</BUTTON>&nbsp;
							<BUTTON NAME="Btn2" CLASS="CLSMBTN" ONCLICK="vbscript:BtnT()" >�������</BUTTON>&nbsp;
							<BUTTON NAME="Btn3" CLASS="CLSMBTN" ONCLICK="vbscript:BtnP()" >���Ű���</BUTTON>&nbsp;
							<BUTTON NAME="Btn4" CLASS="CLSMBTN" ONCLICK="vbscript:BtnQ()" >ǰ������</BUTTON>
						</TD>
						<TD ALIGN =right>
							 <A href = "VBSCRIPT:JumpChgCheck1()">ǰ��/�԰ݺ����Ƿڵ��</A>
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
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtStatus"     TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtInternalCd" TAG="24">

<INPUT TYPE=HIDDEN NAME="txtUpdMode"     TAG="24">

<INPUT TYPE=HIDDEN NAME="htxtRDt"        TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtRGrade"     TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtRDesc"      TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtRPerson"    TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtRPersonNm"  TAG="24">

<INPUT TYPE=HIDDEN NAME="htxtTDt"        TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtTGrade"     TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtTDesc"      TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtTPerson"    TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtTPersonNm"  TAG="24">

<INPUT TYPE=HIDDEN NAME="htxtPDt"        TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtPGrade"     TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtPDesc"      TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtPPerson"    TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtPPersonNm"  TAG="24">

<INPUT TYPE=HIDDEN NAME="htxtQDt"        TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtQGrade"     TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtQDesc"      TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtQPerson"    TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtQPersonNm"  TAG="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

