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
<SCRIPT LANGUAGE="VBScript"   SRC="B82102ma1.vbs"></SCRIPT>
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" CLASS="CLSMTAB" ALIGN="center"><FONT COLOR=white><%=Request("strASPMnuMnuNm")%></font></td>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenReReqRef">재의뢰내역</A>&nbsp;</A></TD>
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
												<TD CLASS=TD5 NOWRAP>의뢰번호</TD>
				            			        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtarReqNo" ALT="의뢰번호" TYPE="Text" SiZE=18 MAXLENGTH=18   tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenReqNo()"></TD>
				            			        <TD CLASS=TD5 NOWRAP>품목코드</TD>
				            				    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtarItemCd" ALT="품목코드" TYPE="Text" SiZE=15 MAXLENGTH=18   tag="24XXXU"> 
				            							             <INPUT NAME="txtarItemNm" ALT="품목명"   TYPE="Text" SiZE=25   tag="24XXXU"></TD>

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
                                            <TD CLASS=TD5 NOWRAP>의뢰번호</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtReqNo" ALT="의뢰번호" TYPE="Text" SiZE=23 MAXLENGTH=18   tag="24XXXU">    <INPUT NAME="txtStatus" ALT="Status" TYPE="Text" SiZE=10 MAXLENGTH=18 style="font-weight:bold;text-align:center;" tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>완료요청일자</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82102ma1_fpDateTime1_txtEndReqDt.js'></script></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>품목계정</TD>
                                            <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct"  CLASS=cboNormal TAG="24" ALT="품목계정"><OPTION VALUE=""></OPTION></SELECT></TD>
                                            <TD CLASS=TD5 NOWRAP>품목구분</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemKind"   ALT="품목구분" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemKindNm" ALT="품목구분명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>대분류</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl1" ALT="대분류" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemLvl1Nm" ALT="대분류명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>중분류</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl2" ALT="중분류" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemLvl2Nm" ALT="중분류명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                                            
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>소분류</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemLvl3" ALT="소분류" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtItemLvl3Nm" ALT="소분류명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>Serial No</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSerialNo" ALT="SerialNo" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU"></TD>
                                        </TR>
                                        <TR>
                                           <TD CLASS=TD5 NOWRAP>파생여부</TD>  
                                           <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoDerive" ID="rdoDerive1" Value="Y" CLASS="RADIO" tag="24"><LABEL FOR="rdoDerive1">예</LABEL>
                                                                <INPUT TYPE="RADIO" NAME="rdoDerive" ID="rdoDerive2" Value="N" CLASS="RADIO" tag="24" CHECKED><LABEL FOR="rdoDerive2">아니오</LABEL>
                                                                <INPUT TYPE= HIDDEN NAME="hrdoDerive"  SIZE= 10 MAXLENGTH=10  TAG="24" ALT="파생여부"></TD>                     
                                            <TD CLASS=TD5 NOWRAP>기준품목코드</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBasicItem"   ALT="기준품목코드" TYPE="Text" SiZE=15 MAXLENGTH=18   tag="24XXXX">
                                                                 <INPUT NAME="txtBasicItemNm" ALT="기준코드명"   TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                        </TR> 
                                        <TR>                        
                                            <TD CLASS=TD5 NOWRAP>이슈부여</TD>
                                            <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemVer"  CLASS=cboNormal TAG="24" ALT="이슈부여"><OPTION VALUE=""></OPTION></SELECT></TD>
                                            <TD CLASS=TD5 NOWRAP>품목명</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm" ALT="품목명"   TYPE="Text" SiZE=40 MAXLENGTH=50  tag="24XXXU"></TD>
                                        </TR>    
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>품목정식명</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm2" ALT="보조품목명" TYPE="Text" SiZE=40   tag="24XXXX"></TD>
                                            <TD CLASS=TD5 NOWRAP>규격</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpec" ALT="규격" TYPE="Text" SiZE=37 tag="24XXXU"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>상세규격</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpec2" ALT="상세규격" TYPE="Text" SiZE=40   tag="24XXXX"></TD>
                                            <TD CLASS=TD5 NOWRAP>재고단위</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemUnit" ALT="재고단위" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU"></TD>
                                        </TR>
                                         <TR>
                                            <TD CLASS=TD5 NOWRAP>조달구분</TD>
                                            <TD CLASS=TD6 NOWRAP><SELECT NAME="cboPurType"  CLASS=cboNormal TAG="24" ALT="조달구분"><OPTION VALUE=""></OPTION></SELECT></TD>
                                            <TD CLASS=TD5 NOWRAP>통합구매구분</TD>  
                                            <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoUnifyPurFlg" ID="rdoUnifyPurFlg1" Value="Y" CLASS="RADIO" tag="24X"><LABEL FOR="rdoUnifyPurFlg1">예</LABEL>
                                                                 <INPUT TYPE="RADIO" NAME="rdoUnifyPurFlg" ID="rdoUnifyPurFlg2" Value="N" CLASS="RADIO" tag="24X" CHECKED><LABEL FOR="rdoUnifyPurFlg2">아니오</LABEL>
                                                                 <INPUT TYPE= HIDDEN NAME="hrdoUnifyPurFlg"  SIZE= 10 MAXLENGTH=10  TAG="24" ALT="통합구매구분"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>공급처</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPurVendor" ALT="공급처" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtPurVendorNm" ALT="공급처명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>구매그룹</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPurGroup" ALT="구매그룹" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtPurGroupNm" ALT="구매그룹명" TYPE="Text" SiZE=20   tag="24XXXU"></TD>                                                            
                                        </TR>                                        
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>의뢰자</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtReqId" ALT="의뢰자" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="24XXXU">
                                                                 <INPUT NAME="txtReqIdNm" ALT="의뢰자명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>의뢰일자</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82102ma1_fpDateTime1_txtReqDt.js'></script></TD>                                                            
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>Net중량</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82102ma1_OBJECT1_txtNetWeight.js'></script>&nbsp;
                                                                 <INPUT TYPE=TEXT NAME="txtNetWeightUnit" SIZE=5 MAXLENGTH=3 tag="24XXXU" ALT="Net중량단위"></TD>
                                            <TD CLASS=TD5 NOWRAP>Gross중량</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82102ma1_OBJECT2_txtGrossWeight.js'></script>&nbsp;
                                                                 <INPUT TYPE=TEXT NAME="txtGrossWeightUnit" SIZE=5 MAXLENGTH=3 tag="24XXXU" ALT="Gross중량단위"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>CBM(부피)</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82102ma1_OBJECT3_txtCBM.js'></script></TD>
                                            <TD CLASS=TD5 NOWRAP>CBM정보</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCBMInfo" SIZE=40 MAXLENGTH=50 tag="24" ALT="CBM정보"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>HS코드</TD>
                                            <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHSCd" SIZE=10 MAXLENGTH=10 tag="24XXXU" ALT="HS코드">
                                                                 <INPUT NAME="txtHSNm" ALT="HS코드명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                            <TD CLASS=TD5 NOWRAP>유효기간</TD>
                                            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82102ma1_OBJECT4_txtValidFromDt.js'></script> &nbsp;~&nbsp;
                                                                 <script language =javascript src='./js/b82102ma1_OBJECT5_txtValidToDt.js'></script></TD>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>도면번호</TD>
                                            <TD CLASS=TD6 ColSpan=3><INPUT NAME="txtDocNo"  ALT="도면번호" TYPE="Text" SiZE=40 MAXLENGTH=60   tag="24XXXX">
                                                                    <INPUT NAME="txtFileNm" ALT="파일명" TYPE="Text"   SiZE=45 MAXLENGTH=100  tag="24XXXX">
                                                                    <INPUT style="FONT-SIZE: 9pt; WIDTH: 100px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 19px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbScript:OpenDocFile()' type=button value="도면파일보기" id=button1 name=button1></TD>
										</TR>                                        
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>의뢰사유</TD>
                                            <TD CLASS=TD6 ColSpan=3><TEXTAREA  NAME="txtReqReason" tag="24Xxx" rows = 2 cols=90  ALT="의뢰사유"></TEXTAREA>
                                                                    <INPUT TYPE=HIDDEN NAME="htxtReqReason"  SIZE= 50 MAXLENGTH=200  TAG="24" ALT="의뢰사유"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS=TD5 NOWRAP>비고</TD>
                                            <TD CLASS=TD6 ColSpan=3><TEXTAREA  NAME="txtRemark" tag="24Xxx" rows = 1 cols=90  ALT="비고"></TEXTAREA>
                                                                    <INPUT TYPE=HIDDEN NAME="htxtRemark"  SIZE= 50 MAXLENGTH=100  TAG="24" ALT="비고"></TD>
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
							<BUTTON NAME="Btn1" CLASS="CLSMBTN" ONCLICK="vbscript:BtnR()" >접수검토</BUTTON>&nbsp;
							<BUTTON NAME="Btn2" CLASS="CLSMBTN" ONCLICK="vbscript:BtnT()" >기술검토</BUTTON>&nbsp;
							<BUTTON NAME="Btn3" CLASS="CLSMBTN" ONCLICK="vbscript:BtnP()" >구매검토</BUTTON>&nbsp;
							<BUTTON NAME="Btn4" CLASS="CLSMBTN" ONCLICK="vbscript:BtnQ()" >품질검토</BUTTON>
						</TD>
						<TD ALIGN =right>
							 <A href = "VBSCRIPT:JumpChgCheck1()">품목신규의뢰등록</A>
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

<INPUT TYPE=HIDDEN NAME="txtUpdMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="htxtInternalCd" TAG="24">

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
	<IFRAME NAME="FR_ATTWIZ" SRC="../../Notice/FrAttwiz.html" MARGINWIDTH=0 MARGINHEIGHT=0 WIDTH=0 HEIGHT=0 ></IFRAME>
</DIV>
</BODY>
</HTML>

