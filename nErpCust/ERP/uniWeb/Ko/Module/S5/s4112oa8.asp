<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 납품처별 미출고집계 출력 
'*  3. Program ID           : s4112oa8
'*  4. Program Name         : 
'*  5. Program Desc         : 납품처별 미출고집계 출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/05
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : 손범열 
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Dim IsOpenPop          

'=========================================
Sub InitVariables()
    IsOpenPop = False
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtDueFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtDueToDt.Text = EndDate
	
    Call SetFocusToDocument("M")	
	frm1.txtDueFromDt.focus 
End Sub

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'=========================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "b_plant A"	
	arrParam(2) = Trim(frm1.txtPlant_cd.value)						
	arrParam(3) = ""											
	arrParam(4) = "Exists (SELECT * FROM b_item_by_plant B Where A.plant_cd=B.plant_cd) "			
	arrParam(5) = "공장"									
	
	arrField(0) = "ED15" & parent.gColSep & "A.plant_cd"				
	arrField(1) = "ED30" & parent.gColSep & "A.plant_nm"				
		    
	arrHeader(0) = "공장"									
	arrHeader(1) = "공장명"									

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
				
	IsOpenPop = False

	frm1.txtPlant_cd.focus
	If arrRet(0) <> "" Then
		frm1.txtPlant_cd.value = arrRet(0)
		frm1.txtPlant_nm.value = arrRet(1)
	End If

End Function
	
'=========================================
Function OpenConPop1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "납품처"							
	arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"		
	arrParam(2) = Trim(frm1.txtFromShipToParty.value)		
	arrParam(3) = Trim(frm1.txtFromShipToPartyNm.value)		
	arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SSH", "''", "S") & " " _
					& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE IN (" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"
	arrParam(5) = "납품처"							
		
	arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"	
	arrField(1) = "PARTNER.BP_NM"						
	arrField(2) = "PARTNER_FTN.BP_CD"					
	arrField(3) = "PARTNER_FTN.PARTNER_FTN"				
	arrField(4) = "PARTNER_FTN.USAGE_FLAG"				
	    
	arrHeader(0) = "납품처"							
	arrHeader(1) = "납품처명"						
	arrHeader(2) = "거래처코드"						
	arrHeader(3) = "거래처타입"						
	arrHeader(4) = "사용여부"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtFromShipToParty.focus
	If arrRet(0) <> "" Then
		frm1.txtFromShipToParty.Value	= arrRet(0)
		frm1.txtFromShipToPartyNm.Value	= arrRet(1)
	End If

End Function
	
'=========================================
Function OpenConPop2()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "납품처"							
	arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"		
	arrParam(2) = Trim(frm1.txtToShipToParty.value)		
	arrParam(3) = Trim(frm1.txtToShipToPartyNm.value)		
	arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SSH", "''", "S") & " " _
					& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE IN (" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"
	arrParam(5) = "납품처"							
		
	arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"			
	arrField(1) = "PARTNER.BP_NM"						
	arrField(2) = "PARTNER_FTN.BP_CD"					
	arrField(3) = "PARTNER_FTN.PARTNER_FTN"				
	arrField(4) = "PARTNER_FTN.USAGE_FLAG"				
	    
	arrHeader(0) = "납품처"							
	arrHeader(1) = "납품처명"						
	arrHeader(2) = "거래처코드"						
	arrHeader(3) = "거래처타입"						
	arrHeader(4) = "사용여부"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtToShipToParty.focus
	If arrRet(0) <> "" Then
		frm1.txtToShipToParty.Value	= arrRet(0)
		frm1.txtToShipToPartyNm.Value	= arrRet(1)
	End If

End Function

'=========================================
Function OpenConPop3()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "영업그룹"						
	arrParam(1) = "B_SALES_GRP"		                    
	arrParam(2) = Trim(frm1.txtSales_Grp.value)		
	arrParam(3) = Trim(frm1.txtSales_Grp_Nm.value)		
	arrParam(4) = ""	                            
	arrParam(5) = "영업그룹"						
	
	arrField(0) = "SALES_GRP"					        
	arrField(1) = "SALES_GRP_NM"					        
    
	arrHeader(0) = "영업그룹"						
	arrHeader(1) = "영업그룹명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtSales_Grp.focus
	
	If arrRet(0) <> "" Then
		frm1.txtSales_GRP.Value		= arrRet(0)
		frm1.txtSales_Grp_Nm.Value	= arrRet(1)
	End If

End Function

'=========================================
Function OpenDNType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	
	    arrParam(0) = "출하형태"					        
		arrParam(1) = "b_minor A, I_MOVETYPE_CONFIGURATION B"	
		arrParam(2) = Trim(frm1.txtDN_TYPE.value)		        
		arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & " "	
		arrParam(5) = "출하형태"					         

		arrField(0) = "A.MINOR_CD"							     
		arrField(1) = "A.MINOR_NM"							      

		arrHeader(0) = "출하형태"					         
		arrHeader(1) = "출하형태명"					          


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtDN_TYPE.focus
	If arrRet(0) <> "" Then
		frm1.txtDN_TYPE.Value	= arrRet(0)
		frm1.txtDN_TYPE_NM.Value= arrRet(1)
	End If

End Function

'=========================================
Sub Form_Load()

    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 

End Sub

'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
  
End Sub

'=========================================
Sub txtDueFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtDueFromDt.Focus
    End If
End Sub

'=========================================
Sub txtDueToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtDueToDt.Focus
    End If
End Sub

'=========================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'=========================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================
Function BtnPreview_OnClick()
	Call BtnPrint("N")
End Function

'========================================
Function BtnPrint_OnClick()
	Call BtnPrint("Y")
End Function

'=========================================
Function BtnPrint(ByVal pvStrPrint) 
	
    If Not chkField(Document, "1") Then	Exit Function

	If ValidDateCheck(frm1.txtDueFromDt, frm1.txtDueToDt) = False Then Exit Function

	Dim iStrUrl
    
	If Trim(frm1.txtFromShipToParty.value) = "" Then
		iStrUrl = "FromShipToParty|%"
	Else
		iStrUrl = "FromShipToParty|" & Replace(UCase(Trim(frm1.txtFromShipToParty.value)), "'" ,  "''")
	End If

	If Trim(frm1.txtToShipToParty.value) = "" Then
		iStrUrl = iStrUrl & "|ToShipToParty|%"
	Else
		iStrUrl = iStrUrl & "|ToShipToParty|" & Replace(UCase(Trim(frm1.txtToShipToParty.value)), "'" ,  "''")
	End If

	If Trim(frm1.txtPlant_cd.value) = "" Then
		iStrUrl = iStrUrl & "|Plant|%"
	Else
		iStrUrl = iStrUrl & "|Plant|" & Replace(UCase(Trim(frm1.txtPlant_cd.value)), "'" ,  "''")
	End If

	If Trim(frm1.txtDN_TYPE.value) = "" Then
		iStrUrl = iStrUrl & "|Mov_Type|%"
	Else
		iStrUrl = iStrUrl & "|Mov_Type|" & Replace(UCase(Trim(frm1.txtDN_TYPE.value)), "'" ,  "''")
	End If


	If Trim(frm1.txtSales_Grp.value) = "" Then
		iStrUrl = iStrUrl & "|Sales_Grp|%"
	Else
		iStrUrl = iStrUrl & "|Sales_Grp|" & Replace(UCase(Trim(frm1.txtSales_Grp.value)), "'" ,  "''")
	End If

	iStrUrl = iStrUrl & "|fromDlvyDt|" & UniConvDateToYYYYMMDD(frm1.txtDueFromDt.Text,parent.gDateFormat,parent.gServerDateType)
	iStrUrl = iStrUrl & "|toDlvyDt|" & UniConvDateToYYYYMMDD(frm1.txtDueToDt.Text,parent.gDateFormat,parent.gServerDateType)

	OBjName = AskEBDocumentName("s4112oa8","ebr")    

	If pvStrPrint = "N" Then
		' 미리보기 
		Call FncEBRPreview(ObjName, iStrUrl)
	Else
		' 출력 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTABP"><font color=white>납품처별미출고집계현황</font></td>
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
									<TD CLASS="TD5" NOWRAP>납기일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s4112oa8_fpDateTime1_txtDueFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s4112oa8_fpDateTime2_txtDueToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
					                </TD>
				                 </TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFromShipToParty" ALT="납품처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop1">&nbsp;<INPUT NAME="txtFromShipToPartyNm" TYPE="Text" SIZE=30 tag="14"></TD>								
							    </TR>
							    <TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;~&nbsp;<INPUT NAME="txtToShipToParty" ALT="납품처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop2">&nbsp;<INPUT NAME="txtToShipToPartyNm" TYPE="Text" SIZE=30 tag="14"></TD>								
							    </TR>

								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPlant_cd" ALT="공장" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShipToParty" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlant_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출하형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDN_TYPE" ALT="출하형태" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDN_TYPE" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDNType()">&nbsp;<INPUT NAME="txtDN_TYPE_NM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
							    <TR>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Grp" ALT="영업그룹" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSALES_GRP" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop3">&nbsp;<INPUT NAME="txtSales_Grp_Nm" TYPE="Text" SIZE=30 tag="14"></TD>								
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
						<BUTTON NAME="BtnPreview" CLASS="CLSSBTN" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="BtnPrint" CLASS="CLSSBTN" Flag=1>인쇄</BUTTON>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
	
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
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

