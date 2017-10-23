<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 미출고현황출력 
'*  3. Program ID           : s3112oa1
'*  4. Program Name         : 미출고현황 
'*  5. Program Desc         : 미출고현황 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/01/15
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : 손범열 
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'*                            2002/12/17 Include 성능향상 강준구 
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

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          

'=====================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub
'=====================================================
Sub SetDefaultVal()

	frm1.txtPlant.focus 
	frm1.txtDueFromDt.Text = StartDate
	frm1.txtDueToDt.Text = EndDate

End Sub

'=====================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'=====================================================
Function OpenConPop1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"						
	arrParam(1) = "B_PLANT"		                    
	arrParam(2) = Trim(frm1.txtPlant.value)		    
	arrParam(3) = Trim(frm1.txtPlant_NM.value)		
	arrParam(4) = ""
	arrParam(5) = "공장"						
	
	arrField(0) = "PLANT_CD"					        
	arrField(1) = "PLANT_NM"					        
    
	arrHeader(0) = "공장"						
	arrHeader(1) = "공장명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtPlant.focus 

	If arrRet(0) <> "" Then
		frm1.txtPlant.value	= arrRet(0)
		frm1.txtPlant_NM.value	= arrRet(1)
	End If

End Function
'=====================================================
Function OpenConPop2()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "납품처"						
	arrParam(1) = "B_BIZ_PARTNER"		                    
	arrParam(2) = Trim(frm1.txtShip_To_Party.value)		
	arrParam(3) = Trim(frm1.txtShip_To_Party_Nm.value)		
	arrParam(4) = "BP_TYPE IN(" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"
	arrParam(5) = "납품처"						
	
	arrField(0) = "BP_CD"					        
	arrField(1) = "BP_NM"					        
    
	arrHeader(0) = "납품처"						
	arrHeader(1) = "납품처명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtShip_To_Party.focus

	If arrRet(0) <> "" Then
		frm1.txtShip_To_Party.Value		= arrRet(0)
		frm1.txtShip_To_Party_Nm.Value	= arrRet(1)
	End If

End Function

'=====================================================
Function OpenConPop3()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "영업조직"						
	arrParam(1) = "B_SALES_ORG"		                    
	arrParam(2) = Trim(frm1.txtSales_org.value)		
	arrParam(3) = Trim(frm1.txtSales_Org_Nm.value)		
	arrParam(4) = ""	                            
	arrParam(5) = "영업조직"						
	
	arrField(0) = "SALES_ORG"					        
	arrField(1) = "SALES_ORG_NM"					        
    
	arrHeader(0) = "영업조직"						
	arrHeader(1) = "영업조직명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtSales_org.focus 

	If arrRet(0) <> "" Then
		frm1.txtSales_org.Value		= arrRet(0)
		frm1.txtSales_Org_Nm.Value		= arrRet(1)
	End If

End Function

'=======================================================
Sub Form_Load()

    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 

End Sub
'=======================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
   
End Sub

'=======================================================
Sub txtDueFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtDueFromDt.Focus
    End If
End Sub

'=======================================================
Sub txtDueToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtDueToDt.Focus
    End If
End Sub
'=====================================================
 Function FncPrint() 
	Call parent.FncPrint()
End Function
'=====================================================
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

'========================================
Function BtnPrint(ByVal pvStrPrint) 
    If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(frm1.txtDueFromDt, frm1.txtDueToDt) = False Then Exit Function
	
	Dim iStrUrl
	
	If UCase(frm1.txtPlant.value) = "" Then
		iStrUrl = "PLANT|%"
	Else
		iStrUrl = "PLANT|" & Replace(UCase(Trim(frm1.txtPlant.value)), "'", "''")
	End If

    If UCase(frm1.txtShip_To_Party.value) = "" Then
		iStrUrl = iStrUrl & "|SHIP_TO_PARTY|%"
	Else
		iStrUrl = iStrUrl & "|SHIP_TO_PARTY|" & Replace(UCase(Trim(frm1.txtShip_To_Party.value)), "'" , "''")
	End If
	
	If UCase(frm1.txtSales_org.value) = "" Then
		iStrUrl = iStrUrl & "|SALES_ORG|%"
	Else
		iStrUrl = iStrUrl & "|SALES_ORG|" & Replace(UCase(Trim(frm1.txtSales_org.value)), "'" , "''")
	End If

	iStrUrl = iStrUrl & "|FROM_DLVY_DT|" &UniConvDateToYYYYMMDD(frm1.txtDueFromDt.Text,parent.gDateFormat,parent.gServerDateType)
	iStrUrl = iStrUrl & "|TO_DLVY_DT|" &UniConvDateToYYYYMMDD(frm1.txtDueToDt.Text,parent.gDateFormat,parent.gServerDateType)

	OBjName = AskEBDocumentName("s3112oa1","ebr")    

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>미출고현황출력</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop1">&nbsp;<INPUT NAME="txtPlant_NM" TYPE="Text" SIZE=30 tag="14"></TD>								
							    </TR>
							    
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShip_To_Party" ALT="납품처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop2">&nbsp;<INPUT NAME="txtShip_To_Party_Nm" TYPE="Text" SIZE=30 tag="14"></TD>								
							    </TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>납기일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s3112oa1_fpDateTime1_txtDueFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s3112oa1_fpDateTime2_txtDueToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
					                </TD>
				                 </TR>			
							    <TR>
									<TD CLASS=TD5 NOWRAP>영업조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_org" ALT="영업조직" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnITEM_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop3">&nbsp;<INPUT NAME="txtSales_Org_Nm" TYPE="Text" SIZE=30 tag="14"></TD>								
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
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

