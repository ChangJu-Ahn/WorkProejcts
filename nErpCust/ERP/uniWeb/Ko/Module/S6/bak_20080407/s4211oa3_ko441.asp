<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S4211OA3
'*  4. Program Name         : 통관관리대장 출력 
'*  5. Program Desc         : 통관관리대장 출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/07/18
'*  8. Modified date(Last)  : 2000/07/18
'*  9. Modifier (First)     : Cho Sung Hyun
'* 10. Modifier (Last)      : 손범열 
'* 11. Comment              :
'**********************************************************************************************

%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit																	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = parent.UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = parent.UnIDateAdd("m", -1, EndDate, parent.gDateFormat)
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim IsOpenPop 

'==========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE     
    lgBlnFlgChgValue = False        
    lgIntGrpCount = 0                                       
End Sub

'==========================================================================================================
Sub SetDefaultVal()
	frm1.txtApplicant.focus 
	frm1.txtCCFromDt.text = StartDate
	frm1.txtCCToDt.text = EndDate
End Sub
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'==========================================================================================================
Function OpenConPop()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "수입자"						
	arrParam(1) = "B_BIZ_PARTNER"					
	arrParam(2) = Trim(frm1.txtApplicant.value)		
	arrParam(3) = ""								
	arrParam(4) = "BP_TYPE IN (" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"					
	arrParam(5) = "수입자"						
			
	arrField(0) = "BP_CD"							
	arrField(1) = "BP_NM"							
		    
	arrHeader(0) = "수입자"						
	arrHeader(1) = "수입자명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtApplicant.focus 

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop(arrRet)
	End If

End Function
'==========================================================================================================
Function SetConPop(Byval arrRet)
	With frm1	
		.txtApplicant.Value		= arrRet(0)
		.txtApplicant.focus
	End With
End Function

'==========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables														
    Call GetValue_ko441()
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										
    
    frm1.txtApplicant.focus 
    
End Sub

'==========================================================================================================
Sub txtCCFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtCCFromDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtCCFromDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub
'==========================================================================================================
Sub txtCCToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtCCToDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtCCToDt.Focus
    End If   
    lgBlnFlgChgValue = True 
End Sub

'==========================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'==========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     
End Function

'==========================================================================================================
Function FncQuery() 
    FncQuery = true
End Function

'==========================================================================================================
Function BtnPrint() 
	Dim strUrl
	Dim ObjName    
	Dim vargBizArea,vargPlant,vargSalesGrp,vargSalesOrg    

	If ValidDateCheck(frm1.txtCCFromDt, frm1.txtCCToDt) = False Then Exit Function
    
    If Not chkField(Document, "1") Then									
       Exit Function
    End If

	dim var1, var2 ,var3
	
	
	If UCase(frm1.txtApplicant.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtApplicant.value)), "" ,  "SNM")
	End If
 	    var2 = UniConvDateToYYYYMMDD(frm1.txtCCFromDt.text,parent.gDateFormat,parent.gServerDateType)
     	var3 = UniConvDateToYYYYMMDD(frm1.txtCCToDt.text,parent.gDateFormat,parent.gServerDateType)

	If lgBACd <> "" Then
		vargBizArea = " AND S_CC_HDR.BIZ_AREA =  " & FilterVar(Trim(UCase(lgBACd)), "" ,  "S")
	Else
		vargBizArea = ""
	End If
	If lgPLCd <> "" Then
		vargPlant = " AND S_CC_DTL.PLANT_CD = " & FilterVar(Trim(UCase(lgPLCd)), "" ,  "S")
	Else
		vargPlant = ""
	End If
	If lgSGCd <> "" Then
		vargSalesGrp = " AND S_CC_HDR.SALES_GRP = " & FilterVar(Trim(UCase(lgSGCd)), "" ,  "S")
	Else
		vargSalesGrp = ""
	End If
	If lgSOCd <> "" Then
		vargSalesOrg = " AND S_CC_HDR.SALES_ORG = " & FilterVar(Trim(UCase(lgSOCd)), "" ,  "S")
	Else
		vargSalesOrg = ""
	End If


	strUrl = strUrl & "ConApplicant|" & var1 & "|CCFromDt|" & var2 & "|CCToDt|" & var3 
	strUrl = strUrl & "|gBizArea|" & vargBizArea 
	strUrl = strUrl & "|gPlant|" & vargPlant 
	strUrl = strUrl & "|gSalesGrp|" & vargSalesGrp
	strUrl = strUrl & "|gSalesOrg|" & vargSalesOrg
 	
	ObjName = AskEBDocumentName("S4211oa3_KO441", "ebr")
	call FncEBRprint(EBAction, ObjName, strUrl)
	
End Function
'==========================================================================================================
Function BtnPreview() 
    
	Dim ObjName
	Dim vargBizArea,vargPlant,vargSalesGrp,vargSalesOrg 
	
	If ValidDateCheck(frm1.txtCCFromDt, frm1.txtCCToDt) = False Then Exit Function

    
    If Not chkField(Document, "1") Then									
       Exit Function
    End If

	Dim var1, var2, var3
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader

	If UCase(frm1.txtApplicant.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtApplicant.value)), "" ,  "SNM")
	End If

		var2 = UniConvDateToYYYYMMDD(frm1.txtCCFromDt.text,parent.gDateFormat,parent.gServerDateType)
     	var3 = UniConvDateToYYYYMMDD(frm1.txtCCToDt.text,parent.gDateFormat,parent.gServerDateType)

	If lgBACd <> "" Then
		vargBizArea = " AND S_CC_HDR.BIZ_AREA =  " & FilterVar(Trim(UCase(lgBACd)), "" ,  "S")
	Else
		vargBizArea = ""
	End If
	If lgPLCd <> "" Then
		vargPlant = " AND S_CC_DTL.PLANT_CD = " & FilterVar(Trim(UCase(lgPLCd)), "" ,  "S")
	Else
		vargPlant = ""
	End If
	If lgSGCd <> "" Then
		vargSalesGrp = " AND S_CC_HDR.SALES_GRP = " & FilterVar(Trim(UCase(lgSGCd)), "" ,  "S")
	Else
		vargSalesGrp = ""
	End If
	If lgSOCd <> "" Then
		vargSalesOrg = " AND S_CC_HDR.SALES_ORG = " & FilterVar(Trim(UCase(lgSOCd)), "" ,  "S")
	Else
		vargSalesOrg = ""
	End If

	
	strUrl = strUrl & "ConApplicant|" & var1 & "|CCFromDt|" & var2 & "|CCToDt|" & var3
	strUrl = strUrl & "|gBizArea|" & vargBizArea 
	strUrl = strUrl & "|gPlant|" & vargPlant 
	strUrl = strUrl & "|gSalesGrp|" & vargSalesGrp
	strUrl = strUrl & "|gSalesOrg|" & vargSalesOrg
	
	ObjName = AskEBDocumentName("S4211oa3_KO441", "ebr")
	Call FncEBRPreview(ObjName, strUrl)	
		
End Function
'==========================================================================================================
Function FncExit()
 FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관관리대장출력</font></td>
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
									<TD CLASS=TD5 NOWRAP>수입자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" ALT="수입자" SIZE=10 MAXLENGTH=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPop" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPop"><div style="display:none"><input type="text" name="none"></div></TD>								                      
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>작성일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s4211oa3_fpDateTime1_txtCCFromDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s4211oa3_fpDateTime2_txtCCToDt.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    
						<TR>
						    <TD WIDTH=10>&nbsp;</TD>
							<TD>
							    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
							    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON>
							</TD>
						</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> SRC= "../../blank.htm" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
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
