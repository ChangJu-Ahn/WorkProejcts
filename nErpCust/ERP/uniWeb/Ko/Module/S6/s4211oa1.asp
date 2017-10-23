<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업관리 
'*  2. Function Name        : 
'*  3. Program ID           : S4211OA1
'*  4. Program Name         : Commercial Invoice 출력 
'*  5. Program Desc         : Commercial Invoice 출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/07/19
'*  8. Modified date(Last)  : 2000/07/19
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = parent.UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = parent.UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_CCQRY_ID = "s4211ob1.asp"
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim IsOpenPop          
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
        
End Sub
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtShipper.focus 
	frm1.txtIvDate.text = EndDate
	frm1.txtLCDate.text = EndDate
	frm1.txtSailing.text = EndDate
End Sub
'=========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenCCRef()
	Dim arrRet
	Dim iCalledAspName
	Dim IntRetCD			
		
	frm1.txtHCCRef.value = ""
	
	' Popup의 title을 "통관참조"로 보여주기 위해 Z_PR_ASPNAME에 "S4211PA2"추가 
	' 실제로 실행되는 프로그램은 "S4211PA1"임 
	iCalledAspName = AskPRAspName("S4211PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S4211PA2", "X")			
		Exit Function
	End If
				
		
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	If arrRet(0) = "" Then
		Exit Function
	Else
	    frm1.txtHCCRef.value = "Ref"
		Call SetCCRef(arrRet)
	End If
End Function	
'=========================================================================================================
Function OpenConPop()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

		arrParam(0) = "은행"							
		arrParam(1) = "B_Bank"								
		arrParam(2) = Trim(frm1.txtLCBank.value)												
		arrParam(3) = Trim(frm1.txtLCBankNm.value)													
		arrParam(4) = ""									
		arrParam(5) = "은행"							
		
		arrField(0) = "Bank_cd"								
		arrField(1) = "BANK_NM"							
	    
		arrHeader(0) = "은행"							
		arrHeader(1) = "은행명"							

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtLCBank.focus 

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop(arrRet)
	End If

End Function
'=========================================================================================================
Function SetConPop(Byval arrRet)
	With frm1	
		.txtLCBank.Value = arrRet(0)
		.txtLCBankNm.Value = arrRet(1)
		.txtLCBank.focus
	End With

End Function
'=========================================================================================================
Function SetCCRef(strRet)
		
	frm1.txtHCCNo.value = strRet(0)

	Dim strVal

	Call LayerShowHide(1)

	strVal = BIZ_PGM_CCQRY_ID & "?txtCCNo=" & Trim(frm1.txtHCCNo.value)	

	Call RunMyBizASP(MyBizASP, strVal)									

End Function
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables														
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										
    
    frm1.txtShipperNm.focus

End Sub

'=========================================================================================================
Sub txtIvDate_DblClick(Button)
    If Button = 1 Then
        frm1.txtIvDate.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtIvDate.Focus
    End If
    lgBlnFlgChgValue = True
End Sub
'=========================================================================================================

Sub txtLCDate_DblClick(Button)
    If Button = 1 Then
        frm1.txtLCDate.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtLCDate.Focus
    End If
    lgBlnFlgChgValue = True
End Sub
'=========================================================================================================
Sub txtSailing_DblClick(Button)
    If Button = 1 Then
        frm1.txtSailing.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtSailing.Focus
    End If
    lgBlnFlgChgValue = True
	End Sub
'=========================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'=========================================================================================================
 Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                   
End Function


'=========================================================================================================
Function BtnPrint() 
	Dim strUrl
	Dim ObjName
    
    If frm1.txtIvNo.value = "" Then
			Call DisplayMsgBox("205151",  "x", "통관", "x")
	Exit Function
	End If  
   
	Dim var1, var2, var3, var4, var5, var6, var7, var8, var9, var10, var11, var12, var13, var14, var15, var16, var17
	
	If Trim(frm1.txtShipperNm.value) = "" Then
		var1 = "%"	
	Else
		var1 = FilterVar(Trim(frm1.txtShipperNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtMessersNm.value) = "" Then
		var2 = "SAME AS ABOVE"	
	Else
		var2 = FilterVar(Trim(frm1.txtMessersNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtPartyNm.value) = "" Then
		var3 = "%"	
	Else
		var3 = FilterVar(Trim(frm1.txtPartyNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtLoadingPortNm.value) = "" Then
		var4 = "%"	
	Else
		var4 = FilterVar(Trim(frm1.txtLoadingPortNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtDestinationNm.value) = "" Then
		var5 = "%"	
	Else
		var5 = FilterVar(Trim(frm1.txtDestinationNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtCarrier.value) = "" Then
		var6 = "%"	
	Else
		var6 = FilterVar(Trim(frm1.txtCarrier.value), "" ,  "SNM")
	End If

	 	var7 = UniConvDateToYYYYMMDD(frm1.txtSailing.text,parent.gDateFormat,parent.gServerDateType)

	
	If Trim(frm1.txtIvNo.value) = "" Then
		var8 = "%"	
	Else
		var8 = FilterVar(Trim(frm1.txtIvNo.value), "" ,  "SNM")
	End If

	 	var9 = UniConvDateToYYYYMMDD(frm1.txtIvDate.text,parent.gDateFormat,parent.gServerDateType)

	
	If UCase(frm1.txtLCNo.value) = "" Then
		var10 = "%"	
	Else
		var10 = FilterVar(Trim(UCase(frm1.txtLCNo.value)), "" ,  "SNM")
	End If

	 	var11 = UniConvDateToYYYYMMDD(frm1.txtLCDate.text,parent.gDateFormat,parent.gServerDateType)

	If Trim(frm1.txtLCBankNm.value) = "" Then
		var12 = "%"	
	Else
		var12 = FilterVar(Trim(frm1.txtLCBankNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtRemarks1.value) = "" Then
		var13 = "%"	
	Else
		var13 = FilterVar(Trim(frm1.txtRemarks1.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtRemarks2.value) = "" Then
		var14 = "%"	
	Else
		var14 = FilterVar(Trim(frm1.txtRemarks2.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtRemarks3.value) = "" Then
		var15 = "%"	
	Else
		var15 = FilterVar(Trim(frm1.txtRemarks3.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtHCCNo.value) = "" Then
		var16 = "%"	
	Else
		var16 = FilterVar(Trim(frm1.txtHCCNo.value), "" ,  "SNM")
	End If

	strUrl = strUrl & "Shipper|" & var1
    strUrl = strUrl & "|AccMessers|" & var2 & "|Party|" & var3 & "|LoadPort|" & var4 & "|Destination|" & var5
    strUrl = strUrl & "|Carrier|" & var6 & "|Sailing|" & var7 & "|InvoiceNo|" & var8
    strUrl = strUrl & "|InvoiceDate|" & var9 & "|LCNo|" & var10 & "|LCDate|" & var11
    strUrl = strUrl & "|LCBank|" & var12 & "|Remarks1|" & var13 & "|Remarks2|" & var14 & "|Remarks3|" & var15
    strUrl = strUrl & "|CCNo|" & var16 


	ObjName = AskEBDocumentName("s4211oa1", "ebr")
	call FncEBRprint(EBAction, ObjName, strUrl)

		
End Function

'=========================================================================================================

Function BtnPreview() 

	Dim ObjName

    If frm1.txtIvNo.value = "" Then
		Call DisplayMsgBox("205151",  "x", "통관", "x")
	Exit Function
	End If

	Dim var1, var2, var3, var4, var5, var6, var7, var8, var9, var10, var11, var12, var13, var14, var15, var16, var17
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
		
	
	If Trim(frm1.txtShipperNm.value) = "" Then
		var1 = "%"	
	Else
		var1 = FilterVar(Trim(frm1.txtShipperNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtMessersNm.value) = "" Then
		var2 = "SAME AS ABOVE"	
	Else
		var2 = FilterVar(Trim(frm1.txtMessersNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtPartyNm.value) = "" Then
		var3 = "%"	
	Else
		var3 = FilterVar(Trim(frm1.txtPartyNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtLoadingPortNm.value) = "" Then
		var4 = "%"	
	Else
		var4 = FilterVar(Trim(frm1.txtLoadingPortNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtDestinationNm.value) = "" Then
		var5 = "%"	
	Else
		var5 = FilterVar(Trim(frm1.txtDestinationNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtCarrier.value) = "" Then
		var6 = "%"	
	Else
		var6 = FilterVar(Trim(frm1.txtCarrier.value), "" ,  "SNM")
	End If

	 	var7 = UniConvDateToYYYYMMDD(frm1.txtSailing.text,parent.gDateFormat,parent.gServerDateType)

	
	If Trim(frm1.txtIvNo.value) = "" Then
		var8 = "%"	
	Else
		var8 = FilterVar(Trim(frm1.txtIvNo.value), "" ,  "SNM")
	End If

	 	var9 = UniConvDateToYYYYMMDD(frm1.txtIvDate.text,parent.gDateFormat,parent.gServerDateType)

	
	If UCase(frm1.txtLCNo.value) = "" Then
		var10 = "%"	
	Else
		var10 = FilterVar(Trim(UCase(frm1.txtLCNo.value)), "" ,  "SNM")
	End If

	 	var11 = UniConvDateToYYYYMMDD(frm1.txtLCDate.text,parent.gDateFormat,parent.gServerDateType)

	If Trim(frm1.txtLCBankNm.value) = "" Then
		var12 = "%"	
	Else
		var12 = FilterVar(Trim(frm1.txtLCBankNm.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtRemarks1.value) = "" Then
		var13 = "%"	
	Else
		var13 = FilterVar(Trim(frm1.txtRemarks1.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtRemarks2.value) = "" Then
		var14 = "%"	
	Else
		var14 = FilterVar(Trim(frm1.txtRemarks2.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtRemarks3.value) = "" Then
		var15 = "%"	
	Else
		var15 = FilterVar(Trim(frm1.txtRemarks3.value), "" ,  "SNM")
	End If

	If Trim(frm1.txtHCCNo.value) = "" Then
		var16 = "%"	
	Else
		var16 = FilterVar(Trim(frm1.txtHCCNo.value), "" ,  "SNM")
	End If



	strUrl = strUrl & "Shipper|" & var1
    strUrl = strUrl & "|AccMessers|" & var2 & "|Party|" & var3 & "|LoadPort|" & var4 & "|Destination|" & var5
    strUrl = strUrl & "|Carrier|" & var6 & "|Sailing|" & var7 & "|InvoiceNo|" & var8
    strUrl = strUrl & "|InvoiceDate|" & var9 & "|LCNo|" & var10 & "|LCDate|" & var11
    strUrl = strUrl & "|LCBank|" & var12 & "|Remarks1|" & var13 & "|Remarks2|" & var14 & "|Remarks3|" & var15
    strUrl = strUrl & "|CCNo|" & var16 

    
		ObjName = AskEBDocumentName("s4211oa1", "ebr")
		Call FncEBRPreview(ObjName, strUrl)		
		
End Function
'=========================================================================================================
Function FncExit()
	FncExit = True
End Function
'=========================================================================================================

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Commercial Invoice출력</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenCCRef">통관참조</A></TD>
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
									<TD CLASS=TD5 NOWRAP>SHIPPER</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShipper" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="SHIPPER">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtShipperNm" ALT="SHIPPER" SIZE=20 MAXLENGTH=30 TAG="11"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>For Acc. Risk of Mssers.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMessers" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="FOR ACC RISK OF MESSERS">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT ALT="FOR ACC RISK OF MESSERS" NAME="txtMessersNm" SIZE=20 MAXLENGTH=30 TAG="11"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>NOTIFY PARTY</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtParty" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="NOTIFY PARTY">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPartyNm" ALT="NOTIFY PARTY" SIZE=20 MAXLENGTH=30 TAG="11"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>PORT OF LOADING</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoadingPort" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="PORT OF LOADING">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" ALT="PORT OF LOADING" SIZE=20 MAXLENGTH=50 TAG="11"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>DESTINATION</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDestination" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="DESTINATION">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtDestinationNm" ALT="DESTINATION" SIZE=20 MAXLENGTH=50 TAG="11"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>CARRIER</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCarrier" SIZE=18 MAXLENGTH=50 TAG="11XXXU" ALT="CARRIER"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>SAILING ON OR ABOUT</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211oa1_fpDateTime2_txtSailing.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>INVOICE NO</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=35 MAXLENGTH=25 TAG="14XXXU" ALT="INVOICE NO"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>INVOICE DATE</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211oa1_fpDateTime2_txtIvDate.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>L/C NO</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo" ALT="L/C NO" TYPE="Text" MAXLENGTH="25" SIZE=35 tag="14XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>L/C DATE</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211oa1_fpDateTime2_txtLCDate.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>L/C ISSUING BANK</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCBank" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="L/C ISSUING BANK"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCBank" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPop">&nbsp;<INPUT TYPE=TEXT NAME="txtLCBankNm" ALT="L/C ISSUING BANK" SIZE=20 MAXLENGTH=20 TAG="11"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>REMARKS</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemarks1" ALT="REMARKS" TYPE="Text" MAXLENGTH=60 SIZE=80  tag="11"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemarks2" ALT="REMARKS" TYPE="Text" MAXLENGTH=60 SIZE=80  tag="11"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemarks3" ALT="REMARKS" TYPE="Text" MAXLENGTH=60 SIZE=80  tag="11"></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> SRC= "../../blank.htm" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHCCNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLCNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHCCRef" tag="24">
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
