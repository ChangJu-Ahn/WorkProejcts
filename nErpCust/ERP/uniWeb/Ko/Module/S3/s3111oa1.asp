<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3111oa1.asp	
'*  4. Program Name         : 수주대장출력 
'*  5. Program Desc         : 수주대장출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Dim IsOpenPop          

'===========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size             
End Sub

'===========================================================================================================
Sub SetDefaultVal()
	frm1.txtSales_Grp.focus 
	frm1.txtSOFromDt.Text = StartDate
	frm1.txtSOToDt.Text = EndDate	
End Sub

'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'===========================================================================================================
Function OpenConPop1()
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

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop1(arrRet)
	End If

End Function

'===========================================================================================================
Function OpenConPop2()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

		arrParam(0) = "주문처"							
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtSold_to_party.value)		
		arrParam(3) = Trim(frm1.txtSold_to_partyNm.value)	
		arrParam(4) = "BP_TYPE <= " & FilterVar("CS", "''", "S") & ""						
		arrParam(5) = "주문처"	                        
		
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
	   								
		arrHeader(0) = "주문처"							
	    arrHeader(1) = "주문처명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

    frm1.txtSold_to_party.focus

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop2(arrRet)
	End If

End Function

'===========================================================================================================
Function OpenConPop3()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

		arrParam(0) = "수주형태"							
		arrParam(1) = "S_SO_TYPE_CONFIG"						
		arrParam(2) = Trim(frm1.txtSo_Type.value)		
		arrParam(3) = Trim(frm1.txtSo_Type_Nm.value)	
		arrParam(4) = ""						
		arrParam(5) = "수주형태"	                        
		
		arrField(0) = "SO_TYPE"								
		arrField(1) = "SO_TYPE_NM"								
	   								
		arrHeader(0) = "수주형태"							
	    arrHeader(1) = "수주형태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

    frm1.txtSo_Type.focus 

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop3(arrRet)
	End If

End Function


'===========================================================================================================
Function SetConPop1(Byval arrRet)
	With frm1	
		.txtSales_Grp.Value		= arrRet(0)
		.txtSales_Grp_Nm.Value	= arrRet(1)
	End With
End Function

'===========================================================================================================
Function SetConPop2(Byval arrRet)
	With frm1	
		.txtSold_to_party.Value		= arrRet(0)
		.txtSold_to_partyNm.Value	= arrRet(1)
	End With
End Function

'===========================================================================================================
Function SetConPop3(Byval arrRet)
	With frm1	
		.txtSo_Type.Value		= arrRet(0)
		.txtSo_Type_Nm.Value	= arrRet(1)
	End With
End Function

'===========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
    <% '----------  Coding part  -------------------------------------------------------------%>
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
End Sub

'===========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'===========================================================================================================
Sub txtSOFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSOFromDt.Action = 7
    End If
End Sub

Sub txtSOToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSOToDt.Action = 7
    End If
End Sub

'===========================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'===========================================================================================================
 Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                  
End Function

'===========================================================================================================
Function FncQuery() 
 FncQuery = true   
End Function

'===========================================================================================================
Function BtnPrint() 
	Dim strUrl
	
	If ValidDateCheck(frm1.txtSOFromDt, frm1.txtSOToDt) = False Then Exit Function
    
    If Not chkField(Document, "1") Then								
       Exit Function
    End If

    '--출력조건을 지정하는 부분 수정 
    dim var1, var2 ,var3, var4,var5	
	
	If UCase(frm1.txtSales_Grp.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtSales_Grp.value)), "" ,  "SNM")
	End If
	
	If UCase(frm1.txtSold_to_party.value) = "" Then
		var2 = "%"
	Else
		var2 = FilterVar(Trim(UCase(frm1.txtSold_to_party.value)), "" ,  "SNM")
	End If
    
    If UCase(frm1.txtSo_Type.value) = "" Then
		var3 = "%"
	Else
		var3 = FilterVar(Trim(UCase(frm1.txtSo_Type.value)), "" ,  "SNM")
	End If

 	var4 = UniConvDateToYYYYMMDD(frm1.txtSOFromDt.Text,parent.gDateFormat, parent.gServerDateType)
	
	var5 = UniConvDateToYYYYMMDD(frm1.txtSOToDt.Text,parent.gDateFormat, parent.gServerDateType)
	
	'--출력조건을 지정하는 부분 수정 - 끝 
	
	strUrl = strUrl & "SALES_GRP|" & var1
	strUrl = strUrl & "|SOLD_TO_PARTY|" & var2
	strUrl = strUrl & "|SO_TYPE|" & var3
	strUrl = strUrl & "|AFromSODt|" & var4
	strUrl = strUrl & "|AToSODt|" & var5 


'----------------------------------------------------------------
' Print 함수에서 호출 
'----------------------------------------------------------------
	ObjName = AskEBDocumentName("s3111oa1","ebr")
	
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------
		
End Function


'===========================================================================================================
Function BtnPreview() 

	If ValidDateCheck(frm1.txtSOFromDt, frm1.txtSOToDt) = False Then Exit Function

    If Not chkField(Document, "1") Then								
       Exit Function
    End If

	

	Dim var1, var2, var3, var4, var5
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
		
	If UCase(frm1.txtSales_Grp.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtSales_Grp.value)), "" ,  "SNM")
	End If
	
	If UCase(frm1.txtSold_to_party.value) = "" Then
		var2 = "%"
	Else
		var2 = FilterVar(Trim(UCase(frm1.txtSold_to_party.value)), "" ,  "SNM")
	End If
    
    If UCase(frm1.txtSo_Type.value) = "" Then
		var3 = "%"
	Else
		var3 = FilterVar(Trim(UCase(frm1.txtSo_Type.value)), "" ,  "SNM")
	End If

 	var4 = UniConvDateToYYYYMMDD(frm1.txtSOFromDt.Text,parent.gDateFormat,parent.gServerDateType)
	
	var5 = UniConvDateToYYYYMMDD(frm1.txtSOToDt.Text,parent.gDateFormat,parent.gServerDateType)

	strUrl = strUrl & "SALES_GRP|" & var1
	strUrl = strUrl & "|SOLD_TO_PARTY|" & var2
	strUrl = strUrl & "|SO_TYPE|" & var3
	strUrl = strUrl & "|AFromSODt|" & var4
	strUrl = strUrl & "|AToSODt|" & var5 
	
	ObjName = AskEBDocumentName("s3111oa1","ebr")
	
	Call FncEBRPreview(ObjName, strUrl)	
		
End Function

'===========================================================================================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수주대장출력</font></td>
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
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Grp" ALT="영업그룹" TYPE="Text" MAXLENGTH=4 SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSales_Grp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop1()">&nbsp;<INPUT NAME="txtSales_Grp_Nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
								</TR>
								
								<TR>
									<TD CLASS=TD5 NOWRAP>주문처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSold_to_party" ALT="주문처" TYPE="Text" MAXLENGTH=10 SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSold_to_party" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop2()">&nbsp;<INPUT NAME="txtSold_to_partyNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>수주형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSo_Type" ALT="수주형태" TYPE="Text" MAXLENGTH=4 SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSo_Type" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop3()">&nbsp;<INPUT NAME="txtSo_Type_NM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								 <TR>
									<TD CLASS="TD5" NOWRAP>수주일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s3111oa1_fpDateTime1_txtSOFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s3111oa1_fpDateTime2_txtSOToDt.js'></script>
												</TD>
								            </TR>
						               </TABLE>
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
						    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON></TD>
						</TD>
					</TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD valign=top>
		    
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                        FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1" ></IFRAME>
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
