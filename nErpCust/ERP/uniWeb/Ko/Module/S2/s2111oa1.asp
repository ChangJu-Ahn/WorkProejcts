<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2111OA1
'*  4. Program Name         : 조직별 판매계획대실적출력 
'*  5. Program Desc         : 조직별 판매계획대실적출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/05/09
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Son Bum Yeol
'* 10. Modifier (Last)      : Son Bum Yeol
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

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop          

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
             
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
    frm1.txtCUR.value = Parent.gCurrency
    frm1.txtSALES_ORG.focus
	
End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
End Sub

'========================================================================================================= 
Function OpenConPop1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	<% '납품처 %>
	arrParam(0) = "영업조직"					    <%' 팝업 명칭 %>
	arrParam(1) = "B_SALES_ORG"							<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtSALES_ORG.value)		    <%' Code Condition%>
	arrParam(3) = ""                                  	<%' Name Cindition%>
	arrParam(4) = "END_ORG_FLAG=" & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG=" & FilterVar("Y", "''", "S") & " " <%' Where Condition%>	
	arrParam(5) = "영업조직"					    <%' TextBox 명칭 %>
		
	arrField(0) = "SALES_ORG"			                <%' Field명(0)%>
	arrField(1) = "SALES_ORG_NM"						<%' Field명(1)%>
	
	    
	arrHeader(0) = "영업조직"						<%' Header명(0)%>
	arrHeader(1) = "영업조직명"						<%' Header명(1)%>
	

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	frm1.txtSALES_ORG.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop1(arrRet)
	End If

End Function

'========================================================================================================= 
Function OpenConPop2()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	<% '납품처 %>
	arrParam(0) = "화폐"					    <%' 팝업 명칭 %>
	arrParam(1) = "B_CURRENCY"						<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtCUR.value)		    <%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""                                <%' Where Condition%>
	arrParam(5) = "화폐"					    <%' TextBox 명칭 %>
		
	arrField(0) = "CURRENCY"			            <%' Field명(0)%>
	arrField(1) = "CURRENCY_DESC"					<%' Field명(1)%>
	
	    
	arrHeader(0) = "화폐"						<%' Header명(0)%>
	arrHeader(1) = "화폐명"						<%' Header명(1)%>
	

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	frm1.txtCUR.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop2(arrRet)
	End If

End Function

'========================================================================================================= 
Function OpenConPop3()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	<% '납품처 %>
	arrParam(0) = "계획구분"					    <%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"								<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtPLAN_FLAG.value)		    <%' Code Condition%>
	arrParam(3) = Trim(frm1.txtPLAN_FLAG_NM.value)		<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("s4089", "''", "S") & ""                                    <%' Where Condition%>
	arrParam(5) = "계획구분"					    <%' TextBox 명칭 %>
		
	arrField(0) = "MINOR_CD"			                <%' Field명(0)%>
	arrField(1) = "MINOR_NM"						    <%' Field명(1)%>
	
	    
	arrHeader(0) = "계획구분"						<%' Header명(0)%>
	arrHeader(1) = "계획구분명"						<%' Header명(1)%>
	

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	frm1.txtPLAN_FLAG.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop3(arrRet)
	End If

End Function

'========================================================================================================= 
Function OpenConPop4()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	<% '납품처 %>
	arrParam(0) = "거래구분"					    <%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"								<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtEXPORT_FLAG.value)		<%' Code Condition%>
	arrParam(3) = Trim(frm1.txtEXPORT_FLAG_NM.value)	<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("S4225", "''", "S") & ""                    <%' Where Condition%>
	arrParam(5) = "거래구분"					    <%' TextBox 명칭 %>
		
	arrField(0) = "MINOR_CD"			                <%' Field명(0)%>
	arrField(1) = "MINOR_NM"						    <%' Field명(1)%>
	
	    
	arrHeader(0) = "거래구분"						<%' Header명(0)%>
	arrHeader(1) = "거래구분명"						<%' Header명(1)%>
	

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	frm1.txtEXPORT_FLAG.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop4(arrRet)
	End If

End Function

'========================================================================================================= 
Function OpenConPop5()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	<% '계획차수 %>
	arrParam(0) = "계획차수"					    <%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"					            <%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtPLAN_SEQ.value)			<%' Code Condition%>
	arrParam(3) = Trim(frm1.txtPLAN_SEQ_NM.value)		<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD='S2001'"                    <%' Where Condition%>
	arrParam(5) = "계획차수"					    <%' TextBox 명칭 %>
		
	arrField(0) = "MINOR_CD"			                <%' Field명(0)%>
	arrField(1) = "MINOR_NM"						    <%' Field명(1)%>
	
	    
	arrHeader(0) = "계획차수"						<%' Header명(0)%>
	arrHeader(1) = "계획차수명"						<%' Header명(1)%>
	

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	frm1.txtPLAN_SEQ.focus  
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop5(arrRet)
	End If

End Function

'========================================================================================================= 
Function SetConPop1(Byval arrRet)
	With frm1	
		.txtSALES_ORG.Value	= arrRet(0)
		.txtSALES_ORG_NM.Value= arrRet(1)
	End With

End Function

'========================================================================================================= 
Function SetConPop2(Byval arrRet)
	With frm1	
		.txtCUR.Value	= arrRet(0)

	End With

End Function

'========================================================================================================= 
Function SetConPop3(Byval arrRet)
	With frm1	
		.txtPLAN_FLAG.Value	= arrRet(0)
		.txtPLAN_FLAG_NM.Value= arrRet(1)
	End With

End Function

'========================================================================================================= 
Function SetConPop4(Byval arrRet)
	With frm1	
		.txtEXPORT_FLAG.Value	= arrRet(0)
		.txtEXPORT_FLAG_NM.Value= arrRet(1)
	End With

End Function

'========================================================================================================= 
Function SetConPop5(Byval arrRet)
	With frm1	
		.txtPLAN_SEQ.Value	= arrRet(0)
		.txtPLAN_SEQ_NM.Value= arrRet(1)
	End With

End Function

'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														'⊙: Load table , B_numeric_format

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
    <% '----------  Coding part  -------------------------------------------------------------%>
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
    
    frm1.txtSP_YEAR.Value = Year(UniConvDateToYYYYMMDD(EndDate,Parent.gDateFormat,Parent.gServerDateType))

End Sub

'========================================================================================================= 
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================= 
Sub txtSP_YEAR_onKeyPress()
	Call NumericCheck()
End Sub

Sub txtPLAN_SEQ_onKeyPress()
	Call NumericCheck()
End Sub

'========================================================================================================= 
Function NumericCheck()

	Dim objEl, KeyCode
	
	Set objEl = window.event.srcElement
	KeyCode = window.event.keycode

	Select Case KeyCode
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
	Case Else
		window.event.keycode = 0
	End Select

End Function


'========================================================================================================= 
 Function FncPrint() 
	Call parent.FncPrint()
End Function
'========================================================================================================= 
 Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================================================================================= 
Function FncQuery() 
FncQuery = true    
End Function

'========================================================================================================= 
Function BtnPrint() 

    Dim strUrl
	Dim ObjName    
	Dim var1, var2, var3,  var4, var5
	
	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

    <%'--출력조건을 지정하는 부분 수정 %>
	
	
	
    If Trim(frm1.txtSALES_ORG.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtSALES_ORG.value)), "" ,  "SNM")
	End If
    
	If Trim(frm1.txtSP_YEAR.value) = "" Then
		var2 = "%"
	Else
		var2 = UCase(frm1.txtSP_YEAR.value)
	End If
   
    
	 If Trim(frm1.txtPLAN_FLAG.value) = "" Then
		var3 = "%"
	Else
		var3 = FilterVar(Trim(UCase(frm1.txtPLAN_FLAG.value)), "" ,  "SNM")
	End If
	
	 If Trim(frm1.txtEXPORT_FLAG.value) = "" Then
		var4 = "%"
	Else
		var4 = FilterVar(Trim(UCase(frm1.txtEXPORT_FLAG.value)), "" ,  "SNM")
	End If
	
        var5 = FilterVar(Trim(UCase(frm1.txtPLAN_SEQ.value)), "" ,  "SNM")

	  
	
	<%'--출력조건을 지정하는 부분 수정 - 끝 %>
	
'    On Error Resume Next                                                    '☜: Protect system from crashing
    
    
    If frm1.Rb_WK1.checked = True Then
     

		<%'--출력조건을 지정하는 부분 수정 %>
		strUrl = strUrl & "SALES_ORG|" & var1
		strUrl = strUrl & "|SP_YEAR|" & var2
		strUrl = strUrl & "|PLAN_FLAG|" & var3
		strUrl = strUrl & "|EXPORT_FLAG|" & var4
		strUrl = strUrl & "|PLAN_SEQ|" & var5  

		<%'--출력조건을 지정하는 부분 수정 - 끝 %>
	
	'----------------------------------------------------------------
	' Print 함수에서 호출 
	'----------------------------------------------------------------
		ObjName = AskEBDocumentName("S2111OG1", "ebr")
		call FncEBRprint(EBAction, ObjName, strUrl)				
	'----------------------------------------------------------------

	
	ElseIf frm1.Rb_WK2.checked = True Then

		<%'--출력조건을 지정하는 부분 수정 %>
		strUrl = strUrl & "SALES_ORG|" & var1
		strUrl = strUrl & "|SP_YEAR|" & var2
		strUrl = strUrl & "|PLAN_FLAG|" & var3
		strUrl = strUrl & "|EXPORT_FLAG|" & var4
		strUrl = strUrl & "|PLAN_SEQ|" & var5  

		<%'--출력조건을 지정하는 부분 수정 - 끝 %>
	
	'----------------------------------------------------------------
	' Print 함수에서 호출 
	'----------------------------------------------------------------
		ObjName = AskEBDocumentName("S2111OG2", "ebr")
		call FncEBRprint(EBAction, ObjName, strUrl)	
	'----------------------------------------------------------------
 
    End If	
    
End Function

'========================================================================================================= 
Function BtnPreview() 


    Dim strUrl
    Dim ObjName
	Dim var1, var2, var3,  var4, var5	

    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If


<%
'	특수문자를 넘겨줄때는 아스키 코드값으로 변환을 해주어야 한다는군요 
'	"%" ---> %
'	""  ---> %32 로 바꾸어 주셔야 합니다.
'	아스키코드 25는 %이고 32는 space입니다.
'	SQL 7.0에서는 ""과 " "를 같이 인식하더군요 
%>
    If Trim(frm1.txtSALES_ORG.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtSALES_ORG.value)), "" ,  "SNM")
	End If
    
	If Trim(frm1.txtSP_YEAR.value) = "" Then
		var2 = "%"
	Else
		var2 = UCase(frm1.txtSP_YEAR.value)
	End If
   
    
	 If Trim(frm1.txtPLAN_FLAG.value) = "" Then
		var3 = "%"
	Else
		var3 = FilterVar(Trim(UCase(frm1.txtPLAN_FLAG.value)), "" ,  "SNM")
	End If
	
	 If Trim(frm1.txtEXPORT_FLAG.value) = "" Then
		var4 = "%"
	Else
		var4 = FilterVar(Trim(UCase(frm1.txtEXPORT_FLAG.value)), "" ,  "SNM")
	End If
	
        var5 = FilterVar(Trim(UCase(frm1.txtPLAN_SEQ.value)), "" ,  "SNM")

    
    If frm1.Rb_WK1.checked = True Then
		
	 
		strUrl = strUrl & "SALES_ORG|" & var1
		strUrl = strUrl & "|SP_YEAR|" & var2
		strUrl = strUrl & "|PLAN_FLAG|" & var3
		strUrl = strUrl & "|EXPORT_FLAG|" & var4
		strUrl = strUrl & "|PLAN_SEQ|" & var5	
  
		ObjName = AskEBDocumentName("S2111OG1", "ebr")
		Call FncEBRPreview(ObjName, strUrl)	
		
	ElseIf frm1.Rb_WK2.checked = True Then
	
		strUrl = strUrl & "SALES_ORG|" & var1
		strUrl = strUrl & "|SP_YEAR|" & var2
		strUrl = strUrl & "|PLAN_FLAG|" & var3
		strUrl = strUrl & "|EXPORT_FLAG|" & var4
		strUrl = strUrl & "|PLAN_SEQ|" & var5	

		ObjName = AskEBDocumentName("S2111OG2", "ebr")
		Call FncEBRPreview(ObjName, strUrl)	
        
    End If
    
End Function

'========================================================================================================= 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>조직별 판매계획대실적</font></td>
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
									<TD CLASS=TD5 NOWRAP>영업조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSALES_ORG" ALT="영업조직" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSALES_ORG" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop1()">&nbsp;<INPUT NAME="txtSALES_ORG_NM" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계획년도</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSP_YEAR" ALT="계획년도" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="12XXXU"></TD>
						        </TR>							
								<TR>
									<TD CLASS=TD5 NOWRAP>화폐</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCUR" ALT="화폐" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="14XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계획구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPLAN_FLAG" ALT="계획구분" TYPE="Text" MAXLENGTH="1" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPLAN_FLAG" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop3()">&nbsp;<INPUT NAME="txtPLAN_FLAG_NM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>거래구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEXPORT_FLAG" ALT="거래구분" TYPE="Text" MAXLENGTH="1" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEXPORT_FLAG" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop4()">&nbsp;<INPUT NAME="txtEXPORT_FLAG_NM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계획차수</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPLAN_SEQ" ALT="계획차수" TYPE="Text" MAXLENGTH="2" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btntxtPLAN_SEQ" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop5()">&nbsp;<INPUT NAME="txtPLAN_SEQ_NM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>집계형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK1 tag="12" Checked><LABEL FOR=Rb_WK1>품목그룹별</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2 tag="12"><LABEL FOR=Rb_WK2>품목별</LABEL>&nbsp;
														   
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX="-1">
    <input type="hidden" name="dbname" TABINDEX="-1">
    <input type="hidden" name="filename" TABINDEX="-1">
    <input type="hidden" name="condvar" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
<!-- End of Print HTML Code -->

</BODY>
</HTML>
