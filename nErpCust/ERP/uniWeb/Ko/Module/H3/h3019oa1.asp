<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--
======================================================================================================
'*  1. Module명          : 인사/급여관리 
'*  2. Function명        : 근무이력관리 
'*  3. Program ID        : h3019oa1.asp
'*  4. Program 이름      : 인원현황출력 
'*  5. Program 설명      : 인원현황출력 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2001/05/23
'*  8. 최종 수정년월일   : 2004/06/10
'*  9. 최초 작성자       : Shin Kwang- Ho/mok young bin
'* 10. 최종 작성자       : Lee SiNa
'* 11. 전체 comment      :
=======================================================================================================-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
Dim lsInternal_cd

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    
End Sub

Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtBas_dt.focus

	frm1.txtBas_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtBas_dt.Month = strMonth 
	frm1.txtBas_dt.Day = strDay	
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "OA") %>
End Sub

Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

	Call ggoOper.FormatDate(frm1.txtBas_dt, parent.gDateFormat, 1)

    Call InitVariables  
    Call FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
    Call SetToolbar("10000000000011")
    
    Set gActiveElement = document.activeElement	
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub



'========================================================================================================
' Name : txtBas_dt_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
' 날짜에서 엔터키 입력시 미리보기 실행 
Sub txtBas_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtBas_dt.Action = 7
		frm1.txtBas_dt.focus
	End If
End Sub

Function FncQuery()
	' 엔터키 입력시 미리보기 실행 
    FncBtnPreview()
End Function

Function SetPrintCond(StrEbrFile, strUrl)

	Dim ObjName
	Dim bas_dt , basyymmdd
    
    SetPrintCond = False
    
	bas_dt = UniConvDateAToB(frm1.txtBas_dt.Text, parent.gDateFormat, parent.gAPDateFormat)
	
	If frm1.txtPrt_sel(0).checked Then
        StrEbrFile = "h3019oa1_1"
    ElseIf frm1.txtPrt_sel(1).checked Then
        StrEbrFile = "h3019oa1_2"
    Elseif frm1.txtPrt_sel(2).checked Then 
		StrEbrFile = "h3019oa1_3"
    Elseif frm1.txtPrt_sel(3).checked Then 
		StrEbrFile = "h3019oa1_4"
    Elseif frm1.txtPrt_sel(4).checked Then 
    	StrEbrFile = "h3019oa1_5"
    Else
    	StrEbrFile = "h3019oa1_6"
    End If

    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Call BtnDisabled(0)
	   Exit Function
    End If    

    basyymmdd = frm1.txtBas_dt.year & Right("0" & frm1.txtBas_dt.month , 2) & Right("0" & frm1.txtBas_dt.day , 2)

	Call BtnDisabled(1)
	
	strUrl = "bas_dt|" & basyymmdd

	SetPrintCond = True

End Function

Function FncBtnPrint() 

    Dim StrEbrFile, strUrl
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
	
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
	call FncEBCPrint(EBAction,ObjName,strUrl)	

End Function

Function FncBtnPreview() 
    
    Dim StrEbrFile, strUrl

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If

	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
	call FncEBCPreview(ObjName , strUrl)
	
End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

Function FncExit()
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

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
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>인원현황출력</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>									
							<TR>
								<TD CLASS=TD5  NOWRAP>기준일</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h3019oa1_fpDateTime1_txtBas_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>출력선택</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtPrt_sel VALUE = "1" ID=Rb_tot Checked><LABEL FOR=Rb_tot>근속년수</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtPrt_sel VALUE = "2" ID=Rb_dur><LABEL FOR=Rb_dur>연령별</LABEL></TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtPrt_sel VALUE = "3" ID=Rb_dept><LABEL FOR=Rb_dept>부서인원</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtPrt_sel VALUE = "4" ID=Rb_dept1><LABEL FOR=Rb_dept>부서/학력</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtPrt_sel VALUE = "5" ID=Rb_dept2><LABEL FOR=Rb_dept>직급/학력</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtPrt_sel VALUE = "6" ID=Rb_ocpt><LABEL FOR=Rb_ocpt>직종별</LABEL></TD>
							</TR>	
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
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
		<TD >
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="strUrl" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1" >	
</FORM>
</BODY>
</HTML>

