<!-- #Include file="../inc/IncServer.asp" -->
<!--#Include file="../inc/incServerAdoDb.asp" -->
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/ccm.vbs"></SCRIPT>

<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<TITLE>공지사항</TITLE>
<Script Language=vbscript>

	Function FncWrite()

		Dim arrRet

		If IsOpenPop = True Then Exit Function

		IsOpenPop = True	
		
		arrRet = window.showModalDialog("frwrite.asp?strMode=" & UID_M0001,Array(window.parent), _
			"dialogWidth=600px; dialogHeight=550px; center: Yes; help: No; resizable: No; status: No;")

		If arrRet = True Then
			MyBizASP.location.reload						
		End If
				
		IsOpenPop = False

	End Function

	Function FncModify()

		Dim arrRet
		
		If IsOpenPop = True Then Exit Function
		
		IsOpenPop = True
		
		arrRet = window.showModalDialog("frwrite.asp?strMode=" & UID_M0002 & "&intKeyNo=" & MyBizAsp.frTitle.intKeyNo, Array(window.parent), _
			"dialogWidth=600px; dialogHeight=550px; center: Yes; help: No; resizable: No; status: No;")	

		If arrRet = True Then
			'MyBizASP.frTitle.location.reload	' Title Refresh
			'MyBizASP.location.reload
			MyBizASP.frames("frTitle").document.URL = "frtitle.asp?page=" & CStr(MyBizAsp.frTitle.intNowPage)
		End If
		
		IsOpenPop = False

	End Function

	Function FncDelete()	
		
		If DisplayMsgBox("210034", VB_YES_NO, "x", "x") <> vbYes Then '삭제하시겠습니까?
		   Exit Function
		End If   	
	    
	    'MyBizASP.location.href = "frwriteBiz.asp?txtMode=" & UID_M0003  & "&txtKeyNo=" & MyBizAsp.frTitle.intKeyNo
	    MyBizASPForDelete.location.href = "frwriteBiz.asp?txtMode=" & UID_M0003  & "&txtKeyNo=" & MyBizAsp.frTitle.intKeyNo	    
	    
	End Function
	
	Public Function FncExit()
	    FncExit = True
	End Function

	Sub Form_Load()
	    gFocusSkip = True
	End Sub

</Script>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../CShared/image/table/seltab_up_left.gif"  width="9" height="23"></td>
								<td background="../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공지사항</font></td>
								<td background="../../CShared/image/table/seltab_up_right.gif" align="right"  width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* Align=right><A onclick="vbscript:FncWrite()">등록</A>&nbsp;|&nbsp;<A onclick="vbscript:FncModify()">수정</A>&nbsp;|&nbsp;
					<A onclick="vbscript:FncDelete()">삭제</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>			
	<TR>
		<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
			<IFRAME NAME="MyBizASP" SRC="Notice1.asp" WIDTH=100% HEIGHT=98% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
			<IFRAME NAME="MyBizASPForDelete" SRC="..\blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>