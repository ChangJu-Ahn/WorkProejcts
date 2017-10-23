<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Common Module
*  2. Function Name        : Common Function
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
<TITLE>- SQL 구문</TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================-->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================-->
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">
   
Sub Form_Load()
	Call MM_preloadImages("../../CShared/image/Query.gif","../../CShared/image/OK.gif","../../CShared/image/Cancel.gif")
	document.frm.STRSQL.value =  "SELECT " & opener.document.frm1.txtSelect.value & vbCrlf
	
	If Trim(opener.document.frm1.txtFrom.value) <> "" Then
		document.frm.STRSQL.value =  document.frm.STRSQL.value & "FROM " & opener.document.frm1.txtFrom.value & vbCrlf
	End If
	
	If Trim(opener.document.frm1.txtWhere.value) <> "" Then
		document.frm.STRSQL.value =  document.frm.STRSQL.value & "WHERE " & opener.document.frm1.txtWhere.value & vbCrlf
	End If
	
	If Trim(opener.document.frm1.txtEtc.value) <> "" Then
		document.frm.STRSQL.value =  document.frm.STRSQL.value & "" & opener.document.frm1.txtEtc.value & vbCrlf
	End If	
	
End Sub

Function CancelClick()
	Self.Close()
End Function
    
</SCRIPT>
<SCRIPT LANGUAGE="JAVASCRIPT">
function ViewSQL()
{
	clipboardData.setData("Text",eval("document.frm.STRSQL.value"));
	alert("Clipboard에 복사되었습니다.");
}

</SCRIPT>
</HEAD>
<BODY SCROLL=NO TABINDEX="-1" ONLOAD="Form_Load()">
<form name=frm>
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD STYLE="PADDING-RIGHT: 5px;WIDTH: 14%;BACKGROUND-COLOR: #e7e5ce;" ALIGN=LEFT>&nbsp;SQL 구문</SPAN></TD>
			</TR>
		</TABLE>
	</TD>
	</TR>
	<TR HEIGHT=* valign="top">
		<TD>
		<TEXTAREA NAME="STRSQL" style="width=100%;border=0;height=100%"></TEXTAREA>
		</TD>
	</TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<BUTTON NAME="btnExeStdCost" CLASS="CLSSBTN" onclick="ViewSQL()" Flag=1>Copy</BUTTON>				</TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
</TABLE>
</form>
</BODY>
</HTML>


