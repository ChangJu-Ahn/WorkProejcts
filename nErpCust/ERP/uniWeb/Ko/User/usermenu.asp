<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : User Menu
*  2. Function Name        : User Menu
*  3. Program ID           : UserMenu
*  4. Program Name         : UserMenu
*  5. Program Desc         : UserMenu
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     : 
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request.Cookies("unierp")("gLogoName")%> Menu</TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../inc/SCACheckUS.inc"  -->	
<!-- #Include file="../inc/IncServer.asp" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->

<Script Language="VBScript" SRC="../inc/Ccm.vbs">       </Script>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/Common.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/Event.vbs">     </SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/Operation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/Variables.vbs"> </SCRIPT>
<Script Language="VBScript" SRC="../inc/incUni2KTV.vbs"></Script>
<Script Language="JavaScript" SRC="../inc/incImage.js"> </SCRIPT>


<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'==========================================================================================================
Const BIZ_PGM_ID = "BizUserMenu.asp"
Const BIZ_PGM_USERMENU_ID = "LoadAllMenu.asp"

Const USR_MNU_NM  = "사용자 메뉴"
Const DEFOLDER ="새폴더"

Dim BIZ_DOC 
Dim BIZ_DOC_DOM
Dim USR_DOC 

Dim CliBizPath
Dim CliUsrPath

Function BizLoad(gCon)

	Dim Searchnode 
	Dim StrDate	
	Dim BizXmlDoc 
	Dim PI  
	Dim ChkXmlDoc

	StrDate = DBDateCheck(gusrid,gCon)
	ChkXmlDoc = frm2.uniXMLTree.GetBizXml(CliBizPath,strDate)
	
	if ChkXmlDoc = "1" then ' File Not found or Lower version
       BIZ_DOC =  BizXMLLoad(gusrid,glang,gCon)

       Set BIZ_DOC_DOM = CreateObject("MSXML2.DOMDocument")		
       BIZ_DOC_DOM.async = False 		
       BIZ_DOC_DOM.LoadXML(BIZ_DOC)
	Else 
	     Biz_DOC =""
	End If

End Function 

Function UsrLoad(gCon)
		
	USR_DOC =  UsrXMLLoad(gusrid,glang,gCon)
			    
End function 
    
Function DBDateCheck(UID,gCon) 

	Dim httpDate
	Set httpDate = createObject("MSXML2.XMLHTTP")			
	
	httpDate.open "post", "MenuLoad.asp", false	
	httpDate.setRequestHeader "SOAPAction","GetDate"
	httpDate.setRequestHeader "uid",UID
	httpDate.setRequestHeader "Connect",gCon

	httpDate.send
	
	DBDateCheck =  httpDate.responseText
		
	Set httpDate = Nothing 
		
	    
End Function
	
'Loading and save

Function BizXMLLoad(UID,LANG,gCon)
		
	Dim httpBiz
	Set httpBiz = createObject("MSXML2.XMLHTTP")			
		
	httpBiz.open "post", "MenuLoad.asp", false	
	httpBiz.setRequestHeader "SOAPAction","BIZ"
	httpBiz.setRequestHeader "uid",UID
	httpBiz.setRequestHeader "lang",gLANG
	httpBiz.setRequestHeader "Connect",gCon
		
	httpBiz.send
	
	BizXMLLoad =  httpBiz.responseText
		
	Set httpBiz = Nothing 
		
End Function

Function UsrXMLLoad(UID,LANG,gCon)
		
	Dim httpDoc		
	Set httpDoc = createObject("MSXML2.XMLHTTP")			
		
	httpDoc.open "post", "MenuLoad.asp", false	
	httpDoc.setRequestHeader "SOAPAction","USR"
	httpDoc.setRequestHeader "uid",UID
	httpDoc.setRequestHeader "lang",gLANG
	httpDoc.setRequestHeader "Connect",gCon
		
	httpDoc.send
		
	UsrXMLLoad =  httpDoc.responseText
		
	Set httpDoc = Nothing 
		
End Function 
	
Sub MnuFileNm()
	
	Dim CliTmp
	Call GetGlobalVar()
	
	CliTmp = frm2.uniXMLTree.GetSysTmp

	CliBizPath =  CliTmp & gCompany + "_" & gDBServer + "_" + gDatabase + "_" + gLang +"_" + gUsrId + "_Bizmnu.xml"
	CliUsrPath =  CliTmp & gCompany + "_" & gDBServer + "_" + gDatabase + "_" + gLang + "_" + gUsrId + "_Usrmnu.xml"
	
	
End sub


Sub Form_Load()
	
	Dim NodX, lHwnd
	
	Call GetGlobalVar
	
	
	frm2.uniXMLTree.HideSelection = false
	
	Call MnuFileNm()
	Call BizLoad(gADODBConnString)' Biz Menu Loading
	Call UsrLoad(gADODBConnString)' Usr Menu Loading  
		
	With frm2
		
		.uniXMLTree.Usr_Mnu_Init = USR_MNU_NM
		.uniXMLTree.XMLBizPath = CliBizPath
		.uniXMLTree.XMLUsrPath = CliUsrPath
		
		If glang ="JA" Then 			
			.uniXMLTree.SetTvCharSet =128
		End If
		
		'Option
		'.uniXMLTree.settvfont ="MS Sans Serif"
		'.uniXMLTree.SetTvFontSize ="10"
		'Option
		
		.uniXMLTree.XMLBizTreeDoc_Init = BIZ_DOC
		.uniXMLTree.XMLUsrTreeDoc_Init = USR_DOC		
		.uniXMLTree.Default_Folder = DEFOLDER
		.uniXMLTree.XMLColSep = gColSep
		.uniXMLTree.XMLRowSep = gRowSep
		.uniXMLTree.SetMoveMsg = "해당 위치로는 이동할 수 없습니다."
		.uniXMLTree.SetProductNm = "uniERP"
		
		.uniXMLTree.OpenTitle   = "열기"
		.uniXMLTree.CloseTitle   = "닫기"
		.uniXMLTree.AddTitle    = "폴더 추가"
		.uniXMLTree.DeleteTitle = "삭제"
		.uniXMLTree.RenameTitle = "이름 변경"		
		.uniXMLTree.SetTvFont     = gFontName   'new 20050111
		.uniXMLTree.SetTvFontSize = gFontSize   'new 20050111
		.uniXMLTree.SetTvAppearance = 0
		
		.uniXMLTree.Make_Tree
		
		
	End With 
	
End Sub

Function button1_onclick()	     

	Call parent.frToolbar.OpenUserMenu()

End Function

Sub uniXMLTree_DblClick(nodeX)
    
    Dim GoCode
    DIm SNode
    
    Dim uID
    Dim mID
    Dim mIDNM
    
    GoCode = Split(nodex.key,"^")
    
    mID = GoCode(0)    

    Set Snode = BIZ_DOC_DOM.selectSingleNode("/BMNU/row[@MNU_ID='" & mID & "']")
	
    If TypeName(Snode) = "Nothing" Then
       Call parent.frToolbar.DBGo(mID,True)
       Exit Sub
    End If

    uID   = Snode.getAttribute("UPPER_MNU_ID")
    mIDNM = Snode.getAttribute("MNU_NM")
    
    Call parent.frToolbar.DBGo(mID,True)
    
End Sub

Sub uniXMLTree_Dbsave(StrVal)
	
	'msgbox strval 
	frm2.txtMode.value = UID_M0002
	frm2.txtMulti.value = StrVal
	
	Call ExecMyBizASP(frm2, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
End Sub 

Sub AddMenuPopup(StrMod , StrID , StrUppId, StrNm , StrType , StrLvl , StrSeq )
	
	frm2.uniXMLTree.AppUsrMenu StrMod,StrID,StrUppId,StrNm,StrType,StrLvl,StrSeq
	
End Sub 

Sub RefreshUsrXml()
	
	Call UsrLoad(gADODBConnString)' Usr Menu Loading  
	Frm2.uniXMLTree.XMLUsrTreeDoc_Init = USR_DOC
	frm2.uniXMLTree.Make_Tree(True)		
	
End Sub 
</SCRIPT>

<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<BODY SCROLL=no rightmargin=0 bgColor="#FFFFFF">
<FORM NAME=frm2 TARGET="MyBizASP2" METHOD="POST">
<TABLE HEIGHT=100% WIDTH=100%  cellspacing= 0 cellpadding= 0 >
	<TR STYLE='BACKGROUND-IMAGE: url(../../CShared/Image/left_mn_bg.gif);'>
		<TD WIDTH=100% HEIGHT=36 COLSPAN=3>
			<TABLE WIDTH=100% border=0 cellspacing=0 cellpadding=0 height=100%>
				<TR >
				    <% If gUserIdKind = "U" Then  %>
					<TD WIDTH="*">&nbsp;&nbsp;<img src=../image/bi.jpg border=0 ></TD>
					<% Else %>
					<TD WIDTH="*">&nbsp;&nbsp;<img src=../image/biC.jpg border=0 ></TD>
					<% End If %>
					<TD STYLE="TEXT-ALIGN: right; TEXT-VLIGN: center">
					<IMG SRC="../../CShared/Image/x.gif" NAME=tbHide alt="숨기기" CLASS="enableIMG" style='cursor:hand' onclick="vbscript:button1_onclick" language="javascript" onMouseUp="javascript:MM_swapImage('tbHide','','../../CShared/Image/x.gif',1)" onMouseDown="javascript:MM_swapImage('tbHide','','../../CShared/Image/x_dn.gif',1)"></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	
	<TR>
		<TD HEIGHT=* WIDTH=4 >
		<TD HEIGHT=* WIDTH=100% style='    BORDER-RIGHT: #dddfe6 1px solid;    BORDER-TOP: #dddfe6 1px solid;    BORDER-LEFT: #dddfe6 1px solid;    BORDER-BOTTOM: #dddfe6 1px solid;    BACKGROUND-COLOR: #dddfe6    '>
			<script language =javascript src='./js/usermenu_uniXMLTree_N186801648.js'></script>
		</TD>
		<TD HEIGHT=* WIDTH=5 Class="Toolbar2">
	</TR>
	<TR Class="Toolbar">
		<TD HEIGHT=1 WIDTH=100% COLSPAN=3>
		<IFRAME ID="MyBizASP2" NAME="MyBizASP2" SRC='../blank.htm' width=100% height=0 STYLE="display: ''"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<TEXTAREA NAME=txtMulti STYLE="display: none">
</TEXTAREA>
</FORM>
</BODY>
</HTML>

