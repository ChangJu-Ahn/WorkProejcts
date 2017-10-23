<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incServer.asp"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/CommStyleSheet.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1804mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1      = "e1803ma1.asp"

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXCOLS = 9

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
<!-- #Include file="../../inc/incGrid.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim Grid1

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    if  pOpt = "Q" then
        lgKeyStream       = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtLang_cd.Value) & gColSep
    else
        lgKeyStream       = Trim(frm1.txtEmp_no.Value) & gColSep
    end if
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
'  Call CommonQueryRs(" rTrim(LANG_CD),LANG_NM "," B_LANGUAGE "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'    iCodeArr = lgF0
'    iNameArr = lgF1
    iCodeArr = gLang & chr(11) 
    iNameArr = gLang & chr(11) 
    Call SetCombo2(frm1.txtlang_cd, iCodeArr, iNameArr, Chr(11))    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = 10
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call InitComboBox()
    Call LockField(Document)	
    Call LayerShowHide(0)

    Call InitGrid()

    Call SetToolBar("10000")

'    frm1.txtlang_cd.value = parent.txtLang.value
'    frm1.txtlang_cd.focus()
    Call DbQuery(1)
End Sub
'========================================================================================
' Function Name : Window_onUnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing
    Set Grid1 = Nothing
End Sub

Function DbQuery(ppage)

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    'If Grid1.ChkChange() Then Exit Function
    Call ClearField(document,2)
'    Call ElementVisible(window.parent.document.all("RunQuery"), 0)


    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    frm1.GRID_PAGE.VALUE = ppage

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)


End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status

'    Call ElementVisible(window.parent.document.all("RunQuery"), 0)


End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel

	Dim strRes_no

    DbSave = False                                                          
    
    Call LayerShowHide(1)

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
       For lRow = 1 To Grid1.SheetMaxrows
           Select Case document.all(CStr(lRow)).value
               Case UpdateFlag                                      '☜: Update
                    strVal = strVal & "U" & gColSep
                    strVal = strVal & lRow & gColSep
                    strVal = strVal & document.all("SPREADCELL_EMP_NO" & CStr(lRow)).value  & gColSep
                    strVal = strVal & document.all("SPREADCELL_DILIG_STRT_DT" & CStr(lRow)).value  & gColSep
                    strVal = strVal & document.all("SPREADCELL_DILIG_CD" & CStr(lRow)).value  & gColSep
                    strVal = strVal & document.all("SPREADCELL_APP_YN" & CStr(lRow)).value  & gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & gColSep
                                                  strDel = strDel & lRow & gColSep
                                                  strDel = strDel & .txtEmp_no.value & gColSep
                    '.vspdData.Col = C_FAMILY_NM	: strDel = strDel & Trim(.vspdData.Text) & gColSep
                    '.vspdData.Col = C_REL_CD	: strDel = strDel & Trim(.vspdData.Text) & gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = "UID_M0002"
       .txtUpdtUserId.value  = .txtEmp_no.value
       .txtInsrtUserId.value = .txtEmp_no.value
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Dim curpage

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    'Call InitVariables
    'curpage = frm1.GRID_PAGE.VALUE

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call DbQuery(frm1.GRID_PAGE.VALUE)

    'frm1.GRID_PAGE.VALUE = curpage
    'Call ShowData(Source,Source.GRID_PAGE.Value-1)
End Function


Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
End Function


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Function DoubleGetRow(pRow)
    Dim objList
    Dim elmCnt

    Dim menu_id
    Dim lang_cd
    Dim strVal

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    menu_id = ""
    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_menu_id" & pRow then
               menu_id = objList.value
    		End if
    		If objList.name = "SPREADCELL_lang_cd" & pRow then
               lang_cd = objList.value
    		End if
    	Next
    End With

    If  menu_id <> "" then
        strVal = BIZ_PGM_ID1 & "?menu_id=" & menu_id
        strVal = strVal & "&lang_cd=" & lang_cd
		Call CommonQueryRs(" MENU_NAME "," E11000T "," Menu_id = " & FilterVar("E1803MA1", "''", "S") & " AND LANG_CD =  " & FilterVar(gLang , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		parent.txtTitle.value = replace(lgF0,chr(11),"")
        document.location = strVal

    end if

	DoubleGetRow = True
End Function

Sub MouseRow(pRow)
	If frm1.grid_totpages.value = "" Then Exit Sub
    Dim objList   

	Grid1.ActiveRow = pRow	
	Set objList = window.event.srcElement	
	
	If  UCase(objList.getAttribute("flag")) = "SPREADCELL" then
        if objList.value = "" then            
             objList.style.cursor = "auto"
        else
             objList.style.cursor = "hand"
        end if
    End If        

End Sub
'========================================================================================================

Sub Query_OnClick()
    Call DbQuery()
End Sub

Sub Print_onClick()
    Call SubPrint(MyBizASP)
End Sub


Sub GRID_PAGE_OnChange()
End Sub

Sub DELETE_OnClick()
    Call Grid1.DeleteClick()
End Sub

Sub CANCEL_OnClick()
    Call Grid1.CancelClick()
End Sub

</SCRIPT>

<!-- #Include file="../../inc/uniSimsClassID.inc" --> 


<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
</HEAD>
<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR height=35 valign=middle>
	    	<TD class=base1 colspan=2>언어:<SELECT NAME="txtlang_cd" ALT=언어 STYLE="WIDTH: 150px" TAG="12"></SELECT></TD>
	    	<TD></TD>
        </TR>

        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=680 border=0 bgcolor=#ffffff>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 border=0 bgcolor=#ffffff>
                                <TR bgcolor=#d0d6e4 height=20>
		                        	<TD></TD>
		                        	<TD class=TDFAMILY_TITLE1>언어</TD>
		                        	<TD class=TDFAMILY_TITLE1>메뉴ID</TD>
		                        	<TD class=TDFAMILY_TITLE1>메뉴명</TD>
		                        	<TD class=TDFAMILY_TITLE1>프로그램</TD>
		                            <TD class=TDFAMILY_TITLE1>메뉴타입</TD>
		                            <TD class=TDFAMILY_TITLE1>사용</TD>
		                        	<TD class=TDFAMILY_TITLE1>레벨</TD>
		                        	<TD class=TDFAMILY_TITLE1>상위</TD>
                                </TR>
<%            
        For i=1 To 10
            Response.Write "<TR bgcolor=#E9EDF9 height=20 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<TD><INPUT name='" & i & "'  tag='25X' flag='SPREADCELL' style='WIDTH: 30px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_lang_cd" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 60px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_menu_id" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 100px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_menu_name" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 200px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_href" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 95px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_menu_level" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 80px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH: 58px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH: 78px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH: 58px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "</TR>"
        Next
%>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>
        <TR height=20>
            <TD width=13></TD>
            <TD VALIGN=center ALIGN=right>
                        <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../../../Cshared/Image/uniSIMS/gprev.jpg border=0 ></A>&nbsp;
                        <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../../../Cshared/Image/uniSIMS/gnext.jpg border=0 ></A>&nbsp;&nbsp;
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 width=500 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=auto noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">
    <INPUT TYPE=hidden NAME="txtMaxRows"> 
    <TEXTAREA style="display: none" name=txtSpread></TEXTAREA>

    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
</FORM>	

</BODY>
</HTML>
