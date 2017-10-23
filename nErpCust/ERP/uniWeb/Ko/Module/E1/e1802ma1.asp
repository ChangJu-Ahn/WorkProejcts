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

Const BIZ_PGM_ID      = "e1802mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1      = "e1801ma1.asp"

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
        if frm1.txtuse_yn0.checked = true then
            lgKeyStream = lgKeyStream & "" & gColSep
        elseif frm1.txtuse_yn1.checked = true then
            lgKeyStream = lgKeyStream & "Y" & gColSep
        else
            lgKeyStream = lgKeyStream & "N" & gColSep
        end if
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


    Call LayerShowHide(0)

    Call InitGrid()

    Call SetToolBar("10000")
    Call LockField(Document)

    Call DbQuery(1)
End Sub
'========================================================================================
' Function Name : Window_onUnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
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
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
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

    Dim emp_no
    Dim strVal

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    emp_no = ""
    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_emp_no" & pRow then
               emp_no = objList.value
    		End if
    	Next
    End With

    If  emp_no <> "" then
        strVal = BIZ_PGM_ID1 & "?emp_no=" & emp_no
		strVal = strVal& "&updateok=ok"
		Call CommonQueryRs(" MENU_NAME "," E11000T "," Menu_id = " & FilterVar("E1801MA1", "''", "S") & " AND LANG_CD =  " & FilterVar(gLang , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
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
        <TR height=45 valign=middle>
            <TD class=base1 colspan=2>사용여부:
                <INPUT TYPE="RADIO" NAME="txtuse_yn" tag="12" CHECKED ID="txtuse_yn0" VALUE=A STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #FFFFFF"><LABEL FOR="txtuse_yn0">전체</LABEL>
    	        <INPUT TYPE="RADIO" NAME="txtuse_yn" tag="12" ID="txtuse_yn1" VALUE=Y STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #FFFFFF"><LABEL FOR="txtuse_yn1">사용</LABEL>
                <INPUT TYPE="RADIO" NAME="txtuse_yn" tag="12" ID="txtuse_yn2" VALUE=N STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #FFFFFF"><LABEL FOR="txtuse_yn2">미사용</LABEL>
            </TD>
        </TR>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=680 border=0 bgcolor=#ffffff>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width="100%" border=0 bgcolor=#ffffff>
                                <TR bgcolor=#d0d6e4 height=20>
		                        	<TD></TD>
		                        	<TD class=TDFAMILY_TITLE1>사용자ID</TD>
		                        	<TD class=TDFAMILY_TITLE1>성명</TD>
		                            <TD class=TDFAMILY_TITLE1>사번</TD>
		                            <TD class=TDFAMILY_TITLE1>부서명</TD>
		                        	<TD class=TDFAMILY_TITLE1>레벨</TD>
		                        	<TD class=TDFAMILY_TITLE1>자료권한</TD>
		                        	<TD class=TDFAMILY_TITLE1>사용</TD>
		                        	<TD class=TDFAMILY_TITLE1>등록일</TD>
		                        	<TD class=TDFAMILY_TITLE1></TD>
                                </TR>
<%            
        For i=1 To 10
            Response.Write "<TR bgcolor=#E9EDF9 height=20 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<TD><INPUT name='" & i & "'  tag='25X' flag='SPREADCELL' style='WIDTH: 30px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_ID" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 100px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_NAME" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 120px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_emp_no" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_deptnm" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 120px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_pro_auth' tag='25X' flag='SPREADCELL' style='WIDTH: 80px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_dept_auth' tag='25X' flag='SPREADCELL' style='WIDTH:  90px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 60px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 70px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
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
    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
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
