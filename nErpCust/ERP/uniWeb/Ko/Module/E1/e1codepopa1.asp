<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incServer.asp"  -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<TITLE><%=gLogoName%>-코드팝업</TITLE>
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

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1codepopb1.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXCOLS = 3

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
<!-- #Include file="../../inc/incGrid.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim Grid1

Dim arrParam				'--- First Parameter Group		
Dim arrParent

Dim arrReturn				'--- Return Parameter Group

	arrParent = window.dialogArguments
	arrParam = arrParent(0)

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
        lgKeyStream = Trim(frm1.txtMajor.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & Trim(frm1.txtminor.Value) & gColSep
        lgKeyStream = lgKeyStream & "" & gColSep
    else
        'lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        'lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
    end if
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
     
    frm1.txtMajor.value = arrParam(0)
	frm1.txtMinor.value = arrParam(1)
	frm1.txtCode2.value = arrParam(2)

    Call LayerShowHide(0)
    Call InitGrid()
    Call LockField(Document)
'    Call DbQuery(1)
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
    frm1.GRID_PAGE.VALUE = 1
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)

End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status

    'Call ElementVisible(window.parent.document.all("RunQuery"), 0)

End Function



Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Function DoubleGetRow(pRow)
    Dim objList
    Dim elmCnt

    Dim txttrip_strt_dt
    Dim txttrip_cd
    Dim strVal

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    Redim arrReturn(3)

    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_minor" & pRow then
    		    if objList.value = "" then
    		        exit function
    		    else
		            arrReturn(0) = objList.value
		        end if
    		End if
    		If objList.name = "SPREADCELL_name" & pRow then
    		    if objList.value = "" then
    		        exit function
    		    else
    		        arrReturn(1) = objList.value
    		    end if
    		End if
    		If objList.name = "SPREADCELL_code2" & pRow then
    		    if objList.value = "" then
    		        exit function
    		    else
    		        arrReturn(2) = objList.value
    		    end if
    		End if
    	Next
    End With

	Self.Returnvalue = arrReturn

	DoubleGetRow = True

    Self.Close()
End Function

'========================================================================================================

Sub Query_OnClick()
    Call DbQuery(1)
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

Function txtMinor_onKeyDown()
	Dim CuEvObj,KeyCode	
		Set CuEvObj = window.event.srcElement		
		KeyCode = window.event.keycode
		Select Case KeyCode
			Case 13 'enter key
			    Call DbQuery(1)
		End Select		
		txtMinor_onKeyDown	= True	
End Function

</SCRIPT>

<!-- #Include file="../../inc/uniSimsClassID.inc" --> 


<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
</HEAD>

<BODY>

<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=442 border=0>
        <TR height=5><TD colspan=3></TD></TR>
        <TR height=13>
            <TD background=../../../Cshared/Image/uniSIMS/body1left.jpg width=13></TD>
            <TD background=../../../Cshared/Image/uniSIMS/body1.jpg></TD>
            <TD background=../../../Cshared/Image/uniSIMS/body1right.jpg width=14></TD>
        </TR>
        <TR height=7>
            <TD background=../../../Cshared/Image/uniSIMS/bodyleft.jpg width=13></TD>
            <TD background=../../../Cshared/Image/uniSIMS/body.jpg></TD>
            <TD background=../../../Cshared/Image/uniSIMS/bodyright.jpg width=14></TD>
        </TR>
        <TR>
            <TD background=../../../Cshared/Image/uniSIMS/bodyleft.jpg width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 border=0 bgcolor=#ffffff>
                    <TR height=36 valign=center>
                        <TD class=base1>코드:<INPUT class=base1 NAME="txtMinor" MAXLENGTH=13 SiZE=12 tag=12></TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 border=0 bgcolor=#ffffff>
                                <TR bgcolor=#d0d6e4 height=19>
		                        	<TD></TD>
		                        	<TD class=TDFAMILY_TITLE1>코드</TD>
		                        	<TD class=TDFAMILY_TITLE1>코드명</TD>
		                        	<TD class=TDFAMILY_TITLE1></TD>
                                </TR>
<%            
        For i=1 To 10
            Response.Write "<TR bgcolor=#E9EDF9 height=19 ondblclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<TD><INPUT name='" & i & "'  tag='25X' flag='SPREADCELL' style='WIDTH:  30px; TEXT-ALIGN: center' ></TD>"
            Response.Write "<TD><INPUT type=text name='SPREADCELL_emp_no" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 80px; TEXT-ALIGN: center'></TD>"
            Response.Write "<TD><INPUT name='SPREADCELL_name" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 300px; TEXT-ALIGN: left'></TD>"
            Response.Write "</TR>"
        Next        
%>

                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD background=../../../Cshared/Image/uniSIMS/bodyright.jpg width=14></TD>
        </TR>
        <TR height=20>
            <TD background=../../../Cshared/Image/uniSIMS/bodyleft.jpg width=13></TD>
            <TD VALIGN=center ALIGN=right>
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" href="#" ><IMG alt="이전페이지" src=../../../Cshared/Image/uniSIMS/gprev.jpg border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" href="#" ><IMG alt="다음페이지" src=../../../Cshared/Image/uniSIMS/gnext.jpg border=0 ></A>&nbsp;&nbsp;
            </TD>
            <TD background=../../../Cshared/Image/uniSIMS/bodyright.jpg width=14></TD>
        </TR>
        <TR height=5>
            <TD background=../../../Cshared/Image/uniSIMS/body2left.jpg width=13></TD>
            <TD background=../../../Cshared/Image/uniSIMS/body2.jpg></TD>
            <TD background=../../../Cshared/Image/uniSIMS/body2right.jpg width=14></TD>
        </TR>
        <TR height=10><TD colspan=3></TD></TR>
        <TR>
    		<TD vAlign=middle align=right colspan=3>
    		    <A onclick="VBSCRIPT:CALL DBQuery(1)" href="#" name=submit><img SRC="../../../Cshared/Image/uniSIMS/ret1.jpg" WIDTH=28 HEIGHT=27 border=0 alt='조회' onMouseOver="javascript:this.src='../../../Cshared/Image/uniSIMS/ret2.jpg';" onMouseOut="javascript:this.src='../../../Cshared/Image/uniSIMS/ret1.jpg';"></A>
    		    <A onclick="VBSCRIPT: self.close()" href="#"><img SRC="../../../Cshared/Image/uniSIMS/exit1.jpg" WIDTH=28 HEIGHT=27 border=0 alt='취소' onMouseOver="javascript:this.src='../../../Cshared/Image/uniSIMS/exit2.jpg';" onMouseOut="javascript:this.src='../../../Cshared/Image/uniSIMS/exit1.jpg';"></A>
    		</TD>
        </TR>
    </TABLE>
    <TABLE cellPadding=0 width=500 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">

    <INPUT TYPE=hidden NAME="txtMajor">
    <INPUT TYPE=hidden NAME="txtCode2">
    
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
</FORM>	
</BODY>
</HTML>

