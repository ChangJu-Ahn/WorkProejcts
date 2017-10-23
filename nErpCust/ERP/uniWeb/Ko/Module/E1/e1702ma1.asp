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
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

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

Const BIZ_PGM_ID      = "e1702mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1701ma1.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXCOLS = 10

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
<!-- #Include file="../../inc/incGrid.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim Grid1
dim fDiligAuth,fAuthCheck
<% EndDate   = GetSvrDate %>

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
		if  Trim(parent.txtEmp_no2.Value) = "" then
		    lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
		else
		    lgKeyStream = Trim(parent.txtEmp_no2.Value) & gColSep
		end if
    else     
            lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep   
    end if 
        
        lgKeyStream = lgKeyStream & Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtfrom.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtto.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(fDiligAuth) & gColSep        
        lgKeyStream = lgKeyStream & Trim(fAuthCheck) & gColSep     
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
    call FncGetDiligAuth(fDiligAuth,fAuthCheck)
    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If    
    Call LayerShowHide(0)

    Call InitGrid()

    Call SetToolBar("10000")

	frm1.txtto.Value   = UniConvDateAToB("<%=EndDate%>",gServerDateFormat,gDateFormat)
	frm1.txtfrom.Value =  uniDateAdd("d", -7 ,frm1.txtto.value, gDateFormat) 

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
    Dim strDate
    Err.Clear                                                                    '☜: Clear err status


   	With Frm1

        if  Date_chk(.txtfrom.value, strDate) = True then
            .txtfrom.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtfrom.focus()
            exit function
        end if
        if  Date_chk(.txtto.value, strDate) = True then
            .txtto.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtto.focus()
            exit function
        end if    
		If CompareDateByFormat(.txtfrom.value,.txtto.value,"출장시작일","출장종료일","970025", gDateFormat, gComDateType,TRUE) = False Then
		    .txtfrom.focus()
		    exit function
		END IF
    End With	
	If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    	
    DbQuery = False                                                              '☜: Processing is NG
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
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

'	frm1.txtto.Value   = UniConvDateAToB("<%=EndDate%>",gServerDateFormat,gDateFormat)
'	frm1.txtfrom.Value =  uniDateAdd("d", -7 ,frm1.txtto.value, gDateFormat) 
    
'    frm1.txtfrom.value = UNIDateClientFormat(StartDate)
'    frm1.txtto.value = UNIDateClientFormat(EndDate)
    Call ClearField(Document,2)                                                                    '☜: Clear err status

'    Call ElementVisible(window.parent.document.all("RunQuery"), 0)

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG

	Call LayerShowHide(1)

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
'		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    'Call InitVariables

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call DbQuery(1)
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

    Dim txttrip_strt_dt
    Dim txttrip_cd
    Dim txtapp_yn
    Dim strVal

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    txttrip_strt_dt = ""
    txttrip_cd = ""
    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_trip_strt_DT" & pRow then
               txttrip_strt_dt = objList.value
    		End if
    		If objList.name = "SPREADCELL_trip_CD" & pRow then
               txttrip_cd = objList.value
    		End if
    		If objList.name = "SPREADCELL_app_yn" & pRow then
                If objList.value = "승인" Then
                 '   Call DisplayMsgBox("800472","X","X","X")
				'	Exit Function
					txtapp_yn = "Y"
				Else
					If objList.value = "반려" Then
						txtapp_yn = "R"
					Else
						txtapp_yn = "N"
					End IF
				End if
    		End if
    	Next
    End With

    If txttrip_strt_dt <> "" and txttrip_cd <> "" then

        strVal = BIZ_PGM_ID1 & "?trip_strt_dt=" & txttrip_strt_dt
        strVal = strVal & "&trip_cd=" & txttrip_cd & "&app_yn=" & txtapp_yn
        
		Call CommonQueryRs(" MENU_NAME "," E11000T "," Menu_id = " & FilterVar("E1701MA1", "''", "S") & " AND LANG_CD =  " & FilterVar(gLang , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
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
Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function
'========================================================================================================
' Name : FncGetDiligAuth()
' Desc : developer describe this line 
'========================================================================================================
Function FncGetDiligAuth(fDiligAuth,fAuthCheck)
    fDiligAuth = ""
    fAuthCheck = ""
    Call CommonQueryRs(" internal_cd,internal_auth "," e11090t "," emp_no =  " & FilterVar(parent.txtEmp_no.Value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    fDiligAuth = replace(lgF0,chr(11),chr(12))
    fDiligAuth = replace(fDiligAuth," ","")    
    fAuthCheck = replace(lgF1,chr(11),chr(12))
    fAuthCheck = replace(fAuthCheck," ","")      
End Function



</SCRIPT>
<!-- #Include file="../../inc/uniSimsClassID.inc" --> 



<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST" >
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=722 border=0 bgcolor=#ffffff>
                    <TR height=30 valign=center  height=19>
                        <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 TAG=14></TD>
                        <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10 TAG=14></TD>
                        <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10 TAG=14></TD>
                        <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15 TAG=14></TD>
                    </TR>
                    <TR height=25 valign=top>
                        <TD class=base1 colspan=4>기간:
                            <INPUT class=SINPUTTEST_STYLE ID="txtfrom" NAME="txtfrom" MAXLENGTH=10 SiZE=10 tag="12D" ondblclick="VBScript:Call OpenCalendar('txtfrom',3)">&nbsp;~
                            <INPUT class=SINPUTTEST_STYLE ID="txtto" NAME="txtto" MAXLENGTH=10 SiZE=10 tag="12D" ondblclick="VBScript:Call OpenCalendar('txtto',3)">
                        </TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width="100%" border=0 bgcolor=#ffffff>
                                <TR bgcolor=#d0d6e4 height=20>
		                        	<TD></TD>
		                        	<TD class=TDFAMILY_TITLE1 colspan=2>출장기간</TD>
		                        	<TD class=TDFAMILY_TITLE1>출장</TD>
		                        	<TD class=TDFAMILY_TITLE1>출장내용</TD>
		                        	<TD class=TDFAMILY_TITLE1>출장비</TD>
		                        	<TD class=TDFAMILY_TITLE1>비고</TD>
		                        	<TD class=TDFAMILY_TITLE1>승인자</TD>
		                        	<TD class=TDFAMILY_TITLE1>승인</TD>
                                </TR>
						<%            
						    For i=1 To 10
						        Response.Write "<TR bgcolor=#E9EDF9 height=20 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
						        Response.Write "<TD><INPUT name='" & i & "'  tag='25X' flag='SPREADCELL' style='WIDTH:  30px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        Response.Write "<TD><INPUT name='SPREADCELL_trip_strt_DT" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 75px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        Response.Write "<TD><INPUT name='SPREADCELL_trip_end_DT" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 75px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						        Response.Write "<TD><INPUT name='SPREADCELL_trip_CD" & i & "' tag='25x' type=hidden flag='SPREADCELL' style='WIDTH: 0px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'>"
						        Response.Write "<INPUT name='SPREADCELL" & i & "' tag='25x' flag='SPREADCELL' style='WIDTH: 100px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						    	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH: 170px; TEXT-ALIGN: left' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						    	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH:  80px; TEXT-ALIGN: right' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						    	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH: 100px; TEXT-ALIGN: LEFT' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						    	Response.Write "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH: 70px;' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
						    	Response.Write "<TD><INPUT name='SPREADCELL_app_yn" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 60px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
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
    <TABLE cellSpacing=2 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">
    
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
</FORM>	

</BODY>
</HTML>
