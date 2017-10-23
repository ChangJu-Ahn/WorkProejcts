<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>

<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/incServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->

<Script Language="VBScript">
Option Explicit 

Const BIZ_PGM_ID      = "e1804mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1      = "e1803ma1.asp"
Const C_SHEETMAXCOLS = 9

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        lgKeyStream       = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtLang_cd.Value) & gColSep
    else
        lgKeyStream       = Trim(frm1.txtEmp_no.Value) & gColSep
    end if
End Sub   
     
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    iCodeArr = gLang & chr(11) 
    iNameArr = gLang & chr(11) 
    Call SetCombo2(frm1.txtlang_cd, iCodeArr, iNameArr, Chr(11))    
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = 10
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call InitComboBox()
    Call LockField(Document)	
    Call LayerShowHide(0)

    Call InitGrid()

    Call SetToolBar("10000")
    Call DbQuery(1)
End Sub

'========================================================================================
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing
    Set Grid1 = Nothing
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    frm1.GRID_PAGE.VALUE = ppage

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)
End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
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
'========================================================================================================
Function DbSaveOk()
    Dim curpage

    Call DbQuery(frm1.GRID_PAGE.VALUE)
End Function

'========================================================================================================
' Function Name : GetRow
'========================================================================================================
Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
End Function

'========================================================================================================
' Function Name : DoubleGetRow
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

'========================================================================================================
' Function Name : MouseRow
'========================================================================================================
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
'                        5.5 Tag Event
'========================================================================================================
Sub Query_OnClick()
    Call DbQuery()
End Sub

Sub GRID_PAGE_OnChange()
End Sub

</SCRIPT>

<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 


</HEAD>
<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
                            <tr> 
								<td width="100" height="30" bgcolor="D4E5E8" class=base1 valign=middle>언어
								</td>
								<td width="626" bgcolor="FFFFFF" align=left colspan=3>&nbsp;&nbsp;
								    <SELECT Name="txtlang_cd" tabindex=-1 class=form01 STYLE="width: 100px;">
								    </SELECT>
								</td>
                            </tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="5"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
							<TR> 
		                    	<TD class=TDFAMILY_TITLE1></TD>
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
							     Response.Write "<TR bgcolor=#F8F8F8 height=24 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
							     Response.Write "<TD><INPUT class=listrow01 name='" & i & "' tag='2'  flag='SPREADCELL' style='WIDTH: 30px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_lang_cd" & i & "' tag='2' flag='SPREADCELL' style='WIDTH: 50px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_menu_id" & i & "' tag='2' flag='SPREADCELL' style='WIDTH: 80px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_menu_name" & i & "' tag='2' flag='SPREADCELL' style='WIDTH:  180px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_href" & i & "' tag='2' flag='SPREADCELL' style='WIDTH:  105px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_menu_level" & i & "' tag='2' flag='SPREADCELL' style='WIDTH: 85px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' tag='2' flag='SPREADCELL' style='WIDTH: 55px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' tag='2' flag='SPREADCELL' style='WIDTH: 83px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' tag='2' flag='SPREADCELL' style='WIDTH: 55px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "</TR>"
							Next
							%>
                           </table>
                        </td>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
            <TD height=5></TD>
        </TR>
        <TR height=20>
            <TD VALIGN=center ALIGN=center>
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../ESSimage/button_07.gif border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
            </TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 border=0>
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
