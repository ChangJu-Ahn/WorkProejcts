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

Const BIZ_PGM_ID      = "e1702mb1.asp"						           '��: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1701ma1.asp"						           '��: Biz Logic ASP Name
Const C_SHEETMAXCOLS = 10
Const C_SHEETMAXROWS = 9

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1
dim fDiligAuth,fAuthCheck
<% EndDate   = GetSvrDate %>

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
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
End Sub        

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = C_SHEETMAXROWS
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
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
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
    Set Grid1 = Nothing
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)
    Dim strVal
    Dim strDate
    Err.Clear                                                                    '��: Clear err status

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
		If CompareDateByFormat(.txtfrom.value,.txtto.value,"���������","����������","970025", gDateFormat, gComDateType,TRUE) = False Then
		    .txtfrom.focus()
		    exit function
		END IF
    End With	
	If Not chkFieldLength(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If
    	
    DbQuery = False                                                              '��: Processing is NG
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQuery = True                                                               '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '��: Clear err status

    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)

End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '��: Clear err status
End Function

'========================================================================================================
' Name : DbSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG

	Call LayerShowHide(1)

	With Frm1
		.txtMode.value        = "UID_M0002"                                        '��: Save
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '��: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
'========================================================================================================
Function DbSaveOk()
    Call DbQuery(1)
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
                If objList.value = "����" Then
                 '   Call DisplayMsgBox("800472","X","X","X")
				'	Exit Function
					txtapp_yn = "Y"
				Else
					If objList.value = "�ݷ�" Then
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
    Call DbQuery(1)
End Sub

Sub GRID_PAGE_OnChange()
End Sub

Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function

'========================================================================================================
' Name : FncGetDiligAuth()
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
                        <td><table width="733" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
                            <tr> 
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">���</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">����</td>
								<td width="86" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">����</td>
								<td width="100" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">�μ�</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
							<tr height=25 valign=top>
							    <TD width="60" height="30" bgcolor="D4E5E8" class=base1 valign=center>�Ⱓ</TD>
							    <TD bgcolor="FFFFFF" class="base2" align=left valign=center colspan=7>&nbsp;&nbsp;
							        <INPUT ID="txtfrom" NAME="txtfrom" MAXLENGTH=16 SiZE=12 tag="12" ondblclick="VBScript:Call OpenCalendar('txtfrom',3)" style='font-family: "����"; font-size: 9pt; color: 002232; padding-left: 12px;'>&nbsp;~
							        <INPUT ID="txtto" NAME="txtto" MAXLENGTH=16 SiZE=12 tag="12" ondblclick="VBScript:Call OpenCalendar('txtto',3)" style='font-family: "����"; font-size: 9pt; color: 002232; padding-left: 12px; valign:center;'>
							    </TD>
							</tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
								<TR> 
								    <TD class=TDFAMILY_TITLE1></TD>
		                        	<TD class=TDFAMILY_TITLE1 colspan=2>����Ⱓ</TD>
		                        	<TD class=TDFAMILY_TITLE1>����</TD>
		                        	<TD class=TDFAMILY_TITLE1>���峻��</TD>
		                        	<TD class=TDFAMILY_TITLE1>�����</TD>
		                        	<TD class=TDFAMILY_TITLE1>���</TD>
		                        	<TD class=TDFAMILY_TITLE1>������</TD>
		                        	<TD class=TDFAMILY_TITLE1>����</TD>
                                </TR>
							    <% 
                                For i=1 To 9
                                    Response.Write "<TR bgcolor=#F8F8F8 height=24 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
                                    Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH:  30px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly ></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_trip_strt_DT" & i & "' flag='SPREADCELL' style='WIDTH: 73px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_trip_end_DT" & i & "' flag='SPREADCELL' style='WIDTH: 73px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_trip_CD" & i & "' type=hidden flag='SPREADCELL' style='WIDTH:   0px; text-align: left;'>"
                                	Response.Write "    <INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 80px; text-align: left;'  onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 170px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH:  70px; text-align: right;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 100px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH:  65px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_app_yn" & i & "' flag='SPREADCELL' style='WIDTH: 60px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
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
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="����������" src=../ESSimage/button_07.gif border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="����������" src=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
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
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
 </FORM>	
</BODY>
</HTML>
