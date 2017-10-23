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
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "e1603mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXCOLS = 12

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
        if frm1.txtapp_n.checked = true then
            lgKeyStream = lgKeyStream & "N" & gColSep
        elseif frm1.txtapp_y.checked = true then
            lgKeyStream = lgKeyStream & "Y" & gColSep
        elseif frm1.txtapp_r.checked = true then
            lgKeyStream = lgKeyStream & "R" & gColSep
        elseif frm1.txtapp_a.checked = true then
            lgKeyStream = lgKeyStream & "A" & gColSep
        else
            lgKeyStream = lgKeyStream & "N" & gColSep
        end if
    else
        lgKeyStream       = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
    end if
End Sub        

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = 10
    Grid1.MaxQueryRows = 10
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                   '☜: Clear err status
    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call LayerShowHide(0)
    Call InitGrid()      
    Call SetToolBar("10010")
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
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    Call ClearField(document,2)

    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    frm1.GRID_PAGE.VALUE = ppage

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Dim lRow, iRet
    Err.Clear                                                                    '☜: Clear err status
 
    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)
    With Frm1
       For lRow = 1 To Grid1.SheetMaxrows
			if document.all(CStr(lRow)).value = "" Then
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).style.visibility = "hidden"
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).style.visibility = "hidden"	
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(2).style.visibility = "hidden"	
			Else
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).style.visibility = "visible"
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).style.visibility = "visible"	
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(2).style.visibility = "visible"	
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(2).disabled = False
				
				if  document.all("SPREADCELL_APP_YN1_" & CStr(lRow)).value = "Y" then
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).checked=true
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).disabled = True
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).disabled = True
				elseif   document.all("SPREADCELL_APP_YN1_" & CStr(lRow)).value = "R" then
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).checked = True
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).disabled = True
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).disabled = True
				else
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).checked = False
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).checked = False
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).disabled = False
					document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).disabled = False
				end if
			End IF
			
     Next
	End With     
End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
    Call Grid1.Clear(frm1,frm1.grid_page.value)
End Function

'========================================================================================================
' Name : DbSave
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
	Dim strA
	Dim strOK
 
'수정된 내용있나 검색.
    With Frm1
       For lRow = 1 To Grid1.SheetMaxrows
			If document.all(CStr(lRow)).value = "수정" Then
					strOK = "Y"
			End If
       Next
	End With

	If strOK <> "Y" Then
	
        Call DisplayMsgBox("800380","X","X","X")
        '메시지 내용 : 수정된 내용이 없습니다 
		exit function
	
	End IF	

    DbSave = False                                                          
    Call LayerShowHide(1)
    
    strVal = ""
    strDel = ""
    lGrpCnt = 1
 
	With Frm1
 
       For lRow = 1 To Grid1.SheetMaxrows
           Select Case document.all(CStr(lRow)).value
               Case UpdateFlag                                      '☜: Update
                    IF  document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).checked = true THEN
                        strVal = strVal & "U" & gColSep
                        strVal = strVal & lRow & gColSep
                        strVal = strVal & document.all("SPREADCELL_EMP_NO" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_STRT_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_END_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_CD" & CStr(lRow)).value  & gColSep
                        strVal = strVal & "Y" & gColSep
                        strVal = strVal & document.all("SPREADCELL_remark" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_HOUR" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_MIN" & CStr(lRow)).value  & gRowSep
                        lGrpCnt = lGrpCnt + 1
                    ElseIF  document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).checked = true THEN
                        strVal = strVal & "U" & gColSep
                        strVal = strVal & lRow & gColSep
                        strVal = strVal & document.all("SPREADCELL_EMP_NO" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_STRT_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_END_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_CD" & CStr(lRow)).value  & gColSep
                        strVal = strVal & "R" & gColSep
                        strVal = strVal & document.all("SPREADCELL_remark" & CStr(lRow)).value  & gRowSep
                        lGrpCnt = lGrpCnt + 1
                    ElseIF  document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(2).checked = true THEN
                        strVal = strVal & "D" & gColSep
                        strVal = strVal & lRow & gColSep
                        strVal = strVal & document.all("SPREADCELL_EMP_NO" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_STRT_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_END_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_CD" & CStr(lRow)).value  & gColSep
                        strVal = strVal & "Y" & gColSep
                        strVal = strVal & document.all("SPREADCELL_remark" & CStr(lRow)).value  & gRowSep
                        lGrpCnt = lGrpCnt + 1
                    Else
                        Call DisplayMsgBox("800094","X","X","X")
                        document.all("SPREADCELL_APP_YN_2" & CStr(lRow))(0).checked = false
                        document.all("SPREADCELL_APP_YN_2" & CStr(lRow))(1).checked = false
                    End If               
               Case DeleteFlag                                      '☜: Delete
                    strDel = strDel & "D" & gColSep
                    strDel = strDel & lRow & gColSep
                    strDel = strDel & parent.txtEmp_no.value & gColSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

       .txtMode.value        = "UID_M0002"
       .txtUpdtUserId.value  = parent.txtEmp_no.value
       .txtInsrtUserId.value = parent.txtEmp_no.value
	   .txtMaxRows.value     = lGrpCnt-1	
 	   
	   .txtSpread.value      = strDel & strVal
	End With
 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
 	
exit function
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Function Name : DbSaveOk
'========================================================================================================
Function DbSaveOk()
    Dim curpage
 
    Call DbQuery(frm1.grid_page.VALUE)

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

    Dim txtDilig_dt
    Dim txtDilig_cd
    Dim strVal

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    txtDilig_dt = ""
    txtDilig_cd = ""
    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_DILIG_DT" & pRow then
               txtDilig_dt = objList.value
    		End if
    		If objList.name = "SPREADCELL_DILIG_CD" & pRow then
               txtDilig_cd = objList.value
    		End if
    	Next
    End With

    If txtDilig_dt <> "" and txtDilig_cd <> "" then
        strVal = BIZ_PGM_ID1 & "?Dilig_dt=" & txtDilig_dt
        strVal = strVal & "&Dilig_cd=" & txtDilig_cd
        document.location = strVal
    End if
	DoubleGetRow = True
End Function

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
    <TABLE cellSpacing=0 cellPadding=0 border=0 width=732>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
                            <tr> 
								<TD width="100" height="30" bgcolor="D4E5E8" class=base1>승인여부
								</TD>
								<TD width="626" height="30" bgcolor="FFFFFF" class="base2" align=left valign=center colspan=3>&nbsp;&nbsp;
								    <INPUT TYPE="RADIO" NAME="txtapp_yn" CLASS="radio_title" ID="txtapp_a" VALUE=A ><LABEL FOR="txtapp_a">전체</LABEL>
    							    <INPUT TYPE="RADIO" NAME="txtapp_yn" CLASS="radio_title" ID="txtapp_y" VALUE=Y ><LABEL FOR="txtapp_y">승인</LABEL>
								    <INPUT TYPE="RADIO" NAME="txtapp_yn" CLASS="radio_title" CHECKED ID="txtapp_n" VALUE=N ><LABEL FOR="txtapp_n">미처리</LABEL>
    							    <INPUT TYPE="RADIO" NAME="txtapp_yn" CLASS="radio_title" ID="txtapp_r" VALUE=R ><LABEL FOR="txtapp_R">반려</LABEL>
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
		                    	<TD class=TDFAMILY_TITLE1>사번</TD>
		                    	<TD class=TDFAMILY_TITLE1>성명</TD>
		                    	<TD class=TDFAMILY_TITLE1 colspan=2>근태기간</TD>
		                    	<TD class=TDFAMILY_TITLE1 colspan=2>시간</TD>
		                    	<TD class=TDFAMILY_TITLE1>근태</TD>		                    		                    	
		                    	<TD class=TDFAMILY_TITLE1>사유</TD>
		                    	<TD class=TDFAMILY_TITLE1>승인자</TD>
		                    	<TD class=TDFAMILY_TITLE1>승인/반려/취소</TD>
                            </TR>
							<% 
							For i=1 To 10
							     Response.Write "<TR bgcolor=#F8F8F8 height=24 onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
							     Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH: 40px; text-align: center;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_EMP_NO" & i & "' flag='SPREADCELL' style='WIDTH: 65px; text-align: left;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' flag='SPREADCELL' style='WIDTH: 80px; text-align: left;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_DILIG_STRT_DT" & i & "' flag='SPREADCELL' style='WIDTH: 70px; text-align: center;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_DILIG_END_DT" & i & "' flag='SPREADCELL' style='WIDTH: 70px; text-align: center;' readonly></TD>"
							     Response.Write "<TD align=right><INPUT class=listrow01 name='SPREADCELL_DILIG_HOUR" & i & "' flag='SPREADCELL' style='WIDTH: 20px; text-align: right;' readonly></TD>"
							     Response.Write "<TD align=right><INPUT class=listrow01 name='SPREADCELL_DILIG_MIN" & i & "' flag='SPREADCELL' style='WIDTH:  20px; text-align: right;' readonly></TD>"
							     
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_DILIG_CD" & i & "' type=hidden flag='SPREADCELL' style='WIDTH: 0px; text-align: center;'>"
							     Response.Write "    <INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 75px; text-align: left;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_remark" & i & "' flag='SPREADCELL' style='WIDTH: 115px; text-align: left;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' flag='SPREADCELL' style='WIDTH: 70px; text-align: center;' readonly></TD>"
								 Response.Write "<TD align=center><INPUT class=listrow01 type='hidden' name='SPREADCELL_APP_YN1_" & i & "' flag='SPREADCELL' size=60 style='WIDTH: 60px;'>"
								 Response.Write "    <INPUT class=listrow01 type=RADIO name='SPREADCELL_APP_YN2_" & i & "' value='Y'  maxlength=1 style='WIDTH: 21px; visibility : hidden; text-align: center;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled>"
								 Response.Write "/   <INPUT class=listrow01 type=RADIO name='SPREADCELL_APP_YN2_" & i & "' value='R'  maxlength=1 style='WIDTH: 21px; visibility : hidden; text-align: center;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled>"
								 Response.Write "/   <INPUT class=listrow01 type=RADIO name='SPREADCELL_APP_YN2_" & i & "' value='C'  maxlength=1 style='WIDTH: 21px; visibility : hidden; text-align: center;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled>"
							     Response.Write "</TD>"
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
    <TABLE cellSpacing=0 cellPadding=0 border=0>    
    <TEXTAREA style="display: none" name=txtSpread></TEXTAREA>
    </TABLE>    
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
 </FORM>	
</BODY>
</HTML>
