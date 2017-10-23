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

Const BIZ_PGM_ID      = "e1703mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXCOLS = 11
Const C_SHEETMAXROWS = 9

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
        else
            lgKeyStream = lgKeyStream & "A" & gColSep
        end if
    else
        lgKeyStream       = Trim(frm1.txtEmp_no.Value) & gColSep
    end if
End Sub        

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS+1
    Grid1.SheetMaxrows = C_SHEETMAXROWS
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

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

Function DbQuery(ppage)

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    Call ClearField(Document,2)

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
    Dim lRow, iRet
    Err.Clear                                                                    '☜: Clear err status

    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)
	
    With Frm1
       For lRow = 1 To Grid1.SheetMaxrows
			if document.all(CStr(lRow)).value = "" Then
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).style.visibility = "hidden"
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).style.visibility = "hidden"	
			Else
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).style.visibility = "visible"
				document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).style.visibility = "visible"	
				
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
     
    if .txtapp_a.checked = true then
		call SetToolBar("10010")
	end if
	
	if .txtapp_y.checked = true then
		call SetToolBar("10000")
	end if
		
    if .txtapp_n.checked = true then
		call SetToolBar("10010")
	end if
	
	if .txtapp_r.checked = true then
		call SetToolBar("10000")
	end if

	End With
End Function

'========================================================================================================
' Name : DbQueryFail
'========================================================================================================
Function DbQueryFail()
	Dim lRow, iRet
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
    Call Grid1.Clear(frm1,frm1.GRID_PAGE.VALUE)
End Function

'========================================================================================================
' Name : DbSave
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim lStartRow   
    Dim lEndRow     
	Dim strVal, strDel

	Dim strRes_no

	Dim strA
	Dim strOK
'----------------------------------------------------------------------	

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
'----------------------------------------------------------------------	
    DbSave = False                                                          
    
    Call LayerShowHide(1)

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With Frm1
       For lRow = 1 To Grid1.SheetMaxrows
           Select Case document.all(CStr(lRow)).value

               Case UpdateFlag                                      '☜: Update
                    'IF  UCASE(document.all("SPREADCELL_APP_YN" & CStr(lRow)).value) = "Y" THEN
                    IF  document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).checked = true THEN
                        strVal = strVal & "U" & gColSep
                        strVal = strVal & lRow & gColSep
                        strVal = strVal & document.all("SPREADCELL_EMP_NO" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_strt_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & "Y" & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_cd" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_end_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_loc" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_amt" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_app_emp_no" & CStr(lRow)).value  & gRowSep
                        
                        lGrpCnt = lGrpCnt + 1
                    ElseIF  document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).checked = true THEN
                        strVal = strVal & "U" & gColSep
                        strVal = strVal & lRow & gColSep
                        strVal = strVal & document.all("SPREADCELL_EMP_NO" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_strt_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & "R" & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_cd" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_end_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_loc" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_trip_amt" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_app_emp_no" & CStr(lRow)).value  & gRowSep
                        
                        lGrpCnt = lGrpCnt + 1
                    Else
                        Call DisplayMsgBox("800094","X","X","X")
                        document.all("SPREADCELL_APP_YN" & CStr(lRow)).value = "N"
                        document.all("SPREADCELL_APP_YN" & CStr(lRow)).focus
                    End If
               Case DeleteFlag                                      '☜: Delete
                    strDel = strDel & "D" & gColSep
                    strDel = strDel & lRow & gColSep
                    strDel = strDel & .txtEmp_no.value & gColSep
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
' Function Name : DoubleGetRow
'========================================================================================================
Function DoubleGetRow(pRow)
    Dim objList
    Dim elmCnt
	dim sm_name, sm_emp_no
    Dim txttrip_strt_dt, txttrip_end_dt
    Dim txttrip_cd
    DIM LOC
    DIM APP_YN
    Dim Trip_amt
    Dim strVal
	dim iRet
	dim remark


	DoubleGetRow = False
	Grid1.ActiveRow = pRow
	sm_name         = ""
    sm_emp_no       = ""
    txttrip_strt_dt = ""
    txttrip_end_dt  = ""
    txttrip_cd      = ""
    Trip_amt        = ""
    APP_YN          = ""
    remark			= ""
    
    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		
    		If objList.name = "SPREADCELL_EMP_NO" & pRow AND objList.value <> "" then
               sm_emp_no = objList.value
    		End if
    		
    		If objList.name = "SPREADCELL_name" & pRow  AND objList.value <> "" then
               sm_name = objList.value
    		End if
    		
    		If objList.name = "SPREADCELL_trip_strt_DT" & pRow AND objList.value <> "" then
               txttrip_strt_dt = objList.value
    		End if
    		
    		If objList.name = "SPREADCELL_trip_end_DT" & pRow AND objList.value <> "" then
               txttrip_end_dt = objList.value
    		End if
    		If objList.name = "SPREADCELL_trip_CD" & pRow AND objList.value <> "" then
               txttrip_cd = objList.value
    		End if
    		
    		If objList.name = "SPREADCELL_trip_loc" & pRow AND objList.value <> "" then
               LOC = objList.value
    		End if
    		
    		If objList.name = "SPREADCELL_trip_amt" & pRow AND objList.value <> "" then
               Trip_amt = objList.value
    		End if
    		
    		If objList.name = "SPREADCELL_APP_YN1_" & pRow AND objList.value <> "" then
               APP_YN = objList.value
    		End if
    		
    		If objList.name = "SPREADCELL_remark_" & pRow AND objList.value <> "" then
               remark = objList.value
    		End if    		
    		
    	Next
    End With
    strval = "성명 : " & sm_name & vbcr & "사번 :" & sm_emp_no & vbcr
    strval = strval  & "출장기간 : " & txttrip_strt_dt & " ~ " & txttrip_end_dt & vbCr
    strval = strval & "내용 : " & LOC & vbCR
    strval = strval & "출장비 : " & Trip_amt & vbCR
    strval = strval & "비고 : " & remark & vbCR
    strVal = strVal & "--------------------------------------------" & vbCR & vbCR
    strVal = strVal & "승인하시겠습니까?" & vbCR
    strVal = strVal & "승인 : Yes, 반려 : No, 보류 : Cancel"
    
    if APP_YN = "N" then
		iRet = msgbox (strval , vbYesNoCancel, "출장승인")
		if iRet = vbYes then
			document.all("SPREADCELL_APP_YN2_" & CStr(pRow))(0).checked = true
			Call grid1.SetUpdateFlag(pRow)
			
		elseif iRet = vbNO then
			document.all("SPREADCELL_APP_YN2_" & CStr(pRow))(1).checked = true
			Call grid1.SetUpdateFlag(pRow)
		else
			
		end if 
	end if
	DoubleGetRow = True
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub Query_OnClick()
    Call DbQuery(1)
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
                        <td><table width="733" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
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
		                        <TD NOWRAP class=TDFAMILY_TITLE1>사번</TD>
		                        <TD NOWRAP class=TDFAMILY_TITLE1>성명</TD>
		                        <TD NOWRAP class=TDFAMILY_TITLE1 colspan=2>출장기간</TD>
		                        <TD NOWRAP class=TDFAMILY_TITLE1>출장</TD>
		                        <TD NOWRAP class=TDFAMILY_TITLE1>내용</TD>
		                        <TD NOWRAP class=TDFAMILY_TITLE1>출장비</TD>
		                        <TD NOWRAP class=TDFAMILY_TITLE1>승인자</TD>
		                        <TD NOWRAP class=TDFAMILY_TITLE1>승인/반려</TD>
                                </TR>
							<% 
							For i=1 To 9
							     Response.Write "<TR bgcolor=#F8F8F8 height=24 ondblclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
							     Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH: 40px; text-align: center;'  readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_EMP_NO" & i & "' flag='SPREADCELL' style='WIDTH: 80px; text-align: left;')' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_name" & i & "' flag='SPREADCELL' style='WIDTH: 70px; text-align: center;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_trip_strt_DT" & i & "' flag='SPREADCELL' style='WIDTH:  65px; text-align: center;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_trip_end_DT" & i & "' flag='SPREADCELL' style='WIDTH:  65px; text-align: center;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_trip_CD" & i & "' type=hidden flag='SPREADCELL' style='WIDTH: 0px; text-align: center;'>"
							     Response.Write "    <INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 90px; text-align: left;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_trip_loc" & i & "' flag='SPREADCELL' style='WIDTH: 120px; text-align: left;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_trip_amt" & i & "' flag='SPREADCELL' style='WIDTH: 80px; text-align: right;' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_app_emp_no" & i & "' flag='SPREADCELL' style='WIDTH: 70px; text-align: center;' readonly></TD>"
								 Response.Write "<TD align=center><INPUT class=listrow01 type='hidden' name='SPREADCELL_APP_YN1_" & i & "' flag='SPREADCELL' size=10 style='WIDTH: 50px;'>"
								 Response.Write "    <INPUT class=listrow01 type=RADIO name='SPREADCELL_APP_YN2_" & i & "' value='Y'  maxlength=1 style='WIDTH: 21px; visibility : hidden; text-align: center;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled>"
								 Response.Write "/   <INPUT class=listrow01 type=RADIO name='SPREADCELL_APP_YN2_" & i & "' value='R'  maxlength=1 style='WIDTH: 21px; visibility : hidden; text-align: center;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled></TD>"
						    	 Response.Write "<TD align=center><INPUT class=listrow01 type='hidden' name='SPREADCELL_remark_" & i & "' flag='SPREADCELL' size=1>"
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
    <TEXTAREA style="display: none" name=txtSpread></TEXTAREA>
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
 </FORM>	
</BODY>
</HTML>
