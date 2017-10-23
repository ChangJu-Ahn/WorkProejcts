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
Const BIZ_PGM_ID      = "e1607mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXCOLS = 11

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
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
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
    Grid1.MaxQueryRows = 10
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Function Name : GridDsplay
' Function Desc : This method initializes spread sheet column
'========================================================================================================
Function GridDsplay()
	Dim i
	    For i=1 To 10
	        document.writeln "<TR bgcolor=#E9EDF9 height=20 onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
	        document.writeln "<TD><INPUT name='" & i & "'  tag='25X' flag='SPREADCELL' style='WIDTH: 40px; TEXT-ALIGN: center' ></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_EMP_NO" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 85px; TEXT-ALIGN: left')'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH: 110px; TEXT-ALIGN: left'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_DILIG_DT" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH:  100px; TEXT-ALIGN: center'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_DILIG_HOUR" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH:  15px; TEXT-ALIGN: center'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_DILIG_MIN" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH:  15px; TEXT-ALIGN: center'></TD>"	        
	        document.writeln "<TD><INPUT name='SPREADCELL_DILIG_CD" & i & "' tag='25x' type=hidden flag='SPREADCELL' style='WIDTH: 0px; TEXT-ALIGN: center'>"
	        document.writeln "<INPUT name='SPREADCELL" & i & "' tag='25x' flag='SPREADCELL' style='WIDTH: 80px; TEXT-ALIGN: left'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_remark" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 117px; TEXT-ALIGN: LEFT'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL' tag='25X' flag='SPREADCELL' style='WIDTH: 80px;'></TD>"
	    	document.writeln "<TD align=center><input type='hidden' name='SPREADCELL_APP_YN1_" & i & "' flag='SPREADCELL' size=1>"
	    	document.writeln "    <INPUT type=RADIO name='SPREADCELL_APP_YN2_" & i & "' tag='25X' value='Y'  maxlength=1 style='WIDTH: 15px; TEXT-ALIGN: center; visibility : hidden;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled>"
	    	document.writeln "/   <INPUT type=RADIO name='SPREADCELL_APP_YN2_" & i & "' tag='25X' value='R'  maxlength=1 style='WIDTH: 15px; TEXT-ALIGN: center; visibility : hidden;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled>"
	        document.writeln "</TD>"
	        document.writeln "</TR>"
	    Next
End Function


'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                   '☜: Clear err status
    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call InitComboBox()

    Call LayerShowHide(0)

    Call InitGrid()

    Call SetToolBar("10010")

    Call LockField(Document)

    Call DbQuery(1)

End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
    Set Grid1 = Nothing
End Sub

Function DbQuery(ppage)

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    Call ClearField(document,2)

    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    frm1.GRID_PAGE.VALUE = ppage

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

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
	End With     
End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
    Call Grid1.Clear(frm1,frm1.grid_page.value)
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
                    IF  document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(0).checked = true THEN
                    '----------------- ☜:Status : Yes -------------------------------
                        strVal = strVal & "U" & gColSep
                        strVal = strVal & lRow & gColSep
                        strVal = strVal & document.all("SPREADCELL_EMP_NO" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_HOUR" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_MIN" & CStr(lRow)).value  & gColSep                        
                        strVal = strVal & document.all("SPREADCELL_DILIG_CD" & CStr(lRow)).value  & gColSep
                        strVal = strVal & "Y" & gColSep
                        strVal = strVal & document.all("SPREADCELL_remark" & CStr(lRow)).value  & gRowSep
                        lGrpCnt = lGrpCnt + 1
                    ElseIF  document.all("SPREADCELL_APP_YN2_" & CStr(lRow))(1).checked = true THEN
                    '----------------- ☜:Status : Return ----------------------------
                        strVal = strVal & "U" & gColSep
                        strVal = strVal & lRow & gColSep
                        strVal = strVal & document.all("SPREADCELL_EMP_NO" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_DT" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_HOUR" & CStr(lRow)).value  & gColSep
                        strVal = strVal & document.all("SPREADCELL_DILIG_MIN" & CStr(lRow)).value  & gColSep                        
                        strVal = strVal & document.all("SPREADCELL_DILIG_CD" & CStr(lRow)).value  & gColSep
                        strVal = strVal & "R" & gColSep
                        strVal = strVal & document.all("SPREADCELL_remark" & CStr(lRow)).value  & gRowSep
                        lGrpCnt = lGrpCnt + 1
                    Else
                        Call DisplayMsgBox("800094","X","X","X")
                        document.all("SPREADCELL_APP_YN_2" & CStr(lRow))(0).checked = false
                        document.all("SPREADCELL_APP_YN_2" & CStr(lRow))(1).checked = false
'                        document.all("SPREADCELL_APP_YN_2" & CStr(lRow)).focus
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
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Dim curpage
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call DbQuery(frm1.grid_page.VALUE)

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
'                        5.5 Tag Event
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
<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
<TD>
    <TABLE cellSpacing=0 cellPadding=0 width=747 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=641 border=0 bgcolor=#ffffff>
                    <TR height=45 valign=middle>
                        <TD class=base1 colspan=2>승인여부:
                            <INPUT TYPE="RADIO" NAME="txtapp_yn" tag="12" ID="txtapp_a" VALUE=A STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #FFFFFF"><LABEL FOR="txtapp_a">전체</LABEL>
    					    <INPUT TYPE="RADIO" NAME="txtapp_yn" tag="12" ID="txtapp_y" VALUE=Y STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #FFFFFF"><LABEL FOR="txtapp_y">승인</LABEL>
                            <INPUT TYPE="RADIO" NAME="txtapp_yn" tag="12" CHECKED ID="txtapp_n" VALUE=N STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #FFFFFF"><LABEL FOR="txtapp_n">미처리</LABEL>
    					    <INPUT TYPE="RADIO" NAME="txtapp_yn" tag="12" ID="txtapp_r" VALUE=R STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #FFFFFF"><LABEL FOR="txtapp_R">반려</LABEL>
                        </TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width="100%" border=0 bgcolor=#ffffff>
                                <TR bgcolor=#d0d6e4 height=20 valign=middle>
		                        	<TD></TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>사번</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>성명
		                        	</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1 colspan=1>근무날짜</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1 colspan=1>시간</TD>	
		                        	<TD NOWRAP class=TDFAMILY_TITLE1 colspan=1>분</TD>			                        		                        	
		                        	<TD NOWRAP class=TDFAMILY_TITLE1 colspan=1>근태</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>사유</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>승인자</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>승인/반려</TD>
                                </TR>
								<script language=vbscript>    Call GridDsplay()  </script>
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
                        <A onclick="Vbscript: Call Grid1.Prepages()"  onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../../../Cshared/Image/uniSIMS/gprev.jpg border=0 ></A>&nbsp;
                        <A onclick="Vbscript: Call Grid1.Nextpages()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../../../Cshared/Image/uniSIMS/gnext.jpg border=0 ></A>&nbsp;&nbsp;
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
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
</TD>
</FORM>
</TR>
</TABLE>
</BODY>
</HTML>
</TD>
</FORM>
</TR>
</TABLE>
</BODY>
</HTML>
