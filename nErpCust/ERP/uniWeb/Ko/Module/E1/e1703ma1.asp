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

Const BIZ_PGM_ID      = "e1703mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------
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
        if frm1.txtapp_n.checked = true then
            lgKeyStream = lgKeyStream & "N" & gColSep
        elseif frm1.txtapp_y.checked = true then
            lgKeyStream = lgKeyStream & "Y" & gColSep
        elseif frm1.txtapp_r.checked = true then
            lgKeyStream = lgKeyStream & "R" & gColSep
        else
            lgKeyStream = lgKeyStream & "A" & gColSep
        end if
        'lgKeyStream = lgKeyStream & Trim(frm1.txtfrom.Value) & gColSep
        'lgKeyStream = lgKeyStream & Trim(frm1.txtto.Value) & gColSep
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
    Grid1.MaxCols = C_SHEETMAXCOLS+1
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

    Call SetToolBar("10010")
    
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
    Call ClearField(Document,2)
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
	

'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)

End Function

Function DbQueryFail()
	Dim lRow, iRet
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
    
    Call Grid1.Clear(frm1,frm1.GRID_PAGE.VALUE)
    

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
                    '.vspdData.Col = C_FAMILY_NM	: strDel = strDel & Trim(.vspdData.Text) & gColSep
                    '.vspdData.Col = C_REL_CD	: strDel = strDel & Trim(.vspdData.Text) & gRowSep
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


'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncPrev() 
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


Function GridDsplay()
	Dim i
	    For i=1 To 10
	        document.writeln "<TR bgcolor=#E9EDF9 height=19 ondblclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
	        document.writeln "<TD><INPUT name='" & i & "'  tag='25X' flag='SPREADCELL' style='WIDTH: 30px; TEXT-ALIGN: center' ></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_EMP_NO" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 80px; TEXT-ALIGN: left'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_name" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 60px; TEXT-ALIGN: left' onClick='vbscript: Call GetRow(1)'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_trip_strt_DT" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH:  65px; TEXT-ALIGN: center'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_trip_end_DT" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH:  65px; TEXT-ALIGN: center'></TD>"
	        document.writeln "<TD><INPUT name='SPREADCELL_trip_CD" & i & "' tag='25x' type=hidden flag='SPREADCELL' style='WIDTH: 0px; TEXT-ALIGN: center'>"
	        document.writeln "<INPUT name='SPREADCELL" & i & "' tag='25x' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: left'></TD>"
	    	document.writeln "<TD><INPUT name='SPREADCELL_trip_loc" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 120px; TEXT-ALIGN: left'></TD>"
	    	document.writeln "<TD><INPUT name='SPREADCELL_trip_amt" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 80px; TEXT-ALIGN: right'></TD>"
	    	document.writeln "<TD><INPUT name='SPREADCELL_app_emp_no" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 80px;'></TD>"
	    	  document.writeln "<TD align=center><input type='hidden' name='SPREADCELL_APP_YN1_" & i & "' flag='SPREADCELL' size=1>"
	    	  document.writeln "    <INPUT type=RADIO name='SPREADCELL_APP_YN2_" & i & "' tag='25X' value='Y'  maxlength=1 style='WIDTH: 15px; TEXT-ALIGN: center; visibility : hidden;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled>"
	    	  document.writeln "/   <INPUT type=RADIO name='SPREADCELL_APP_YN2_" & i & "' tag='25X' value='R'  maxlength=1 style='WIDTH: 15px; TEXT-ALIGN: center; visibility : hidden;' onChange='vbscript: Call grid1.SetUpdateFlag(" & i & ")' disabled>"
	    	document.writeln "<TD align=center><input type='hidden' name='SPREADCELL_remark_" & i & "' flag='SPREADCELL' size=1>"
	        document.writeln "</TD>"
	        document.writeln "</TR>"
	    Next
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
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
    strval = strval & "출장비 : " & UNIConvNum(Trip_amt,0) & vbCR
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
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 border=0 width=722 bgcolor=#ffffff>
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
                            <TABLE cellSpacing=1 cellPadding=0 border=0 bgcolor=#ffffff>
                                <TR bgcolor=#d0d6e4 height=19 valign=middle>
		                        	<TD></TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>사번</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>성명</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1 colspan=2>출장기간</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>출장</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>내용</TD>
		                        	<TD NOWRAP class=TDFAMILY_TITLE1>출장비</TD>
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
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()"><IMG alt="이전페이지" src=../../../Cshared/Image/uniSIMS/gprev.jpg border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()"><IMG alt="다음페이지" src=../../../Cshared/Image/uniSIMS/gnext.jpg border=0 ></A>&nbsp;&nbsp;
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
</TD>
</FORM>	
</TR>
</TABLE>
</BODY>
</HTML>
