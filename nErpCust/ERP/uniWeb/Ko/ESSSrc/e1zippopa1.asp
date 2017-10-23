<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>

<HTML>
<HEAD>

<!-- #Include file="../ESSinc/incServer.asp"  -->
<TITLE><%=gLogoName%>-우편번호팝업</TITLE>

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit  

Const BIZ_PGM_ID      = "e1zippopb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXCOLS = 3
Const C_SHEETMAXROWS = 7

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1

Dim arrParam				'--- First Parameter Group		
Dim arrParent

Dim arrReturn				'--- Return Parameter Group

Dim CFlag : CFlag = True
	arrParent = window.dialogArguments
	arrParam = arrParent(0)

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
    if  pOpt = "Q" then
        lgKeyStream = Trim(frm1.txtzip_cd.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & Trim(frm1.txtaddress.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtnat_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & "" & gColSep
    else
    end if
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

'========================================================================================
' Function Name : Form_Load
'========================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    
	frm1.txtaddress.value = arrParam(1)

    Call LockField(Document)

    Dim strzip_cd
    strzip_cd = Replace(arrParam(0), "-", "")
	If strzip_cd="" Then
		frm1.txtzip_cd.value = ""
	Else
		frm1.txtzip_cd.value = Mid(strzip_cd, 1, 3) & "-" & Mid(strzip_cd,4,3)
	End If

    frm1.txtnat_cd.value = arrParam(2)
    Call LayerShowHide(0)
    Call InitGrid()
    Call DbQuery(1)
End Sub

'========================================================================================
' Function Name : Window_onUnLoad
'========================================================================================
Sub Form_UnLoad()
    Set Grid1 = Nothing
    If CFlag Then
        call ExitClick()
    End If    
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    frm1.GRID_PAGE.VALUE = 1
    if  frm1.txtzip_cd.value = "" AND frm1.txtaddress.value = "" then
        'Call DisplayMsgBox("800094","X","X","X")
        frm1.txtzip_cd.focus()
        exit function
    end if

    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
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

'========================================================================================
' Function Name : DoubleGetRow
'========================================================================================
Function DoubleGetRow(pRow)
    Dim objList
    Dim elmCnt

    Dim txttrip_strt_dt
    Dim txttrip_cd
    Dim strVal

	DoubleGetRow = False
    CFlag = False
	Grid1.ActiveRow = pRow

    Redim arrReturn(3)

    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_zip_cd" & pRow then
    		    if objList.value = "" then
    		        exit function
    		    else
		            arrReturn(0) = objList.value
		        end if
    		End if
    		If objList.name = "SPREADCELL_address" & pRow then
    		    if objList.value = "" then
    		        exit function
    		    else
		            arrReturn(1) = objList.value
		        end if
    		End if
    	Next
    End With

	Self.Returnvalue = arrReturn

	DoubleGetRow = True

    Self.Close()
End Function

'========================================================================================
' Function Name : ExitClick
'========================================================================================
Function ExitClick()
    Redim arrReturn(3)
    arrReturn(0) = ""
    arrReturn(1) = ""
	Self.Returnvalue = arrReturn
    Self.Close()
End Function

'========================================================================================
' Function Name : MouseRow
'========================================================================================
Sub MouseRow(pRow)
    on Error Resume Next
    Err.Clear
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
	Set objList = Nothing

End Sub

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Sub GRID_PAGE_OnChange()
End Sub

Function Document_onClick()
Dim Evobj

    Set Evobj = window.event.srcElement
    
    If IsNull(Evobj.id) Then
        CFlag = True
        Exit Function
    Else
        If UCase(Evobj.getAttribute("flag")) = "SPREADCELL" Then            
            CFlag = False
        Else
            CFlag = True
        End If        
    End IF
    Set Evobj = nothing
    
    Document_onClick = True
End Function

Function Document_onKeyDown()

    Dim CuEvObj,KeyCode

	On Error Resume Next
	Err.Clear 

	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode

	Select Case KeyCode
		Case 27
            self.close()
	End Select		
	
	Window_onKeyDown	= True	
End Function

Function txtzip_cd_onKeyDown()
	Dim CuEvObj,KeyCode	
		Set CuEvObj = window.event.srcElement		
		KeyCode = window.event.keycode
		Select Case KeyCode
			Case 13 'enter key
			    Call DbQuery(1)
		End Select		
		txtzip_cd_onKeyDown	= True	
End Function

Function txtaddress_onKeyDown()
	Dim CuEvObj,KeyCode	
		Set CuEvObj = window.event.srcElement		
		KeyCode = window.event.keycode
		Select Case KeyCode
			Case 13 'enter key
			    Call DbQuery(1)
		End Select		
		txtaddress_onKeyDown	= True	
End Function
</SCRIPT>

<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>

<BODY>

<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=460 border=0>
		<tr> 
		  <td width="10" height="10"></td>
		  <td></td>
		  <td width="10"></td>
		</tr>
		<tr> 
		  <td width="10" height="10"></td>
		  <td align="right">
			<table border="0" cellspacing="0" cellpadding="0">
				<TR>
				  <TD class=ftgray><font color="057279">우편번호</font> :
					<INPUT class=form01 NAME="txtzip_cd" MAXLENGTH=12 SiZE=12></TD>
				  <TD class=ftgray><font color="057279">&nbsp;주소</font> :
					<INPUT class=form01 NAME="txtaddress" MAXLENGTH=20 SiZE=20></TD>
				</TR>
		    </table></td>
		  <td></td>
		</tr>
        <TR>
           <td width="10" height="5"></td>
		   <td></td>
		   <td></td>
        </TR>
        <TR>
		   <td width="10"></td>
           <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
				<TR> 
				    <TD class=TDFAMILY_TITLE1></TD>
		        	<TD class=TDFAMILY_TITLE1>우편번호</TD>
		        	<TD class=TDFAMILY_TITLE1>주소</TD>
		        	<TD class=TDFAMILY_TITLE1></TD>
                </TR>
				<%   
		        For i=1 To 7
		            Response.Write "<TR bgcolor=#F8F8F8 height=24 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
		            Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH:  30px; TEXT-ALIGN: center' onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
		            Response.Write "<TD><INPUT class=listrow01 type=text name='SPREADCELL_zip_cd" & i & "' flag='SPREADCELL' style='WIDTH: 80px; TEXT-ALIGN: center'onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
		            Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_address" & i & "' flag='SPREADCELL' style='WIDTH: 325px; TEXT-ALIGN: left'onMouseOver='vbscript: Call MouseRow(" & i & ")'></TD>"
		            Response.Write "</TR>"
		        Next
				%>
               </table>
           </td>
		   <td></td>
        </TR>
        <TR>
            <TD width="10" height=10></TD>
			<td></td>
			<td></td>
        </TR>
        <TR height=20>
            <TD width="10"></TD>
            <TD VALIGN=center ALIGN=center>
                <A onclick="VBSCRIPT: CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../ESSimage/button_07.gif border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
            </TD>
            <TD></TD>
        </TR>
		<tr>
		  <td width="10" height="35" background="../../CShared/ESSimage/popup_bg_01.gif"></td>
		  <td align="center" valign="bottom" background="../../CShared/ESSimage/popup_bg_01.gif">
    		<A onclick="VBSCRIPT:CALL DBQuery(1)" href="#" name=submit><img id=button1 name=button1 SRC="../ESSimage/button_01.gif" border=0 alt='조회' onMouseOver="javascript:this.src='../ESSimage/button_r_01.gif';" onMouseOut="javascript:this.src='../ESSimage/button_01.gif';"></A>
    		<A onclick="VBSCRIPT:CALL ExitClick()" href="#"><img SRC="../ESSimage/button_03.gif" border=0 alt='취소' onMouseOver="javascript:this.src='../ESSimage/button_r_03.gif';" onMouseOut="javascript:this.src='../ESSimage/button_03.gif';"></A>
		  </td>
		  <td background="../../CShared/ESSimage/popup_bg_01.gif"></td>
		</tr>
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

    <INPUT TYPE=hidden NAME="txtnat_cd">
    
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
	<DIV ID="MousePT" NAME="MousePT">
        <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../ESSinc/cursor.htm"></iframe>
    </DIV>
</FORM>	
</BODY>
</HTML>

