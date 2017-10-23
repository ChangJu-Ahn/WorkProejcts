<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>

<HTML>
<HEAD>

<!-- #Include file="../ESSinc/incServer.asp"  -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<TITLE><%=gLogoName%>-코드팝업</TITLE>

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<Script Language="VBScript">
Option Explicit    

Const BIZ_PGM_ID      = "e1codepopb1.asp"						           '☆: Biz Logic ASP Name
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
        lgKeyStream = Trim(frm1.txtTable.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & Trim(frm1.txtCol1.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtCol2.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtMinor.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtWhere.Value) & gColSep
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
'========================================================================================================
' Function Name : GridDsplay
'========================================================================================================
Function GridDsplay()
	Dim i
	
	For i=1 To C_SHEETMAXROWS
    
        document.writeln "<TR bgcolor=#F8F8F8 height=19 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
        document.writeln "<TD><INPUT name='" & i & "'  tag='25X' flag='SPREADCELL' style='WIDTH:  30px; TEXT-ALIGN: center' ></TD>"
        document.writeln "<TD><INPUT type=text	name='SPREADCELL_minor" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 80px; TEXT-ALIGN: center'></TD>"
        document.writeln "<TD><INPUT			name='SPREADCELL_name" & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 300px; TEXT-ALIGN: left'></TD>"
        document.writeln "</TR>"
	    
	Next
End Function
'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    Call LockField(Document)
     
    frm1.txtTable.value	= arrParam(0)
	frm1.txtCol1.value	= arrParam(1)
	frm1.txtCol2.value	= arrParam(2)
	frm1.txtMinor.value = arrParam(3)
	frm1.txtWhere.value = arrParam(4)
	
    Call LayerShowHide(0)
    Call InitGrid()
    Call DbQuery(1)
    
 
End Sub
'========================================================================================
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
    Set Grid1 = Nothing
    If CFlag Then
        call POPClose()
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
    'If Grid1.ChkChange() Then Exit Function
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
    Err.Clear             
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
' Function Name : GetRow
'========================================================================================
Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
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
'    		If objList.name = "SPREADCELL_code2" & pRow then
 '   		    if objList.value = "" then
  '  		        exit function
   ' 		    else
    '		        arrReturn(2) = objList.value
    '		    end if
    '		End if
    	Next
    End With

	Self.Returnvalue = arrReturn

	DoubleGetRow = True

    Self.Close()
End Function

'========================================================================================
' Function Name : POPClose
'========================================================================================
Function POPClose()
    Redim arrReturn(3)
	Self.Returnvalue = arrReturn
    Self.Close()
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Sub GRID_PAGE_OnChange()
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

Function txtName_onKeyDown()
	Dim CuEvObj,KeyCode	
		Set CuEvObj = window.event.srcElement		
		KeyCode = window.event.keycode
		Select Case KeyCode
			Case 13 'enter key
			    Call DbQuery(1)
		End Select		
		txtName_onKeyDown	= True	
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

</SCRIPT>

<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>

<BODY leftmargin="0" topmargin="0" marginwidth="0">

<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=540 border=0>
		<tr> 
		  <td width="10" height="10"></td>
		  <td></td>
		  <td width="10"></td>
		</tr>
		<tr>
		  <td width="10"></td>
		  <td><table width="100%" height="30" border="0" cellpadding="3" cellspacing="1" bgcolor="DDDDDD">
		      <tr> 
		        <td bgcolor="F5F5F5"> 
		          <!------------------  Title S ----------------------->
		          <table width="100%" border="0" cellspacing="0" cellpadding="1">
		            <tr> 
		              <td height="30" align="center" bgcolor="#FFFFFF"> 
		                <!-------- 사번, 성명 S ------->
		                <table border="0" cellspacing="0" cellpadding="1">
		                  <tr> 
		                    <td><img src="../../CShared/ESSimage/icon_03.gif" width="10" height="12"></td>
		                    <td class="ftgray">코드</td>
		                    <td width="5"></td>
		                    <td><input name="txtMinor" type="text" class="form01" style="width:100px"></td>
		                    <td width="20"></td>
		                    <td><img src="../../CShared/ESSimage/icon_03.gif" width="10" height="12"></td>
		                    <td class="ftgray">코드명</td>
		                    <td width="5"></td>
		                    <td><input name="txtName" type="text" class="form01" style="width:100px"></td>
		                  </tr>
		                </table>
		                <!--------  사번,성명 E-------->
		              </td>
		            </tr>
		          </table>
		          <!--------------------- Title E ----------------------->
		        </td>
		      </tr>
		    </table>
		  </td>
		  <td></td>
		</tr>
        <TR>
           <td height="10"></td>
		   <td></td>
		   <td></td>
        </TR>
        <TR>
		   <td></td>
           <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
					<TR> 
					    <TD class=TDFAMILY_TITLE1></TD>
		            	<TD class=TDFAMILY_TITLE1>코드</TD>
		            	<TD class=TDFAMILY_TITLE1>코드명</TD>
		            	<TD class=TDFAMILY_TITLE1></TD>
                    </TR>
					<script language=vbscript>    Call GridDsplay()  </script>
               </table>
           </td>
		   <td></td>
        </TR>
        <TR>
            <TD height=10></TD>
			<td></td>
			<td></td>
        </TR>
        <TR height=20>
            <TD></TD>
            <TD VALIGN=center ALIGN=center>
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../ESSimage/button_07.gif border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
            </TD>
            <TD></TD>
        </TR>
		<tr>
		  <td height="35" background="../../CShared/ESSimage/popup_bg_01.gif"></td>
		  <td align="center" valign="bottom" background="../../CShared/ESSimage/popup_bg_01.gif">
    		<A onclick="VBSCRIPT:CALL DBQuery(1)" href="#" name=submit><img id=button1 name=button1 SRC="../ESSimage/button_01.gif" border=0 alt='조회' onMouseOver="javascript:this.src='../ESSimage/button_r_01.gif';" onMouseOut="javascript:this.src='../ESSimage/button_01.gif';"></A>
    		<A onclick="VBSCRIPT:CALL POPClose()" href="#"><img SRC="../ESSimage/button_03.gif" border=0 alt='취소' onMouseOver="javascript:this.src='../ESSimage/button_r_03.gif';" onMouseOut="javascript:this.src='../ESSimage/button_03.gif';"></A>
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


    <INPUT TYPE=hidden NAME="txtTable">
    <INPUT TYPE=hidden NAME="txtCol1">    
    <INPUT TYPE=hidden NAME="txtCol2"> 
    <INPUT TYPE=hidden NAME="txtWhere">

    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
    <DIV ID="MousePT" NAME="MousePT">
        <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../ESSinc/cursor.htm"></iframe>
    </DIV>
</FORM>	
</BODY>
</HTML>

